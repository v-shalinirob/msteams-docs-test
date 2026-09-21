---
title: Connect agent and tab authentication
description: Learn how connected authentication links an agent or bot sign-in to a Microsoft identity for seamless access to an associated tab.
ms.topic: how-to
ms.localizationpriority: medium
ms.date: 09/18/2026
---

# Connect agent and tab authentication

Connected authentication combines sign-in and authentication flow for a **Teams app that includes an agent or bot and a tab**, with one coordinated authentication experience.

> [!IMPORTANT]
> Connected authentication is a one-way flow from the agent or bot to the tab. Signing in to the agent can authenticate the associated tab after account linking. Signing in to the tab doesn't sign the user in to the agent.

## User experience

Connected authentication streamlines the authentication flows for an agent or bot and an associated tab through account linking. Users sign in to the conversational capability first and can then link that account to their Microsoft identity. After linking, the hosted experience can authenticate the user through their active Microsoft session, reducing repeated prompts across app capabilities.

[Placeholder: Screenshots of connected auth pop-up.]

**Key highlights for users**:

- **Unified user experience**: Users authenticate once and gain access to all app capabilities, reducing confusion and repetitive logins.
- **Consistent Onboarding**: Connected authentication flow ensures that all users meet minimum setup requirements before accessing app features. Following onboarding, the user experiences increased reliability and lesser support issues.
- **Persistent Login**: Account linking with Entra Nested app authentication (NAA) ensures that the user stays logged in, even if the primary login method expires.
- **Seamless Access**: Connected authentication achieves smoother app transactions and interactions as bot and tab capabilities recognize the user through the linked tokens.

## Connected authentication at runtime

The connected authentication flow works as follows:

:::image type="content" source="../../assets/images/authentication/connected-authentication/authentication-flow.png" alt-text="This image shows the authentication flow for connected authentication.":::

1. The user opens the agent or bot and is prompted to sign in with the app's identity provider.
1. After sign-in succeeds, the user chooses whether to link that account to their Microsoft identity for access to the associated tab.
1. If the user chooses to link the accounts, they review and accept any required Microsoft identity permissions.
1. After linking succeeds, the user can open the associated tab without another sign-in prompt.

If the user skips linking or later revokes it, the agent or bot remains independently authenticated, while the tab uses its existing sign-in flow or the app restarts account linking from the agent. If consent, Conditional Access, or reauthentication is required later, the app displays a Microsoft identity prompt.

## Implement connected authentication

Coordinate the App manifest, Teams SDK agent sign-in, NAA token acquisition, and backend identity linking so the tab can authenticate the linked user.

### Prerequisites

Before you implement connected authentication, you need:

- A Teams app with a personal agent or bot and a tab.
- A Teams SDK TypeScript project using `@microsoft/teams.apps`, `@microsoft/teams.api`, and related Teams SDK packages.
- An Azure Bot resource with an OAuth connection for your identity provider.
- A Microsoft Entra app registration configured for NAA.
- An identity provider that supports account linking and Authorization Code flow with PKCE.
- A public HTTPS origin that hosts your agent endpoint, account-linking page, and OAuth bridge endpoints.
- An account-linking URL, such as `https://app.contoso.com/authTab`.

For information about registering the trusted broker redirect and acquiring NAA tokens, see [Nested app authentication](nested-authentication.md).

### Configure the app manifest

Use App manifest version 1.22 or later to add `nestedAppAuthInfo`. The following example uses version 1.23:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/teams/v1.23/MicrosoftTeams.schema.json",
  "manifestVersion": "1.23",
  "bots": [
    {
      "botId": "${{ENTRA_APP_ID}}",
      "scopes": ["personal"],
      "isNotificationOnly": false
    }
  ],
  "validDomains": [
    "app.contoso.com"
  ],
  "webApplicationInfo": {
    "id": "${{ENTRA_APP_ID}}",
    "resource": "api://botid-${{ENTRA_APP_ID}}",
    "nestedAppAuthInfo": [
      {
        "redirectUri": "brk-multihub://app.contoso.com",
        "scopes": ["User.Read"]
      }
    ]
  }
}
```

- `bots[0].botId`: Identifies the agent or bot registration. Set it to the Microsoft Entra application client ID to route Teams activities to the registered agent or bot.
- `validDomains`: Allows Teams to load the linking page. Add the host name from `ACCOUNT_LINKING_URL`, such as `app.contoso.com`, so the page can open in Teams.
- `webApplicationInfo.id`: Identifies the app requesting Microsoft tokens. Set it to the same Microsoft Entra application client ID used at runtime to connect the App manifest to the NAA token request.
- `webApplicationInfo.resource`: Identifies the agent or bot API resource. Set the application ID URI, such as `api://botid-${{ENTRA_APP_ID}}`, to associate the Teams app with its protected API.
- `nestedAppAuthInfo.redirectUri`: Registers the trusted NAA broker redirect. Set the SPA redirect to `brk-multihub://<app-host-name>` without a path so Microsoft 365 hosts can broker NAA authentication.
- `nestedAppAuthInfo.scopes`: Declares permissions requested during NAA authentication. Add the exact runtime scopes, such as `User.Read`, to enable token prefetch and Microsoft Entra consent validation.

### Configure Teams SDK authentication

Set the  provider's OAuth connection as the default connection for the Teams SDK app:

```typescript
import { App, ExpressAdapter } from '@microsoft/teams.apps';

const connectionName = process.env.CONNECTION_NAME || 'Auth0';
const httpServerAdapter = new ExpressAdapter();

const app = new App({
  applicationIdUri: process.env.RESOURCE_URI,
  httpServerAdapter,
  oauth: {
    defaultConnectionName: connectionName,
  },
});
```

- `applicationIdUri`: Identifies the protected agent or bot resource. Set it to the app registration URI from `RESOURCE_URI` to associate Teams SDK authentication with the registered API.
- `httpServerAdapter`: Hosts Teams routes in the web server. Set it to an `ExpressAdapter` instance to receive sign-in activities and serve account-linking endpoints.
- `oauth.defaultConnectionName`: Selects the agent's  OAuth connection. Set it to the Azure Bot OAuth connection name, such as `Auth0`, that handles the initial sign-in.

Start sign-in when the user sends a message and handle the successful sign-in event:

```typescript
app.on('message', async ({ send, signin, isSignedIn }) => {
  if (!isSignedIn) {
    await send(`Sign in with ${connectionName} to continue.`);
    await signin();
    return;
  }

  await send('You are signed in.');
});

app.event('signin', async ({ send }) => {
  await send('Sign-in succeeded. Complete account linking to use the tab.');
});
```

- `isSignedIn`: Indicates whether the agent user authenticated. Use the value supplied by Teams SDK for the current activity to avoid starting another sign-in for an authenticated user.
- `signin()`: Starts the configured  OAuth sign-in. Call it without a connection name to use the default connection and authenticate the primary account before account linking.
- `app.event('signin', ...)`: Handles successful  provider authentication. Register a handler for the `signin` event so the agent can continue into connected authentication.

### Return the account-linking URL

Handle `signin.verify-state` to exchange the state code for the -provider token. Create a short-lived linking session that binds the Teams channel and user to the account-linking request. Return the session-specific account-linking URL in the invoke response:

```typescript
import { InvokeResponse } from '@microsoft/teams.api';
 
const accountLinkingUrl = process.env.ACCOUNT_LINKING_URL;
 
app.on('signin.verify-state', async (context) => {
  const state = context.activity.value.state;
  if (!state || !accountLinkingUrl) {
    return { status: 404 };
  }
 
  try {
    await context.api.users.getToken({
      channelId: context.activity.channelId,
      userId: context.activity.from.id,
      connectionName,
      code: state,
    });
  } catch {
    context.log.error('Failed to verify the sign-in state.');
    return { status: 412 };
  }
 
  // Bind the linking request to the verified user and conversation.
  const sessionId = await createLinkingSession(
    context.activity.channelId,
    context.activity.from.id
  );
  const url = new URL(accountLinkingUrl);
  url.searchParams.set('session', sessionId);
 
  const response: InvokeResponse<'signin/verifyState'> = { status: 200 };
  Object.assign(response, {
    body: {
      composeExtension: {
        text: url.toString(),
        channelData: { accountLinkingUrl: url.toString() },
      },
    },
  });
  return response;
});
```

- `state`: Supplies the one-time sign-in verification code. Use `context.activity.value.state` so Teams SDK can exchange the completed OAuth sign-in for the -provider token.
- `context.api.users.getToken`: Verifies the completed -provider sign-in. Set `channelId` and `userId` from the activity, use the configured `connectionName`, and pass `state` as `code`.
- `createLinkingSession`: Correlates linking with the verified user. Pass the activity's channel and user IDs, and persist a random, short-lived, single-use session ID for the remaining linking operations.
- `ACCOUNT_LINKING_URL`: Identifies the app-hosted linking experience. Set it to an HTTPS URL whose host is included in `validDomains`, such as `https://app.contoso.com/authTab`.
- `session`: Binds the linking page to the request. Add the generated session ID as a query parameter without placing access tokens or identity tokens in the URL.
- `channelData.accountLinkingUrl`: Opens the connected-authentication dialog in Teams. Set it to the session-specific account-linking URL returned in the successful invoke response.

> [!NOTE]
> The Teams SDK TypeScript definitions currently declare the `signin/verifyState` response body as `void`. The example assigns the connected-authentication response payload after creating a typed invoke response.

### Acquire the Microsoft identity with NAA

Initialize MSAL for NAA on the account-linking page. Attempt silent token acquisition first and use an interactive prompt only when required:

```typescript
import { app as teamsApp } from '@microsoft/teams-js';
import { createNestablePublicClientApplication } from '@azure/msal-browser';
 
await teamsApp.initialize();
 
const client = await createNestablePublicClientApplication({
  auth: {
    clientId: naaClientId,
    authority: `https://login.microsoftonline.com/${naaTenantId}`,
    redirectUri: 'brk-multihub://app.contoso.com',
    supportsNestedAppAuth: true,
  },
});
 
const request = { scopes: ['User.Read'] };
const { accessToken } = await client
  .acquireTokenSilent(request)
  .catch(() => client.acquireTokenPopup(request));
```

- `teamsApp.initialize()`: Initializes the page in the Teams host. Call it before creating the MSAL client so NAA can use the host authentication broker.
- `clientId`: Identifies the NAA Microsoft Entra application. Set `naaClientId` to the same application client ID as `webApplicationInfo.id` in the App manifest.
- `authority`: Selects the Microsoft Entra tenant for authentication. Replace `naaTenantId` with the tenant ID supported by the app's account configuration.
- `redirectUri`: Returns control through the trusted NAA broker. Set it to the same `brk-multihub://<app-host-name>` URI declared in `nestedAppAuthInfo`.
- `supportsNestedAppAuth`: Enables brokered authentication in Microsoft 365 hosts. Set it to `true` when creating the nestable public client application.
- `request.scopes`: Requests permissions for the Microsoft identity. Use the minimum scopes declared in `nestedAppAuthInfo`, such as `User.Read`.
- `acquireTokenSilent`: Attempts authentication without prompting the user. Call it first to reuse the active Microsoft session and cached consent.
- `acquireTokenPopup`: Handles authentication that requires user interaction. Use it when silent acquisition can't satisfy consent, Conditional Access, or reauthentication.

The App manifest and runtime values for the client ID, redirect URI, and scopes must match.

### Complete account linking

Your account-linking page and backend must complete these operations:

1. Post the NAA token to your backend over HTTPS with the short-lived linking-session ID.
1. Start the  provider's Authorization Code flow with PKCE for the secondary Microsoft connection.
1. Correlate the authorization request with the linking session using an integrity-protected, single-use value.
1. Exchange the one-time authorization code at the token endpoint.
1. Retrieve the primary -provider token with Teams SDK:

   ```typescript
   const primaryToken = await app.api.users.getToken({
     channelId: session.channelId,
     userId: session.userId,
     connectionName,
   });
   ```

   - `channelId`: Selects the channel for the primary token. Use the channel ID stored in the linking session to retrieve the token for the verified conversation.
   - `userId`: Selects the user for the primary token. Use the user ID stored in the linking session to prevent linking another user's token.
   - `connectionName`: Selects the  provider token to retrieve. Use the same OAuth connection as agent sign-in to supply the primary identity for linking.

1. Validate both identities immediately before linking.
1. Call your identity provider's account-linking API.
1. Delete the linking session, temporary token, and authorization code.
1. Return success to the account-linking page and close the Teams dialog.

The following table shows example app-hosted endpoints for completing the connected authentication flow:

| Endpoint | Purpose |
| --- | --- |
| `GET /authTab` | Renders the account-linking page for a valid linking session. |
| `POST /api/setAuthToken` | Accepts the NAA token over HTTPS and binds it to the linking session. |
| `GET /api/authorize` | Validates the  provider callback and issues a short-lived, single-use authorization code. |
| `POST /api/token` | Exchanges the one-time code for the NAA access token used by the  provider's custom connection. |
| `POST /api/linkAccounts` | Verifies the primary and secondary identities and links them in the identity provider. |

### Test connected authentication

Test at least the following scenarios:

| Scenario | Expected result |
| --- | --- |
| First agent sign-in | The  provider authenticates the user and Teams opens the account-linking page. |
| Account-linking consent | NAA obtains the requested Microsoft token and the backend links the verified identities. |
| Tab open after linking | The tab authenticates silently with the linked Microsoft identity. |
| Different device with an active Teams session | The linked Microsoft identity authenticates the user even when the original -provider session isn't available. |
| User skips linking | The agent remains signed in, but the tab can require its existing sign-in flow. |
| Expired linking session | The backend rejects the request and asks the user to start sign-in again. |
| Concurrent linking attempts | Each attempt remains bound to the correct user, conversation, and one-time correlation value. |
| Revoked consent or Conditional Access | The app requests interaction and handles denial without exposing tokens. |

### Troubleshoot connected authentication

| Problem | Resolution |
| --- | --- |
| Teams rejects the account-linking URL | Add the exact host name to `validDomains`, regenerate the app package, and upload the updated package. |
| NAA can't find the application | Ensure that `webApplicationInfo.id`, the runtime client ID, and the Microsoft Entra app registration are the same. |
| NAA doesn't use a prefetched token | Ensure that the client ID, broker redirect, scopes, and optional claims exactly match the runtime request. |
| The  provider rejects the callback | Register the exact HTTPS account-linking callback and its origin in the provider's application settings. |
| Authorization fails after leaving the Teams webview | Don't depend on third-party cookies for correlation. Use an integrity-protected, single-use correlation value. |
| The tab prompts again after successful linking | Verify that the secondary Microsoft identity is linked to the primary  account and that the tab uses the NAA connection. |

## Design guidelines and best practices

Follow these guidelines when you design and deploy connected authentication:

- **Keep authentication states independent**: Don't infer the agent's authentication state from the tab's state.
- **Preserve token boundaries**: Teams SDK retrieves the primary -provider token for the agent, while MSAL acquires the Microsoft Entra token for the account-linking page. Link verified identity records in your backend. Don't pass an agent token to the tab or expose tokens in URLs.
- **Make account linking clear and optional**: Explain why the Microsoft account is requested and how linking affects the tab. Allow the user to continue or skip linking, and handle cancellation and failure.
- **Correlate every attempt**: Bind each short-lived linking session to the user and conversation that completed sign-in. Use an integrity-protected, single-use correlation value across the NAA, PKCE, and linking operations.
- **Require verified identities**: Don't link accounts based only on identifiers supplied by the client. Require recent authentication for both accounts and validate token issuer, audience, signature, tenant, expiration, nonce, and scopes.
- **Protect authentication endpoints**: Validate the OAuth client at the token endpoint and add cross-site request forgery, replay, rate-limit, and abuse protections.
- **Protect authentication data**: Store linking sessions and one-time codes in an encrypted, durable store with expiration and atomic consumption. Never log access tokens, ID tokens, authorization codes, cookies, or client secrets.
- **Plan for account recovery**: Provide secure account unlinking and recovery, and handle revoked consent without treating the agent and tab as sharing one authentication session.
- **Use production infrastructure**: Keep secrets in a managed secret store, rotate them regularly, and use a permanent app-owned HTTPS origin instead of a development tunnel.
- **Complete security review**: Complete threat modeling, privacy review, consent review, and penetration testing before deployment.

> [!CAUTION]
> The sample implementation used for the code snippets stores linking data in memory and includes a single-pending-session fallback for local testing. Don't use either approach in a concurrent or multi-user deployment.

## Error codes

Connected authentication doesn't define a standardized set of error codes. The following status and error codes are application-defined responses used in the sample or responses returned by the configured identity provider:

| Status code | Error code | Description | Developer action |
| --- | --- | --- | --- |
| HTTP `400` | Invalid token or request | The token submission, authorization request, authorization code, or account-linking request is invalid. | Validate required request values and reject malformed input. If the authorization code or linking session expired, discard it and ask the user to start sign-in again. |
| HTTP `404` | Missing sign-in state | The `signin/verifyState` activity doesn't contain the state value required to complete  sign-in. | Confirm that the agent starts sign-in through its configured OAuth connection and that the activity includes `value.state`. Ask the user to restart sign-in instead of continuing without state. |
| HTTP `410` | Expired linking session | The account-linking session expired or no longer exists. | Delete any temporary tokens and authorization codes associated with the session, then ask the user to restart sign-in from the agent. |
| HTTP `412` | Sign-in state verification failed | Teams SDK couldn't exchange the sign-in state for the -provider token. | Verify the OAuth connection name and provider configuration. Treat the state as expired or invalid and ask the user to start a new sign-in attempt. |
| HTTP `500` | Account linking failed | An unexpected error prevented the backend from linking the accounts. | Log a correlation identifier without logging tokens, return a generic failure message, and investigate identity validation, storage, and provider communication before retrying. |
| HTTP `502` | Identity provider rejected request | The identity provider returned an unsuccessful response to the account-linking request. | Inspect the upstream status, verify the provider endpoint and request, and retry only if the failure is transient. Don't return provider tokens or sensitive response details to the client. |
| HTTP `503` | Account linking not configured | The backend doesn't have the -provider configuration required to link accounts. | Configure the provider domain, OAuth connection, credentials, and account-linking permissions before enabling the flow. |
| HTTP `401` or `403` | Identity provider authorization failed | The primary -provider token has the wrong audience or lacks permission to link identities. | Configure the OAuth connection to request the provider's account-management API audience and identity-linking scopes, then have the user sign in again to obtain a new token. |

## Code sample

<!-- Add the TypeScript sample link when the sample is published. -->

| Sample name | Description | TypeScript |
| --- | --- | --- |
| Connected authentication with Auth0 | This sample shows how to link an agent's Auth0 identity to a Microsoft identity for seamless tab authentication. | Coming soon |

## Next step

Configure and test [Nested app authentication](nested-authentication.md) for the tab.

## See also

- [Authenticate users in Microsoft Teams](authentication.md)
- [Add authentication to a Teams agent or bot](../../bots/how-to/authentication/add-authentication.md)
- [Nested app authentication](nested-authentication.md)
- [App manifest schema](/microsoftteams/platform/resources/schema/manifest-schema)
- [Teams SDK documentation](https://microsoft.github.io/teams-sdk/)
- [Auth0 user account linking](https://auth0.com/docs/manage-users/user-accounts/user-account-linking/link-user-accounts)
