---
title: Connect agent and tab authentication
description: Learn how to link an agent's external identity to a Microsoft identity so users can access an associated tab without signing in again.
ms.topic: how-to
ms.localizationpriority: medium
ms.date: 09/15/2026
---

# Connect agent and tab authentication

Connected authentication links the external identity that a user selects when signing in to an agent or bot with the user's Microsoft identity in Teams. After the accounts are linked, your tab can use nested app authentication (NAA) to sign in the user without another interactive prompt.

Use this flow for a Teams app that has:

* A conversational agent or bot built with Teams SDK.
* A tab or other app-hosted web experience.
* An external identity provider, such as Auth0.
* A backend that can securely link two verified identities.

> [!IMPORTANT]
> Connected authentication is a one-way flow from the agent or bot to the tab. Signing in to the agent can authenticate the associated tab after account linking. Signing in to the tab doesn't sign the user in to the agent.

## User experience

Consider a user who accesses a multi-capability Teams app that includes an agent or bot and a tab. The user signs in to the agent or bot with the app's external identity provider and receives a one-time option to link that account to their Microsoft account. If the user skips account linking, the agent or bot remains signed in, while the tab uses its existing sign-in flow. The following flow uses Auth0 as an example external identity provider:

1. The user opens the agent chat, and the Teams SDK agent starts the external OAuth flow for the user to sign in through Auth0.
1. After sign-in succeeds, Teams sends a `signin/verifyState` activity to the agent. The agent verifies the sign-in state, creates a short-lived session that identifies the user and conversation, and returns the app-hosted account-linking URL.
1. Teams opens the account-linking page in a dialog. The page explains that linking the accounts allows the associated tab to sign in without another prompt.
1. If the user continues, the page uses nested app authentication (NAA) to acquire a Microsoft Entra access token for the signed-in Teams user and request any required permissions.
1. The page completes an Authorization Code flow with Proof Key for Code Exchange (PKCE). The app backend correlates the linking session, validates the tokens and callback, exchanges the one-time authorization code, and links the verified Microsoft identity to the user's primary Auth0 account.
1. After linking succeeds, Teams closes the dialog. The tab uses NAA and the linked Microsoft identity for silent authentication and token renewal.

If consent, Conditional Access, or reauthentication is required later, the app displays the Microsoft identity prompt instead of treating the user as signed out.

Connected authentication links identity records; it doesn't combine or expose access tokens across the agent and tab. Continue to use each token only for its intended resource and audience.

## Prerequisites

Before you implement connected authentication, you need:

* A Teams app with a personal agent or bot and a tab.
* A Teams SDK TypeScript project using `@microsoft/teams.apps`, `@microsoft/teams.api`, and related Teams SDK packages.
* An Azure Bot resource with an OAuth connection for your external identity provider.
* A Microsoft Entra app registration configured for NAA.
* An external identity provider that supports account linking and Authorization Code flow with PKCE.
* A public HTTPS origin that hosts your agent endpoint, account-linking page, and OAuth bridge endpoints.
* An account-linking URL, such as `https://app.contoso.com/authTab`.

For information about registering the trusted broker redirect and acquiring NAA tokens, see [Nested app authentication](nested-authentication.md).

## Implement connected authentication

You implement one coordinated authentication flow instead of separate onboarding flows for the agent and tab. First, configure NAA in Microsoft Entra ID and the App manifest for the tab, and configure the external OAuth connection used by the Teams SDK agent. After the agent verifies the external sign-in, return an app-hosted account-linking URL that explains the flow, requests consent, and handles success, cancellation, and failure. The account-linking page then acquires the Microsoft identity with NAA. Your backend correlates the agent sign-in, NAA token, PKCE transaction, and linking request without exposing tokens in URLs, verifies both identities, and links them in the external identity provider. The tab can then use the linked Microsoft identity for subsequent authentication.

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

* `bots[0].botId`: Identifies the agent or bot registration. Set it to the Microsoft Entra application client ID to route Teams activities to the registered agent or bot.
* `validDomains`: Allows Teams to load the linking page. Add the host name from `ACCOUNT_LINKING_URL`, such as `app.contoso.com`, so the page can open in Teams.
* `webApplicationInfo.id`: Identifies the app requesting Microsoft tokens. Set it to the same Microsoft Entra application client ID used at runtime to connect the App manifest to the NAA token request.
* `webApplicationInfo.resource`: Identifies the agent or bot API resource. Set the application ID URI, such as `api://botid-${{ENTRA_APP_ID}}`, to associate the Teams app with its protected API.
* `nestedAppAuthInfo.redirectUri`: Registers the trusted NAA broker redirect. Set the SPA redirect to `brk-multihub://<app-host-name>` without a path so Microsoft 365 hosts can broker NAA authentication.
* `nestedAppAuthInfo.scopes`: Declares permissions requested during NAA authentication. Add the exact runtime scopes, such as `User.Read`, to enable token prefetch and Microsoft Entra consent validation.

### Configure Teams SDK authentication

Set the external provider's OAuth connection as the default connection for the Teams SDK app:

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

* `applicationIdUri`: Identifies the protected agent or bot resource. Set it to the app registration URI from `RESOURCE_URI` to associate Teams SDK authentication with the registered API.
* `httpServerAdapter`: Hosts Teams routes in the web server. Set it to an `ExpressAdapter` instance to receive sign-in activities and serve account-linking endpoints.
* `oauth.defaultConnectionName`: Selects the agent's external OAuth connection. Set it to the Azure Bot OAuth connection name, such as `Auth0`, that handles the initial sign-in.

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

* `isSignedIn`: Indicates whether the agent user authenticated. Use the value supplied by Teams SDK for the current activity to avoid starting another sign-in for an authenticated user.
* `signin()`: Starts the configured external OAuth sign-in. Call it without a connection name to use the default connection and authenticate the primary account before account linking.
* `app.event('signin', ...)`: Handles successful external provider authentication. Register a handler for the `signin` event so the agent can continue into connected authentication.

### Return the account-linking URL

Handle `signin.verify-state` to exchange the state code for the external-provider token. Create a short-lived linking session that binds the Teams channel and user to the account-linking request. Return the session-specific account-linking URL in the invoke response:

```typescript
import { randomUUID } from 'node:crypto';
import { ChannelID, InvokeResponse } from '@microsoft/teams.api';

interface LinkingSession {
  readonly channelId: ChannelID;
  readonly expiresAt: number;
  readonly userId: string;
}

const linkingSessions = new Map<string, LinkingSession>();
const accountLinkingUrl = process.env.ACCOUNT_LINKING_URL;

function createLinkingSession(channelId: ChannelID, userId: string): string {
  const sessionId = randomUUID();
  linkingSessions.set(sessionId, {
    channelId,
    userId,
    expiresAt: Date.now() + 10 * 60 * 1000,
  });
  return sessionId;
}

app.on('signin.verify-state', async (context) => {
  const state = context.activity.value.state;
  if (!state) {
    context.log.warn('The sign-in activity did not include state.');
    return { status: 404 };
  }

  try {
    await context.api.users.getToken({
      channelId: context.activity.channelId,
      userId: context.activity.from.id,
      connectionName,
      code: state,
    });
  } catch (error) {
    context.log.error('Failed to verify the sign-in state.');
    return { status: 412 };
  }

  if (!accountLinkingUrl) {
    context.log.warn('ACCOUNT_LINKING_URL is not configured.');
    return { status: 200 };
  }

  const sessionId = createLinkingSession(
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
        channelData: {
          accountLinkingUrl: url.toString(),
        },
      },
    },
  });
  return response;
});
```

* `state`: Carries the one-time sign-in verification code. Use `context.activity.value.state` from Teams so Teams SDK can exchange the completed OAuth sign-in.
* `channelId`: Identifies the conversation channel for token retrieval. Set it to `context.activity.channelId` to bind the linking session to the correct conversation.
* `userId`: Identifies the user completing connected authentication. Set it to `context.activity.from.id` to bind token retrieval and account linking to one user.
* `connectionName`: Selects the completed external OAuth connection. Use the same connection configured as `defaultConnectionName` to retrieve the primary external-provider token.
* `code`: Supplies the state code for token exchange. Set it to the `state` value from the invoke activity to complete the external OAuth sign-in.
* `ACCOUNT_LINKING_URL`: Locates the app-hosted linking experience. Set an HTTPS URL, such as `https://app.contoso.com/authTab`, to tell Teams which page to open.
* `session`: Correlates the page with the verified sign-in. Set it to a short-lived, random linking-session ID that connects NAA and PKCE operations to the correct user.
* `channelData.accountLinkingUrl`: Returns the linking page location to Teams. Set it to the account-linking URL containing the session ID to open the connected authentication dialog.

> [!NOTE]
> The Teams SDK TypeScript definitions currently declare the `signin/verifyState` response body as `void`. The example assigns the connected-authentication response payload after creating a typed invoke response.

### Acquire the Microsoft identity with NAA

Initialize MSAL for NAA on the account-linking page. Attempt silent token acquisition first and use an interactive prompt only when required:

```typescript
import { createNestablePublicClientApplication } from '@azure/msal-browser';

interface AcquireNaaTokenOptions {
  readonly clientId: string;
  readonly redirectUri: string;
  readonly scopes: readonly string[];
  readonly tenantId: string;
}

export async function acquireNaaAccessToken({
  clientId,
  redirectUri,
  scopes,
  tenantId,
}: AcquireNaaTokenOptions): Promise<string> {
  const client = await createNestablePublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      supportsNestedAppAuth: true,
      redirectUri,
    },
  });

  const request = { scopes: [...scopes] };
  const token = await client
    .acquireTokenSilent(request)
    .catch(() => client.acquireTokenPopup(request));
  return token.accessToken;
}
```

* `clientId`: Identifies the NAA Microsoft Entra application. Set it to the same client ID as `webApplicationInfo.id` to request the Microsoft identity configured for the Teams app.
* `authority`: Selects the Microsoft Entra tenant authority. Set it to `https://login.microsoftonline.com/<tenant-id>` to direct authentication to the intended tenant.
* `supportsNestedAppAuth`: Enables brokered authentication inside Microsoft 365 hosts. Set it to `true` to activate NAA behavior in the MSAL public client.
* `redirectUri`: Identifies the trusted broker return location. Use the same `brk-multihub://<app-host-name>` URI as the App manifest so Teams can return the NAA result.
* `scopes`: Specifies Microsoft permissions requested by the page. Use the exact App manifest scopes, such as `User.Read`, to control consent and the NAA token permissions.

Initialize TeamsJS before you initialize MSAL on the account-linking page:

```typescript
await microsoftTeams.app.initialize();

const naaAccessToken = await acquireNaaAccessToken({
  clientId: naaClientId,
  redirectUri: naaRedirectUri,
  tenantId: naaTenantId,
  scopes: ['User.Read'],
});
```

* `naaClientId`: Identifies the Microsoft Entra application registration. Set it to the same client ID used in the App manifest to keep the runtime NAA request aligned.
* `naaRedirectUri`: Supplies the registered trusted broker redirect. Set it to `brk-multihub://<app-host-name>` to return authentication control to the embedded linking page.
* `naaTenantId`: Selects the Teams user's Microsoft Entra tenant. Set the tenant ID for the supported account configuration to acquire the identity used for linking.
* `scopes`: Requests permissions required by the linking flow. Add the minimum permissions declared in the App manifest to produce the NAA token submitted to the backend.

The app manifest values and runtime values for the client ID, redirect URI, and scopes must match.

### Complete account linking

Your account-linking page and backend must complete these operations:

1. Post the NAA token to your backend over HTTPS with the short-lived linking-session ID.
1. Start the external provider's Authorization Code flow with PKCE for the secondary Microsoft connection.
1. Correlate the authorization request with the linking session using an integrity-protected, single-use value.
1. Exchange the one-time authorization code at the token endpoint.
1. Retrieve the primary external-provider token with Teams SDK:

   ```typescript
   const primaryToken = await app.api.users.getToken({
     channelId: session.channelId,
     userId: session.userId,
     connectionName,
   });
   ```

   * `channelId`: Selects the channel for the primary token. Use the channel ID stored in the linking session to retrieve the token for the verified conversation.
   * `userId`: Selects the user for the primary token. Use the user ID stored in the linking session to prevent linking another user's token.
   * `connectionName`: Selects the external provider token to retrieve. Use the same OAuth connection as agent sign-in to supply the primary identity for linking.

1. Validate both identities immediately before linking.
1. Call your identity provider's account-linking API.
1. Delete the linking session, temporary token, and authorization code.
1. Return success to the account-linking page and close the Teams dialog.

The following table shows the minimum app-hosted endpoints used by the sample:

| Endpoint | Purpose |
| --- | --- |
| `GET /authTab` | Renders the account-linking page for a valid linking session. |
| `POST /api/setAuthToken` | Accepts the NAA token over HTTPS and binds it to the linking session. |
| `GET /api/authorize` | Validates the external provider callback and issues a short-lived, single-use authorization code. |
| `POST /api/token` | Exchanges the one-time code for the NAA access token used by the external provider's custom connection. |
| `POST /api/linkAccounts` | Verifies the primary and secondary identities and links them in the identity provider. |

### Authentication at run time

The user's linking state determines the authentication experience at run time:

| State | Runtime behavior |
| --- | --- |
| Agent isn't signed in | The Teams SDK agent starts the external OAuth flow by calling `signin()`. |
| Agent is signed in, but accounts aren't linked | The agent returns the account-linking URL after `signin.verify-state`. Teams opens the app-hosted page so the user can continue or skip linking. |
| Account linking is in progress | The page acquires the Microsoft Entra token with NAA, completes the external provider's PKCE flow, and asks the backend to link both verified identities. |
| Accounts are linked | The tab starts NAA token acquisition. MSAL first calls `acquireTokenSilent`, and the external identity provider resolves the Microsoft identity to the linked primary app account. |
| User interaction is required | MSAL calls `acquireTokenPopup` for consent, Conditional Access, or reauthentication. |
| Linking was skipped or revoked | The agent remains independently authenticated. The tab uses its existing sign-in flow or the app restarts account linking from the agent. |

After linking succeeds, the tab repeats the NAA-based authorization flow whenever it needs a token. If the Teams user has a valid Microsoft session, MSAL renews the token silently and the tab doesn't display another sign-in prompt.

## Design guidelines and best practices

Follow these guidelines when you design and deploy connected authentication:

* **Keep the flow one-way**: Don't infer that the agent is signed in from the tab's authentication state. Connected authentication doesn't support the reverse tab-to-agent direction.
* **Preserve token boundaries**: Teams SDK retrieves the primary external-provider token for the agent, while MSAL acquires the Microsoft Entra token for the account-linking page. Link verified identity records in your backend. Don't pass an agent token to the tab or expose tokens in URLs.
* **Make account linking clear and optional**: Explain why the Microsoft account is requested and how linking affects the tab. Allow the user to continue or skip linking, and handle cancellation and failure.
* **Correlate every attempt**: Bind each short-lived linking session to the user and conversation that completed sign-in. Use an integrity-protected, single-use correlation value across the NAA, PKCE, and linking operations.
* **Require verified identities**: Don't link accounts based only on identifiers supplied by the client. Require recent authentication for both accounts and validate token issuer, audience, signature, tenant, expiration, nonce, and scopes.
* **Protect authentication endpoints**: Validate the OAuth client at the token endpoint and add cross-site request forgery, replay, rate-limit, and abuse protections.
* **Protect authentication data**: Store linking sessions and one-time codes in an encrypted, durable store with expiration and atomic consumption. Never log access tokens, ID tokens, authorization codes, cookies, or client secrets.
* **Plan for account recovery**: Provide secure account unlinking and recovery, and handle revoked consent without treating the agent and tab as sharing one authentication session.
* **Use production infrastructure**: Keep secrets in a managed secret store, rotate them regularly, and use a permanent app-owned HTTPS origin instead of a development tunnel.
* **Complete security review**: Complete threat modeling, privacy review, consent review, and penetration testing before deployment.

> [!CAUTION]
> The sample implementation used for the code snippets stores linking data in memory and includes a single-pending-session fallback for local testing. Don't use either approach in a concurrent or multi-user deployment.

## Test connected authentication

Test at least the following scenarios:

| Scenario | Expected result |
| --- | --- |
| First agent sign-in | The external provider authenticates the user and Teams opens the account-linking page. |
| Account-linking consent | NAA obtains the requested Microsoft token and the backend links the verified identities. |
| Tab open after linking | The tab authenticates silently with the linked Microsoft identity. |
| Different device with an active Teams session | The linked Microsoft identity authenticates the user even when the original social-provider session isn't available. |
| User skips linking | The agent remains signed in, but the tab can require its existing sign-in flow. |
| Expired linking session | The backend rejects the request and asks the user to start sign-in again. |
| Concurrent linking attempts | Each attempt remains bound to the correct user, conversation, and one-time correlation value. |
| Revoked consent or Conditional Access | The app requests interaction and handles denial without exposing tokens. |

## Troubleshoot connected authentication

| Problem | Resolution |
| --- | --- |
| Teams rejects the account-linking URL | Add the exact host name to `validDomains`, regenerate the app package, and upload the updated package. |
| NAA can't find the application | Ensure that `webApplicationInfo.id`, the runtime client ID, and the Microsoft Entra app registration are the same. |
| NAA doesn't use a prefetched token | Ensure that the client ID, broker redirect, scopes, and optional claims exactly match the runtime request. |
| The external provider rejects the callback | Register the exact HTTPS account-linking callback and its origin in the provider's application settings. |
| Authorization fails after leaving the Teams webview | Don't depend on third-party cookies for correlation. Use an integrity-protected, single-use correlation value. |
| Account linking returns `401` or `403` | Verify that the primary external-provider token has the audience and account-linking permissions required by the provider. |
| The tab prompts again after successful linking | Verify that the secondary Microsoft identity is linked to the primary external account and that the tab uses the NAA connection. |

## Next step

Configure and test [Nested app authentication](nested-authentication.md) for the tab.

## See also

* [Authenticate users in Microsoft Teams](authentication.md)
* [Add authentication to a Teams agent or bot](../../bots/how-to/authentication/add-authentication.md)
* [Nested app authentication](nested-authentication.md)
* [App manifest schema](/microsoftteams/platform/resources/schema/manifest-schema)
* [Teams SDK documentation](https://microsoft.github.io/teams-sdk/)
* [Auth0 user account linking](https://auth0.com/docs/manage-users/user-accounts/user-account-linking/link-user-accounts)
