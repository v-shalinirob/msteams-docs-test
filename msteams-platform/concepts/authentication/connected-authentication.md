---
title: Connect agent and tab authentication
description: Learn how to link an agent's external identity to a Microsoft identity so users can access an associated tab without signing in again.
ms.topic: how-to
ms.localizationpriority: medium
ms.date: 09/11/2026
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

## How connected authentication works

The following sequence uses Auth0 as an example external identity provider:

1. The user signs in to the agent through an Auth0 OAuth connection.
1. Teams sends a `signin/verifyState` activity to the agent.
1. The agent verifies the sign-in state and returns an app-hosted account-linking URL.
1. Teams opens the account-linking page in a dialog.
1. The account-linking page uses NAA to get a Microsoft Entra access token for the signed-in Teams user.
1. The page uses an authorization code with Proof Key for Code Exchange (PKCE) to authenticate the Microsoft identity with Auth0.
1. The app backend verifies both identities and links the Microsoft identity to the primary Auth0 account.
1. The tab uses the linked Microsoft identity for silent authentication and token renewal.

| Component | Responsibility |
| --- | --- |
| Teams SDK agent | Starts the external OAuth flow, verifies the sign-in state, creates a short-lived linking session, and returns the account-linking URL. |
| Account-linking page | Gets user consent, acquires the NAA token, completes the PKCE flow, and submits the verified secondary identity. |
| App backend | Correlates a linking attempt, validates tokens and callbacks, exchanges one-time codes, and links the accounts. |
| External identity provider | Authenticates the primary identity and maintains the linked identity record. |
| Microsoft Entra ID and NAA | Authenticate the Teams user and provide a token for the requested scopes. |

## User experience

Consider a user who signs in to an agent and then opens the app's tab. Connected authentication presents a clear, one-time option to link their Microsoft account. After linking, the active Teams session gives the user seamless access to the tab, including on another device where the original external-provider session isn't available. The user can also skip linking and continue with the agent.

The user completes the connected authentication flow as follows:

1. The user installs the app and opens the agent chat.
1. The agent asks the user to sign in with the app's external identity provider.
1. After sign-in succeeds, Teams opens the app-hosted account-linking page in a dialog.
1. The page explains that linking the Microsoft account allows the associated tab to sign in without another prompt.
1. If the user continues, NAA requests the required Microsoft Entra permissions. The external identity provider then links the verified Microsoft identity to the user's primary app account.
1. Teams closes the dialog after linking succeeds. The user can open the tab without signing in again.

If the user skips account linking, the agent remains signed in, but the tab must use its existing sign-in flow. If consent, Conditional Access, or reauthentication is required later, the app displays the Microsoft identity prompt instead of treating the user as signed out.

## Developer experience

You implement one coordinated authentication flow instead of independent onboarding flows for the agent and tab. Your app must:

1. Configure the external OAuth connection used by the Teams SDK agent.
1. Configure NAA in Microsoft Entra ID and the App manifest for the tab.
1. Host an account-linking page that explains the flow, requests consent, and handles success, cancellation, and failure.
1. Return the account-linking URL after the agent verifies the external sign-in.
1. Correlate the agent sign-in, NAA token, PKCE transaction, and linking request without exposing tokens in URLs.
1. Verify both identities and link them in the external identity provider.
1. Use the linked Microsoft identity for subsequent tab authentication.

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

Implement connected authentication by configuring the App manifest and Teams SDK authentication, returning the account-linking URL, acquiring the Microsoft identity, and linking the verified accounts.

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

Ensure that:

* `bots[0].botId` and `webApplicationInfo.id` match the Microsoft Entra application client ID used at runtime.
* `nestedAppAuthInfo[0].redirectUri` is registered as a single-page application redirect URI in Microsoft Entra ID.
* The broker redirect contains only the host name. Don't add the account-linking path.
* `nestedAppAuthInfo[0].scopes` exactly matches the scopes requested by the account-linking page.
* `validDomains` includes the host name for the account-linking URL.

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
