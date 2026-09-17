---
title: Send and receive files and inline images
description: Learn how agents send and receive files and inline images in Microsoft Teams using Teams SDK and Microsoft Graph.
ms.date: 09/17/2026
author: nickwalkmsft
ms.author: nickwalk
ms.reviewer: nickwalk
ms.localizationpriority: medium
ms.topic: how-to
ms.owner: angovil
---
# Send and receive files and inline images

Agents can work with document files and inline images in Microsoft Teams conversations. A document file is stored in OneDrive or SharePoint and usually appears as a file card. An inline image renders directly in the conversation and doesn't appear in the **Files** tab.

Use Teams SDK to receive files in personal chats, send files through file consent, and receive or send inline images. Use Microsoft Graph when your app must work with stored files across personal chats, group chats, and channels.

> [!IMPORTANT]
>
> Agents don't support document-file send and receive workflows in Government Community Cloud High (GCC High), Department of Defense (DoD), or Teams operated by 21Vianet. In these environments, use a base64 image attachment to include an inline image in a message.

## User experience

Users can:

* Attach a document to a personal chat with an agent.
* Accept or decline a file that an agent offers.
* Paste an image that renders within a message.
* View an image that an agent sends from a hosted URL or as base64 data.

Files and inline images behave differently:

| Document file | Inline image |
| --- | --- |
| Stored in OneDrive or SharePoint. | Rendered as part of the message. |
| Appears as a file card or stored item. | Doesn't appear in the **Files** tab. |
| Received as `file.download.info` metadata. | Received as an `image/*` attachment. |
| Sent through file consent or Microsoft Graph. | Sent through an image attachment or HTML content. |

## Developer experience

Choose an approach based on the content and conversation scope:

| Requirement | Recommended approach |
| --- | --- |
| Receive a document in a personal chat | Teams SDK `context.Files` |
| Receive an image pasted into a message | Inspect `Activity.Attachments` |
| Send a document in a personal chat | Teams SDK file consent |
| Send or retrieve stored files in other scopes | Microsoft Graph |
| Send an image beside message text | Image attachment |
| Position an image within formatted text | Base64 image in HTML/XML content |
| Add an image to interactive content | Adaptive Card `Image` element |

`Activity.Attachments` contains non-text message content, including files, inline images, Adaptive Cards, mentions, link previews, and HTML layout information. The `context.Files` accessor filters this collection and exposes supported document files as lazy `IncomingFile` handles.

The following table summarizes scope and permission requirements:

| Operation | Personal chat | Group chat | Channel | Requirement |
| --- | --- | --- | --- | --- |
| Receive files with `context.Files` | Supported | Not explicitly supported | Not explicitly supported | Traditional bots use a pre-authorized URL. Agentic users require Graph file permissions. |
| Send files with file consent | Supported | Not supported | Not supported | Set `supportsFiles` to `true` for traditional bots. |
| Send or retrieve files with Graph | Supported | Supported | Supported | Configure the appropriate OneDrive or SharePoint permissions. |
| Receive inline images | Supported through activity attachments | Use activity attachments | Use activity attachments | Use the app's authenticated HTTP client. |
| Send inline images | Supported | Supported | Supported | No Graph permission or user sign-in is required. |

## Implement

### Configure file support

For a traditional bot, set `supportsFiles` to `true` in the bot entry of the app manifest:

```json
{
  "bots": [
    {
      "botId": "${{BOT_ID}}",
      "scopes": ["personal"],
      "supportsFiles": true
    }
  ]
}
```

Key values:

* `botId`: Set `${{BOT_ID}}` to identify your bot registration.
* `scopes`: Add `personal` to enable one-to-one file interactions.
* `supportsFiles`: Set `true` to expose the file attachment control.

Without `supportsFiles: true`, users can't attach files in a personal chat with the bot and `context.Files.ListAsync()` returns an empty collection. This setting doesn't grant Microsoft Graph permissions.

For an agentic user, configure and obtain administrator consent for a Graph file permission on the agent blueprint. The agentic user retrieves file content with its own identity. For more information, see [inheritable permissions](/entra/agent-id/concept-inheritable-permissions).

### Receive files in personal chat

When a user attaches a document in a personal chat, Teams stores it in OneDrive or SharePoint. The message activity contains file metadata, and `context.Files` provides lazy access to the file bytes.

Use `ListAsync()` to access every supported file attached to the current message:

```csharp
teamsApp.OnMessage(async (context, cancellationToken) =>
{
    IList<IncomingFile> files =
        await context.Files.ListAsync(cancellationToken);

    if (files.Count == 0)
    {
        await context.ReplyAsync(
            "Attach a file and I will read it.",
            cancellationToken);
        return;
    }

    string names = string.Join(", ", files.Select(file => file.Name));
    await context.ReplyAsync(
        $"You sent {files.Count} file(s): {names}",
        cancellationToken);
});
```

Key APIs and values:

* `OnMessage`: Register the handler that processes incoming message activities.
* `context.Files.ListAsync`: Return file metadata without downloading file bytes.
* `cancellationToken`: Propagate cancellation through listing and reply operations.
* `file.Name`: Read the uploader-provided name for display only.

`ListAsync()` preserves attachment order. It returns an empty collection for activities without supported files and skips malformed file entries. Use `FirstAsync()` when your handler expects one file:

```csharp
IncomingFile? file =
    await context.Files.FirstAsync(cancellationToken);

if (file is not null)
{
    await context.ReplyAsync(
        $"Reading {file.Name}...",
        cancellationToken);
}
```

Key APIs and values:

* `FirstAsync`: Return the first supported file or `null`.
* `file.Name`: Display the received file's uploader-provided name.

An `IncomingFile` includes:

* `UniqueId`: OneDrive or SharePoint drive-item ID when available.
* `Name`: Uploader-provided filename, including its file extension.
* `Extension`: Platform-provided extension without the leading period.
* `ContentType`: MIME type when provided by the source.
* `Scope`: Conversation scope where the file was received.
* `Source`: SDK source that identified the incoming file.
* `ContentUrl`: Browsable storage URL, not necessarily a download URL.
* `Raw`: Original attachment metadata for protocol-level diagnostics.

> [!CAUTION]
>
> Treat `Name` as untrusted input. Before writing a file, replace it with a safe application-generated name or sanitize it and verify that the resolved destination remains inside an application-controlled directory.

### Read a received file

Use `DownloadAsync()` to download a file into a reusable in-memory copy:

```csharp
IncomingFile? file =
    await context.Files.FirstAsync(cancellationToken);

if (file is not null)
{
    DownloadedFile downloaded =
        await file.DownloadAsync(cancellationToken);

    await context.ReplyAsync(
        $"Downloaded {downloaded.Filename} " +
        $"({downloaded.Bytes.Length} bytes, {downloaded.ContentType}).",
        cancellationToken);
}
```

Key APIs and values:

* `DownloadAsync`: Fetch and buffer one reusable copy of bytes.
* `downloaded.Bytes`: Access the downloaded file's complete binary content.
* `downloaded.ContentType`: Use the resolved MIME type for processing.
* `downloaded.Filename`: Use the resolved name for display purposes.

Other read options include:

* `TextAsync()`: Download and decode text as UTF-8 by default.
* `StreamAsync()`: Process large files without buffering them completely.
* `SaveAsAsync(path)`: Stream bytes directly to a safe local path.

`TextAsync()` replaces invalid byte sequences with `U+FFFD`. Supply an `Encoding` for non-UTF-8 text and process binary files as bytes or streams.

An `IncomingFile` doesn't cache bytes. Every call to `DownloadAsync()`, `TextAsync()`, `StreamAsync()`, or `SaveAsAsync()` performs another network request. Download once and reuse `DownloadedFile` when you need the same content more than once:

```csharp
DownloadedFile downloaded =
    await file.DownloadAsync(cancellationToken);

string text = downloaded.Text();
byte[] bytes = downloaded.Bytes;

await downloaded.SaveAsAsync(
    "./downloads/copy.bin",
    cancellationToken);
```

Key APIs and values:

* `downloaded.Text()`: Decode the buffered copy without another download.
* `downloaded.Bytes`: Reuse buffered bytes for binary processing.
* `SaveAsAsync`: Save buffered bytes without fetching the file again.

For a traditional bot, the activity contains a short-lived, pre-authorized `downloadUrl`. For an agentic user, the SDK uses the attachment `ContentUrl` and the agentic user's identity to retrieve the file through Microsoft Graph. The same file-read APIs apply to both routes.

### Access the raw file attachment

Use `IncomingFile.Raw` only when you need the original protocol payload:

```csharp
using System.Text.Json;

IncomingFile? file =
    await context.Files.FirstAsync(cancellationToken);

if (file is not null)
{
    string wire = JsonSerializer.Serialize(file.Raw);
    await context.ReplyAsync(
        $"Raw attachment: {wire}",
        cancellationToken);
}
```

Key APIs and values:

* `file.Raw`: Access original metadata for diagnostics or troubleshooting.
* `JsonSerializer.Serialize`: Convert metadata to inspect its protocol shape.

Remove sensitive URLs and identifiers before logging raw attachment data. Inline images, cards, mentions, link previews, HTML attachments, and malformed file entries aren't returned through `context.Files`; access them through `context.Activity.Attachments`.

### Send files in personal chat

The file-consent workflow is available only in personal chats:

1. Send a `FileConsentCard`.
1. Receive a `fileConsent/invoke` activity.
1. If the user accepts, upload the bytes to `uploadUrl` with HTTP `PUT`.
1. Send a `FileInfoCard` that links to the uploaded file.
1. If the user declines, discard the pending content.

The following message requests permission to upload a file:

:::image type="content" source="../../assets/images/bots/bot-file-consent-card.png" alt-text="Consent card requesting permission to upload a file." lightbox="../../assets/images/bots/bot-file-consent-card.png" border="true":::

On mobile, the consent request appears as follows:

<img src="../../assets/images/bots/mobile-bot-file-consent-card.png" alt="Consent card requesting permission to upload a file on mobile." width="350"/>

```json
{
  "attachments": [
    {
      "contentType": "application/vnd.microsoft.teams.card.file.consent",
      "name": "file_example.txt",
      "content": {
        "description": "This is your monthly expense report.",
        "sizeInBytes": 1029393,
        "acceptContext": {
          "fileId": "expense-report-2026-09"
        },
        "declineContext": {
          "fileId": "expense-report-2026-09"
        }
      }
    }
  ]
}
```

Key properties and values:

* `contentType`: Use the file-consent card content type shown.
* `name`: Set the filename that Teams displays to users.
* `description`: Add a brief purpose for the offered file.
* `sizeInBytes`: Set the exact file size in bytes.
* `acceptContext`: Add identifiers needed after the user accepts.
* `declineContext`: Add identifiers needed after the user declines.

Keep pending file bytes outside the context object. Apply expiration and cleanup policies to pending uploads.

When the user accepts, Teams sends `fileConsent/invoke` with `action` set to `accept` and an `uploadInfo.uploadUrl`. If the user declines, `action` is `decline`.

Use the upload URL to transfer the file bytes:

```typescript
async function uploadToOneDrive(url: string, content: Buffer): Promise<void> {
  const fileSize = content.length;
  const response = await axios.put(url, content, {
    headers: {
      'Content-Type': 'application/octet-stream',
      'Content-Length': fileSize.toString(),
      'Content-Range': `bytes 0-${fileSize - 1}/${fileSize}`
    }
  });

  if (![200, 201].includes(response.status)) {
    throw new Error(`Upload failed with status ${response.status}`);
  }
}
```

Key parameters and values:

* `url`: Use the `uploadInfo.uploadUrl` returned after acceptance.
* `content`: Provide the exact bytes associated with consent.
* `Content-Type`: Set `application/octet-stream` for binary file transfer.
* `Content-Length`: Set the decimal length of uploaded bytes.
* `Content-Range`: Describe the uploaded byte range and total.

After a successful upload, send a file information attachment:

```json
{
  "attachments": [
    {
      "contentType": "application/vnd.microsoft.teams.card.file.info",
      "contentUrl": "https://contoso.sharepoint.com/personal/user/Documents/Applications/file_example.txt",
      "name": "file_example.txt",
      "content": {
        "uniqueId": "1150D938-8870-4044-9F2C-5BBDEBA70C8C",
        "fileType": "txt"
      }
    }
  ]
}
```

Key properties and values:

* `contentType`: Use the file-information card content type shown.
* `contentUrl`: Set the stored file URL for user access.
* `uniqueId`: Set the OneDrive or SharePoint drive-item ID.
* `fileType`: Set the platform-reported file extension without punctuation.

### Use Microsoft Graph for stored files

Use Microsoft Graph when your app must send or retrieve stored files outside the personal-chat file-consent workflow:

* Use a user's OneDrive for personal and group-chat files.
* Use the team's SharePoint site for channel files.
* Obtain the required storage access through OAuth 2.0.
* Post a message attachment that references an existing stored file.

For more information, see [send chat message file attachments](/graph/api/chatmessage-post?view=graph-rest-beta&preserve-view=true&tabs=http#example-4-file-attachments) and [OneDrive and SharePoint APIs](/onedrive/developer/rest-api/).

### Receive inline images

An inline image isn't exposed through `context.Files`. An inbound message commonly includes an `image/*` attachment with the authenticated download URL and a `text/html` attachment that preserves the image position.

Use the `image/*` attachment as the canonical source. Don't use the `<img src>` URL from the HTML attachment to download the image.

```csharp
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Teams.Apps.Clients;
using Microsoft.Teams.Apps.Schema;

HttpClient imageClient = app.Services
    .GetRequiredService<IHttpClientFactory>()
    .CreateClient(nameof(ApiClient));

teams.OnMessage(async (context, cancellationToken) =>
{
    TeamsAttachment? image =
        context.Activity.Attachments?.FirstOrDefault(
            attachment =>
                attachment.ContentType is not null &&
                attachment.ContentType.Value.StartsWith(
                    "image/",
                    StringComparison.OrdinalIgnoreCase) &&
                attachment.ContentUrl is not null);

    if (image?.ContentUrl is not Uri contentUrl)
    {
        return;
    }

    using HttpResponseMessage response =
        await imageClient.GetAsync(
            contentUrl,
            HttpCompletionOption.ResponseHeadersRead,
            cancellationToken);

    response.EnsureSuccessStatusCode();

    byte[] bytes =
        await response.Content.ReadAsByteArrayAsync(
            cancellationToken);

    // Process the image bytes.
});
```

Key APIs and values:

* `nameof(ApiClient)`: Select the SDK client with platform authentication.
* `ContentType`: Match an `image/*` MIME type case-insensitively.
* `ContentUrl`: Use the canonical authenticated image download URL.
* `ResponseHeadersRead`: Begin processing without buffering response content first.
* `EnsureSuccessStatusCode`: Fail explicitly when image retrieval is unsuccessful.

Process the bytes directly when possible. Convert them to base64 only when a downstream API requires it, and preserve the attachment's actual MIME type.

### Send an inline image

An outbound inline image uses an image MIME type in `ContentType` and either a reachable HTTPS URL or a base64 data URI in `ContentUrl`. Sending doesn't upload the image to OneDrive or SharePoint and requires no Graph permission or user sign-in.

Teams supports inline images with the following limits:

| Limit | Value |
| --- | --- |
| Maximum dimensions | 1024 x 1024 pixels |
| Maximum image size | 1 MB |
| Supported formats | PNG, JPEG, and GIF |
| Animated GIF | Not supported |
| SDK validation | Not performed |

> [!WARNING]
>
> An image outside the supported dimensions, size, or format can be accepted by the send API but fail to render in the Teams client.

#### Send an image from a hosted URL

```csharp
TeamsAttachment image = TeamsAttachment.CreateBuilder()
    .WithContentType(new AttachmentContentType("image/png"))
    .WithContentUrl(new Uri("https://contoso.com/charts/weekly.png"))
    .WithName("weekly.png")
    .Build();

await context.SendAsync(
    new MessageActivityInput()
        .WithText("Here is the latest chart:")
        .AddAttachment(image),
    cancellationToken);
```

Key properties and values:

* `ContentType`: Set the MIME type matching hosted image bytes.
* `ContentUrl`: Set a publicly reachable HTTPS image URL.
* `Name`: Add an optional display name for the image.
* `AddAttachment`: Attach the image to the outgoing message.

The URL must be reachable by the Teams client without agent credentials. Prefer this approach for images that aren't small.

#### Send image bytes as base64

```csharp
byte[] bytes = await File.ReadAllBytesAsync(
    "./charts/weekly.png",
    cancellationToken);

string dataUri =
    $"data:image/png;base64,{Convert.ToBase64String(bytes)}";

TeamsAttachment image = TeamsAttachment.CreateBuilder()
    .WithContentType(new AttachmentContentType("image/png"))
    .WithContentUrl(new Uri(dataUri))
    .WithName("weekly.png")
    .Build();

await context.SendAsync(
    new MessageActivityInput()
        .WithText("Here is the latest chart:")
        .AddAttachment(image),
    cancellationToken);
```

Key properties and values:

* `dataUri`: Prefix base64 bytes with the matching MIME type.
* `ContentType`: Match the data URI and actual image format.
* `ContentUrl`: Wrap the complete data URI in `Uri`.
* `Name`: Add an optional filename matching the image format.

Base64 increases payload size by approximately one-third. Use it for small generated images and use a hosted URL for larger images.

#### Position an image within message text

Use HTML/XML content to control image placement and dimensions:

```csharp
string encoded = Convert.ToBase64String(
    await File.ReadAllBytesAsync(
        "./charts/weekly.png",
        cancellationToken));

await context.SendAsync(
    new MessageActivityInput()
        .WithText(
            $"<div>Revenue is up." +
            $"<img src=\"data:image/png;base64,{encoded}\"/>" +
            $"Questions?</div>")
        .WithTextFormat(TextFormats.Xml),
    cancellationToken);
```

Key properties and values:

* `src`: Use a base64 data URI for embedded images.
* `TextFormats.Xml`: Enable HTML elements within the message body.
* `height` and `width`: Add optional dimensions to the image element.

An HTTPS `<img src>` isn't uploaded or rewritten. Use a hosted image attachment instead. Every embedded image counts toward the message payload, and exceeding the limit can reject the entire message.

For multiple standalone images, add each attachment and select `List`, `Carousel`, or `Grid` through `AttachmentLayoutType`. For an image within interactive content, use an Adaptive Card `Image` element. For streamed responses, add attachments only to the final message because intermediate typing activities don't carry attachments.

## Handle errors

Handle known SDK file errors separately from transport and service failures:

| Status code | Error code | Description | Developer action |
| --- | --- | --- | --- |
| Not applicable | `FileUrlExpiredException` | The pre-authorized file URL expired before retrieval or re-read. | Ask the user to attach the file again. Download once and reuse `DownloadedFile`. |
| Not applicable | `FileCredentialException` | No Graph credential is available for agentic-user file retrieval. | Verify blueprint permissions, administrator consent, and credential configuration. |
| HTTP 401 | `FileAccessException` | Microsoft Graph rejected the token used to retrieve the file. | Refresh or correct the token and verify the selected actor. |
| HTTP 403 | `FileAccessException` | The identity lacks consent or access to the requested item. | Verify consent, sharing, blueprint permissions, and item access. |
| Not applicable | `FileScopeNotSupportedException` | The high-level file API received an unsupported conversation scope. | Use personal chat or retrieve the stored file through Microsoft Graph. |
| Not applicable | `FileException` | A known SDK file operation failed for another reason. | Show a general user message and log sanitized diagnostics. |
| HTTP 5xx | Transport or service error | Microsoft Graph or storage service couldn't complete the request. | Retry according to service guidance and preserve the original exception. |

A `403` can represent missing consent, missing item access, or a nonexistent item. The SDK doesn't expose a separate file-not-found error for these responses.

```csharp
try
{
    DownloadedFile downloaded =
        await file.DownloadAsync(cancellationToken);
}
catch (FileUrlExpiredException)
{
    await context.ReplyAsync(
        "That file link expired. Attach the file again.",
        cancellationToken);
}
catch (FileCredentialException)
{
    await context.ReplyAsync(
        "The agent isn't configured to access this file.",
        cancellationToken);
}
catch (FileAccessException error)
{
    await context.ReplyAsync(
        $"The storage service denied access ({error.Status}).",
        cancellationToken);
}
catch (FileScopeNotSupportedException)
{
    await context.ReplyAsync(
        "File download isn't supported in this conversation.",
        cancellationToken);
}
catch (FileException)
{
    await context.ReplyAsync(
        "The file couldn't be read.",
        cancellationToken);
}
```

Key types and values:

* `FileUrlExpiredException`: Handle expired short-lived file download URLs.
* `FileCredentialException`: Handle missing agentic-user Graph credential configuration.
* `FileAccessException.Status`: Distinguish rejected tokens from denied access.
* `FileScopeNotSupportedException`: Handle unsupported group or channel retrieval.
* `FileException`: Provide fallback handling for known SDK failures.

## Code sample

The following sample placeholder is reserved for the end-to-end TypeScript implementation:

| Sample name | Description | TypeScript |
| --- | --- | --- |
| File and inline-image handling | Receive and send files and inline images with an agent. | Link to be added |

## Design guidelines and best practices

* Clearly tell users whether content appears as a stored file or an inline image.
* Acknowledge successful file receipt, upload, or image processing.
* Explain why the agent requests file consent before sending a document.
* Handle a declined consent request without repeatedly prompting the user.
* Prefer `context.Files` over manually parsing received document attachments.
* Download once and reuse `DownloadedFile` when processing content repeatedly.
* Stream large files instead of buffering the complete file in memory.
* Sanitize uploader-provided filenames and use application-controlled destinations.
* Validate actual content instead of trusting extensions or reported MIME types.
* Use the authenticated SDK HTTP client for inbound inline-image URLs.
* Search all attachments instead of assuming the first is an image.
* Preserve the actual image MIME type when receiving or sending images.
* Prefer hosted URLs for larger images and keep base64 images small.
* Validate image size, dimensions, and format before sending.
* Don't log file bytes, credentials, or pre-authorized upload and download URLs.
* Apply expiration and cleanup policies to pending file-consent uploads.
* Propagate cancellation and handle HTTP, credential, scope, and expiration failures.

## See also

* [Teams SDK overview](/microsoftteams/platform/teams-sdk/why)
* [File and Image Handling](https://microsoft.github.io/teams-sdk/csharp/in-depth-guides/file-handling/)
* [Receiving Files](https://microsoft.github.io/teams-sdk/csharp/in-depth-guides/file-handling/receiving-files/)
* [Receiving Inline Images](https://microsoft.github.io/teams-sdk/csharp/in-depth-guides/file-handling/receiving-inline-images/)
* [Sending Inline Images](https://microsoft.github.io/teams-sdk/csharp/in-depth-guides/file-handling/sending-inline-images/)
* [Sending messages](/microsoftteams/platform/teams-sdk/essentials/sending-messages/overview?pivots=csharp)
* [User authentication](/microsoftteams/platform/teams-sdk/in-depth-guides/user-authentication?tabs=portal&pivots=csharp)
