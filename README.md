# Bonita Connector for Microsoft Teams

A multi-module Bonita BPM connector providing 14 operations for Microsoft Teams integration via the Microsoft Graph API. Supports messaging, channel management, team operations, online meetings, and file uploads.

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Microsoft Teams Configuration](#microsoft-teams-configuration)
- [Build](#build)
- [Import into Bonita Studio](#import-into-bonita-studio)
- [Connector Configuration](#connector-configuration)
- [Operations Reference](#operations-reference)
- [Troubleshooting](#troubleshooting)
- [Development](#development)

---

## Overview

This connector enables Bonita BPM processes to interact with Microsoft Teams through the Microsoft Graph API v1.0. It provides 14 distinct operations organized into five categories:

| Category | Operations |
|----------|-----------|
| **Messaging** | Send Channel Message, Send Chat Message, Reply Message, Send Adaptive Card |
| **Channels** | Create Channel, List Channels, Delete Channel |
| **Teams** | List Teams, Get Team |
| **Members** | Add Member, Remove Member |
| **Meetings** | Create Meeting, Get Meeting |
| **Files** | Upload File |

Each operation is packaged as both an individual connector module and as part of a single aggregated JAR (`-all` module) for convenience.

---

## Prerequisites

| Requirement | Version |
|-------------|---------|
| Java (JDK) | 17 or later |
| Maven | 3.9+ |
| Bonita Studio | 2025.2+ (Runtime 10.2.0) |
| Microsoft 365 | Business Basic or higher (with Teams enabled) |
| Azure AD | App registration with appropriate permissions |

---

## Microsoft Teams Configuration

### Step 1: Register an Azure AD Application

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations** > **New registration**
3. Fill in:
   - **Name**: `Bonita MS Teams Connector` (or your preferred name)
   - **Supported account types**: "Accounts in this organizational directory only"
   - **Redirect URI**: Leave blank (not needed for client credentials)
4. Click **Register**
5. On the app overview page, note down:
   - **Application (client) ID** -- this is your `clientId`
   - **Directory (tenant) ID** -- this is your `tenantId`

### Step 2: Create a Client Secret

1. In your app registration, go to **Certificates & secrets** > **Client secrets** > **New client secret**
2. Add a description (e.g., "Bonita Connector") and choose an expiry period
3. Click **Add**
4. **Copy the secret value immediately** -- this is your `clientSecret` (it won't be shown again)

### Step 3: Grant Microsoft Graph API Permissions

Go to **API permissions** > **Add a permission** > **Microsoft Graph**.

#### Application Permissions (for App-Only operations)

These permissions allow the connector to act without a signed-in user. An Azure AD admin must grant consent.

| Permission | Used by |
|------------|---------|
| `ChannelMessage.Send` | Send Channel Message, Send Adaptive Card |
| `Chat.ReadWrite.All` | Send Chat Message |
| `ChannelMessage.Read.All` | Reply Message |
| `Channel.Create` | Create Channel |
| `Channel.ReadBasic.All` | List Channels |
| `Channel.Delete.All` | Delete Channel |
| `Team.ReadBasic.All` | List Teams, Get Team |
| `TeamMember.ReadWrite.All` | Add Member, Remove Member |
| `OnlineMeetings.ReadWrite.All` | Create Meeting, Get Meeting |
| `Files.ReadWrite.All` | Upload File |

#### Delegated Permissions (for user-context operations)

Some operations (Chat Message, Reply, Meetings) can also use delegated permissions with a refresh token:

| Permission | Used by |
|------------|---------|
| `Chat.ReadWrite` | Send Chat Message |
| `ChannelMessage.Send` | Reply Message |
| `OnlineMeetings.ReadWrite` | Create Meeting, Get Meeting |

### Step 4: Grant Admin Consent

1. In **API permissions**, click **Grant admin consent for [your organization]**
2. Confirm by clicking **Yes**
3. All permissions should show a green checkmark under "Status"

### Step 5: Collect Your Credentials

You need three values to configure every connector:

| Parameter | Where to find it |
|-----------|-----------------|
| `tenantId` | Azure AD > App registration > Overview > "Directory (tenant) ID" |
| `clientId` | Azure AD > App registration > Overview > "Application (client) ID" |
| `clientSecret` | Azure AD > App registration > Certificates & secrets > Client secret "Value" |

For delegated operations, you also need:
- `userId` -- The Object ID of the user (found in Azure AD > Users > select user > Object ID)

---

## Build

```bash
# Full build with tests
mvn clean verify

# Build without tests (faster)
mvn clean install -DskipTests

# Unit tests only
mvn test

# Integration tests (WireMock-based, no external dependencies)
mvn verify -pl bonita-connector-msteams-all

# E2E tests (requires real MS Teams credentials)
mvn verify -Pe2e -pl bonita-connector-msteams-all

# Mutation testing
mvn org.pitest:pitest-maven:mutationCoverage

# Coverage report
mvn jacoco:report
# Open target/site/jacoco/index.html in a browser
```

The build produces:
- `bonita-connector-msteams-all/target/bonita-connector-msteams-all-1.0.0.jar` -- aggregated JAR with all 14 operations
- `bonita-connector-msteams-all/target/bonita-connector-msteams-all-1.0.0-bonita.jar` -- shaded JAR ready for Bonita Studio import (includes all dependencies)

---

## Import into Bonita Studio

### Option A: Import the All-in-One JAR (Recommended)

1. Build the project:
   ```bash
   mvn clean install -DskipTests
   ```
2. Open Bonita Studio 2025.2+
3. Go to **Extensions** > **Import extension...**
4. Select the file:
   ```
   bonita-connector-msteams-all/target/bonita-connector-msteams-all-1.0.0-bonita.jar
   ```
5. Click **Import**
6. The 14 connectors will appear under the **msteams** category when configuring service tasks

### Option B: Import Individual Connectors

Each operation module also produces its own `-bonita.jar`:
```
bonita-connector-msteams-send-channel-message/target/bonita-connector-msteams-send-channel-message-1.0.0-bonita.jar
bonita-connector-msteams-list-teams/target/bonita-connector-msteams-list-teams-1.0.0-bonita.jar
...
```

Import only the connectors you need using the same **Extensions** > **Import extension** workflow.

### Verify the Import

1. Create or open a process
2. Add a **Service Task**
3. Go to the task's **Execution** > **Connectors in** tab
4. Click **Add...** and look for the **msteams** category
5. You should see all 14 connector definitions listed

---

## Connector Configuration

### Common Parameters (All Connectors)

Every connector requires these authentication parameters on the **Connection** page:

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `tenantId` | String | Yes | Azure AD tenant ID |
| `clientId` | String | Yes | Azure AD application (client) ID |
| `clientSecret` | String | Yes | Azure AD client secret value |
| `connectTimeout` | Integer | No | Connection timeout in ms (default: 30000) |
| `readTimeout` | Integer | No | Read timeout in ms (default: 60000) |

For delegated operations (Chat Message, Reply, Create/Get Meeting), additional parameters:

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `userId` | String | Yes* | User Object ID for delegated operations |
| `refreshToken` | String | No | OAuth refresh token for delegated auth |
| `authMode` | String | No | Authentication mode override |

### Using Process Variables

For best practice, store authentication credentials in process variables and reference them using Groovy expressions:

1. Create process-level variables: `msTeamsTenantId`, `msTeamsClientId`, `msTeamsClientSecret`
2. In each connector's Connection page, use expressions like:
   ```groovy
   msTeamsTenantId
   ```
3. Set variable values at process instantiation or via a configuration form

### Output Parameters

All connectors produce at least these two outputs:

| Output | Type | Description |
|--------|------|-------------|
| `success` | Boolean | `true` if the operation completed successfully |
| `errorMessage` | String | Error details if `success` is `false`, empty string otherwise |

Use a connector output operation to store these in process variables for error handling downstream.

---

## Operations Reference

### 1. Send Channel Message

Sends a text or HTML message to a Teams channel.

| | |
|-|-|
| **Connector ID** | `msteams-send-channel-message` |
| **Graph API** | `POST /teams/{teamId}/channels/{channelId}/messages` |
| **Auth Mode** | App-only |
| **Permission** | `ChannelMessage.Send` |

**Inputs:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `teamId` | String | Yes | Team ID |
| `channelId` | String | Yes | Channel ID |
| `messageContent` | String | Yes | Message body (max 28,000 chars) |
| `contentType` | String | No | `text` (default) or `html` |
| `subject` | String | No | Message subject line |
| `importance` | String | No | `normal` (default), `high`, or `urgent` |

**Outputs:** `messageId`, `webUrl`, `createdDateTime`

---

### 2. Send Chat Message

Sends a message to a 1:1 or group chat.

| | |
|-|-|
| **Connector ID** | `msteams-send-chat-message` |
| **Graph API** | `POST /chats/{chatId}/messages` |
| **Auth Mode** | Delegated |
| **Permission** | `Chat.ReadWrite` |

**Inputs:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `chatId` | String | Yes | Chat thread ID |
| `messageContent` | String | Yes | Message body |
| `contentType` | String | No | `text` (default) or `html` |
| `importance` | String | No | `normal` (default), `high`, or `urgent` |

**Outputs:** `messageId`, `chatId`, `createdDateTime`

---

### 3. Reply Message

Replies to an existing channel message.

| | |
|-|-|
| **Connector ID** | `msteams-reply-message` |
| **Graph API** | `POST /teams/{teamId}/channels/{channelId}/messages/{messageId}/replies` |
| **Auth Mode** | Delegated |
| **Permission** | `ChannelMessage.Send` |

**Inputs:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `teamId` | String | Yes | Team ID |
| `channelId` | String | Yes | Channel ID |
| `parentMessageId` | String | Yes | ID of the message to reply to |
| `messageContent` | String | Yes | Reply body |
| `contentType` | String | No | `text` (default) or `html` |

**Outputs:** `replyId`, `createdDateTime`

---

### 4. Send Adaptive Card

Sends an Adaptive Card (v1.4) to a channel.

| | |
|-|-|
| **Connector ID** | `msteams-send-adaptive-card` |
| **Graph API** | `POST /teams/{teamId}/channels/{channelId}/messages` |
| **Auth Mode** | App-only |
| **Permission** | `ChannelMessage.Send` |

**Inputs:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `teamId` | String | Yes | Team ID |
| `channelId` | String | Yes | Channel ID |
| `cardJson` | String | Yes | Full Adaptive Card JSON payload |
| `subject` | String | No | Message subject |
| `importance` | String | No | `normal`, `high`, or `urgent` |

**Outputs:** `messageId`, `webUrl`, `createdDateTime`

---

### 5. Create Channel

Creates a new channel in a team.

| | |
|-|-|
| **Connector ID** | `msteams-create-channel` |
| **Graph API** | `POST /teams/{teamId}/channels` |
| **Auth Mode** | App-only |
| **Permission** | `Channel.Create` |

**Inputs:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `teamId` | String | Yes | Team ID |
| `displayName` | String | Yes | Channel name |
| `description` | String | No | Channel description |
| `membershipType` | String | No | `standard` (default) or `private` |

**Outputs:** `channelId`, `displayName`, `webUrl`

---

### 6. List Channels

Lists all channels in a team.

| | |
|-|-|
| **Connector ID** | `msteams-list-channels` |
| **Graph API** | `GET /teams/{teamId}/channels` |
| **Auth Mode** | App-only |
| **Permission** | `Channel.ReadBasic.All` |

**Inputs:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `teamId` | String | Yes | Team ID |

**Outputs:** `channels` (List), `totalCount` (Integer)

---

### 7. Delete Channel

Deletes a channel from a team.

| | |
|-|-|
| **Connector ID** | `msteams-delete-channel` |
| **Graph API** | `DELETE /teams/{teamId}/channels/{channelId}` |
| **Auth Mode** | App-only |
| **Permission** | `Channel.Delete.All` |

**Inputs:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `teamId` | String | Yes | Team ID |
| `channelId` | String | Yes | Channel ID to delete |

**Outputs:** `success`, `errorMessage` (standard only)

---

### 8. List Teams

Lists all teams in the organization.

| | |
|-|-|
| **Connector ID** | `msteams-list-teams` |
| **Graph API** | `GET /groups` (filtered for teams) |
| **Auth Mode** | App-only |
| **Permission** | `Team.ReadBasic.All` |

**Inputs:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `filter` | String | No | OData filter expression |
| `maxResults` | Integer | No | Maximum number of teams to return |
| `appOnly` | Boolean | No | Force app-only auth mode |

**Outputs:** `teams` (List), `totalCount` (Integer)

---

### 9. Get Team

Gets details of a specific team.

| | |
|-|-|
| **Connector ID** | `msteams-get-team` |
| **Graph API** | `GET /teams/{teamId}` |
| **Auth Mode** | App-only |
| **Permission** | `Team.ReadBasic.All` |

**Inputs:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `teamId` | String | Yes | Team ID |

**Outputs:** `teamId`, `displayName`, `description`, `visibility`, `isArchived` (Boolean), `webUrl`

---

### 10. Add Member

Adds a user as a member to a team.

| | |
|-|-|
| **Connector ID** | `msteams-add-member` |
| **Graph API** | `POST /teams/{teamId}/members` |
| **Auth Mode** | App-only |
| **Permission** | `TeamMember.ReadWrite.All` |

**Inputs:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `teamId` | String | Yes | Team ID |
| `userPrincipalName` | String | Yes | User's UPN (e.g., `user@contoso.com`) |
| `role` | String | No | `member` (default) or `owner` |

**Outputs:** `membershipId`, `displayName`

---

### 11. Remove Member

Removes a member from a team.

| | |
|-|-|
| **Connector ID** | `msteams-remove-member` |
| **Graph API** | `DELETE /teams/{teamId}/members/{membershipId}` |
| **Auth Mode** | App-only |
| **Permission** | `TeamMember.ReadWrite.All` |

**Inputs:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `teamId` | String | Yes | Team ID |
| `membershipId` | String | Yes | Membership ID (from Add Member output) |

**Outputs:** `success`, `errorMessage` (standard only)

---

### 12. Create Meeting

Creates an online meeting.

| | |
|-|-|
| **Connector ID** | `msteams-create-meeting` |
| **Graph API** | `POST /users/{userId}/onlineMeetings` |
| **Auth Mode** | Delegated / App-only |
| **Permission** | `OnlineMeetings.ReadWrite.All` |

**Inputs:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `userId` | String | Yes | Organizer's user Object ID |
| `subject` | String | Yes | Meeting subject |
| `startDateTime` | String | Yes | ISO 8601 datetime (e.g., `2026-01-15T10:00:00.000Z`) |
| `endDateTime` | String | Yes | ISO 8601 datetime |
| `allowedPresenters` | String | No | `everyone`, `organization`, `roleIsPresenter`, `organizer` |
| `lobbyBypassScope` | String | No | `organizer`, `organization`, `organizationAndFederated`, `everyone` |
| `recordAutomatically` | Boolean | No | Auto-record the meeting |

**Outputs:** `meetingId`, `joinWebUrl`, `meetingCode`, `subject`, `startDateTime`, `endDateTime`

---

### 13. Get Meeting

Gets details of an online meeting.

| | |
|-|-|
| **Connector ID** | `msteams-get-meeting` |
| **Graph API** | `GET /users/{userId}/onlineMeetings/{meetingId}` |
| **Auth Mode** | Delegated / App-only |
| **Permission** | `OnlineMeetings.ReadWrite.All` |

**Inputs:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `userId` | String | Yes | Organizer's user Object ID |
| `meetingId` | String | Yes | Meeting ID (from Create Meeting output) |

**Outputs:** `meetingId`, `joinWebUrl`, `subject`, `startDateTime`, `endDateTime`, `participants` (Map)

---

### 14. Upload File

Uploads a file to a channel's file tab (SharePoint).

| | |
|-|-|
| **Connector ID** | `msteams-upload-file` |
| **Graph API** | `PUT /drives/{driveId}/items/{folderId}:/{fileName}:/content` |
| **Auth Mode** | App-only |
| **Permission** | `Files.ReadWrite.All` |

**Inputs:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `teamId` | String | Yes | Team ID |
| `channelId` | String | Yes | Channel ID |
| `fileName` | String | Yes | Target file name (max 255 chars) |
| `fileContent` | String | Yes | File content as Base64-encoded string |
| `contentType` | String | No | MIME type (default: `application/octet-stream`) |

**Outputs:** `fileId`, `fileName`, `webUrl`, `size` (Long)

---

## Troubleshooting

### Authentication Errors

| Error | Cause | Solution |
|-------|-------|----------|
| `AADSTS7000215: Invalid client secret` | Client secret is wrong or expired | Create a new client secret in Azure AD and update the connector |
| `AADSTS700016: Application not found` | Wrong `clientId` or `tenantId` | Verify the values match your app registration |
| `AADSTS65001: User or admin has not consented` | Missing admin consent | Go to API Permissions and click "Grant admin consent" |

### Permission Errors

| Error | Cause | Solution |
|-------|-------|----------|
| `403 Forbidden - Insufficient privileges` | Missing API permission | Add the required permission in Azure AD and grant admin consent |
| `Authorization_RequestDenied` | App-only permissions not granted | Ensure Application permissions (not Delegated) are configured |

### Connector Errors

| Error | Cause | Solution |
|-------|-------|----------|
| `Failed to connect: Failed to authenticate` | Network or credential issue | Check firewall rules; verify credentials |
| `messageContent exceeds maximum length` | Message > 28,000 chars | Shorten the message or split into multiple messages |
| `fileContent must be valid Base64 encoded data` | Invalid Base64 input | Ensure file content is properly Base64-encoded |
| `Could not resolve driveId or folderId` | Channel has no Files tab | Ensure the channel has been initialized (open Files tab once in Teams) |
| `connectTimeout` / `readTimeout` | Network latency | Increase timeout values (default: 30s connect, 60s read) |

### Finding IDs

| ID | How to find it |
|----|---------------|
| `teamId` | Teams Admin Center > Teams > select team > copy from URL; or use List Teams connector |
| `channelId` | Use List Channels connector; or right-click channel in Teams > "Get link to channel" |
| `chatId` | Graph Explorer: `GET /me/chats` |
| `userId` | Azure AD > Users > select user > Object ID |
| `membershipId` | Output from Add Member connector |
| `meetingId` | Output from Create Meeting connector |

---

## Development

### Project Structure

```
bonita-connector-msteams/
├── pom.xml                                    # Parent POM (Java 17, dependency management)
├── bonita-connector-msteams-common/           # Shared: AbstractMSTeamsConnector, MSTeamsClient,
│   └── src/main/java/.../                     #   MSTeamsConfiguration, MSTeamsException,
│                                              #   RetryPolicy, AdaptiveCardBuilder
├── bonita-connector-msteams-send-channel-message/  # Operation module (1 of 14)
│   ├── src/main/java/.../MSTeamsSendChannelMessageConnector.java
│   ├── src/main/resources/msteams-send-channel-message.def         # Connector definition
│   ├── src/main/resources/msteams-send-channel-message.properties  # i18n labels
│   ├── src/main/resources/msteams.png                              # 16x16 icon
│   ├── src/main/resources-filtered/msteams-send-channel-message.impl  # Implementation descriptor
│   └── src/test/java/.../MSTeamsSendChannelMessageConnectorTest.java
├── bonita-connector-msteams-send-chat-message/     # ... (12 more operation modules)
├── bonita-connector-msteams-reply-message/
├── bonita-connector-msteams-send-adaptive-card/
├── bonita-connector-msteams-create-channel/
├── bonita-connector-msteams-list-channels/
├── bonita-connector-msteams-delete-channel/
├── bonita-connector-msteams-list-teams/
├── bonita-connector-msteams-get-team/
├── bonita-connector-msteams-add-member/
├── bonita-connector-msteams-remove-member/
├── bonita-connector-msteams-create-meeting/
├── bonita-connector-msteams-get-meeting/
├── bonita-connector-msteams-upload-file/
└── bonita-connector-msteams-all/              # Aggregator: shade all into one JAR
    └── src/test/
        ├── java/.../integration/GraphConnectorIT.java    # Assembly integration tests
        ├── java/.../e2e/MSTeamsConnectorE2ETest.java     # E2E tests (14 operations)
        ├── java/.../e2e/ConnectorTestToolkit.java        # Test helper utilities
        └── resources/MSTeamsConnectorTestProcess.proc    # BPMN test process
```

### Connector Lifecycle

Every connector follows the 4-phase Bonita lifecycle:

1. **VALIDATE** (`validateInputParameters`) -- check required inputs, validate formats
2. **CONNECT** (`connect`) -- create `MSTeamsClient`, authenticate via OAuth 2.0 client credentials
3. **EXECUTE** (`executeBusinessLogic` -> `executeOperation`) -- call Microsoft Graph API
4. **DISCONNECT** (`disconnect`) -- close HTTP client resources

### Adding a New Operation

1. Create a new module: `bonita-connector-msteams-{operation}/`
2. Add the module to the parent `pom.xml` `<modules>` section
3. Create the connector class extending `AbstractMSTeamsConnector`
4. Create `.def`, `.impl`, `.properties`, and icon files
5. Add the module as a dependency in `bonita-connector-msteams-all/pom.xml`
6. Add unit tests following the `should_X_when_Y()` naming convention
7. Build and verify: `mvn clean install`

### Running Tests

```bash
# Unit tests (fast, mocked)
mvn test

# Integration tests (WireMock, no external deps)
mvn verify -pl bonita-connector-msteams-all

# E2E tests (real MS Teams tenant required)
export MSTEAMS_TENANT_ID=your-tenant-id
export MSTEAMS_CLIENT_ID=your-client-id
export MSTEAMS_CLIENT_SECRET=your-client-secret
export MSTEAMS_USER_ID=user-object-id
export MSTEAMS_TEAM_ID=your-team-id
export MSTEAMS_CHANNEL_ID=your-channel-id
export MSTEAMS_CHAT_ID=your-chat-id
export MSTEAMS_USER_PRINCIPAL_NAME=user@contoso.com
mvn verify -Pe2e -pl bonita-connector-msteams-all
```

### Code Quality

- **JaCoCo**: 80% line coverage, 70% branch coverage (enforced)
- **PIT Mutation Testing**: 80% mutation threshold
- **Test naming**: `should_X_when_Y()` pattern
- **Error messages**: Truncated to 1,000 characters (Bonita H2 column limit)

### Dependencies

| Dependency | Version | Scope |
|------------|---------|-------|
| Bonita Runtime (`bonita-common`) | 10.2.0 | provided |
| Jackson Databind | 2.17.2 | compile |
| Lombok | 1.18.34 | provided |
| SLF4J | 2.0.13 | provided |
| JUnit 5 | 5.10.2 | test |
| Mockito | 5.11.0 | test |
| AssertJ | 3.25.3 | test |
| WireMock | 3.5.4 | test |
| Testcontainers | 1.19.7 | test |

---

## License

Copyright (C) 2025 BonitaSoft S.A. -- GPLv2
