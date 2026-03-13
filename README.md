# Bonita Connector - Microsoft Teams

Bonita BPM connector for Microsoft Teams integration. Supports 14 operations across messaging, channels, teams, meetings, and file management.

## Operations

| # | Connector ID | Auth Mode | Description |
|---|-------------|-----------|-------------|
| 1 | `msteams-send-channel-message` | Webhook | Send text/HTML message to channel |
| 2 | `msteams-send-adaptive-card` | Webhook | Send Adaptive Card (v1.4) to channel |
| 3 | `msteams-send-chat-message` | Delegated | Send message to a chat |
| 4 | `msteams-reply-message` | Delegated | Reply to a channel message |
| 5 | `msteams-list-channels` | App-only | List channels in a team |
| 6 | `msteams-create-channel` | App-only | Create a channel |
| 7 | `msteams-delete-channel` | App-only | Delete a channel |
| 8 | `msteams-list-teams` | App-only | List all teams |
| 9 | `msteams-get-team` | App-only | Get team details |
| 10 | `msteams-add-member` | App-only | Add member to team |
| 11 | `msteams-remove-member` | App-only | Remove member from team |
| 12 | `msteams-create-meeting` | Delegated | Create online meeting |
| 13 | `msteams-get-meeting` | Delegated | Get meeting details |
| 14 | `msteams-upload-file` | App-only | Upload file to channel |

## Authentication Modes

### Webhook (Connectors 1-2)
Uses Teams Incoming Webhook URLs. No Azure AD registration required.

### App-only (Connectors 5-11, 14)
Uses Azure AD app registration with client credentials.
- **Required:** `tenantId`, `clientId`, `clientSecret`
- **Permissions:** Application permissions in Microsoft Graph

### Delegated (Connectors 3-4, 12-13)
Uses Azure AD with user-delegated permissions via refresh token.
- **Required:** `tenantId`, `clientId`, `clientSecret`, `refreshToken`
- **Permissions:** Delegated permissions in Microsoft Graph

## Azure AD Setup

1. Register an app in [Azure Portal > App registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Add required API permissions:
   - **App-only:** `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `Channel.Create`, `Channel.Delete.All`, `TeamMember.ReadWrite.All`, `Files.ReadWrite.All`
   - **Delegated:** `Chat.ReadWrite`, `ChannelMessage.Send`, `OnlineMeetings.ReadWrite`
3. Create a client secret
4. For webhook connectors: Configure an [Incoming Webhook](https://learn.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/add-incoming-webhook) in the target channel

## Installation

### Import in Bonita Studio

1. Build the project: `./mvnw clean package`
2. Import the ZIP from `target/bonita-connector-msteams-1.0.0-SNAPSHOT-all.zip`
3. Or import individual connectors: `target/bonita-connector-msteams-1.0.0-SNAPSHOT-{operation}-impl.zip`

## Build

```bash
# Full build
./mvnw clean verify

# Unit tests only
./mvnw test

# Integration tests (WireMock)
./mvnw verify -PITs

# E2E tests (requires Docker)
./mvnw verify -PE2E

# Mutation testing
./mvnw org.pitest:pitest-maven:mutationCoverage

# Coverage report
./mvnw jacoco:report
open target/site/jacoco/index.html
```

## Dependencies

- **Bonita Runtime:** 10.2.0 (provided)
- **MSAL4J:** 1.17.2 (Azure AD auth)
- **Jackson:** 2.17.2 (JSON, provided by Bonita)
- **Java:** 17+

## Project Structure

```
src/main/java/com/bonitasoft/connectors/msteams/
├── common/          # Shared: auth, HTTP clients, retry, card builder
├── messaging/       # Send channel message, adaptive card, chat, reply
├── channels/        # List, create, delete channels
├── teams/           # List, get teams, add/remove members
├── meetings/        # Create, get online meetings
└── files/           # Upload files to channels
```

## License

Copyright (C) 2025 BonitaSoft S.A. — GPLv2
