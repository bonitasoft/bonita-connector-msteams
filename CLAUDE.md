# Bonita Connector MS Teams — Instructions for Claude Code

## What This Project Is

Bonita BPM connector for Microsoft Teams. Single-module Maven project with 14 connector operations across 3 auth modes (Webhook, App-only, Delegated).

## Architecture

- **Single-module** Maven project (not multi-module)
- Base package: `com.bonitasoft.connectors.msteams`
- Common classes in `.common` package
- Operations in `.messaging`, `.channels`, `.teams`, `.meetings`, `.files`

## Key Patterns

### Connector Lifecycle (4 phases)
1. **VALIDATE** — `validateInputParameters()` → `validateConnectionParameters()` + `validateOperationParameters()`
2. **CONNECT** — `connect()` → create client, authenticate
3. **EXECUTE** — `executeBusinessLogic()` → call API, set outputs
4. **DISCONNECT** — `disconnect()` → close client

### Class Hierarchy
```
AbstractMsTeamsConnector (base)
├── AbstractWebhookConnector (webhook auth, MsTeamsWebhookClient)
│   ├── SendChannelMessageConnector
│   └── SendAdaptiveCardConnector
└── AbstractGraphConnector (MSAL4J auth, MsTeamsGraphClient)
    ├── ListChannelsConnector, CreateChannelConnector, DeleteChannelConnector
    ├── ListTeamsConnector, GetTeamConnector, AddMemberConnector, RemoveMemberConnector
    ├── SendChatMessageConnector, ReplyMessageConnector (delegated)
    ├── CreateMeetingConnector, GetMeetingConnector (delegated)
    └── UploadFileConnector
```

### Output Convention
Every connector sets `success` (Boolean) and `errorMessage` (String) outputs.

## Build Commands

```bash
./mvnw clean verify          # Full build
./mvnw test                  # Unit tests
./mvnw verify -PITs          # Integration tests
./mvnw verify -PE2E          # E2E tests (Docker required)
```

## File Locations

| Type | Location |
|------|----------|
| Java sources | `src/main/java/com/bonitasoft/connectors/msteams/` |
| .def/.impl files | `src/main/resources-filtered/` (Maven-filtered) |
| .properties | `src/main/resources-filtered/` |
| Assembly XMLs | `src/assembly/` |
| Groovy script | `src/script/dependencies-as-var.groovy` |
| Unit tests | `src/test/java/.../` (*Test.java) |
| Integration tests | `src/test/java/.../integration/` (*IT.java) |

## Rules

- Java 17 features (records, pattern matching, text blocks)
- `@Slf4j` for logging (Lombok)
- Test naming: `should_X_when_Y()`
- EMF .def rules: no `<label>`/`<description>` children, self-closing `<category>`, `inputName` as attribute
- .impl files MUST be in `resources-filtered/` (Maven filtering)
- Standard outputs: `success` + `errorMessage` on every connector
