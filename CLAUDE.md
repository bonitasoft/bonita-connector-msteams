# Bonita MS Teams Connector - Project Instructions

## Skill: Bonita Connector Expert

### Architecture
- Multi-module Maven project: parent POM + common + 14 operations + all aggregator
- Java 17 with Bonita Runtime 10.2.0
- Microsoft Graph API v1.0 with App-Only (client_credentials) authentication
- Java HttpClient for HTTP + Jackson for JSON

### Build Commands
```bash
# Full build (produces -bonita.jar per module)
mvn clean install -DskipTests

# Single module (requires prior install of dependencies)
mvn clean package -DskipTests -pl bonita-connector-msteams-send-channel-message -am

# Run unit tests
mvn test

# Verify JAR contents
jar tf bonita-connector-msteams-all/target/bonita-connector-msteams-all-*-bonita.jar | grep -E '\.(def|impl|properties|png)$'
```

### Import Protocol (Bonita Studio 2025.2+)
1. Build: `mvn clean install -DskipTests`
2. Locate: `bonita-connector-msteams-all/target/bonita-connector-msteams-all-*-bonita.jar`
3. Studio: Extensions > Import extension > select the `-bonita.jar`
4. Verify: connectors appear under "MS Teams" category in task configuration

### Connector File Structure (per module)

```
src/main/resources/
  {id}.def              # EMF ConnectorDefinition XML (namespace 6.1)
  {id}.properties       # i18n labels
  msteams.png           # 16x16 binary PNG icon

src/main/resources-filtered/
  {id}.impl             # Implementation descriptor (namespace 6.0, Maven-filtered)
```

### EMF .def Critical Rules
- Namespace: `http://www.bonitasoft.org/ns/connector/definition/6.1`
- NEVER use `<label>` or `<description>` child elements (causes eResource() null error)
- `inputName` MUST be an attribute on `<widget>`, never a child element
- `<category>` MUST be self-closing: `<category icon="x.png" id="y"/>`

### .impl Critical Rules
- Namespace: `http://www.bonitasoft.org/ns/connector/implementation/6.0`
- `definitionId` and `definitionVersion` MUST match .def `<id>` and `<version>` exactly
- `${connector-dependencies}` placeholder resolved by Groovy script at build time
- MUST be in `src/main/resources-filtered/` (not `src/main/resources/`)

### Java Lifecycle (4 phases)
1. `validateInputParameters()` - Validate all inputs, throw `ConnectorValidationException`
2. `connect()` - Initialize MSTeamsClient, authenticate with Microsoft Graph
3. `executeBusinessLogic()` - Call `executeOperation(MSTeamsClient)`, set outputs
4. `disconnect()` - Close resources

### MSTeamsClient API
```java
client.get(path)                          // GET → JsonNode
client.post(path, jsonBody)               // POST → JsonNode
client.put(path, byte[], contentType)     // PUT → JsonNode (binary upload)
client.delete(path)                       // DELETE → JsonNode
client.getObjectMapper()                  // ObjectMapper for JSON building
client.getConfiguration()                 // MSTeamsConfiguration record
client.authenticate()                     // OAuth token acquisition
```

### Authentication
- **App-Only**: client_credentials grant (tenantId + clientId + clientSecret)
- Token acquired via POST to `https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token`
- Token refresh and retry logic built into MSTeamsClient

### Error Prevention Checklist (before marking done)
1. `mvn clean install -DskipTests` succeeds
2. `jar tf *-bonita.jar` contains .def, .impl, .properties, .png at root
3. .impl inside JAR has resolved `<jarDependencies>` (not `${connector-dependencies}`)
4. .def has NO `<label>`/`<description>` child elements
5. All widget `inputName` are attributes, not children
6. `<category>` is self-closing
7. Icon is valid 16x16 PNG binary
8. `definitionId` in .impl matches `<id>` in .def
9. `implementationClassname` exists as .class in the JAR
10. Tests pass: `mvn test`

### Quality Targets
- Test naming: `should_X_when_Y()`
- Error messages truncated to 1000 chars (H2 column limit)

### Module Dependencies
```
common (base classes, MSTeamsClient, MSTeamsException, no .def)
  ├── send-channel-message  → common + Jackson
  ├── send-chat-message     → common + Jackson
  ├── reply-message         → common + Jackson
  ├── send-adaptive-card    → common + Jackson
  ├── create-channel        → common + Jackson
  ├── list-channels         → common + Jackson
  ├── delete-channel        → common + Jackson
  ├── list-teams            → common + Jackson
  ├── get-team              → common + Jackson
  ├── add-member            → common + Jackson
  ├── remove-member         → common + Jackson
  ├── create-meeting        → common + Jackson
  ├── get-meeting           → common + Jackson
  ├── upload-file           → common + Jackson
  └── all                   → depends on ALL above (single fat JAR)
```

### Operations Summary
| Operation | Graph API Endpoint | Method |
|-----------|-------------------|--------|
| Send Channel Message | /teams/{id}/channels/{id}/messages | POST |
| Send Chat Message | /chats/{id}/messages | POST |
| Reply Message | /teams/{id}/channels/{id}/messages/{id}/replies | POST |
| Send Adaptive Card | /teams/{id}/channels/{id}/messages | POST |
| Create Channel | /teams/{id}/channels | POST |
| List Channels | /teams/{id}/channels | GET |
| Delete Channel | /teams/{id}/channels/{id} | DELETE |
| List Teams | /groups or /me/joinedTeams | GET |
| Get Team | /teams/{id} | GET |
| Add Member | /teams/{id}/members | POST |
| Remove Member | /teams/{id}/members/{id} | DELETE |
| Create Meeting | /users/{id}/onlineMeetings | POST |
| Get Meeting | /users/{id}/onlineMeetings/{id} | GET |
| Upload File | /drives/{id}/items/{id}:/{name}:/content | PUT |
