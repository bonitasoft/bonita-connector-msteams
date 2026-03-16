package com.bonitasoft.connectors.msteams.btt;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicLong;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/**
 * Programmatically builds a Bonita .bar (Business Archive) as a ZIP file
 * containing a process with all 14 MS Teams connector operations wired as automatic tasks.
 *
 * <p>Generates the .bar directly as a ZIP (no dependency on Bonita's ProcessDefinitionBuilder),
 * using the exact XML formats expected by Bonita runtime.</p>
 *
 * <p>The process flow is linear:
 * Start -> sendChannelMessage -> sendChatMessage -> replyMessage -> sendAdaptiveCard
 * -> createChannel -> listChannels -> deleteChannel -> listTeams -> getTeam
 * -> addMember -> removeMember -> createMeeting -> getMeeting -> uploadFile -> End</p>
 */
public final class MSTeamsProcessBarBuilder {

    public static final String PROCESS_NAME = "MSTeamsConnectorTestProcess";
    public static final String PROCESS_VERSION = "1.0";
    private static final String CONNECTOR_VERSION = "1.0.0";
    private static final String BONITA_JAR_NAME = "bonita-connector-msteams-all-1.0.0-bonita.jar";

    private static final AtomicLong ID_COUNTER = new AtomicLong(1000000);

    private MSTeamsProcessBarBuilder() {
        // utility class
    }

    /**
     * Builds the .bar and writes it to the specified file.
     *
     * @param outputFile the target .bar file
     */
    public static void writeToFile(File outputFile) throws IOException {
        outputFile.getParentFile().mkdirs();

        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(outputFile))) {
            // 1. process-design.xml
            zos.putNextEntry(new ZipEntry("process-design.xml"));
            zos.write(buildProcessDesignXml().getBytes(StandardCharsets.UTF_8));
            zos.closeEntry();

            // 2. actorMapping.xml
            zos.putNextEntry(new ZipEntry("actorMapping.xml"));
            zos.write(buildActorMappingXml().getBytes(StandardCharsets.UTF_8));
            zos.closeEntry();

            // 3. form-mapping.xml
            zos.putNextEntry(new ZipEntry("form-mapping.xml"));
            zos.write(buildFormMappingXml().getBytes(StandardCharsets.UTF_8));
            zos.closeEntry();

            // 4. Extract .impl files from the bonita JAR and add under connector/
            File bonitaJar = findBonitaJar();
            List<String> implEntries = extractImplFiles(bonitaJar, zos);

            // 5. Add the shaded bonita JAR itself under classpath/
            zos.putNextEntry(new ZipEntry("classpath/" + bonitaJar.getName()));
            zos.write(java.nio.file.Files.readAllBytes(bonitaJar.toPath()));
            zos.closeEntry();
        }
    }

    // =========================================================================
    // XML Generators
    // =========================================================================

    private static String buildProcessDesignXml() {
        ID_COUNTER.set(1000000);

        StringBuilder xml = new StringBuilder();
        xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        xml.append("<tns:processDefinition id=\"_").append(nextId())
                .append("\" name=\"").append(PROCESS_NAME)
                .append("\" version=\"").append(PROCESS_VERSION)
                .append("\" xmlns:tns=\"http://www.bonitasoft.org/ns/process/client/7.4\">\n");

        // Actor
        String actorId = "_" + nextId();
        xml.append("    <actors>\n");
        xml.append("        <actor id=\"").append(actorId).append("\" name=\"System\" initiator=\"true\">\n");
        xml.append("            <description></description>\n");
        xml.append("        </actor>\n");
        xml.append("    </actors>\n");
        xml.append("    <actorInitiator>").append(actorId).append("</actorInitiator>\n");

        // Flow elements container
        String flowId = "_" + nextId();
        xml.append("    <flowElements id=\"").append(flowId).append("\">\n");

        // Build tasks with connectors
        List<ConnectorDef> connectors = buildConnectorDefs();
        String[] taskNames = {
                "sendChannelMessage", "sendChatMessage", "replyMessage", "sendAdaptiveCard",
                "createChannel", "listChannels", "deleteChannel", "listTeams", "getTeam",
                "addMember", "removeMember", "createMeeting", "getMeeting", "uploadFile"
        };

        // Generate task IDs and transition IDs
        String startId = "_" + nextId();
        String endId = "_" + nextId();
        String[] taskIds = new String[14];
        for (int i = 0; i < 14; i++) {
            taskIds[i] = "_" + nextId();
        }

        // Transition IDs: Start->task0, task0->task1, ..., task13->End = 15 transitions
        String[] transitionIds = new String[15];
        for (int i = 0; i < 15; i++) {
            transitionIds[i] = "_" + nextId();
        }

        // Automatic tasks with connectors
        for (int i = 0; i < 14; i++) {
            xml.append("        <automaticTask id=\"").append(taskIds[i])
                    .append("\" name=\"").append(taskNames[i]).append("\">\n");
            xml.append("            <incomingTransition>").append(transitionIds[i]).append("</incomingTransition>\n");
            xml.append("            <outgoingTransition>").append(transitionIds[i + 1]).append("</outgoingTransition>\n");

            // Connector
            ConnectorDef conn = connectors.get(i);
            String connElemId = "_" + nextId();
            xml.append("            <connector id=\"").append(connElemId)
                    .append("\" name=\"").append(conn.definitionId).append("-connector")
                    .append("\" connectorId=\"").append(conn.definitionId)
                    .append("\" activationEvent=\"ON_ENTER\" version=\"").append(CONNECTOR_VERSION)
                    .append("\" failAction=\"FAIL\">\n");

            // Inputs
            xml.append("                <inputs>\n");
            for (InputMapping input : conn.inputs) {
                String exprId = "_" + nextId();
                xml.append("                    <input name=\"").append(input.connectorInputName).append("\">\n");
                xml.append("                        <expression id=\"").append(exprId)
                        .append("\" name=\"").append(input.processVariableName)
                        .append("\" expressionType=\"TYPE_VARIABLE\" returnType=\"java.lang.String\" interpreter=\"\">\n");
                xml.append("                            <content>").append(input.processVariableName).append("</content>\n");
                xml.append("                        </expression>\n");
                xml.append("                    </input>\n");
            }
            xml.append("                </inputs>\n");

            // Outputs
            xml.append("                <outputs>\n");
            for (OutputMapping output : conn.outputs) {
                String rightId = "_" + nextId();
                xml.append("                    <operation operatorType=\"ASSIGNMENT\">\n");
                xml.append("                        <leftOperand name=\"").append(output.processVariableName)
                        .append("\" type=\"DATA\"/>\n");
                xml.append("                        <rightOperand id=\"").append(rightId)
                        .append("\" name=\"").append(output.connectorOutputName)
                        .append("\" expressionType=\"TYPE_INPUT\" returnType=\"").append(output.returnType)
                        .append("\" interpreter=\"\">\n");
                xml.append("                            <content>").append(output.connectorOutputName).append("</content>\n");
                xml.append("                        </rightOperand>\n");
                xml.append("                    </operation>\n");
            }
            xml.append("                </outputs>\n");

            xml.append("            </connector>\n");
            xml.append("            <dataDefinitions/>\n");
            xml.append("            <businessDataDefinitions/>\n");
            xml.append("            <operations/>\n");
            xml.append("            <boundaryEvents/>\n");
            xml.append("        </automaticTask>\n");
        }

        // Transitions
        xml.append("        <transitions>\n");
        // Start -> task0
        xml.append("            <transition id=\"").append(transitionIds[0])
                .append("\" name=\"Start-&gt;").append(taskNames[0])
                .append("\" source=\"").append(startId)
                .append("\" target=\"").append(taskIds[0]).append("\"/>\n");
        // task[i] -> task[i+1]
        for (int i = 0; i < 13; i++) {
            xml.append("            <transition id=\"").append(transitionIds[i + 1])
                    .append("\" name=\"").append(taskNames[i]).append("-&gt;").append(taskNames[i + 1])
                    .append("\" source=\"").append(taskIds[i])
                    .append("\" target=\"").append(taskIds[i + 1]).append("\"/>\n");
        }
        // task13 -> End
        xml.append("            <transition id=\"").append(transitionIds[14])
                .append("\" name=\"").append(taskNames[13]).append("-&gt;End")
                .append("\" source=\"").append(taskIds[13])
                .append("\" target=\"").append(endId).append("\"/>\n");
        xml.append("        </transitions>\n");

        // Start event
        xml.append("        <startEvent id=\"").append(startId).append("\" name=\"Start\" interrupting=\"true\">\n");
        xml.append("            <outgoingTransition>").append(transitionIds[0]).append("</outgoingTransition>\n");
        xml.append("        </startEvent>\n");

        // End event
        xml.append("        <endEvent id=\"").append(endId).append("\" name=\"End\">\n");
        xml.append("            <incomingTransition>").append(transitionIds[14]).append("</incomingTransition>\n");
        xml.append("        </endEvent>\n");

        // Process-level data definitions
        xml.append("        <dataDefinitions>\n");
        // String variables
        for (String varName : getStringVariables()) {
            String dataId = "_" + nextId();
            xml.append("            <textDataDefinition longText=\"false\" id=\"").append(dataId)
                    .append("\" name=\"").append(varName)
                    .append("\" transient=\"false\" className=\"java.lang.String\"/>\n");
        }
        // Boolean variable: success
        String successId = "_" + nextId();
        xml.append("            <dataDefinition id=\"").append(successId)
                .append("\" name=\"success\" transient=\"false\" className=\"java.lang.Boolean\"/>\n");
        xml.append("        </dataDefinitions>\n");

        xml.append("        <businessDataDefinitions/>\n");
        xml.append("        <documentDefinitions/>\n");
        xml.append("        <documentListDefinitions/>\n");
        xml.append("        <connectors/>\n");
        xml.append("        <elementFinder/>\n");
        xml.append("    </flowElements>\n");

        // String indexes
        xml.append("    <stringIndexes>\n");
        for (int i = 1; i <= 5; i++) {
            xml.append("        <stringIndex index=\"").append(i).append("\"/>\n");
        }
        xml.append("    </stringIndexes>\n");
        xml.append("    <context/>\n");
        xml.append("</tns:processDefinition>\n");

        return xml.toString();
    }

    private static String buildActorMappingXml() {
        return """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <actorMappings:actorMappings xmlns:actorMappings="http://www.bonitasoft.org/ns/actormapping/6.0">
                    <actorMapping name="System">
                        <users/>
                        <groups/>
                        <roles>
                            <role>member</role>
                        </roles>
                        <memberships/>
                    </actorMapping>
                </actorMappings:actorMappings>
                """;
    }

    private static String buildFormMappingXml() {
        return """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <form:formMappingModel xmlns:form="http://www.bonitasoft.org/ns/form/6.0">
                    <form-mappings/>
                </form:formMappingModel>
                """;
    }

    // =========================================================================
    // Connector Definitions
    // =========================================================================

    private static List<ConnectorDef> buildConnectorDefs() {
        List<ConnectorDef> defs = new ArrayList<>();

        // 1. Send Channel Message
        defs.add(new ConnectorDef("msteams-send-channel-message",
                List.of(
                        input("tenantId"), input("clientId"), input("clientSecret"),
                        input("teamId"), input("channelId"), input("messageContent")),
                List.of(
                        boolOutput("success", "success"),
                        strOutput("errorMessage", "errorMessage"),
                        strOutput("messageId", "parentMessageId"))));

        // 2. Send Chat Message
        defs.add(new ConnectorDef("msteams-send-chat-message",
                List.of(
                        input("tenantId"), input("clientId"), input("clientSecret"),
                        input("chatId"), input("messageContent")),
                List.of(
                        boolOutput("success", "success"),
                        strOutput("errorMessage", "errorMessage"))));

        // 3. Reply Message
        defs.add(new ConnectorDef("msteams-reply-message",
                List.of(
                        input("tenantId"), input("clientId"), input("clientSecret"),
                        input("teamId"), input("channelId"), input("parentMessageId"),
                        input("messageContent")),
                List.of(
                        boolOutput("success", "success"),
                        strOutput("errorMessage", "errorMessage"))));

        // 4. Send Adaptive Card
        defs.add(new ConnectorDef("msteams-send-adaptive-card",
                List.of(
                        input("tenantId"), input("clientId"), input("clientSecret"),
                        input("teamId"), input("channelId"), input("cardJson")),
                List.of(
                        boolOutput("success", "success"),
                        strOutput("errorMessage", "errorMessage"))));

        // 5. Create Channel (displayName -> channelDisplayName, description -> channelDescription)
        defs.add(new ConnectorDef("msteams-create-channel",
                List.of(
                        input("tenantId"), input("clientId"), input("clientSecret"),
                        input("teamId"),
                        new InputMapping("displayName", "channelDisplayName"),
                        new InputMapping("description", "channelDescription")),
                List.of(
                        boolOutput("success", "success"),
                        strOutput("errorMessage", "errorMessage"),
                        strOutput("channelId", "createdChannelId"))));

        // 6. List Channels
        defs.add(new ConnectorDef("msteams-list-channels",
                List.of(
                        input("tenantId"), input("clientId"), input("clientSecret"),
                        input("teamId")),
                List.of(
                        boolOutput("success", "success"),
                        strOutput("errorMessage", "errorMessage"))));

        // 7. Delete Channel
        defs.add(new ConnectorDef("msteams-delete-channel",
                List.of(
                        input("tenantId"), input("clientId"), input("clientSecret"),
                        input("teamId"), input("channelId")),
                List.of(
                        boolOutput("success", "success"),
                        strOutput("errorMessage", "errorMessage"))));

        // 8. List Teams
        defs.add(new ConnectorDef("msteams-list-teams",
                List.of(
                        input("tenantId"), input("clientId"), input("clientSecret")),
                List.of(
                        boolOutput("success", "success"),
                        strOutput("errorMessage", "errorMessage"))));

        // 9. Get Team
        defs.add(new ConnectorDef("msteams-get-team",
                List.of(
                        input("tenantId"), input("clientId"), input("clientSecret"),
                        input("teamId")),
                List.of(
                        boolOutput("success", "success"),
                        strOutput("errorMessage", "errorMessage"))));

        // 10. Add Member
        defs.add(new ConnectorDef("msteams-add-member",
                List.of(
                        input("tenantId"), input("clientId"), input("clientSecret"),
                        input("teamId"), input("userPrincipalName")),
                List.of(
                        boolOutput("success", "success"),
                        strOutput("errorMessage", "errorMessage"),
                        strOutput("membershipId", "membershipId"))));

        // 11. Remove Member
        defs.add(new ConnectorDef("msteams-remove-member",
                List.of(
                        input("tenantId"), input("clientId"), input("clientSecret"),
                        input("teamId"), input("membershipId")),
                List.of(
                        boolOutput("success", "success"),
                        strOutput("errorMessage", "errorMessage"))));

        // 12. Create Meeting (subject -> meetingSubject)
        defs.add(new ConnectorDef("msteams-create-meeting",
                List.of(
                        input("tenantId"), input("clientId"), input("clientSecret"),
                        input("userId"),
                        new InputMapping("subject", "meetingSubject"),
                        input("startDateTime"), input("endDateTime")),
                List.of(
                        boolOutput("success", "success"),
                        strOutput("errorMessage", "errorMessage"),
                        strOutput("meetingId", "meetingId"))));

        // 13. Get Meeting
        defs.add(new ConnectorDef("msteams-get-meeting",
                List.of(
                        input("tenantId"), input("clientId"), input("clientSecret"),
                        input("userId"), input("meetingId")),
                List.of(
                        boolOutput("success", "success"),
                        strOutput("errorMessage", "errorMessage"))));

        // 14. Upload File
        defs.add(new ConnectorDef("msteams-upload-file",
                List.of(
                        input("tenantId"), input("clientId"), input("clientSecret"),
                        input("teamId"), input("channelId"), input("fileName"),
                        input("fileContent")),
                List.of(
                        boolOutput("success", "success"),
                        strOutput("errorMessage", "errorMessage"))));

        return defs;
    }

    // =========================================================================
    // Variable Lists
    // =========================================================================

    private static List<String> getStringVariables() {
        return List.of(
                "tenantId", "clientId", "clientSecret", "userId", "teamId", "channelId",
                "chatId", "messageContent", "parentMessageId", "cardJson",
                "channelDisplayName", "channelDescription", "createdChannelId",
                "userPrincipalName", "membershipId", "meetingSubject",
                "startDateTime", "endDateTime", "meetingId", "fileName", "fileContent",
                "errorMessage");
    }

    // =========================================================================
    // JAR / .impl Extraction
    // =========================================================================

    private static File findBonitaJar() {
        Path targetDir = Path.of("target");
        File jarFile = targetDir.resolve(BONITA_JAR_NAME).toFile();
        if (!jarFile.exists()) {
            // Try from module directory
            jarFile = Path.of("bonita-connector-msteams-all", "target", BONITA_JAR_NAME).toFile();
        }
        if (!jarFile.exists()) {
            throw new IllegalStateException(
                    "Bonita JAR not found. Run 'mvn clean install -DskipTests' first. " +
                            "Expected: target/" + BONITA_JAR_NAME);
        }
        return jarFile;
    }

    /**
     * Extracts root-level .impl files from the bonita JAR and adds them
     * under the {@code connector/} directory in the .bar ZIP.
     */
    private static List<String> extractImplFiles(File bonitaJar, ZipOutputStream zos) throws IOException {
        List<String> implNames = new ArrayList<>();

        try (ZipFile zip = new ZipFile(bonitaJar)) {
            Enumeration<? extends ZipEntry> entries = zip.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                String name = entry.getName();
                // Root-level .impl files only (not in subdirectories)
                if (name.endsWith(".impl") && !name.contains("/")) {
                    zos.putNextEntry(new ZipEntry("connector/" + name));
                    try (InputStream is = zip.getInputStream(entry)) {
                        is.transferTo(zos);
                    }
                    zos.closeEntry();
                    implNames.add(name);
                }
            }
        }

        if (implNames.isEmpty()) {
            throw new IllegalStateException(
                    "No .impl files found at root level in " + bonitaJar.getName() +
                            ". Verify the connector modules are built correctly.");
        }

        return implNames;
    }

    // =========================================================================
    // Internal Model
    // =========================================================================

    private record ConnectorDef(String definitionId, List<InputMapping> inputs, List<OutputMapping> outputs) {}

    private record InputMapping(String connectorInputName, String processVariableName) {}

    private record OutputMapping(String connectorOutputName, String processVariableName, String returnType) {}

    /** Creates an InputMapping where connector input name matches process variable name. */
    private static InputMapping input(String name) {
        return new InputMapping(name, name);
    }

    private static OutputMapping strOutput(String connectorOutput, String processVar) {
        return new OutputMapping(connectorOutput, processVar, "java.lang.String");
    }

    private static OutputMapping boolOutput(String connectorOutput, String processVar) {
        return new OutputMapping(connectorOutput, processVar, "java.lang.Boolean");
    }

    private static long nextId() {
        return ID_COUNTER.getAndIncrement();
    }
}
