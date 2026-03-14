package com.bonitasoft.connectors.msteams.integration;

import java.util.HashMap;
import java.util.Map;

import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.ConnectorLifecycleTestHelper;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsConfiguration;
import com.bonitasoft.connectors.msteams.addmember.MSTeamsAddMemberConnector;
import com.bonitasoft.connectors.msteams.createchannel.MSTeamsCreateChannelConnector;
import com.bonitasoft.connectors.msteams.createmeeting.MSTeamsCreateMeetingConnector;
import com.bonitasoft.connectors.msteams.deletechannel.MSTeamsDeleteChannelConnector;
import com.bonitasoft.connectors.msteams.getmeeting.MSTeamsGetMeetingConnector;
import com.bonitasoft.connectors.msteams.getteam.MSTeamsGetTeamConnector;
import com.bonitasoft.connectors.msteams.listchannels.MSTeamsListChannelsConnector;
import com.bonitasoft.connectors.msteams.listteams.MSTeamsListTeamsConnector;
import com.bonitasoft.connectors.msteams.removemember.MSTeamsRemoveMemberConnector;
import com.bonitasoft.connectors.msteams.replymessage.MSTeamsReplyMessageConnector;
import com.bonitasoft.connectors.msteams.sendadaptivecard.MSTeamsSendAdaptiveCardConnector;
import com.bonitasoft.connectors.msteams.sendchannelmessage.MSTeamsSendChannelMessageConnector;
import com.bonitasoft.connectors.msteams.sendchatmessage.MSTeamsSendChatMessageConnector;
import com.bonitasoft.connectors.msteams.uploadfile.MSTeamsUploadFileConnector;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import org.junit.jupiter.api.*;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * Integration test simulating the full MSTeamsConnectorTestProcess.proc flow.
 * <p>
 * Executes all 14 connectors in the same sequence as the .proc diagram:
 * Start → SendChannelMessage → SendChatMessage → ReplyMessage → SendAdaptiveCard →
 * CreateChannel → ListChannels → DeleteChannel → ListTeams → GetTeam →
 * AddMember → RemoveMember → CreateMeeting → GetMeeting → UploadFile → End
 * <p>
 * Process variables are propagated between connectors (e.g., messageId from
 * SendChannelMessage becomes parentMessageId for ReplyMessage).
 * <p>
 * MSTeamsClient is mocked — no real Graph API calls are made.
 */
@Tag("integration")
@ExtendWith(MockitoExtension.class)
@DisplayName("Process Flow Simulation - MSTeamsConnectorTestProcess.proc")
@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
class MSTeamsProcessFlowIT {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    // Simulated process variables (shared across all test steps)
    private static final Map<String, Object> processVars = new HashMap<>();

    @Mock
    private MSTeamsClient client;

    @BeforeAll
    static void initProcessVariables() {
        // Auth variables (shared by all connectors)
        processVars.put("tenantId", "test-tenant-id");
        processVars.put("clientId", "test-client-id");
        processVars.put("clientSecret", "test-client-secret");
        processVars.put("userId", "test-user-id");

        // Messaging variables
        processVars.put("teamId", "team-001");
        processVars.put("channelId", "channel-001");
        processVars.put("chatId", "chat-001");
        processVars.put("messageContent", "Test message from process flow simulation");
        processVars.put("cardJson", """
                {"type":"message","attachments":[{"contentType":"application/vnd.microsoft.card.adaptive",\
                "content":{"type":"AdaptiveCard","version":"1.4","body":[{"type":"TextBlock","text":"Test"}]}}]}""");

        // Channel management variables
        processVars.put("channelDisplayName", "Test-Channel");
        processVars.put("channelDescription", "Created by process flow test");

        // Team membership variables
        processVars.put("userPrincipalName", "testuser@contoso.com");

        // Meeting variables
        processVars.put("meetingSubject", "Process Flow Test Meeting");
        processVars.put("startDateTime", "2026-12-01T10:00:00.000Z");
        processVars.put("endDateTime", "2026-12-01T11:00:00.000Z");

        // File variables
        processVars.put("fileName", "test-file.txt");
        processVars.put("fileContent", "dGVzdCBjb250ZW50"); // base64 "test content"
    }

    private static final MSTeamsConfiguration TEST_CONFIG =
            new MSTeamsConfiguration("test-tenant-id", "test-client-id", "test-client-secret",
                    5000, 10000, true);

    @BeforeEach
    void setupClient() {
        lenient().when(client.getObjectMapper()).thenReturn(MAPPER);
        lenient().when(client.getConfiguration()).thenReturn(TEST_CONFIG);
    }

    // ========================================================================
    // Step 1: Send Channel Message
    // ========================================================================

    @Test
    @Order(1)
    @DisplayName("Step 01 - should_send_channel_message_when_process_starts")
    void should_send_channel_message_when_process_starts() throws Exception {
        // Mock: POST /teams/{teamId}/channels/{channelId}/messages
        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "msg-001");
        response.put("webUrl", "https://teams.microsoft.com/msg/001");
        response.put("createdDateTime", "2026-03-14T10:00:00Z");
        when(client.post(contains("/messages"), anyString())).thenReturn(response);

        var connector = new MSTeamsSendChannelMessageConnector();
        Map<String, Object> inputs = buildInputs(connector,
                "tenantId", "clientId", "clientSecret", "teamId", "channelId",
                "messageContent");
        // contentType is a static value in .proc: expression="text"
        inputs.put("contentType", "text");
        connector.setInputParameters(inputs);

        executeAndVerify(connector);

        // .proc maps: messageId → parentMessageId
        processVars.put("parentMessageId", connector.getOutputs().get("messageId"));
        assertThat(processVars.get("parentMessageId")).isEqualTo("msg-001");
    }

    // ========================================================================
    // Step 2: Send Chat Message
    // ========================================================================

    @Test
    @Order(2)
    @DisplayName("Step 02 - should_send_chat_message_when_channel_message_sent")
    void should_send_chat_message_when_channel_message_sent() throws Exception {
        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "chat-msg-001");
        response.put("createdDateTime", "2026-03-14T10:01:00Z");
        when(client.post(contains("/chats/"), anyString())).thenReturn(response);

        var connector = new MSTeamsSendChatMessageConnector();
        connector.setInputParameters(buildInputs(connector,
                "tenantId", "clientId", "clientSecret", "chatId", "messageContent"));

        executeAndVerify(connector);
    }

    // ========================================================================
    // Step 3: Reply Message
    // ========================================================================

    @Test
    @Order(3)
    @DisplayName("Step 03 - should_reply_message_when_parent_message_exists")
    void should_reply_message_when_parent_message_exists() throws Exception {
        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "reply-001");
        response.put("createdDateTime", "2026-03-14T10:02:00Z");
        when(client.post(contains("/replies"), anyString())).thenReturn(response);

        var connector = new MSTeamsReplyMessageConnector();
        connector.setInputParameters(buildInputs(connector,
                "tenantId", "clientId", "clientSecret", "teamId", "channelId",
                "parentMessageId", "messageContent"));

        executeAndVerify(connector);
        assertThat(connector.getOutputs().get("replyId")).isEqualTo("reply-001");
    }

    // ========================================================================
    // Step 4: Send Adaptive Card
    // ========================================================================

    @Test
    @Order(4)
    @DisplayName("Step 04 - should_send_adaptive_card_when_valid_json")
    void should_send_adaptive_card_when_valid_json() throws Exception {
        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "card-msg-001");
        response.put("webUrl", "https://teams.microsoft.com/card/001");
        response.put("createdDateTime", "2026-03-14T10:03:00Z");
        when(client.post(contains("/messages"), anyString())).thenReturn(response);

        var connector = new MSTeamsSendAdaptiveCardConnector();
        connector.setInputParameters(buildInputs(connector,
                "tenantId", "clientId", "clientSecret", "teamId", "channelId", "cardJson"));

        executeAndVerify(connector);
    }

    // ========================================================================
    // Step 5: Create Channel
    // ========================================================================

    @Test
    @Order(5)
    @DisplayName("Step 05 - should_create_channel_when_valid_team")
    void should_create_channel_when_valid_team() throws Exception {
        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "new-channel-001");
        response.put("displayName", "Test-Channel");
        response.put("webUrl", "https://teams.microsoft.com/channel/001");
        when(client.post(contains("/channels"), anyString())).thenReturn(response);

        var connector = new MSTeamsCreateChannelConnector();
        Map<String, Object> inputs = buildInputs(connector,
                "tenantId", "clientId", "clientSecret", "teamId");
        inputs.put("displayName", processVars.get("channelDisplayName"));
        inputs.put("description", processVars.get("channelDescription"));
        connector.setInputParameters(inputs);

        executeAndVerify(connector);

        // .proc maps: channelId → createdChannelId
        processVars.put("createdChannelId", connector.getOutputs().get("channelId"));
        assertThat(processVars.get("createdChannelId")).isEqualTo("new-channel-001");
    }

    // ========================================================================
    // Step 6: List Channels
    // ========================================================================

    @Test
    @Order(6)
    @DisplayName("Step 06 - should_list_channels_when_valid_team")
    void should_list_channels_when_valid_team() throws Exception {
        ObjectNode response = MAPPER.createObjectNode();
        ArrayNode channels = response.putArray("value");
        ObjectNode ch = channels.addObject();
        ch.put("id", "channel-001");
        ch.put("displayName", "General");
        ch.put("description", "Default channel");
        ch.put("membershipType", "standard");
        when(client.get(contains("/channels"))).thenReturn(response);

        var connector = new MSTeamsListChannelsConnector();
        connector.setInputParameters(buildInputs(connector,
                "tenantId", "clientId", "clientSecret", "teamId"));

        executeAndVerify(connector);
        assertThat(connector.getOutputs().get("totalCount")).isEqualTo(1);
    }

    // ========================================================================
    // Step 7: Delete Channel
    // ========================================================================

    @Test
    @Order(7)
    @DisplayName("Step 07 - should_delete_channel_when_channel_exists")
    void should_delete_channel_when_channel_exists() throws Exception {
        // DELETE returns null (204 No Content)
        when(client.delete(contains("/channels/"))).thenReturn(null);

        var connector = new MSTeamsDeleteChannelConnector();
        Map<String, Object> inputs = buildInputs(connector,
                "tenantId", "clientId", "clientSecret", "teamId");
        // .proc uses: channelId expression="createdChannelId"
        inputs.put("channelId", processVars.get("createdChannelId"));
        connector.setInputParameters(inputs);

        executeAndVerify(connector);
    }

    // ========================================================================
    // Step 8: List Teams
    // ========================================================================

    @Test
    @Order(8)
    @DisplayName("Step 08 - should_list_teams_when_authenticated")
    void should_list_teams_when_authenticated() throws Exception {
        ObjectNode response = MAPPER.createObjectNode();
        ArrayNode teams = response.putArray("value");
        ObjectNode team = teams.addObject();
        team.put("id", "team-001");
        team.put("displayName", "Test Team");
        team.put("description", "A test team");
        team.put("visibility", "public");
        when(client.get(contains("/groups"))).thenReturn(response);

        var connector = new MSTeamsListTeamsConnector();
        connector.setInputParameters(buildInputs(connector,
                "tenantId", "clientId", "clientSecret"));

        executeAndVerify(connector);
        assertThat(connector.getOutputs().get("totalCount")).isEqualTo(1);
    }

    // ========================================================================
    // Step 9: Get Team
    // ========================================================================

    @Test
    @Order(9)
    @DisplayName("Step 09 - should_get_team_when_valid_team_id")
    void should_get_team_when_valid_team_id() throws Exception {
        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "team-001");
        response.put("displayName", "Test Team");
        response.put("description", "A test team");
        response.put("visibility", "public");
        response.put("isArchived", false);
        response.put("webUrl", "https://teams.microsoft.com/team/001");
        when(client.get(contains("/teams/"))).thenReturn(response);

        var connector = new MSTeamsGetTeamConnector();
        connector.setInputParameters(buildInputs(connector,
                "tenantId", "clientId", "clientSecret", "teamId"));

        executeAndVerify(connector);
        assertThat(connector.getOutputs().get("displayName")).isEqualTo("Test Team");
    }

    // ========================================================================
    // Step 10: Add Member
    // ========================================================================

    @Test
    @Order(10)
    @DisplayName("Step 10 - should_add_member_when_valid_user")
    void should_add_member_when_valid_user() throws Exception {
        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "member-001");
        response.put("displayName", "Test User");
        when(client.post(contains("/members"), anyString())).thenReturn(response);

        var connector = new MSTeamsAddMemberConnector();
        connector.setInputParameters(buildInputs(connector,
                "tenantId", "clientId", "clientSecret", "teamId", "userPrincipalName"));

        executeAndVerify(connector);

        // .proc maps: membershipId → membershipId
        processVars.put("membershipId", connector.getOutputs().get("membershipId"));
        assertThat(processVars.get("membershipId")).isEqualTo("member-001");
    }

    // ========================================================================
    // Step 11: Remove Member
    // ========================================================================

    @Test
    @Order(11)
    @DisplayName("Step 11 - should_remove_member_when_membership_exists")
    void should_remove_member_when_membership_exists() throws Exception {
        when(client.delete(contains("/members/"))).thenReturn(null);

        var connector = new MSTeamsRemoveMemberConnector();
        connector.setInputParameters(buildInputs(connector,
                "tenantId", "clientId", "clientSecret", "teamId", "membershipId"));

        executeAndVerify(connector);
    }

    // ========================================================================
    // Step 12: Create Meeting
    // ========================================================================

    @Test
    @Order(12)
    @DisplayName("Step 12 - should_create_meeting_when_valid_schedule")
    void should_create_meeting_when_valid_schedule() throws Exception {
        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "meeting-001");
        response.put("joinWebUrl", "https://teams.microsoft.com/meet/001");
        response.put("subject", "Process Flow Test Meeting");
        response.put("startDateTime", "2026-12-01T10:00:00Z");
        response.put("endDateTime", "2026-12-01T11:00:00Z");
        ObjectNode joinSettings = response.putObject("joinMeetingIdSettings");
        joinSettings.put("joinMeetingId", "join-001");
        when(client.post(contains("/onlineMeetings"), anyString())).thenReturn(response);

        var connector = new MSTeamsCreateMeetingConnector();
        Map<String, Object> inputs = buildInputs(connector,
                "tenantId", "clientId", "clientSecret", "userId");
        inputs.put("subject", processVars.get("meetingSubject"));
        inputs.put("startDateTime", processVars.get("startDateTime"));
        inputs.put("endDateTime", processVars.get("endDateTime"));
        connector.setInputParameters(inputs);

        executeAndVerify(connector);

        // .proc maps: meetingId → meetingId
        processVars.put("meetingId", connector.getOutputs().get("meetingId"));
        assertThat(processVars.get("meetingId")).isEqualTo("meeting-001");
    }

    // ========================================================================
    // Step 13: Get Meeting
    // ========================================================================

    @Test
    @Order(13)
    @DisplayName("Step 13 - should_get_meeting_when_meeting_exists")
    void should_get_meeting_when_meeting_exists() throws Exception {
        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "meeting-001");
        response.put("joinWebUrl", "https://teams.microsoft.com/meet/001");
        response.put("subject", "Process Flow Test Meeting");
        response.put("startDateTime", "2026-12-01T10:00:00Z");
        response.put("endDateTime", "2026-12-01T11:00:00Z");
        response.putObject("participants");
        when(client.get(contains("/onlineMeetings/"))).thenReturn(response);

        var connector = new MSTeamsGetMeetingConnector();
        connector.setInputParameters(buildInputs(connector,
                "tenantId", "clientId", "clientSecret", "userId", "meetingId"));

        executeAndVerify(connector);
        assertThat(connector.getOutputs().get("subject")).isEqualTo("Process Flow Test Meeting");
    }

    // ========================================================================
    // Step 14: Upload File
    // ========================================================================

    @Test
    @Order(14)
    @DisplayName("Step 14 - should_upload_file_when_valid_content")
    void should_upload_file_when_valid_content() throws Exception {
        // Step 1: GET /teams/{teamId}/channels/{channelId}/filesFolder
        ObjectNode folderResponse = MAPPER.createObjectNode();
        ObjectNode parentRef = folderResponse.putObject("parentReference");
        parentRef.put("driveId", "drive-001");
        folderResponse.put("id", "folder-001");
        when(client.get(contains("/filesFolder"))).thenReturn(folderResponse);

        // Step 2: PUT /drives/{driveId}/items/{folderId}:/{fileName}:/content
        ObjectNode uploadResponse = MAPPER.createObjectNode();
        uploadResponse.put("id", "file-001");
        uploadResponse.put("name", "test-file.txt");
        uploadResponse.put("webUrl", "https://sharepoint.com/file/001");
        uploadResponse.put("size", 12L);
        when(client.put(contains("/drives/"), any(byte[].class), anyString())).thenReturn(uploadResponse);

        var connector = new MSTeamsUploadFileConnector();
        connector.setInputParameters(buildInputs(connector,
                "tenantId", "clientId", "clientSecret", "teamId", "channelId",
                "fileName", "fileContent"));

        executeAndVerify(connector);
        assertThat(connector.getOutputs().get("fileId")).isEqualTo("file-001");
    }

    // ========================================================================
    // Process completion verification
    // ========================================================================

    @Test
    @Order(100)
    @DisplayName("should_have_all_process_variables_set_after_complete_flow")
    void should_have_all_process_variables_set_after_complete_flow() {
        assertThat(processVars.get("parentMessageId")).as("parentMessageId from Step 1").isNotNull();
        assertThat(processVars.get("createdChannelId")).as("createdChannelId from Step 5").isNotNull();
        assertThat(processVars.get("membershipId")).as("membershipId from Step 10").isNotNull();
        assertThat(processVars.get("meetingId")).as("meetingId from Step 12").isNotNull();
    }

    @Test
    @Order(101)
    @DisplayName("should_verify_proc_file_is_parseable")
    void should_verify_proc_file_is_parseable() throws Exception {
        var procStream = getClass().getClassLoader()
                .getResourceAsStream("MSTeamsConnectorTestProcess.proc");
        assertThat(procStream)
                .as("MSTeamsConnectorTestProcess.proc should be on classpath")
                .isNotNull();

        // Verify it's valid XML
        var factory = javax.xml.parsers.DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        var doc = factory.newDocumentBuilder().parse(procStream);
        assertThat(doc.getDocumentElement().getLocalName()).isEqualTo("definitions");

        // Verify 14 service tasks
        var serviceTasks = doc.getElementsByTagNameNS(
                "http://www.omg.org/spec/BPMN/20100524/MODEL", "serviceTask");
        assertThat(serviceTasks.getLength()).isEqualTo(14);

        // Verify 15 sequence flows (14 between tasks + start-to-first)
        var flows = doc.getElementsByTagNameNS(
                "http://www.omg.org/spec/BPMN/20100524/MODEL", "sequenceFlow");
        assertThat(flows.getLength()).isEqualTo(15);
    }

    // ========================================================================
    // Helper methods
    // ========================================================================

    /**
     * Builds input parameters map from process variables for the given variable names.
     */
    private Map<String, Object> buildInputs(AbstractMSTeamsConnector connector, String... varNames) {
        var inputs = new HashMap<String, Object>();
        for (String name : varNames) {
            Object value = processVars.get(name);
            if (value != null) {
                inputs.put(name, value);
            }
        }
        return inputs;
    }

    /**
     * Simulates the Bonita connector lifecycle (minus connect) with a mocked client:
     * validate → inject mock client → executeBusinessLogic → disconnect.
     * <p>
     * Uses {@link ConnectorLifecycleTestHelper} to access the protected
     * {@code executeBusinessLogic()} method from the correct package.
     */
    private void executeAndVerify(AbstractMSTeamsConnector connector) throws Exception {
        ConnectorLifecycleTestHelper.executeWithMockClient(connector, client);

        assertThat(connector.getOutputs().get("success"))
                .as("Connector %s should succeed", connector.getClass().getSimpleName())
                .isEqualTo(true);
        assertThat(connector.getOutputs().get("errorMessage"))
                .as("Connector %s should have empty error", connector.getClass().getSimpleName())
                .isEqualTo("");
    }
}
