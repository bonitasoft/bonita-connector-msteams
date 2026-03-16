package com.bonitasoft.connectors.msteams.e2e;

import java.util.Base64;
import java.util.Map;

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

import org.junit.jupiter.api.*;

import static com.bonitasoft.connectors.msteams.e2e.ConnectorTestToolkit.*;
import static org.assertj.core.api.Assertions.assertThat;

/**
 * End-to-end tests for all 14 MS Teams connectors.
 * <p>
 * These tests execute real connector operations against a Microsoft Teams tenant.
 * They require valid Azure AD credentials set via environment variables:
 * <ul>
 *   <li>{@code MSTEAMS_TENANT_ID} - Azure AD tenant ID</li>
 *   <li>{@code MSTEAMS_CLIENT_ID} - Azure AD application (client) ID</li>
 *   <li>{@code MSTEAMS_CLIENT_SECRET} - Azure AD client secret</li>
 *   <li>{@code MSTEAMS_USER_ID} - User ID for delegated operations</li>
 *   <li>{@code MSTEAMS_TEAM_ID} - Test team ID</li>
 *   <li>{@code MSTEAMS_CHANNEL_ID} - Test channel ID</li>
 *   <li>{@code MSTEAMS_CHAT_ID} - Test chat ID</li>
 *   <li>{@code MSTEAMS_USER_PRINCIPAL_NAME} - UPN of a user to add/remove</li>
 * </ul>
 * <p>
 * Enable by removing @Disabled and providing credentials.
 * Run with: {@code mvn verify -Pe2e -pl bonita-connector-msteams-all}
 */
@Tag("provider-api")
@DisplayName("MS Teams Connector E2E Tests")
@Disabled("Requires MS Teams credentials - set MSTEAMS_* environment variables and remove @Disabled")
@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
class MSTeamsConnectorE2ETest {

    // Credentials from environment
    private static final String TENANT_ID = System.getenv("MSTEAMS_TENANT_ID");
    private static final String CLIENT_ID = System.getenv("MSTEAMS_CLIENT_ID");
    private static final String CLIENT_SECRET = System.getenv("MSTEAMS_CLIENT_SECRET");
    private static final String USER_ID = System.getenv("MSTEAMS_USER_ID");

    // Test data from environment
    private static final String TEAM_ID = System.getenv("MSTEAMS_TEAM_ID");
    private static final String CHANNEL_ID = System.getenv("MSTEAMS_CHANNEL_ID");
    private static final String CHAT_ID = System.getenv("MSTEAMS_CHAT_ID");
    private static final String USER_PRINCIPAL_NAME = System.getenv("MSTEAMS_USER_PRINCIPAL_NAME");

    // State shared between tests (set by earlier tests, used by later ones)
    private static String createdMessageId;
    private static String createdChannelId;
    private static String createdMembershipId;
    private static String createdMeetingId;

    private Map<String, Object> appAuth() {
        return authInputs(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    }

    private Map<String, Object> delegatedAuth() {
        return delegatedAuthInputs(TENANT_ID, CLIENT_ID, CLIENT_SECRET, USER_ID);
    }

    // ========================================================================
    // Messaging operations
    // ========================================================================

    @Test
    @Order(1)
    @DisplayName("01 - should_send_channel_message_when_valid_inputs")
    void should_send_channel_message_when_valid_inputs() throws Exception {
        var connector = new MSTeamsSendChannelMessageConnector();
        var inputs = mergeInputs(
                appAuth(),
                sendChannelMessageInputs(TEAM_ID, CHANNEL_ID, "E2E Test Message from Bonita connector")
        );

        var outputs = executeConnector(connector, inputs);

        assertThat(isSuccess(outputs)).isTrue();
        assertThat(outputs.get("messageId")).isNotNull();
        createdMessageId = (String) outputs.get("messageId");
    }

    @Test
    @Order(2)
    @DisplayName("02 - should_send_chat_message_when_valid_chat_id")
    void should_send_chat_message_when_valid_chat_id() throws Exception {
        var connector = new MSTeamsSendChatMessageConnector();
        var inputs = mergeInputs(
                delegatedAuth(),
                sendChatMessageInputs(CHAT_ID, "E2E Test Chat Message from Bonita connector")
        );

        var outputs = executeConnector(connector, inputs);

        assertThat(isSuccess(outputs)).isTrue();
        assertThat(outputs.get("messageId")).isNotNull();
    }

    @Test
    @Order(3)
    @DisplayName("03 - should_reply_to_message_when_parent_message_exists")
    void should_reply_to_message_when_parent_message_exists() throws Exception {
        Assumptions.assumeTrue(createdMessageId != null, "Requires messageId from test 01");

        var connector = new MSTeamsReplyMessageConnector();
        var inputs = mergeInputs(
                delegatedAuth(),
                replyMessageInputs(TEAM_ID, CHANNEL_ID, createdMessageId,
                        "E2E Reply from Bonita connector")
        );

        var outputs = executeConnector(connector, inputs);

        assertThat(isSuccess(outputs)).isTrue();
        assertThat(outputs.get("replyId")).isNotNull();
    }

    @Test
    @Order(4)
    @DisplayName("04 - should_send_adaptive_card_when_valid_json")
    void should_send_adaptive_card_when_valid_json() throws Exception {
        var connector = new MSTeamsSendAdaptiveCardConnector();
        String cardJson = """
                {
                  "type": "message",
                  "attachments": [{
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content": {
                      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                      "type": "AdaptiveCard",
                      "version": "1.4",
                      "body": [{
                        "type": "TextBlock",
                        "text": "Bonita E2E Test Card",
                        "weight": "Bolder",
                        "size": "Medium"
                      }]
                    }
                  }]
                }
                """;
        var inputs = mergeInputs(
                appAuth(),
                sendAdaptiveCardInputs(TEAM_ID, CHANNEL_ID, cardJson)
        );

        var outputs = executeConnector(connector, inputs);

        assertThat(isSuccess(outputs)).isTrue();
        assertThat(outputs.get("messageId")).isNotNull();
    }

    // ========================================================================
    // Channel operations
    // ========================================================================

    @Test
    @Order(5)
    @DisplayName("05 - should_create_channel_when_valid_team")
    void should_create_channel_when_valid_team() throws Exception {
        var connector = new MSTeamsCreateChannelConnector();
        var inputs = mergeInputs(
                appAuth(),
                createChannelInputs(TEAM_ID, "Bonita-E2E-Test-" + System.currentTimeMillis(),
                        "Channel created by Bonita E2E test")
        );

        var outputs = executeConnector(connector, inputs);

        assertThat(isSuccess(outputs)).isTrue();
        assertThat(outputs.get("channelId")).isNotNull();
        createdChannelId = (String) outputs.get("channelId");
    }

    @Test
    @Order(6)
    @DisplayName("06 - should_list_channels_when_valid_team")
    void should_list_channels_when_valid_team() throws Exception {
        var connector = new MSTeamsListChannelsConnector();
        var inputs = mergeInputs(
                appAuth(),
                listChannelsInputs(TEAM_ID)
        );

        var outputs = executeConnector(connector, inputs);

        assertThat(isSuccess(outputs)).isTrue();
        assertThat(outputs.get("totalCount")).isNotNull();
        assertThat((Integer) outputs.get("totalCount")).isGreaterThan(0);
    }

    @Test
    @Order(7)
    @DisplayName("07 - should_delete_channel_when_channel_exists")
    void should_delete_channel_when_channel_exists() throws Exception {
        Assumptions.assumeTrue(createdChannelId != null, "Requires channelId from test 05");

        var connector = new MSTeamsDeleteChannelConnector();
        var inputs = mergeInputs(
                appAuth(),
                deleteChannelInputs(TEAM_ID, createdChannelId)
        );

        var outputs = executeConnector(connector, inputs);

        assertThat(isSuccess(outputs)).isTrue();
    }

    // ========================================================================
    // Team operations
    // ========================================================================

    @Test
    @Order(8)
    @DisplayName("08 - should_list_teams_when_authenticated")
    void should_list_teams_when_authenticated() throws Exception {
        var connector = new MSTeamsListTeamsConnector();
        var inputs = mergeInputs(appAuth(), listTeamsInputs());

        var outputs = executeConnector(connector, inputs);

        assertThat(isSuccess(outputs)).isTrue();
        assertThat(outputs.get("totalCount")).isNotNull();
    }

    @Test
    @Order(9)
    @DisplayName("09 - should_get_team_when_valid_team_id")
    void should_get_team_when_valid_team_id() throws Exception {
        var connector = new MSTeamsGetTeamConnector();
        var inputs = mergeInputs(
                appAuth(),
                getTeamInputs(TEAM_ID)
        );

        var outputs = executeConnector(connector, inputs);

        assertThat(isSuccess(outputs)).isTrue();
        assertThat(outputs.get("displayName")).isNotNull();
    }

    // ========================================================================
    // Member operations
    // ========================================================================

    @Test
    @Order(10)
    @DisplayName("10 - should_add_member_when_valid_user")
    void should_add_member_when_valid_user() throws Exception {
        var connector = new MSTeamsAddMemberConnector();
        var inputs = mergeInputs(
                appAuth(),
                addMemberInputs(TEAM_ID, USER_PRINCIPAL_NAME)
        );

        var outputs = executeConnector(connector, inputs);

        assertThat(isSuccess(outputs)).isTrue();
        assertThat(outputs.get("membershipId")).isNotNull();
        createdMembershipId = (String) outputs.get("membershipId");
    }

    @Test
    @Order(11)
    @DisplayName("11 - should_remove_member_when_membership_exists")
    void should_remove_member_when_membership_exists() throws Exception {
        Assumptions.assumeTrue(createdMembershipId != null, "Requires membershipId from test 10");

        var connector = new MSTeamsRemoveMemberConnector();
        var inputs = mergeInputs(
                appAuth(),
                removeMemberInputs(TEAM_ID, createdMembershipId)
        );

        var outputs = executeConnector(connector, inputs);

        assertThat(isSuccess(outputs)).isTrue();
    }

    // ========================================================================
    // Meeting operations
    // ========================================================================

    @Test
    @Order(12)
    @DisplayName("12 - should_create_meeting_when_valid_schedule")
    void should_create_meeting_when_valid_schedule() throws Exception {
        var connector = new MSTeamsCreateMeetingConnector();
        var inputs = mergeInputs(
                delegatedAuth(),
                createMeetingInputs(
                        "Bonita E2E Test Meeting",
                        "2026-12-01T10:00:00.000Z",
                        "2026-12-01T11:00:00.000Z"
                )
        );

        var outputs = executeConnector(connector, inputs);

        assertThat(isSuccess(outputs)).isTrue();
        assertThat(outputs.get("meetingId")).isNotNull();
        assertThat(outputs.get("joinWebUrl")).isNotNull();
        createdMeetingId = (String) outputs.get("meetingId");
    }

    @Test
    @Order(13)
    @DisplayName("13 - should_get_meeting_when_meeting_exists")
    void should_get_meeting_when_meeting_exists() throws Exception {
        Assumptions.assumeTrue(createdMeetingId != null, "Requires meetingId from test 12");

        var connector = new MSTeamsGetMeetingConnector();
        var inputs = mergeInputs(
                delegatedAuth(),
                getMeetingInputs(createdMeetingId)
        );

        var outputs = executeConnector(connector, inputs);

        assertThat(isSuccess(outputs)).isTrue();
        assertThat(outputs.get("subject")).isEqualTo("Bonita E2E Test Meeting");
    }

    // ========================================================================
    // File operations
    // ========================================================================

    @Test
    @Order(14)
    @DisplayName("14 - should_upload_file_when_valid_content")
    void should_upload_file_when_valid_content() throws Exception {
        var connector = new MSTeamsUploadFileConnector();
        String fileContentBase64 = Base64.getEncoder()
                .encodeToString("Bonita E2E Test File Content".getBytes());
        var inputs = mergeInputs(
                appAuth(),
                uploadFileInputs(TEAM_ID, CHANNEL_ID,
                        "bonita-e2e-test-" + System.currentTimeMillis() + ".txt",
                        fileContentBase64)
        );

        var outputs = executeConnector(connector, inputs);

        assertThat(isSuccess(outputs)).isTrue();
        assertThat(outputs.get("fileId")).isNotNull();
        assertThat(outputs.get("webUrl")).isNotNull();
    }

    // ========================================================================
    // Validation tests (always runnable, no credentials needed)
    // ========================================================================

    @Test
    @Order(100)
    @DisplayName("should_verify_proc_file_exists_in_test_resources")
    void should_verify_proc_file_exists_in_test_resources() {
        var procStream = getClass().getClassLoader()
                .getResourceAsStream("MSTeamsConnectorTestProcess.bpmn");
        assertThat(procStream)
                .as("The .proc test process file should be on the classpath")
                .isNotNull();
    }
}
