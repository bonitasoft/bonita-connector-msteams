package com.bonitasoft.connectors.msteams.integration;

import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsConfiguration;
import com.bonitasoft.connectors.msteams.MSTeamsException;
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

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import org.bonitasoft.engine.connector.ConnectorException;
import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.lang.reflect.InvocationTargetException;
import java.util.HashMap;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.lenient;
import static org.mockito.Mockito.when;

/**
 * Nivel 2 — Mock Integration Tests for MS Teams connectors.
 * <p>
 * Mocks the MSTeamsClient (MS Graph API) with Mockito to test the full
 * connector lifecycle (VALIDATE -> EXECUTE -> outputs) without real credentials.
 */
@Tag("integration")
@DisplayName("Nivel 2 — MS Teams Connector Mock Integration Tests")
@ExtendWith(MockitoExtension.class)
class ConnectorMockIT {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Mock
    private MSTeamsClient mockClient;

    private Map<String, Object> authParams() {
        Map<String, Object> params = new HashMap<>();
        params.put("tenantId", "test-tenant-id");
        params.put("clientId", "test-client-id");
        params.put("clientSecret", "test-client-secret");
        return params;
    }

    /**
     * Execute the connector's protected executeBusinessLogic() via reflection.
     */
    private void executeLogic(AbstractMSTeamsConnector connector) throws Exception {
        try {
            var method = AbstractMSTeamsConnector.class.getDeclaredMethod("executeBusinessLogic");
            method.setAccessible(true);
            method.invoke(connector);
        } catch (InvocationTargetException e) {
            if (e.getCause() instanceof ConnectorException ce) throw ce;
            if (e.getCause() instanceof RuntimeException re) throw re;
            throw new RuntimeException(e.getCause());
        } catch (ReflectiveOperationException e) {
            throw new RuntimeException("Failed to invoke executeBusinessLogic", e);
        }
    }

    @BeforeEach
    void setUp() {
        // Lenient stubs — not every test uses the mock client directly
        MSTeamsConfiguration config = new MSTeamsConfiguration(
                "test-tenant-id", "test-client-id", "test-client-secret",
                30_000, 60_000, false);
        lenient().when(mockClient.getConfiguration()).thenReturn(config);
        lenient().when(mockClient.getObjectMapper()).thenReturn(MAPPER);
    }

    // ---- Test 1: Instantiation of all 14 connectors ----

    @Test
    @DisplayName("should_instantiate_all_14_connectors")
    void should_instantiate_all_14_connectors() {
        AbstractMSTeamsConnector[] connectors = {
                new MSTeamsSendChannelMessageConnector(),
                new MSTeamsSendChatMessageConnector(),
                new MSTeamsReplyMessageConnector(),
                new MSTeamsSendAdaptiveCardConnector(),
                new MSTeamsCreateChannelConnector(),
                new MSTeamsListChannelsConnector(),
                new MSTeamsDeleteChannelConnector(),
                new MSTeamsListTeamsConnector(),
                new MSTeamsGetTeamConnector(),
                new MSTeamsAddMemberConnector(),
                new MSTeamsRemoveMemberConnector(),
                new MSTeamsCreateMeetingConnector(),
                new MSTeamsGetMeetingConnector(),
                new MSTeamsUploadFileConnector()
        };

        assertThat(connectors).hasSize(14);
        for (var connector : connectors) {
            assertThat(connector).isNotNull();
        }
    }

    // ---- Test 2: Validate send channel message with valid inputs ----

    @Test
    @DisplayName("should_validate_send_channel_message_with_valid_inputs")
    void should_validate_send_channel_message_with_valid_inputs() throws ConnectorValidationException {
        var connector = new MSTeamsSendChannelMessageConnector();
        Map<String, Object> params = authParams();
        params.put("teamId", "team-123");
        params.put("channelId", "channel-456");
        params.put("messageContent", "Hello from Bonita!");

        connector.setInputParameters(params);
        connector.validateInputParameters(); // Should not throw
    }

    // ---- Test 3: Execute send channel message with mocked client ----

    @Test
    @DisplayName("should_execute_send_channel_message_with_mocked_client")
    void should_execute_send_channel_message_with_mocked_client() throws Exception {
        var connector = new MSTeamsSendChannelMessageConnector();
        Map<String, Object> params = authParams();
        params.put("teamId", "team-123");
        params.put("channelId", "channel-456");
        params.put("messageContent", "Hello from Bonita!");
        connector.setInputParameters(params);

        // Mock Graph API response
        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "msg-789");
        response.put("webUrl", "https://teams.microsoft.com/msg/789");
        response.put("createdDateTime", "2026-03-17T10:00:00Z");

        when(mockClient.post(eq("/teams/team-123/channels/channel-456/messages"), anyString()))
                .thenReturn(response);

        connector.setClient(mockClient);
        executeLogic(connector);

        Map<String, Object> outputs = connector.getOutputs();
        assertThat(outputs.get("success")).isEqualTo(true);
        assertThat(outputs.get("errorMessage")).isEqualTo("");
        assertThat(outputs.get("messageId")).isEqualTo("msg-789");
        assertThat(outputs.get("webUrl")).isEqualTo("https://teams.microsoft.com/msg/789");
        assertThat(outputs.get("createdDateTime")).isEqualTo("2026-03-17T10:00:00Z");
    }

    // ---- Test 4: Handle API error 401 Unauthorized ----

    @Test
    @DisplayName("should_handle_api_error_401_unauthorized")
    void should_handle_api_error_401_unauthorized() {
        var connector = new MSTeamsSendChannelMessageConnector();
        Map<String, Object> params = authParams();
        params.put("teamId", "team-123");
        params.put("channelId", "channel-456");
        params.put("messageContent", "Hello");
        connector.setInputParameters(params);

        when(mockClient.post(anyString(), anyString()))
                .thenThrow(MSTeamsException.fromGraphError(401, "InvalidAuthenticationToken",
                        "Access token is invalid"));

        connector.setClient(mockClient);
        assertThatThrownBy(() -> executeLogic(connector))
                .isInstanceOf(ConnectorException.class);

        Map<String, Object> outputs = connector.getOutputs();
        assertThat(outputs.get("success")).isEqualTo(false);
        assertThat((String) outputs.get("errorMessage")).contains("Authentication failed");
    }

    // ---- Test 5: Handle API error 429 Rate Limit ----

    @Test
    @DisplayName("should_handle_api_error_429_rate_limit")
    void should_handle_api_error_429_rate_limit() {
        var connector = new MSTeamsSendChannelMessageConnector();
        Map<String, Object> params = authParams();
        params.put("teamId", "team-123");
        params.put("channelId", "channel-456");
        params.put("messageContent", "Hello");
        connector.setInputParameters(params);

        // 429 is retryable, so the client will eventually throw after max retries
        MSTeamsException rateLimitEx = new MSTeamsException(
                "Max retries (5) exceeded for POST: Too many requests",
                429, "activityLimitReached", false, null);

        when(mockClient.post(anyString(), anyString())).thenThrow(rateLimitEx);

        connector.setClient(mockClient);
        assertThatThrownBy(() -> executeLogic(connector))
                .isInstanceOf(ConnectorException.class);

        Map<String, Object> outputs = connector.getOutputs();
        assertThat(outputs.get("success")).isEqualTo(false);
        assertThat((String) outputs.get("errorMessage")).containsAnyOf("429", "Too many", "retries");
    }

    // ---- Test 6: Handle API error 500 Server Error ----

    @Test
    @DisplayName("should_handle_api_error_500_server_error")
    void should_handle_api_error_500_server_error() {
        var connector = new MSTeamsSendChannelMessageConnector();
        Map<String, Object> params = authParams();
        params.put("teamId", "team-123");
        params.put("channelId", "channel-456");
        params.put("messageContent", "Hello");
        connector.setInputParameters(params);

        MSTeamsException serverError = new MSTeamsException(
                "Internal server error", 500, "serverError", false, null);

        when(mockClient.post(anyString(), anyString())).thenThrow(serverError);

        connector.setClient(mockClient);
        assertThatThrownBy(() -> executeLogic(connector))
                .isInstanceOf(ConnectorException.class);

        Map<String, Object> outputs = connector.getOutputs();
        assertThat(outputs.get("success")).isEqualTo(false);
        assertThat((String) outputs.get("errorMessage")).contains("Internal server error");
    }

    // ---- Test 7: Execute list teams with mocked client ----

    @Test
    @DisplayName("should_execute_list_teams_with_mocked_client")
    void should_execute_list_teams_with_mocked_client() throws Exception {
        var connector = new MSTeamsListTeamsConnector();
        Map<String, Object> params = authParams();
        connector.setInputParameters(params);

        // Build mock response with 2 teams
        ObjectNode response = MAPPER.createObjectNode();
        ArrayNode value = MAPPER.createArrayNode();

        ObjectNode team1 = MAPPER.createObjectNode();
        team1.put("id", "team-1");
        team1.put("displayName", "Engineering");
        team1.put("description", "Dev team");
        team1.put("visibility", "private");
        value.add(team1);

        ObjectNode team2 = MAPPER.createObjectNode();
        team2.put("id", "team-2");
        team2.put("displayName", "Sales");
        team2.putNull("description");
        team2.put("visibility", "public");
        value.add(team2);

        response.set("value", value);

        when(mockClient.get(anyString())).thenReturn(response);

        connector.setClient(mockClient);
        executeLogic(connector);

        Map<String, Object> outputs = connector.getOutputs();
        assertThat(outputs.get("success")).isEqualTo(true);
        assertThat(outputs.get("errorMessage")).isEqualTo("");
        assertThat(outputs.get("totalCount")).isEqualTo(2);
        assertThat(outputs.get("teams")).isNotNull();
    }

    // ---- Test 8: Execute create channel with mocked client ----

    @Test
    @DisplayName("should_execute_create_channel_with_mocked_client")
    void should_execute_create_channel_with_mocked_client() throws Exception {
        var connector = new MSTeamsCreateChannelConnector();
        Map<String, Object> params = authParams();
        params.put("teamId", "team-123");
        params.put("displayName", "New Channel");
        params.put("description", "Channel for testing");
        connector.setInputParameters(params);

        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "ch-001");
        response.put("displayName", "New Channel");
        response.put("webUrl", "https://teams.microsoft.com/channel/ch-001");

        when(mockClient.post(eq("/teams/team-123/channels"), anyString()))
                .thenReturn(response);

        connector.setClient(mockClient);
        executeLogic(connector);

        Map<String, Object> outputs = connector.getOutputs();
        assertThat(outputs.get("success")).isEqualTo(true);
        assertThat(outputs.get("channelId")).isEqualTo("ch-001");
        assertThat(outputs.get("displayName")).isEqualTo("New Channel");
        assertThat(outputs.get("webUrl")).isEqualTo("https://teams.microsoft.com/channel/ch-001");
    }

    // ---- Test 9: Set output parameters correctly ----

    @Test
    @DisplayName("should_set_output_parameters_correctly")
    void should_set_output_parameters_correctly() throws Exception {
        var connector = new MSTeamsSendChannelMessageConnector();
        Map<String, Object> params = authParams();
        params.put("teamId", "t1");
        params.put("channelId", "c1");
        params.put("messageContent", "Test message");
        params.put("contentType", "html");
        params.put("subject", "Test Subject");
        params.put("importance", "high");
        connector.setInputParameters(params);

        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "msg-output-test");
        response.put("webUrl", "https://teams.microsoft.com/msg/test");
        response.put("createdDateTime", "2026-03-17T12:00:00Z");

        when(mockClient.post(anyString(), anyString())).thenReturn(response);

        connector.setClient(mockClient);
        executeLogic(connector);

        Map<String, Object> outputs = connector.getOutputs();
        // Standard outputs
        assertThat(outputs.get("success")).isEqualTo(true);
        assertThat(outputs.get("errorMessage")).isEqualTo("");
        // Operation-specific outputs
        assertThat(outputs.get("messageId")).isEqualTo("msg-output-test");
        assertThat(outputs.get("webUrl")).isNotNull();
        assertThat(outputs.get("createdDateTime")).isNotNull();
    }

    // ---- Test 10: Handle empty/null response ----

    @Test
    @DisplayName("should_handle_empty_response")
    void should_handle_empty_response() throws Exception {
        var connector = new MSTeamsListTeamsConnector();
        Map<String, Object> params = authParams();
        connector.setInputParameters(params);

        // Response with empty value array
        ObjectNode response = MAPPER.createObjectNode();
        response.set("value", MAPPER.createArrayNode());

        when(mockClient.get(anyString())).thenReturn(response);

        connector.setClient(mockClient);
        executeLogic(connector);

        Map<String, Object> outputs = connector.getOutputs();
        assertThat(outputs.get("success")).isEqualTo(true);
        assertThat(outputs.get("totalCount")).isEqualTo(0);
    }
}
