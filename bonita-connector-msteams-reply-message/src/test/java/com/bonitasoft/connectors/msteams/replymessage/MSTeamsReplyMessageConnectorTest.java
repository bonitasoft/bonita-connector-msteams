package com.bonitasoft.connectors.msteams.replymessage;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

import java.util.HashMap;
import java.util.Map;

import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.ObjectMapper;

import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

@DisplayName("MSTeamsReplyMessageConnector")
@ExtendWith(MockitoExtension.class)
class MSTeamsReplyMessageConnectorTest {

    private MSTeamsReplyMessageConnector connector;

    @Mock
    private MSTeamsClient mockClient;

    @BeforeEach
    void setUp() {
        connector = new MSTeamsReplyMessageConnector();
    }

    private Map<String, Object> validAuthParams() {
        Map<String, Object> params = new HashMap<>();
        params.put("tenantId", "test-tenant-id");
        params.put("clientId", "test-client-id");
        params.put("clientSecret", "test-client-secret");
        return params;
    }

    private Map<String, Object> validParams() {
        Map<String, Object> params = validAuthParams();
        params.put("teamId", "team-123");
        params.put("channelId", "channel-456");
        params.put("parentMessageId", "msg-parent-789");
        params.put("messageContent", "This is a reply!");
        return params;
    }

    @Nested
    @DisplayName("Validation")
    class Validation {

        @Test
        @DisplayName("should_fail_when_teamId_is_missing")
        void should_fail_when_teamId_is_missing() {
            Map<String, Object> params = validParams();
            params.remove("teamId");
            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("teamId");
        }

        @Test
        @DisplayName("should_fail_when_channelId_is_missing")
        void should_fail_when_channelId_is_missing() {
            Map<String, Object> params = validParams();
            params.remove("channelId");
            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("channelId");
        }

        @Test
        @DisplayName("should_fail_when_parentMessageId_is_missing")
        void should_fail_when_parentMessageId_is_missing() {
            Map<String, Object> params = validParams();
            params.remove("parentMessageId");
            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("parentMessageId");
        }

        @Test
        @DisplayName("should_fail_when_messageContent_is_missing")
        void should_fail_when_messageContent_is_missing() {
            Map<String, Object> params = validParams();
            params.remove("messageContent");
            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("messageContent");
        }

        @Test
        @DisplayName("should_pass_when_all_required_params_provided")
        void should_pass_when_all_required_params_provided() throws ConnectorValidationException {
            connector.setInputParameters(validParams());
            connector.validateInputParameters();
        }
    }

    @Nested
    @DisplayName("Execution")
    class Execution {

        @Test
        @DisplayName("should_call_correct_graph_api_endpoint")
        void should_call_correct_graph_api_endpoint() throws Exception {
            Map<String, Object> params = validParams();
            connector.setInputParameters(params);
            connector.setClient(mockClient);

            String responseJson = "{\"id\":\"reply-1\",\"createdDateTime\":\"2026-03-13T10:00:00Z\"}";
            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(
                    eq("/teams/team-123/channels/channel-456/messages/msg-parent-789/replies"),
                    anyString()))
                    .thenReturn(mapper.readTree(responseJson));

            connector.executeOperation(mockClient);

            verify(mockClient).post(
                    eq("/teams/team-123/channels/channel-456/messages/msg-parent-789/replies"),
                    anyString());

            Map<String, Object> outputs = connector.getOutputs();
            assertThat(outputs.get("replyId")).isEqualTo("reply-1");
            assertThat(outputs.get("createdDateTime")).isEqualTo("2026-03-13T10:00:00Z");
        }

        @Test
        @DisplayName("should_handle_api_error")
        void should_handle_api_error() {
            connector.setInputParameters(validParams());
            connector.setClient(mockClient);
            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(anyString(), anyString())).thenThrow(new MSTeamsException("error"));
            assertThatThrownBy(() -> connector.executeOperation(mockClient)).isInstanceOf(MSTeamsException.class);
        }
    }

    @Nested
    @DisplayName("Connector metadata")
    class Metadata {

        @Test
        @DisplayName("should_return_correct_connector_name")
        void should_return_correct_connector_name() {
            assertThat(connector.getConnectorName()).isEqualTo("MSTeams-ReplyMessage");
        }
    }
}
