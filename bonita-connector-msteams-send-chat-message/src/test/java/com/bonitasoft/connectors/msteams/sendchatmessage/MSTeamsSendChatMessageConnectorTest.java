package com.bonitasoft.connectors.msteams.sendchatmessage;

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

@DisplayName("MSTeamsSendChatMessageConnector")
@ExtendWith(MockitoExtension.class)
class MSTeamsSendChatMessageConnectorTest {

    private MSTeamsSendChatMessageConnector connector;

    @Mock
    private MSTeamsClient mockClient;

    @BeforeEach
    void setUp() {
        connector = new MSTeamsSendChatMessageConnector();
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
        params.put("chatId", "chat-789");
        params.put("messageContent", "Hello Chat!");
        return params;
    }

    @Nested
    @DisplayName("Validation")
    class Validation {

        @Test
        @DisplayName("should_fail_when_chatId_is_missing")
        void should_fail_when_chatId_is_missing() {
            Map<String, Object> params = validParams();
            params.remove("chatId");
            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("chatId");
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

            String responseJson = "{\"id\":\"msg-1\",\"createdDateTime\":\"2026-03-13T10:00:00Z\"}";
            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(eq("/chats/chat-789/messages"), anyString()))
                    .thenReturn(mapper.readTree(responseJson));

            connector.executeOperation(mockClient);

            verify(mockClient).post(eq("/chats/chat-789/messages"), anyString());

            Map<String, Object> outputs = connector.getOutputs();
            assertThat(outputs.get("messageId")).isEqualTo("msg-1");
            assertThat(outputs.get("chatId")).isEqualTo("chat-789");
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

        @Test
        @DisplayName("should_wrap_generic_exception_in_msteams_exception")
        void should_wrap_generic_exception() {
            connector.setInputParameters(validParams());
            connector.setClient(mockClient);
            when(mockClient.getObjectMapper()).thenThrow(new RuntimeException("unexpected"));
            assertThatThrownBy(() -> connector.executeOperation(mockClient))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("Failed to send chat message");
        }

        @Test
        @DisplayName("should_execute_with_html_content_type_and_urgent_importance")
        void should_execute_with_custom_content_type_and_importance() throws Exception {
            Map<String, Object> params = validParams();
            params.put("contentType", "html");
            params.put("importance", "urgent");
            connector.setInputParameters(params);
            connector.setClient(mockClient);

            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(anyString(), anyString()))
                    .thenReturn(mapper.readTree("{\"id\":\"msg-2\",\"createdDateTime\":\"2026-03-13T11:00:00Z\"}"));

            connector.executeOperation(mockClient);
            verify(mockClient).post(anyString(), org.mockito.ArgumentMatchers.contains("\"contentType\":\"html\""));
            verify(mockClient).post(anyString(), org.mockito.ArgumentMatchers.contains("\"importance\":\"urgent\""));
        }
    }

    @Nested
    @DisplayName("Validation extras")
    class ValidationExtras {

        @Test
        @DisplayName("should_fail_when_chatId_is_blank")
        void should_fail_when_chatId_is_blank() {
            Map<String, Object> params = validParams();
            params.put("chatId", "  ");
            connector.setInputParameters(params);
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("chatId");
        }

        @Test
        @DisplayName("should_fail_when_messageContent_is_blank")
        void should_fail_when_messageContent_is_blank() {
            Map<String, Object> params = validParams();
            params.put("messageContent", "  ");
            connector.setInputParameters(params);
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("messageContent");
        }

        @Test
        @DisplayName("should_fail_when_messageContent_exceeds_max_length")
        void should_fail_when_messageContent_exceeds_max_length() {
            Map<String, Object> params = validParams();
            params.put("messageContent", "x".repeat(28_001));
            connector.setInputParameters(params);
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("exceeds maximum length");
        }
    }

    @Nested
    @DisplayName("Connector metadata")
    class Metadata {

        @Test
        @DisplayName("should_return_correct_connector_name")
        void should_return_correct_connector_name() {
            assertThat(connector.getConnectorName()).isEqualTo("MSTeams-SendChatMessage");
        }
    }
}
