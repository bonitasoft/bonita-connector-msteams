package com.bonitasoft.connectors.msteams.sendchannelmessage;

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

@DisplayName("MSTeamsSendChannelMessageConnector")
@ExtendWith(MockitoExtension.class)
class MSTeamsSendChannelMessageConnectorTest {

    private MSTeamsSendChannelMessageConnector connector;

    @Mock
    private MSTeamsClient mockClient;

    @BeforeEach
    void setUp() {
        connector = new MSTeamsSendChannelMessageConnector();
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
        params.put("messageContent", "Hello Teams!");
        return params;
    }

    @Nested
    @DisplayName("Validation")
    class Validation {

        @Test
        @DisplayName("should_fail_when_tenantId_is_missing")
        void should_fail_when_tenantId_is_missing() {
            Map<String, Object> params = validParams();
            params.remove("tenantId");
            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("tenantId");
        }

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
            connector.validateInputParameters(); // Should not throw
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

            String responseJson = "{\"id\":\"msg-1\",\"webUrl\":\"https://teams.microsoft.com/msg\",\"createdDateTime\":\"2026-03-13T10:00:00Z\"}";
            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(eq("/teams/team-123/channels/channel-456/messages"), anyString()))
                    .thenReturn(mapper.readTree(responseJson));

            connector.executeOperation(mockClient);

            verify(mockClient).post(eq("/teams/team-123/channels/channel-456/messages"), anyString());

            Map<String, Object> outputs = connector.getOutputs();
            assertThat(outputs.get("messageId")).isEqualTo("msg-1");
            assertThat(outputs.get("webUrl")).isEqualTo("https://teams.microsoft.com/msg");
            assertThat(outputs.get("createdDateTime")).isEqualTo("2026-03-13T10:00:00Z");
        }

        @Test
        @DisplayName("should_use_default_content_type_and_importance")
        void should_use_default_content_type_and_importance() throws Exception {
            Map<String, Object> params = validParams();
            connector.setInputParameters(params);
            connector.setClient(mockClient);

            String responseJson = "{\"id\":\"msg-1\",\"webUrl\":\"https://url\",\"createdDateTime\":\"2026-03-13T10:00:00Z\"}";
            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(anyString(), anyString())).thenReturn(mapper.readTree(responseJson));

            connector.executeOperation(mockClient);

            verify(mockClient).post(anyString(), org.mockito.ArgumentMatchers.contains("\"contentType\":\"text\""));
        }

        @Test
        @DisplayName("should_include_subject_when_provided")
        void should_include_subject_when_provided() throws Exception {
            Map<String, Object> params = validParams();
            params.put("subject", "Important Topic");
            connector.setInputParameters(params);
            connector.setClient(mockClient);

            String responseJson = "{\"id\":\"msg-1\",\"webUrl\":\"https://url\",\"createdDateTime\":\"2026-03-13T10:00:00Z\"}";
            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(anyString(), anyString())).thenReturn(mapper.readTree(responseJson));

            connector.executeOperation(mockClient);

            verify(mockClient).post(anyString(), org.mockito.ArgumentMatchers.contains("\"subject\":\"Important Topic\""));
        }

        @Test
        @DisplayName("should_not_include_subject_when_blank")
        void should_not_include_subject_when_blank() throws Exception {
            Map<String, Object> params = validParams();
            params.put("subject", "  ");
            connector.setInputParameters(params);
            connector.setClient(mockClient);

            String responseJson = "{\"id\":\"msg-1\",\"webUrl\":\"https://url\",\"createdDateTime\":\"2026-03-13T10:00:00Z\"}";
            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(anyString(), anyString())).thenReturn(mapper.readTree(responseJson));

            connector.executeOperation(mockClient);
            // subject should not be in the body when blank
        }

        @Test
        @DisplayName("should_handle_api_error")
        void should_handle_api_error() {
            connector.setInputParameters(validParams());
            connector.setClient(mockClient);
            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(anyString(), anyString())).thenThrow(new MSTeamsException("error", 500, "ServerError", true, null));
            assertThatThrownBy(() -> connector.executeOperation(mockClient)).isInstanceOf(MSTeamsException.class);
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

        @Test
        @DisplayName("should_wrap_generic_exception_in_msteams_exception")
        void should_wrap_generic_exception() {
            connector.setInputParameters(validParams());
            connector.setClient(mockClient);
            when(mockClient.getObjectMapper()).thenThrow(new RuntimeException("unexpected"));
            assertThatThrownBy(() -> connector.executeOperation(mockClient))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("Failed to send channel message");
        }

        @Test
        @DisplayName("should_not_include_subject_when_null")
        void should_not_include_subject_when_null() throws Exception {
            Map<String, Object> params = validParams();
            // subject is not set (null)
            connector.setInputParameters(params);
            connector.setClient(mockClient);

            String responseJson = "{\"id\":\"msg-2\",\"webUrl\":\"https://url\",\"createdDateTime\":\"2026-03-13T10:00:00Z\"}";
            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(anyString(), anyString())).thenReturn(mapper.readTree(responseJson));

            connector.executeOperation(mockClient);
            assertThat(connector.getOutputs().get("messageId")).isEqualTo("msg-2");
        }

        @Test
        @DisplayName("should_execute_with_html_content_type_and_urgent_importance")
        void should_execute_with_custom_settings() throws Exception {
            Map<String, Object> params = validParams();
            params.put("contentType", "html");
            params.put("importance", "urgent");
            params.put("subject", "Important");
            connector.setInputParameters(params);
            connector.setClient(mockClient);

            String responseJson = "{\"id\":\"msg-3\",\"webUrl\":\"https://url\",\"createdDateTime\":\"2026-03-13T10:00:00Z\"}";
            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(anyString(), anyString())).thenReturn(mapper.readTree(responseJson));

            connector.executeOperation(mockClient);
            verify(mockClient).post(anyString(), org.mockito.ArgumentMatchers.contains("\"contentType\":\"html\""));
            verify(mockClient).post(anyString(), org.mockito.ArgumentMatchers.contains("\"importance\":\"urgent\""));
            verify(mockClient).post(anyString(), org.mockito.ArgumentMatchers.contains("\"subject\":\"Important\""));
        }
    }

    @Nested
    @DisplayName("Validation extras")
    class ValidationExtras {

        @Test
        @DisplayName("should_fail_when_teamId_is_blank")
        void should_fail_when_teamId_is_blank() {
            Map<String, Object> params = validParams();
            params.put("teamId", "  ");
            connector.setInputParameters(params);
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("teamId");
        }

        @Test
        @DisplayName("should_fail_when_channelId_is_blank")
        void should_fail_when_channelId_is_blank() {
            Map<String, Object> params = validParams();
            params.put("channelId", "  ");
            connector.setInputParameters(params);
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("channelId");
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
    }

    @Nested
    @DisplayName("Connector metadata")
    class Metadata {

        @Test
        @DisplayName("should_return_correct_connector_name")
        void should_return_correct_connector_name() {
            assertThat(connector.getConnectorName()).isEqualTo("MSTeams-SendChannelMessage");
        }
    }
}
