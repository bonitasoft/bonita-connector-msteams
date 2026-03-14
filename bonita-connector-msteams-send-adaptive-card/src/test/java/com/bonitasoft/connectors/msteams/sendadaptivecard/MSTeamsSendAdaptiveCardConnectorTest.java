package com.bonitasoft.connectors.msteams.sendadaptivecard;

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

@DisplayName("MSTeamsSendAdaptiveCardConnector")
@ExtendWith(MockitoExtension.class)
class MSTeamsSendAdaptiveCardConnectorTest {

    private MSTeamsSendAdaptiveCardConnector connector;

    @Mock
    private MSTeamsClient mockClient;

    @BeforeEach
    void setUp() {
        connector = new MSTeamsSendAdaptiveCardConnector();
    }

    private static final String VALID_CARD_JSON = """
            {
                "type": "AdaptiveCard",
                "version": "1.4",
                "body": [
                    {
                        "type": "TextBlock",
                        "text": "Hello from Bonita!"
                    }
                ]
            }
            """;

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
        params.put("cardJson", VALID_CARD_JSON);
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
        @DisplayName("should_fail_when_cardJson_is_missing")
        void should_fail_when_cardJson_is_missing() {
            Map<String, Object> params = validParams();
            params.remove("cardJson");
            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("cardJson");
        }

        @Test
        @DisplayName("should_fail_when_cardJson_is_invalid_json")
        void should_fail_when_cardJson_is_invalid_json() {
            Map<String, Object> params = validParams();
            params.put("cardJson", "not valid json {{{");
            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("cardJson must be valid JSON");
        }

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
        @DisplayName("should_fail_when_cardJson_is_blank")
        void should_fail_when_cardJson_is_blank() {
            Map<String, Object> params = validParams();
            params.put("cardJson", "  ");
            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("cardJson");
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
        @DisplayName("should_call_correct_graph_api_endpoint_with_attachment")
        void should_call_correct_graph_api_endpoint_with_attachment() throws Exception {
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
        @DisplayName("should_include_adaptive_card_content_type_in_payload")
        void should_include_adaptive_card_content_type_in_payload() throws Exception {
            Map<String, Object> params = validParams();
            connector.setInputParameters(params);
            connector.setClient(mockClient);

            String responseJson = "{\"id\":\"msg-1\",\"webUrl\":\"https://url\",\"createdDateTime\":\"2026-03-13T10:00:00Z\"}";
            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(anyString(), anyString())).thenReturn(mapper.readTree(responseJson));

            connector.executeOperation(mockClient);

            verify(mockClient).post(anyString(),
                    org.mockito.ArgumentMatchers.contains("application/vnd.microsoft.card.adaptive"));
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
            // subject should not be present when blank
        }

        @Test
        @DisplayName("should_not_include_subject_when_null")
        void should_not_include_subject_when_null() throws Exception {
            Map<String, Object> params = validParams();
            // subject not set at all (null)
            connector.setInputParameters(params);
            connector.setClient(mockClient);

            String responseJson = "{\"id\":\"msg-1\",\"webUrl\":\"https://url\",\"createdDateTime\":\"2026-03-13T10:00:00Z\"}";
            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(anyString(), anyString())).thenReturn(mapper.readTree(responseJson));

            connector.executeOperation(mockClient);
        }

        @Test
        @DisplayName("should_use_custom_importance_when_provided")
        void should_use_custom_importance_when_provided() throws Exception {
            Map<String, Object> params = validParams();
            params.put("importance", "urgent");
            connector.setInputParameters(params);
            connector.setClient(mockClient);

            String responseJson = "{\"id\":\"msg-1\",\"webUrl\":\"https://url\",\"createdDateTime\":\"2026-03-13T10:00:00Z\"}";
            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(anyString(), anyString())).thenReturn(mapper.readTree(responseJson));

            connector.executeOperation(mockClient);

            verify(mockClient).post(anyString(), org.mockito.ArgumentMatchers.contains("\"importance\":\"urgent\""));
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
        void should_wrap_generic_exception_in_msteams_exception() {
            connector.setInputParameters(validParams());
            connector.setClient(mockClient);
            ObjectMapper mapper = new ObjectMapper();
            when(mockClient.getObjectMapper()).thenReturn(mapper);
            when(mockClient.post(anyString(), anyString())).thenThrow(new RuntimeException("unexpected error"));
            assertThatThrownBy(() -> connector.executeOperation(mockClient))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("Failed to send Adaptive Card");
        }
    }

    @Nested
    @DisplayName("Connector metadata")
    class Metadata {

        @Test
        @DisplayName("should_return_correct_connector_name")
        void should_return_correct_connector_name() {
            assertThat(connector.getConnectorName()).isEqualTo("MSTeams-SendAdaptiveCard");
        }
    }
}
