package com.bonitasoft.connectors.msteams.createmeeting;

import com.bonitasoft.connectors.msteams.MSTeamsException;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.when;

import java.util.HashMap;
import java.util.Map;

import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

@DisplayName("MSTeamsCreateMeetingConnector")
@ExtendWith(MockitoExtension.class)
class MSTeamsCreateMeetingConnectorTest {

    private MSTeamsCreateMeetingConnector connector;

    @Mock
    private MSTeamsClient client;

    private final ObjectMapper objectMapper = new ObjectMapper();

    @BeforeEach
    void setUp() {
        connector = new MSTeamsCreateMeetingConnector();
    }

    @Nested
    @DisplayName("Validation")
    class Validation {

        @Test
        @DisplayName("should fail when tenantId is missing")
        void should_fail_when_tenantId_missing() {
            Map<String, Object> params = new HashMap<>();
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("userId", "user-123");
            params.put("subject", "Team Standup");
            params.put("startDateTime", "2024-01-15T10:00:00Z");
            params.put("endDateTime", "2024-01-15T11:00:00Z");

            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("tenantId");
        }

        @Test
        @DisplayName("should fail when subject is missing")
        void should_fail_when_subject_missing() {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("userId", "user-123");
            params.put("startDateTime", "2024-01-15T10:00:00Z");
            params.put("endDateTime", "2024-01-15T11:00:00Z");

            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("subject");
        }

        @Test
        @DisplayName("should fail when startDateTime is missing")
        void should_fail_when_startDateTime_missing() {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("userId", "user-123");
            params.put("subject", "Team Standup");
            params.put("endDateTime", "2024-01-15T11:00:00Z");

            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("startDateTime");
        }

        @Test
        @DisplayName("should fail when endDateTime is missing")
        void should_fail_when_endDateTime_missing() {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("userId", "user-123");
            params.put("subject", "Team Standup");
            params.put("startDateTime", "2024-01-15T10:00:00Z");

            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("endDateTime");
        }

        @Test
        @DisplayName("should pass with all required parameters")
        void should_pass_when_all_required_params_present() throws ConnectorValidationException {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("userId", "user-123");
            params.put("subject", "Team Standup");
            params.put("startDateTime", "2024-01-15T10:00:00Z");
            params.put("endDateTime", "2024-01-15T11:00:00Z");

            connector.setInputParameters(params);

            connector.validateInputParameters(); // Should not throw
        }
    }

    @Nested
    @DisplayName("Execution")
    class Execution {

        @Test
        @DisplayName("should create meeting and set outputs")
        void should_create_meeting_when_valid_inputs() throws Exception {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("userId", "user-123");
            params.put("subject", "Team Standup");
            params.put("startDateTime", "2024-01-15T10:00:00Z");
            params.put("endDateTime", "2024-01-15T11:00:00Z");
            params.put("allowedPresenters", "everyone");
            params.put("lobbyBypassScope", "organization");
            params.put("recordAutomatically", false);

            connector.setInputParameters(params);

            String responseJson = """
                    {
                        "id": "meeting-abc-123",
                        "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/abc",
                        "joinMeetingIdSettings": {
                            "joinMeetingId": "123456789"
                        },
                        "subject": "Team Standup",
                        "startDateTime": "2024-01-15T10:00:00Z",
                        "endDateTime": "2024-01-15T11:00:00Z"
                    }
                    """;
            JsonNode responseNode = objectMapper.readTree(responseJson);

            when(client.getObjectMapper()).thenReturn(objectMapper);
            when(client.post(anyString(), anyString())).thenReturn(responseNode);

            connector.executeOperation(client);

            assertThat(connector.getOutputs().get("meetingId")).isEqualTo("meeting-abc-123");
            assertThat(connector.getOutputs().get("joinWebUrl")).isEqualTo("https://teams.microsoft.com/l/meetup-join/abc");
            assertThat(connector.getOutputs().get("meetingCode")).isEqualTo("123456789");
            assertThat(connector.getOutputs().get("subject")).isEqualTo("Team Standup");
            assertThat(connector.getOutputs().get("startDateTime")).isEqualTo("2024-01-15T10:00:00Z");
            assertThat(connector.getOutputs().get("endDateTime")).isEqualTo("2024-01-15T11:00:00Z");
        }

        @Test
        @DisplayName("should handle API error")
        void should_handle_api_error() {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("userId", "user-123");
            params.put("subject", "Team Standup");
            params.put("startDateTime", "2024-01-15T10:00:00Z");
            params.put("endDateTime", "2024-01-15T11:00:00Z");
            connector.setInputParameters(params);

            when(client.getObjectMapper()).thenReturn(objectMapper);
            when(client.post(anyString(), anyString())).thenThrow(new MSTeamsException("error", 403, "Forbidden", false, null));

            assertThatThrownBy(() -> connector.executeOperation(client)).isInstanceOf(MSTeamsException.class);
        }

        @Test
        @DisplayName("should return correct connector name")
        void should_return_correct_connector_name() {
            assertThat(connector.getConnectorName()).isEqualTo("MSTeams-CreateMeeting");
        }
    }
}
