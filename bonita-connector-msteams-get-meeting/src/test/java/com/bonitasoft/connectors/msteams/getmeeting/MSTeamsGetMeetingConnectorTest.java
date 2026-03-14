package com.bonitasoft.connectors.msteams.getmeeting;

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

@DisplayName("MSTeamsGetMeetingConnector")
@ExtendWith(MockitoExtension.class)
class MSTeamsGetMeetingConnectorTest {

    private MSTeamsGetMeetingConnector connector;

    @Mock
    private MSTeamsClient client;

    private final ObjectMapper objectMapper = new ObjectMapper();

    @BeforeEach
    void setUp() {
        connector = new MSTeamsGetMeetingConnector();
    }

    @Nested
    @DisplayName("Validation")
    class Validation {

        @Test
        @DisplayName("should fail when meetingId is missing")
        void should_fail_when_meetingId_missing() {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("userId", "user-123");

            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("meetingId");
        }

        @Test
        @DisplayName("should fail when tenantId is missing")
        void should_fail_when_tenantId_missing() {
            Map<String, Object> params = new HashMap<>();
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("userId", "user-123");
            params.put("meetingId", "meeting-123");

            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("tenantId");
        }

        @Test
        @DisplayName("should pass with all required parameters")
        void should_pass_when_all_required_params_present() throws ConnectorValidationException {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("userId", "user-123");
            params.put("meetingId", "meeting-123");

            connector.setInputParameters(params);

            connector.validateInputParameters(); // Should not throw
        }
    }

    @Nested
    @DisplayName("Execution")
    class Execution {

        @Test
        @DisplayName("should get meeting and set outputs")
        void should_get_meeting_when_valid_inputs() throws Exception {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("userId", "user-123");
            params.put("meetingId", "meeting-abc-123");

            connector.setInputParameters(params);

            String responseJson = """
                    {
                        "id": "meeting-abc-123",
                        "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/abc",
                        "subject": "Team Standup",
                        "startDateTime": "2024-01-15T10:00:00Z",
                        "endDateTime": "2024-01-15T11:00:00Z",
                        "participants": {
                            "organizer": {
                                "identity": {
                                    "user": {
                                        "id": "user-123",
                                        "displayName": "John Doe"
                                    }
                                }
                            }
                        }
                    }
                    """;
            JsonNode responseNode = objectMapper.readTree(responseJson);

            when(client.getObjectMapper()).thenReturn(objectMapper);
            when(client.get(anyString())).thenReturn(responseNode);

            connector.executeOperation(client);

            assertThat(connector.getOutputs().get("meetingId")).isEqualTo("meeting-abc-123");
            assertThat(connector.getOutputs().get("joinWebUrl")).isEqualTo("https://teams.microsoft.com/l/meetup-join/abc");
            assertThat(connector.getOutputs().get("subject")).isEqualTo("Team Standup");
            assertThat(connector.getOutputs().get("startDateTime")).isEqualTo("2024-01-15T10:00:00Z");
            assertThat(connector.getOutputs().get("endDateTime")).isEqualTo("2024-01-15T11:00:00Z");
            assertThat(connector.getOutputs().get("participants")).isNotNull();
            @SuppressWarnings("unchecked")
            Map<String, Object> participants = (Map<String, Object>) connector.getOutputs().get("participants");
            assertThat(participants).containsKey("organizer");
        }

        @Test
        @DisplayName("should handle API error")
        void should_handle_api_error() {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("userId", "user-123");
            params.put("meetingId", "meeting-abc-123");
            connector.setInputParameters(params);

            when(client.get(anyString())).thenThrow(new MSTeamsException("error", 404, "NotFound", false, null));

            assertThatThrownBy(() -> connector.executeOperation(client)).isInstanceOf(MSTeamsException.class);
        }

        @Test
        @DisplayName("should return correct connector name")
        void should_return_correct_connector_name() {
            assertThat(connector.getConnectorName()).isEqualTo("MSTeams-GetMeeting");
        }

        @Test
        @DisplayName("should fail when userId is missing")
        void should_fail_when_userId_missing() {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("meetingId", "meeting-123");
            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("userId");
        }

        @Test
        @DisplayName("should handle missing participants node")
        void should_handle_missing_participants() throws Exception {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("userId", "user-123");
            params.put("meetingId", "meeting-abc-123");
            connector.setInputParameters(params);

            JsonNode responseNode = objectMapper.readTree("{\"id\":\"m-1\",\"joinWebUrl\":\"url\",\"subject\":\"S\"}");
            when(client.get(anyString())).thenReturn(responseNode);

            connector.executeOperation(client);
            @SuppressWarnings("unchecked")
            Map<String, Object> participants = (Map<String, Object>) connector.getOutputs().get("participants");
            assertThat(participants).isEmpty();
        }
    }
}
