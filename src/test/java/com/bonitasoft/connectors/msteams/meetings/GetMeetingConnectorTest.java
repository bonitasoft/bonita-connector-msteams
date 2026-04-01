package com.bonitasoft.connectors.msteams.meetings;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.*;

import com.bonitasoft.connectors.msteams.common.MsTeamsGraphClient;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.bonitasoft.engine.connector.ConnectorException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.HashMap;
import java.util.Map;

@ExtendWith(MockitoExtension.class)
class GetMeetingConnectorTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Mock
    private MsTeamsGraphClient graphClient;

    private GetMeetingConnector connector;

    @BeforeEach
    void setUp() {
        connector = new GetMeetingConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("refreshToken", "rt");
        inputs.put("userId", "user-1");
        inputs.put("meetingId", "meet-1");
        connector.setInputParameters(inputs);
    }

    @Test
    void should_get_meeting_successfully() throws Exception {
        connector.setGraphClient(graphClient);
        when(graphClient.get("/users/user-1/onlineMeetings/meet-1"))
                .thenReturn(MAPPER.readTree("{\"subject\":\"Sync\",\"joinWebUrl\":\"https://teams.microsoft.com/meet\"}"));

        connector.executeBusinessLogic();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("subject")).isEqualTo("Sync");
    }

    @Test
    void should_set_error_outputs_when_api_fails() {
        connector.setGraphClient(graphClient);
        when(graphClient.get(any())).thenThrow(
                new com.bonitasoft.connectors.msteams.common.MsTeamsException("Meeting not found"));
        org.junit.jupiter.api.Assertions.assertThrows(
                org.bonitasoft.engine.connector.ConnectorException.class,
                () -> connector.executeBusinessLogic());
        assertThat(connector.getOutputs().get("success")).isEqualTo(false);
    }

    @Test
    void should_fail_validation_when_userId_missing() {
        var c = new GetMeetingConnector();
        c.setInputParameters(new HashMap<>(Map.of(
                "tenantId", "t", "clientId", "c", "clientSecret", "s",
                "refreshToken", "rt", "meetingId", "meet-1")));
        org.junit.jupiter.api.Assertions.assertThrows(
                org.bonitasoft.engine.connector.ConnectorValidationException.class,
                c::validateInputParameters);
    }

    @Test
    void should_fail_validation_when_meetingId_missing() {
        var c = new GetMeetingConnector();
        c.setInputParameters(new HashMap<>(Map.of(
                "tenantId", "t", "clientId", "c", "clientSecret", "s",
                "refreshToken", "rt", "userId", "user-1")));
        org.junit.jupiter.api.Assertions.assertThrows(
                org.bonitasoft.engine.connector.ConnectorValidationException.class,
                c::validateInputParameters);
    }
}
