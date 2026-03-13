package com.bonitasoft.connectors.msteams.meetings;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.*;

import com.bonitasoft.connectors.msteams.common.MsTeamsGraphClient;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.bonitasoft.engine.connector.ConnectorException;
import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.HashMap;
import java.util.Map;

@ExtendWith(MockitoExtension.class)
class CreateMeetingConnectorTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Mock
    private MsTeamsGraphClient graphClient;

    private CreateMeetingConnector connector;
    private Map<String, Object> inputs;

    @BeforeEach
    void setUp() {
        connector = new CreateMeetingConnector();
        inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("refreshToken", "rt");
        inputs.put("userId", "user-1");
        inputs.put("subject", "Team Sync");
        inputs.put("startDateTime", "2025-01-15T10:00:00Z");
        inputs.put("endDateTime", "2025-01-15T11:00:00Z");
        connector.setInputParameters(inputs);
    }

    @Test
    void should_pass_validation() throws ConnectorValidationException {
        connector.validateInputParameters();
    }

    @Test
    void should_fail_when_subject_missing() {
        var c = new CreateMeetingConnector();
        var incompleteInputs = new HashMap<>(inputs);
        incompleteInputs.remove("subject");
        c.setInputParameters(incompleteInputs);
        assertThatThrownBy(() -> c.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class);
    }

    @Test
    void should_create_meeting_successfully() throws Exception {
        connector.setGraphClient(graphClient);
        when(graphClient.post(eq("/users/user-1/onlineMeetings"), any()))
                .thenReturn(MAPPER.readTree("{\"id\":\"meet-1\",\"joinWebUrl\":\"https://teams.microsoft.com/meet/123\",\"subject\":\"Team Sync\"}"));

        connector.executeBusinessLogic();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("meetingId")).isEqualTo("meet-1");
        assertThat(connector.getOutputs().get("joinUrl")).isEqualTo("https://teams.microsoft.com/meet/123");
    }
}
