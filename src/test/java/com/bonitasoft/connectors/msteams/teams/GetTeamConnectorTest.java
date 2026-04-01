package com.bonitasoft.connectors.msteams.teams;

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
class GetTeamConnectorTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Mock
    private MsTeamsGraphClient graphClient;

    private GetTeamConnector connector;

    @BeforeEach
    void setUp() {
        connector = new GetTeamConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("teamId", "team-1");
        connector.setInputParameters(inputs);
    }

    @Test
    void should_get_team_successfully() throws Exception {
        connector.setGraphClient(graphClient);
        when(graphClient.get("/teams/team-1"))
                .thenReturn(MAPPER.readTree("{\"displayName\":\"Engineering\",\"description\":\"Dev team\"}"));

        connector.executeBusinessLogic();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("displayName")).isEqualTo("Engineering");
    }

    @Test
    void should_set_error_outputs_when_api_fails() {
        connector.setGraphClient(graphClient);
        when(graphClient.get(any())).thenThrow(
                new com.bonitasoft.connectors.msteams.common.MsTeamsException("Team not found"));
        org.junit.jupiter.api.Assertions.assertThrows(
                org.bonitasoft.engine.connector.ConnectorException.class,
                () -> connector.executeBusinessLogic());
        assertThat(connector.getOutputs().get("success")).isEqualTo(false);
    }

    @Test
    void should_fail_validation_when_teamId_missing() {
        var c = new GetTeamConnector();
        c.setInputParameters(new HashMap<>(Map.of(
                "tenantId", "t", "clientId", "c", "clientSecret", "s")));
        org.junit.jupiter.api.Assertions.assertThrows(
                org.bonitasoft.engine.connector.ConnectorValidationException.class,
                c::validateInputParameters);
    }
}
