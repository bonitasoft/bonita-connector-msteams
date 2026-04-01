package com.bonitasoft.connectors.msteams.teams;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.*;

import com.bonitasoft.connectors.msteams.common.MsTeamsGraphClient;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.bonitasoft.engine.connector.ConnectorException;
import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

@ExtendWith(MockitoExtension.class)
class ListTeamsConnectorTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Mock
    private MsTeamsGraphClient graphClient;

    private ListTeamsConnector connector;

    @BeforeEach
    void setUp() {
        connector = new ListTeamsConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "test-tenant");
        inputs.put("clientId", "test-client");
        inputs.put("clientSecret", "test-secret");
        connector.setInputParameters(inputs);
    }

    @Test
    void should_pass_validation() throws ConnectorValidationException {
        connector.validateInputParameters();
    }

    @Test
    void should_list_teams_successfully() throws Exception {
        connector.setGraphClient(graphClient);
        JsonNode team = MAPPER.readTree("{\"id\":\"t-1\",\"displayName\":\"Engineering\"}");
        when(graphClient.getWithPagination(any())).thenReturn(List.of(team));

        connector.executeBusinessLogic();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("teamCount")).isEqualTo(1);
    }

    @Test
    void should_set_error_outputs_when_api_fails() {
        connector.setGraphClient(graphClient);
        when(graphClient.getWithPagination(any())).thenThrow(
                new com.bonitasoft.connectors.msteams.common.MsTeamsException("Access denied"));
        org.junit.jupiter.api.Assertions.assertThrows(
                org.bonitasoft.engine.connector.ConnectorException.class,
                () -> connector.executeBusinessLogic());
        assertThat(connector.getOutputs().get("success")).isEqualTo(false);
    }
}
