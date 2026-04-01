package com.bonitasoft.connectors.msteams.teams;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.*;

import com.bonitasoft.connectors.msteams.common.MsTeamsGraphClient;
import org.bonitasoft.engine.connector.ConnectorException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.HashMap;
import java.util.Map;

@ExtendWith(MockitoExtension.class)
class RemoveMemberConnectorTest {

    @Mock
    private MsTeamsGraphClient graphClient;

    private RemoveMemberConnector connector;

    @BeforeEach
    void setUp() {
        connector = new RemoveMemberConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("teamId", "team-1");
        inputs.put("membershipId", "mem-1");
        connector.setInputParameters(inputs);
    }

    @Test
    void should_remove_member_successfully() throws ConnectorException {
        connector.setGraphClient(graphClient);
        connector.executeBusinessLogic();
        verify(graphClient).delete("/teams/team-1/members/mem-1");
        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
    }

    @Test
    void should_set_error_outputs_when_api_fails() {
        connector.setGraphClient(graphClient);
        doThrow(new com.bonitasoft.connectors.msteams.common.MsTeamsException("Not found"))
                .when(graphClient).delete(any());
        org.junit.jupiter.api.Assertions.assertThrows(ConnectorException.class,
                () -> connector.executeBusinessLogic());
        assertThat(connector.getOutputs().get("success")).isEqualTo(false);
        assertThat(connector.getOutputs().get("errorMessage")).isEqualTo("Not found");
    }

    @Test
    void should_fail_validation_when_teamId_missing() {
        var c = new RemoveMemberConnector();
        Map<String, Object> params = Map.of(
                "tenantId", "t", "clientId", "c", "clientSecret", "s",
                "membershipId", "mem-1");
        c.setInputParameters(new HashMap<>(params));
        org.junit.jupiter.api.Assertions.assertThrows(
                org.bonitasoft.engine.connector.ConnectorValidationException.class,
                c::validateInputParameters);
    }

    @Test
    void should_fail_validation_when_membershipId_missing() {
        var c = new RemoveMemberConnector();
        Map<String, Object> params = Map.of(
                "tenantId", "t", "clientId", "c", "clientSecret", "s",
                "teamId", "team-1");
        c.setInputParameters(new HashMap<>(params));
        org.junit.jupiter.api.Assertions.assertThrows(
                org.bonitasoft.engine.connector.ConnectorValidationException.class,
                c::validateInputParameters);
    }
}
