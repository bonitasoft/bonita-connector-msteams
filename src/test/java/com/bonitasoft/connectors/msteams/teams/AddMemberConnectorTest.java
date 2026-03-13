package com.bonitasoft.connectors.msteams.teams;

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
class AddMemberConnectorTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Mock
    private MsTeamsGraphClient graphClient;

    private AddMemberConnector connector;
    private Map<String, Object> inputs;

    @BeforeEach
    void setUp() {
        connector = new AddMemberConnector();
        inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("teamId", "team-1");
        inputs.put("userId", "user-1");
        connector.setInputParameters(inputs);
    }

    @Test
    void should_pass_validation() throws ConnectorValidationException {
        connector.validateInputParameters();
    }

    @Test
    void should_fail_when_user_id_missing() {
        var c = new AddMemberConnector();
        var incompleteInputs = new HashMap<>(inputs);
        incompleteInputs.remove("userId");
        c.setInputParameters(incompleteInputs);
        assertThatThrownBy(() -> c.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class);
    }

    @Test
    void should_add_member_successfully() throws Exception {
        connector.setGraphClient(graphClient);
        when(graphClient.post(eq("/teams/team-1/members"), any()))
                .thenReturn(MAPPER.readTree("{\"id\":\"membership-1\"}"));

        connector.executeBusinessLogic();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("membershipId")).isEqualTo("membership-1");
    }
}
