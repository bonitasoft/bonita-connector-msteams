package com.bonitasoft.connectors.msteams.channels;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.Mockito.*;

import com.bonitasoft.connectors.msteams.common.MsTeamsGraphClient;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import org.bonitasoft.engine.connector.ConnectorException;
import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.HashMap;
import java.util.Map;

@ExtendWith(MockitoExtension.class)
class DeleteChannelConnectorTest {

    @Mock
    private MsTeamsGraphClient graphClient;

    private DeleteChannelConnector connector;
    private Map<String, Object> inputs;

    @BeforeEach
    void setUp() {
        connector = new DeleteChannelConnector();
        inputs = new HashMap<>();
        inputs.put("tenantId", "test-tenant");
        inputs.put("clientId", "test-client");
        inputs.put("clientSecret", "test-secret");
        inputs.put("teamId", "team-123");
        inputs.put("channelId", "ch-123");
        connector.setInputParameters(inputs);
    }

    @Nested
    @DisplayName("Phase 1: Validation")
    class ValidationTests {

        @Test
        void should_pass_with_valid_inputs() throws ConnectorValidationException {
            connector.validateInputParameters();
        }

        @Test
        void should_fail_when_channel_id_is_missing() {
            var c = new DeleteChannelConnector();
            var incompleteInputs = new HashMap<>(inputs);
            incompleteInputs.remove("channelId");
            c.setInputParameters(incompleteInputs);
            assertThatThrownBy(() -> c.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }
    }

    @Nested
    @DisplayName("Phase 3: Execute")
    class ExecutionTests {

        @BeforeEach
        void setUpClient() {
            connector.setGraphClient(graphClient);
        }

        @Test
        void should_delete_channel_successfully() throws ConnectorException {
            connector.executeBusinessLogic();
            verify(graphClient).delete("/teams/team-123/channels/ch-123");
            assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        }

        @Test
        void should_handle_delete_error() {
            doThrow(new MsTeamsException.ExecutionException("Not found", 404))
                    .when(graphClient).delete(any());
            assertThatThrownBy(() -> connector.executeBusinessLogic())
                    .isInstanceOf(ConnectorException.class);
        }
    }
}
