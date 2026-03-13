package com.bonitasoft.connectors.msteams.channels;

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
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.HashMap;
import java.util.Map;

@ExtendWith(MockitoExtension.class)
class CreateChannelConnectorTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Mock
    private MsTeamsGraphClient graphClient;

    private CreateChannelConnector connector;
    private Map<String, Object> inputs;

    @BeforeEach
    void setUp() {
        connector = new CreateChannelConnector();
        inputs = new HashMap<>();
        inputs.put("tenantId", "test-tenant");
        inputs.put("clientId", "test-client");
        inputs.put("clientSecret", "test-secret");
        inputs.put("teamId", "team-123");
        inputs.put("displayName", "New Channel");
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
        void should_fail_when_display_name_is_missing() {
            var c = new CreateChannelConnector();
            var incompleteInputs = new HashMap<>(inputs);
            incompleteInputs.remove("displayName");
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
        void should_create_channel_successfully() throws Exception {
            when(graphClient.post(eq("/teams/team-123/channels"), any()))
                    .thenReturn(MAPPER.readTree("{\"id\":\"ch-new\",\"webUrl\":\"https://teams.microsoft.com/channel\"}"));

            connector.executeBusinessLogic();

            assertThat(connector.getOutputs().get("success")).isEqualTo(true);
            assertThat(connector.getOutputs().get("channelId")).isEqualTo("ch-new");
        }
    }
}
