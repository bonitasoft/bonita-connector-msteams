package com.bonitasoft.connectors.msteams.channels;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.Mockito.*;

import com.bonitasoft.connectors.msteams.common.MsTeamsGraphClient;
import com.fasterxml.jackson.databind.JsonNode;
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
import java.util.List;
import java.util.Map;

@ExtendWith(MockitoExtension.class)
class ListChannelsConnectorTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Mock
    private MsTeamsGraphClient graphClient;

    private ListChannelsConnector connector;
    private Map<String, Object> inputs;

    @BeforeEach
    void setUp() {
        connector = new ListChannelsConnector();
        inputs = new HashMap<>();
        inputs.put("tenantId", "test-tenant");
        inputs.put("clientId", "test-client");
        inputs.put("clientSecret", "test-secret");
        inputs.put("teamId", "team-123");
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
        void should_fail_when_team_id_is_missing() {
            var c = new ListChannelsConnector();
            var incompleteInputs = new HashMap<>(inputs);
            incompleteInputs.remove("teamId");
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
        void should_list_channels_successfully() throws Exception {
            JsonNode channel1 = MAPPER.readTree("{\"id\":\"ch-1\",\"displayName\":\"General\"}");
            JsonNode channel2 = MAPPER.readTree("{\"id\":\"ch-2\",\"displayName\":\"Dev\"}");
            when(graphClient.getWithPagination("/teams/team-123/channels"))
                    .thenReturn(List.of(channel1, channel2));

            connector.executeBusinessLogic();

            assertThat(connector.getOutputs().get("success")).isEqualTo(true);
            assertThat(connector.getOutputs().get("channelCount")).isEqualTo(2);
            assertThat(connector.getOutputs().get("channelsJson").toString()).contains("General");
        }
    }
}
