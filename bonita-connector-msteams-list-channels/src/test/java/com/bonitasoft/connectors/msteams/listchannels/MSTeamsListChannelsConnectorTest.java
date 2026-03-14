package com.bonitasoft.connectors.msteams.listchannels;

import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.ObjectMapper;
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

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

@DisplayName("MSTeamsListChannelsConnector")
@ExtendWith(MockitoExtension.class)
class MSTeamsListChannelsConnectorTest {

    private MSTeamsListChannelsConnector connector;
    @Mock private MSTeamsClient mockClient;
    private static final ObjectMapper MAPPER = new ObjectMapper();

    private Map<String, Object> validParams() {
        Map<String, Object> p = new HashMap<>();
        p.put("tenantId", "t"); p.put("clientId", "c"); p.put("clientSecret", "s");
        p.put("teamId", "team-123");
        return p;
    }

    @BeforeEach void setUp() { connector = new MSTeamsListChannelsConnector(); }

    @Nested @DisplayName("Validation")
    class Validation {
        @Test void should_fail_when_teamId_is_missing() {
            Map<String, Object> p = validParams(); p.remove("teamId"); connector.setInputParameters(p);
            assertThatThrownBy(() -> connector.validateInputParameters()).isInstanceOf(ConnectorValidationException.class).hasMessageContaining("teamId");
        }
        @Test void should_pass_validation_when_all_required_params_present() throws ConnectorValidationException {
            connector.setInputParameters(validParams()); connector.validateInputParameters();
        }
    }

    @Nested @DisplayName("Execution")
    class Execution {
        @Test void should_return_correct_connector_name() { assertThat(connector.getConnectorName()).isEqualTo("MSTeams-ListChannels"); }
        @Test @SuppressWarnings("unchecked") void should_execute_successfully() throws Exception {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            when(mockClient.get("/teams/team-123/channels")).thenReturn(MAPPER.readTree(
                    "{\"value\":[{\"id\":\"ch-1\",\"displayName\":\"General\",\"description\":\"desc\",\"membershipType\":\"standard\"},{\"id\":\"ch-2\",\"displayName\":\"Dev\"}]}"));
            connector.executeOperation(mockClient);
            verify(mockClient).get("/teams/team-123/channels");
            List<Map<String, Object>> channels = (List<Map<String, Object>>) connector.getOutputs().get("channels");
            assertThat(channels).hasSize(2);
            assertThat(channels.get(0).get("id")).isEqualTo("ch-1");
            assertThat(connector.getOutputs().get("totalCount")).isEqualTo(2);
        }
        @Test @SuppressWarnings("unchecked") void should_handle_empty_channels() throws Exception {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            when(mockClient.get("/teams/team-123/channels")).thenReturn(MAPPER.readTree("{\"value\":[]}"));
            connector.executeOperation(mockClient);
            List<Map<String, Object>> channels = (List<Map<String, Object>>) connector.getOutputs().get("channels");
            assertThat(channels).isEmpty();
            assertThat(connector.getOutputs().get("totalCount")).isEqualTo(0);
        }
        @Test void should_handle_api_error() {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            when(mockClient.get("/teams/team-123/channels")).thenThrow(new MSTeamsException("error"));
            assertThatThrownBy(() -> connector.executeOperation(mockClient)).isInstanceOf(MSTeamsException.class);
        }
    }
}
