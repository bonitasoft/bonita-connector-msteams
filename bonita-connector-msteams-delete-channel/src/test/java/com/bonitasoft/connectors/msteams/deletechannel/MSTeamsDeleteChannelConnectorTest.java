package com.bonitasoft.connectors.msteams.deletechannel;

import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
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

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

@DisplayName("MSTeamsDeleteChannelConnector")
@ExtendWith(MockitoExtension.class)
class MSTeamsDeleteChannelConnectorTest {

    private MSTeamsDeleteChannelConnector connector;
    @Mock private MSTeamsClient mockClient;

    private Map<String, Object> validParams() {
        Map<String, Object> p = new HashMap<>();
        p.put("tenantId", "t"); p.put("clientId", "c"); p.put("clientSecret", "s");
        p.put("teamId", "team-123"); p.put("channelId", "channel-456");
        return p;
    }

    @BeforeEach void setUp() { connector = new MSTeamsDeleteChannelConnector(); }

    @Nested @DisplayName("Validation")
    class Validation {
        @Test void should_fail_when_auth_missing() {
            Map<String, Object> p = validParams(); p.remove("tenantId"); p.remove("clientId"); p.remove("clientSecret");
            connector.setInputParameters(p);
            assertThatThrownBy(() -> connector.validateInputParameters()).isInstanceOf(ConnectorValidationException.class).hasMessageContaining("Authentication required");
        }
        @Test void should_fail_when_teamId_is_missing() {
            Map<String, Object> p = validParams(); p.remove("teamId"); connector.setInputParameters(p);
            assertThatThrownBy(() -> connector.validateInputParameters()).isInstanceOf(ConnectorValidationException.class).hasMessageContaining("teamId");
        }
        @Test void should_fail_when_channelId_is_missing() {
            Map<String, Object> p = validParams(); p.remove("channelId"); connector.setInputParameters(p);
            assertThatThrownBy(() -> connector.validateInputParameters()).isInstanceOf(ConnectorValidationException.class).hasMessageContaining("channelId");
        }
        @Test void should_pass_validation_when_all_required_params_present() throws ConnectorValidationException {
            connector.setInputParameters(validParams()); connector.validateInputParameters();
        }
    }

    @Nested @DisplayName("Execution")
    class Execution {
        @Test void should_return_correct_connector_name() { assertThat(connector.getConnectorName()).isEqualTo("MSTeams-DeleteChannel"); }
        @Test void should_execute_successfully() throws Exception {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            when(mockClient.delete("/teams/team-123/channels/channel-456")).thenReturn(null);
            connector.executeOperation(mockClient);
            verify(mockClient).delete("/teams/team-123/channels/channel-456");
        }
        @Test void should_handle_api_error() {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            when(mockClient.delete("/teams/team-123/channels/channel-456")).thenThrow(new MSTeamsException("not found", 404, "NotFound", false, null));
            assertThatThrownBy(() -> connector.executeOperation(mockClient)).isInstanceOf(MSTeamsException.class);
        }
    }
}
