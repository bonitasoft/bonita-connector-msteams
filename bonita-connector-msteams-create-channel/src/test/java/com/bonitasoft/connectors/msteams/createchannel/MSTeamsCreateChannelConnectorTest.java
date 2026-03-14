package com.bonitasoft.connectors.msteams.createchannel;

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
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

@DisplayName("MSTeamsCreateChannelConnector")
@ExtendWith(MockitoExtension.class)
class MSTeamsCreateChannelConnectorTest {

    private MSTeamsCreateChannelConnector connector;
    @Mock private MSTeamsClient mockClient;
    private static final ObjectMapper MAPPER = new ObjectMapper();

    private Map<String, Object> validParams() {
        Map<String, Object> p = new HashMap<>();
        p.put("tenantId", "t"); p.put("clientId", "c"); p.put("clientSecret", "s");
        p.put("teamId", "team-123"); p.put("displayName", "General");
        return p;
    }

    @BeforeEach void setUp() { connector = new MSTeamsCreateChannelConnector(); }

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
        @Test void should_fail_when_displayName_is_missing() {
            Map<String, Object> p = validParams(); p.remove("displayName"); connector.setInputParameters(p);
            assertThatThrownBy(() -> connector.validateInputParameters()).isInstanceOf(ConnectorValidationException.class).hasMessageContaining("displayName");
        }
        @Test void should_fail_when_displayName_exceeds_max_length() {
            Map<String, Object> p = validParams(); p.put("displayName", "a".repeat(51)); connector.setInputParameters(p);
            assertThatThrownBy(() -> connector.validateInputParameters()).isInstanceOf(ConnectorValidationException.class).hasMessageContaining("exceeds maximum length");
        }
        @Test void should_pass_validation_when_all_required_params_present() throws ConnectorValidationException {
            connector.setInputParameters(validParams()); connector.validateInputParameters();
        }
    }

    @Nested @DisplayName("Execution")
    class Execution {
        @Test void should_return_correct_connector_name() { assertThat(connector.getConnectorName()).isEqualTo("MSTeams-CreateChannel"); }
        @Test void should_execute_successfully() throws Exception {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            when(mockClient.post(eq("/teams/team-123/channels"), anyString()))
                    .thenReturn(MAPPER.readTree("{\"id\":\"ch-1\",\"displayName\":\"General\",\"webUrl\":\"https://teams.microsoft.com/ch\"}"));
            connector.executeOperation(mockClient);
            verify(mockClient).post(eq("/teams/team-123/channels"), anyString());
            assertThat(connector.getOutputs().get("channelId")).isEqualTo("ch-1");
            assertThat(connector.getOutputs().get("webUrl")).isEqualTo("https://teams.microsoft.com/ch");
        }
        @Test void should_execute_with_description() throws Exception {
            Map<String, Object> p = validParams(); p.put("description", "A test channel");
            connector.setInputParameters(p); connector.setClient(mockClient);
            when(mockClient.post(anyString(), anyString())).thenReturn(MAPPER.readTree("{\"id\":\"ch-2\",\"displayName\":\"General\"}"));
            connector.executeOperation(mockClient);
            verify(mockClient).post(anyString(), org.mockito.ArgumentMatchers.contains("description"));
        }
        @Test void should_handle_api_error() {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            when(mockClient.post(anyString(), anyString())).thenThrow(new MSTeamsException("error", 400, "BadRequest", false, null));
            assertThatThrownBy(() -> connector.executeOperation(mockClient)).isInstanceOf(MSTeamsException.class);
        }
    }
}
