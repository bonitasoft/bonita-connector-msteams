package com.bonitasoft.connectors.msteams.getteam;

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
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

@DisplayName("MSTeamsGetTeamConnector")
@ExtendWith(MockitoExtension.class)
class MSTeamsGetTeamConnectorTest {

    private MSTeamsGetTeamConnector connector;
    @Mock private MSTeamsClient mockClient;
    private static final ObjectMapper MAPPER = new ObjectMapper();

    private Map<String, Object> validParams() {
        Map<String, Object> p = new HashMap<>();
        p.put("tenantId", "t"); p.put("clientId", "c"); p.put("clientSecret", "s");
        p.put("teamId", "team-123");
        return p;
    }

    @BeforeEach void setUp() { connector = new MSTeamsGetTeamConnector(); }

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
        @Test void should_return_correct_connector_name() { assertThat(connector.getConnectorName()).isEqualTo("MSTeams-GetTeam"); }
        @Test void should_execute_successfully() throws Exception {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            when(mockClient.get("/teams/team-123")).thenReturn(MAPPER.readTree(
                    "{\"id\":\"team-123\",\"displayName\":\"My Team\",\"description\":\"Test\",\"visibility\":\"public\",\"isArchived\":false,\"webUrl\":\"https://url\"}"));
            connector.executeOperation(mockClient);
            verify(mockClient).get("/teams/team-123");
            assertThat(connector.getOutputs().get("teamId")).isEqualTo("team-123");
            assertThat(connector.getOutputs().get("displayName")).isEqualTo("My Team");
            assertThat(connector.getOutputs().get("description")).isEqualTo("Test");
            assertThat(connector.getOutputs().get("visibility")).isEqualTo("public");
            assertThat(connector.getOutputs().get("isArchived")).isEqualTo(false);
        }
        @Test void should_handle_null_description() throws Exception {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            when(mockClient.get("/teams/team-123")).thenReturn(MAPPER.readTree("{\"id\":\"team-123\",\"displayName\":\"T\",\"description\":null}"));
            connector.executeOperation(mockClient);
            assertThat(connector.getOutputs().get("description")).isNull();
        }
        @Test void should_handle_api_error() {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            when(mockClient.get("/teams/team-123")).thenThrow(new MSTeamsException("error", 404, "NotFound", false, null));
            assertThatThrownBy(() -> connector.executeOperation(mockClient)).isInstanceOf(MSTeamsException.class);
        }
        @Test void should_handle_response_without_optional_fields() throws Exception {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            when(mockClient.get("/teams/team-123")).thenReturn(MAPPER.readTree("{\"other\":\"x\"}"));
            connector.executeOperation(mockClient);
            assertThat(connector.getOutputs().get("teamId")).isNull();
            assertThat(connector.getOutputs().get("displayName")).isNull();
            assertThat(connector.getOutputs().get("description")).isNull();
            assertThat(connector.getOutputs().get("visibility")).isNull();
            assertThat(connector.getOutputs().get("isArchived")).isEqualTo(false);
            assertThat(connector.getOutputs().get("webUrl")).isNull();
        }
        @Test void should_handle_archived_team() throws Exception {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            when(mockClient.get("/teams/team-123")).thenReturn(MAPPER.readTree(
                    "{\"id\":\"team-123\",\"displayName\":\"Archived\",\"isArchived\":true,\"visibility\":\"private\",\"webUrl\":\"https://url\"}"));
            connector.executeOperation(mockClient);
            assertThat(connector.getOutputs().get("isArchived")).isEqualTo(true);
            assertThat(connector.getOutputs().get("visibility")).isEqualTo("private");
            assertThat(connector.getOutputs().get("webUrl")).isEqualTo("https://url");
        }
        @Test void should_fail_when_teamId_is_blank() {
            Map<String, Object> p = validParams(); p.put("teamId", "  "); connector.setInputParameters(p);
            assertThatThrownBy(() -> connector.validateInputParameters()).isInstanceOf(ConnectorValidationException.class).hasMessageContaining("teamId");
        }
    }
}
