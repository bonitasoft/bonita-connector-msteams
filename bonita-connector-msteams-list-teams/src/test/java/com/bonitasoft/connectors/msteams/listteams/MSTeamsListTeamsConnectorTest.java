package com.bonitasoft.connectors.msteams.listteams;

import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsConfiguration;
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
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

@DisplayName("MSTeamsListTeamsConnector")
@ExtendWith(MockitoExtension.class)
class MSTeamsListTeamsConnectorTest {

    private MSTeamsListTeamsConnector connector;
    @Mock private MSTeamsClient mockClient;
    private static final ObjectMapper MAPPER = new ObjectMapper();

    private Map<String, Object> validParams() {
        Map<String, Object> p = new HashMap<>();
        p.put("tenantId", "t"); p.put("clientId", "c"); p.put("clientSecret", "s");
        return p;
    }

    @BeforeEach void setUp() { connector = new MSTeamsListTeamsConnector(); }

    @Nested @DisplayName("Validation")
    class Validation {
        @Test void should_fail_when_auth_missing() {
            Map<String, Object> p = validParams(); p.remove("tenantId"); p.remove("clientId"); p.remove("clientSecret");
            connector.setInputParameters(p);
            assertThatThrownBy(() -> connector.validateInputParameters()).isInstanceOf(ConnectorValidationException.class).hasMessageContaining("Authentication required");
        }
        @Test void should_pass_validation_when_all_required_params_present() throws ConnectorValidationException {
            connector.setInputParameters(validParams()); connector.validateInputParameters();
        }
    }

    @Nested @DisplayName("Execution")
    class Execution {
        @Test void should_return_correct_connector_name() { assertThat(connector.getConnectorName()).isEqualTo("MSTeams-ListTeams"); }
        @Test @SuppressWarnings("unchecked") void should_execute_successfully_app_only() throws Exception {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            var config = new MSTeamsConfiguration("t", "c", "s", 30000, 60000, true);
            when(mockClient.getConfiguration()).thenReturn(config);
            when(mockClient.get(anyString())).thenReturn(MAPPER.readTree(
                    "{\"value\":[{\"id\":\"t-1\",\"displayName\":\"Team A\",\"description\":\"Desc\",\"visibility\":\"public\"}]}"));
            connector.executeOperation(mockClient);
            verify(mockClient).get(anyString());
            List<Map<String, Object>> teams = (List<Map<String, Object>>) connector.getOutputs().get("teams");
            assertThat(teams).hasSize(1);
            assertThat(teams.get(0).get("displayName")).isEqualTo("Team A");
            assertThat(connector.getOutputs().get("totalCount")).isEqualTo(1);
        }
        @Test @SuppressWarnings("unchecked") void should_execute_successfully_delegated() throws Exception {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            var config = new MSTeamsConfiguration("t", "c", "s", 30000, 60000, false);
            when(mockClient.getConfiguration()).thenReturn(config);
            when(mockClient.get(anyString())).thenReturn(MAPPER.readTree("{\"value\":[]}"));
            connector.executeOperation(mockClient);
            List<Map<String, Object>> teams = (List<Map<String, Object>>) connector.getOutputs().get("teams");
            assertThat(teams).isEmpty();
        }
        @Test void should_handle_api_error() {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            var config = new MSTeamsConfiguration("t", "c", "s", 30000, 60000, true);
            when(mockClient.getConfiguration()).thenReturn(config);
            when(mockClient.get(anyString())).thenThrow(new MSTeamsException("error"));
            assertThatThrownBy(() -> connector.executeOperation(mockClient)).isInstanceOf(MSTeamsException.class);
        }

        @Test @SuppressWarnings("unchecked") void should_use_filter_when_provided_app_only() throws Exception {
            Map<String, Object> p = validParams();
            p.put("filter", "displayName eq 'MyTeam'");
            connector.setInputParameters(p); connector.setClient(mockClient);
            var config = new MSTeamsConfiguration("t", "c", "s", 30000, 60000, true);
            when(mockClient.getConfiguration()).thenReturn(config);
            when(mockClient.get(anyString())).thenReturn(MAPPER.readTree("{\"value\":[]}"));
            connector.executeOperation(mockClient);
            verify(mockClient).get(org.mockito.ArgumentMatchers.contains("and"));
        }

        @Test @SuppressWarnings("unchecked") void should_not_use_filter_when_blank() throws Exception {
            Map<String, Object> p = validParams();
            p.put("filter", "  ");
            connector.setInputParameters(p); connector.setClient(mockClient);
            var config = new MSTeamsConfiguration("t", "c", "s", 30000, 60000, true);
            when(mockClient.getConfiguration()).thenReturn(config);
            when(mockClient.get(anyString())).thenReturn(MAPPER.readTree("{\"value\":[]}"));
            connector.executeOperation(mockClient);
        }

        @Test @SuppressWarnings("unchecked") void should_handle_null_value_in_response() throws Exception {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            var config = new MSTeamsConfiguration("t", "c", "s", 30000, 60000, true);
            when(mockClient.getConfiguration()).thenReturn(config);
            when(mockClient.get(anyString())).thenReturn(MAPPER.readTree("{\"other\":\"data\"}"));
            connector.executeOperation(mockClient);
            List<Map<String, Object>> teams = (List<Map<String, Object>>) connector.getOutputs().get("teams");
            assertThat(teams).isEmpty();
        }

        @Test @SuppressWarnings("unchecked") void should_handle_missing_fields_in_team() throws Exception {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            var config = new MSTeamsConfiguration("t", "c", "s", 30000, 60000, true);
            when(mockClient.getConfiguration()).thenReturn(config);
            when(mockClient.get(anyString())).thenReturn(MAPPER.readTree(
                    "{\"value\":[{\"id\":\"t-1\"}]}"));
            connector.executeOperation(mockClient);
            List<Map<String, Object>> teams = (List<Map<String, Object>>) connector.getOutputs().get("teams");
            assertThat(teams).hasSize(1);
            assertThat(teams.get(0).get("id")).isEqualTo("t-1");
            assertThat(teams.get(0).get("displayName")).isNull();
            assertThat(teams.get(0).get("description")).isNull();
            assertThat(teams.get(0).get("visibility")).isNull();
        }

        @Test @SuppressWarnings("unchecked") void should_handle_null_description() throws Exception {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            var config = new MSTeamsConfiguration("t", "c", "s", 30000, 60000, true);
            when(mockClient.getConfiguration()).thenReturn(config);
            when(mockClient.get(anyString())).thenReturn(MAPPER.readTree(
                    "{\"value\":[{\"id\":\"t-1\",\"displayName\":\"A\",\"description\":null,\"visibility\":\"public\"}]}"));
            connector.executeOperation(mockClient);
            List<Map<String, Object>> teams = (List<Map<String, Object>>) connector.getOutputs().get("teams");
            assertThat(teams.get(0).get("description")).isNull();
        }

        @Test @SuppressWarnings("unchecked") void should_use_custom_max_results() throws Exception {
            Map<String, Object> p = validParams();
            p.put("maxResults", 50);
            connector.setInputParameters(p); connector.setClient(mockClient);
            var config = new MSTeamsConfiguration("t", "c", "s", 30000, 60000, true);
            when(mockClient.getConfiguration()).thenReturn(config);
            when(mockClient.get(anyString())).thenReturn(MAPPER.readTree("{\"value\":[]}"));
            connector.executeOperation(mockClient);
            verify(mockClient).get(org.mockito.ArgumentMatchers.contains("$top=50"));
        }
    }
}
