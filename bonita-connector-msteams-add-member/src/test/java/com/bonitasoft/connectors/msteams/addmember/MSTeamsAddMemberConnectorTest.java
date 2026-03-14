package com.bonitasoft.connectors.msteams.addmember;

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

@DisplayName("MSTeamsAddMemberConnector")
@ExtendWith(MockitoExtension.class)
class MSTeamsAddMemberConnectorTest {

    private MSTeamsAddMemberConnector connector;
    @Mock private MSTeamsClient mockClient;
    private static final ObjectMapper MAPPER = new ObjectMapper();

    private Map<String, Object> validParams() {
        Map<String, Object> params = new HashMap<>();
        params.put("tenantId", "test-tenant");
        params.put("clientId", "test-client");
        params.put("clientSecret", "test-secret");
        params.put("teamId", "team-123");
        params.put("userPrincipalName", "user@domain.com");
        return params;
    }

    @BeforeEach
    void setUp() { connector = new MSTeamsAddMemberConnector(); }

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
        @Test void should_fail_when_userPrincipalName_is_missing() {
            Map<String, Object> p = validParams(); p.remove("userPrincipalName"); connector.setInputParameters(p);
            assertThatThrownBy(() -> connector.validateInputParameters()).isInstanceOf(ConnectorValidationException.class).hasMessageContaining("userPrincipalName");
        }
        @Test void should_pass_validation_when_all_required_params_present() throws ConnectorValidationException {
            connector.setInputParameters(validParams()); connector.validateInputParameters();
        }
        @Test void should_pass_with_owner_role() throws ConnectorValidationException {
            Map<String, Object> p = validParams(); p.put("role", "owner"); connector.setInputParameters(p); connector.validateInputParameters();
        }
    }

    @Nested @DisplayName("Execution")
    class Execution {
        @Test void should_return_correct_connector_name() {
            assertThat(connector.getConnectorName()).isEqualTo("MSTeams-AddMember");
        }
        @Test void should_execute_successfully() throws Exception {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            when(mockClient.post(eq("/teams/team-123/members"), anyString()))
                    .thenReturn(MAPPER.readTree("{\"id\":\"m-1\",\"displayName\":\"User One\"}"));
            connector.executeOperation(mockClient);
            verify(mockClient).post(eq("/teams/team-123/members"), anyString());
            assertThat(connector.getOutputs().get("membershipId")).isEqualTo("m-1");
            assertThat(connector.getOutputs().get("displayName")).isEqualTo("User One");
        }
        @Test void should_execute_with_owner_role() throws Exception {
            Map<String, Object> p = validParams(); p.put("role", "owner");
            connector.setInputParameters(p); connector.setClient(mockClient);
            when(mockClient.post(eq("/teams/team-123/members"), anyString()))
                    .thenReturn(MAPPER.readTree("{\"id\":\"m-2\",\"displayName\":\"Owner\"}"));
            connector.executeOperation(mockClient);
            verify(mockClient).post(eq("/teams/team-123/members"), org.mockito.ArgumentMatchers.contains("owner"));
        }
        @Test void should_handle_api_error() {
            connector.setInputParameters(validParams()); connector.setClient(mockClient);
            when(mockClient.post(anyString(), anyString())).thenThrow(new MSTeamsException("API error", 403, "Forbidden", false, null));
            assertThatThrownBy(() -> connector.executeOperation(mockClient)).isInstanceOf(MSTeamsException.class);
        }
    }
}
