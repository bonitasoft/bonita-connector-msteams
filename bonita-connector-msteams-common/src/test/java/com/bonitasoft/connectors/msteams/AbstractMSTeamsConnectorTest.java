package com.bonitasoft.connectors.msteams;

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

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.Mockito.*;

@DisplayName("AbstractMSTeamsConnector")
@ExtendWith(MockitoExtension.class)
class AbstractMSTeamsConnectorTest {

    @Mock private MSTeamsClient mockClient;

    private TestConnector connector;

    @BeforeEach
    void setUp() {
        connector = new TestConnector();
    }

    private Map<String, Object> validParams() {
        Map<String, Object> p = new HashMap<>();
        p.put("tenantId", "t");
        p.put("clientId", "c");
        p.put("clientSecret", "s");
        return p;
    }

    @Nested
    @DisplayName("validateInputParameters")
    class Validate {

        @Test
        void should_fail_when_tenantId_is_missing() {
            Map<String, Object> p = validParams();
            p.remove("tenantId");
            connector.setInputParameters(p);
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("Authentication required");
        }

        @Test
        void should_fail_when_clientId_is_blank() {
            Map<String, Object> p = validParams();
            p.put("clientId", "  ");
            connector.setInputParameters(p);
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }

        @Test
        void should_fail_when_clientSecret_is_missing() {
            Map<String, Object> p = validParams();
            p.remove("clientSecret");
            connector.setInputParameters(p);
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }

        @Test
        void should_pass_when_all_auth_params_present() throws ConnectorValidationException {
            connector.setInputParameters(validParams());
            connector.validateInputParameters();
        }

        @Test
        void should_fail_when_connectTimeout_negative() {
            Map<String, Object> p = validParams();
            p.put("connectTimeout", -1);
            connector.setInputParameters(p);
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("connectTimeout");
        }

        @Test
        void should_fail_when_readTimeout_negative() {
            Map<String, Object> p = validParams();
            p.put("readTimeout", -5);
            connector.setInputParameters(p);
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("readTimeout");
        }

        @Test
        void should_accept_positive_timeouts() throws ConnectorValidationException {
            Map<String, Object> p = validParams();
            p.put("connectTimeout", 5000);
            p.put("readTimeout", 10000);
            connector.setInputParameters(p);
            connector.validateInputParameters();
        }

        @Test
        void should_accept_null_timeouts() throws ConnectorValidationException {
            connector.setInputParameters(validParams());
            connector.validateInputParameters();
        }
    }

    @Nested
    @DisplayName("executeBusinessLogic")
    class Execute {

        @Test
        void should_set_success_outputs() throws ConnectorException {
            connector.setInputParameters(validParams());
            connector.setClient(mockClient);
            connector.executeBusinessLogic();
            assertThat(connector.getOutputs().get("success")).isEqualTo(true);
            assertThat(connector.getOutputs().get("errorMessage")).isEqualTo("");
        }

        @Test
        void should_set_error_outputs_on_msteams_exception() {
            connector.failWith = new MSTeamsException("API failed", 500, "ServerError", false, null);
            connector.setInputParameters(validParams());
            connector.setClient(mockClient);
            assertThatThrownBy(() -> connector.executeBusinessLogic())
                    .isInstanceOf(ConnectorException.class);
            assertThat(connector.getOutputs().get("success")).isEqualTo(false);
            assertThat(connector.getOutputs().get("errorMessage").toString()).contains("API failed");
        }

        @Test
        void should_set_error_outputs_on_unexpected_exception() {
            connector.failWith = new RuntimeException("unexpected");
            connector.setInputParameters(validParams());
            connector.setClient(mockClient);
            assertThatThrownBy(() -> connector.executeBusinessLogic())
                    .isInstanceOf(ConnectorException.class);
            assertThat(connector.getOutputs().get("success")).isEqualTo(false);
            assertThat(connector.getOutputs().get("errorMessage").toString()).contains("Unexpected error");
        }
    }

    @Nested
    @DisplayName("disconnect")
    class Disconnect {

        @Test
        void should_close_client() throws ConnectorException {
            connector.setInputParameters(validParams());
            connector.setClient(mockClient);
            connector.disconnect();
            verify(mockClient).close();
        }

        @Test
        void should_handle_null_client() throws ConnectorException {
            connector.setInputParameters(validParams());
            connector.disconnect(); // client is null, should not throw
        }

        @Test
        void should_handle_close_exception() throws Exception {
            connector.setInputParameters(validParams());
            connector.setClient(mockClient);
            doThrow(new RuntimeException("close error")).when(mockClient).close();
            connector.disconnect(); // should not throw, just log
        }
    }

    @Nested
    @DisplayName("Input helpers")
    class InputHelpers {

        @Test
        void should_get_string_input() {
            Map<String, Object> p = new HashMap<>();
            p.put("key", "value");
            connector.setInputParameters(p);
            assertThat(connector.getStringInput("key")).isEqualTo("value");
            assertThat(connector.getStringInput("missing")).isNull();
        }

        @Test
        void should_get_string_input_or_default() {
            Map<String, Object> p = new HashMap<>();
            p.put("key", "value");
            p.put("blank", "  ");
            connector.setInputParameters(p);
            assertThat(connector.getStringInputOrDefault("key", "def")).isEqualTo("value");
            assertThat(connector.getStringInputOrDefault("missing", "def")).isEqualTo("def");
            assertThat(connector.getStringInputOrDefault("blank", "def")).isEqualTo("def");
        }

        @Test
        void should_get_int_input() {
            Map<String, Object> p = new HashMap<>();
            p.put("int", 42);
            p.put("long", 99L);
            p.put("str", "123");
            p.put("bad", "xyz");
            connector.setInputParameters(p);
            assertThat(connector.getIntInput("int")).isEqualTo(42);
            assertThat(connector.getIntInput("long")).isEqualTo(99);
            assertThat(connector.getIntInput("str")).isEqualTo(123);
            assertThat(connector.getIntInput("bad")).isNull();
            assertThat(connector.getIntInput("missing")).isNull();
        }

        @Test
        void should_get_int_input_or_default() {
            Map<String, Object> p = new HashMap<>();
            p.put("key", 10);
            connector.setInputParameters(p);
            assertThat(connector.getIntInputOrDefault("key", 99)).isEqualTo(10);
            assertThat(connector.getIntInputOrDefault("missing", 99)).isEqualTo(99);
        }

        @Test
        void should_get_boolean_input() {
            Map<String, Object> p = new HashMap<>();
            p.put("bool", true);
            p.put("str", "true");
            connector.setInputParameters(p);
            assertThat(connector.getBooleanInput("bool")).isTrue();
            assertThat(connector.getBooleanInput("str")).isTrue();
            assertThat(connector.getBooleanInput("missing")).isNull();
        }

        @Test
        void should_get_boolean_input_or_default() {
            Map<String, Object> p = new HashMap<>();
            p.put("key", false);
            connector.setInputParameters(p);
            assertThat(connector.getBooleanInputOrDefault("key", true)).isFalse();
            assertThat(connector.getBooleanInputOrDefault("missing", true)).isTrue();
        }

        @Test
        void should_get_list_input() {
            Map<String, Object> p = new HashMap<>();
            p.put("list", List.of("a", "b"));
            p.put("notList", "str");
            connector.setInputParameters(p);
            assertThat(connector.<String>getListInput("list")).containsExactly("a", "b");
            assertThat(connector.<String>getListInput("notList")).isNull();
            assertThat(connector.<String>getListInput("missing")).isNull();
        }
    }

    @Nested
    @DisplayName("buildConfiguration")
    class BuildConfig {

        @Test
        void should_build_default_configuration() {
            connector.setInputParameters(validParams());
            MSTeamsConfiguration config = connector.buildConfiguration();
            assertThat(config.tenantId()).isEqualTo("t");
            assertThat(config.connectTimeout()).isEqualTo(30_000);
            assertThat(config.readTimeout()).isEqualTo(60_000);
            assertThat(config.isAppOnly()).isFalse();
        }

        @Test
        void should_build_configuration_with_custom_values() {
            Map<String, Object> p = validParams();
            p.put("connectTimeout", 5000);
            p.put("readTimeout", 15000);
            p.put("appOnly", true);
            connector.setInputParameters(p);
            MSTeamsConfiguration config = connector.buildConfiguration();
            assertThat(config.connectTimeout()).isEqualTo(5000);
            assertThat(config.readTimeout()).isEqualTo(15000);
            assertThat(config.isAppOnly()).isTrue();
        }
    }

    /**
     * Concrete test implementation of AbstractMSTeamsConnector.
     */
    static class TestConnector extends AbstractMSTeamsConnector {
        Exception failWith;

        @Override
        protected void validateOperationInputs(List<String> errors) { }

        @Override
        protected String getConnectorName() { return "Test-Connector"; }

        @Override
        protected void executeOperation(MSTeamsClient client) throws MSTeamsException {
            if (failWith instanceof MSTeamsException mse) throw mse;
            if (failWith instanceof RuntimeException re) throw re;
        }
    }
}
