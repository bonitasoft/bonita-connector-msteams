package com.bonitasoft.connectors.msteams.common;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.bonitasoft.engine.connector.ConnectorException;
import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

@DisplayName("AbstractMsTeamsConnector")
class AbstractMsTeamsConnectorTest {

    private ConcreteConnector connector;
    private final Map<String, Object> inputs = new HashMap<>();

    @BeforeEach
    void setUp() {
        connector = new ConcreteConnector();
        inputs.clear();
    }

    @Nested
    @DisplayName("Input helpers")
    class InputHelpers {

        @Test
        void should_return_null_for_missing_string_input() {
            connector.setInputParameters(inputs);
            assertThat(connector.getStringInput("missing")).isNull();
        }

        @Test
        void should_return_string_value() {
            inputs.put("key", "value");
            connector.setInputParameters(inputs);
            assertThat(connector.getStringInput("key")).isEqualTo("value");
        }

        @Test
        void should_return_null_for_missing_integer_input() {
            connector.setInputParameters(inputs);
            assertThat(connector.getIntegerInput("missing")).isNull();
        }

        @Test
        void should_return_integer_value() {
            inputs.put("num", 42);
            connector.setInputParameters(inputs);
            assertThat(connector.getIntegerInput("num")).isEqualTo(42);
        }

        @Test
        void should_parse_integer_from_string() {
            inputs.put("num", "99");
            connector.setInputParameters(inputs);
            assertThat(connector.getIntegerInput("num")).isEqualTo(99);
        }

        @Test
        void should_return_default_for_missing_integer() {
            connector.setInputParameters(inputs);
            assertThat(connector.getIntegerInputOrDefault("missing", 100)).isEqualTo(100);
        }

        @Test
        void should_return_value_over_default_for_integer() {
            inputs.put("num", 50);
            connector.setInputParameters(inputs);
            assertThat(connector.getIntegerInputOrDefault("num", 100)).isEqualTo(50);
        }

        @Test
        void should_return_null_for_missing_boolean_input() {
            connector.setInputParameters(inputs);
            assertThat(connector.getBooleanInput("missing")).isNull();
        }

        @Test
        void should_return_boolean_value() {
            inputs.put("flag", true);
            connector.setInputParameters(inputs);
            assertThat(connector.getBooleanInput("flag")).isTrue();
        }

        @Test
        void should_parse_boolean_from_string() {
            inputs.put("flag", "true");
            connector.setInputParameters(inputs);
            assertThat(connector.getBooleanInput("flag")).isTrue();
        }

        @Test
        void should_return_default_for_missing_boolean() {
            connector.setInputParameters(inputs);
            assertThat(connector.getBooleanInputOrDefault("missing", true)).isTrue();
        }

        @Test
        void should_return_value_over_default_for_boolean() {
            inputs.put("flag", false);
            connector.setInputParameters(inputs);
            assertThat(connector.getBooleanInputOrDefault("flag", true)).isFalse();
        }
    }

    @Nested
    @DisplayName("Output helpers")
    class OutputHelpers {

        @Test
        void should_set_success_outputs() throws Exception {
            connector.setInputParameters(inputs);
            connector.executeForTest();
            Map<String, Object> outputs = connector.getOutputs();
            assertThat(outputs.get("success")).isEqualTo(true);
            assertThat(outputs.get("errorMessage")).isEqualTo("");
        }

        @Test
        void should_set_error_outputs() {
            connector.setInputParameters(inputs);
            connector.setErrorOutputs("something went wrong");
            Map<String, Object> outputs = connector.getOutputs();
            assertThat(outputs.get("success")).isEqualTo(false);
            assertThat(outputs.get("errorMessage")).isEqualTo("something went wrong");
        }

        @Test
        void should_truncate_long_error_message() {
            connector.setInputParameters(inputs);
            String longMsg = "x".repeat(2000);
            connector.setErrorOutputs(longMsg);
            Map<String, Object> outputs = connector.getOutputs();
            assertThat((String) outputs.get("errorMessage")).hasSize(1000);
        }

        @Test
        void should_handle_null_error_message() {
            connector.setInputParameters(inputs);
            connector.setErrorOutputs(null);
            Map<String, Object> outputs = connector.getOutputs();
            assertThat(outputs.get("errorMessage")).isNull();
        }
    }

    @Nested
    @DisplayName("isNullOrEmpty")
    class NullOrEmpty {

        @Test
        void should_return_true_for_null() {
            assertThat(connector.isNullOrEmpty(null)).isTrue();
        }

        @Test
        void should_return_true_for_empty() {
            assertThat(connector.isNullOrEmpty("")).isTrue();
        }

        @Test
        void should_return_true_for_blank() {
            assertThat(connector.isNullOrEmpty("   ")).isTrue();
        }

        @Test
        void should_return_false_for_value() {
            assertThat(connector.isNullOrEmpty("hello")).isFalse();
        }
    }

    @Nested
    @DisplayName("Validation")
    class Validation {

        @Test
        void should_throw_when_validation_errors_present() {
            var failingConnector = new FailingValidationConnector();
            failingConnector.setInputParameters(inputs);
            assertThatThrownBy(() -> failingConnector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }
    }

    // Concrete impl that always succeeds
    static class ConcreteConnector extends AbstractMsTeamsConnector {
        @Override
        protected void validateConnectionParameters(List<String> errors) {}

        @Override
        protected void validateOperationParameters(List<String> errors) {}

        @Override
        protected void executeBusinessLogic() throws ConnectorException {
            setSuccessOutputs();
        }

        @Override
        public void connect() throws ConnectorException {}

        @Override
        public void disconnect() throws ConnectorException {}
    }

    // Concrete impl that always fails validation
    static class FailingValidationConnector extends ConcreteConnector {
        @Override
        protected void validateConnectionParameters(List<String> errors) {
            errors.add("connection error");
        }

        @Override
        protected void validateOperationParameters(List<String> errors) {
            errors.add("operation error");
        }
    }
}
