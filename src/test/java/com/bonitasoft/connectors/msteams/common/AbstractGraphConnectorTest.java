package com.bonitasoft.connectors.msteams.common;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.Mockito.*;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.bonitasoft.engine.connector.ConnectorException;
import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.junit.jupiter.MockitoExtension;

@DisplayName("AbstractGraphConnector")
@ExtendWith(MockitoExtension.class)
class AbstractGraphConnectorTest {

    private TestableGraphConnector connector;
    private final Map<String, Object> inputs = new HashMap<>();

    @BeforeEach
    void setUp() {
        connector = new TestableGraphConnector();
        inputs.clear();
        inputs.put("tenantId", "test-tenant");
        inputs.put("clientId", "test-client");
        inputs.put("clientSecret", "test-secret");
    }

    private void setInputs() {
        connector.setInputParameters(inputs);
    }

    @Nested
    @DisplayName("Phase 1: Validation")
    class Validation {

        @Test
        @DisplayName("should_pass_validation_when_all_app_only_params_present")
        void should_pass_validation_when_all_app_only_params_present() throws Exception {
            setInputs();
            connector.validateInputParameters();
        }

        @Test
        @DisplayName("should_fail_validation_when_tenantId_missing")
        void should_fail_validation_when_tenantId_missing() {
            inputs.remove("tenantId");
            setInputs();
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }

        @Test
        @DisplayName("should_fail_validation_when_clientId_missing")
        void should_fail_validation_when_clientId_missing() {
            inputs.remove("clientId");
            setInputs();
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }

        @Test
        @DisplayName("should_fail_validation_when_clientSecret_missing")
        void should_fail_validation_when_clientSecret_missing() {
            inputs.remove("clientSecret");
            setInputs();
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }

        @Test
        @DisplayName("should_fail_validation_when_connectTimeout_negative")
        void should_fail_validation_when_connectTimeout_negative() {
            inputs.put("connectTimeout", -1);
            setInputs();
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }

        @Test
        @DisplayName("should_pass_validation_when_connectTimeout_null")
        void should_pass_validation_when_connectTimeout_null() throws Exception {
            setInputs();
            connector.validateInputParameters();
        }

        @Test
        @DisplayName("should_fail_validation_when_delegated_mode_without_refreshToken")
        void should_fail_validation_when_delegated_mode_without_refreshToken() {
            var delegatedConnector = new TestableDelegatedConnector();
            delegatedConnector.setInputParameters(inputs);
            assertThatThrownBy(() -> delegatedConnector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }

        @Test
        @DisplayName("should_pass_validation_when_delegated_mode_with_refreshToken")
        void should_pass_validation_when_delegated_mode_with_refreshToken() throws Exception {
            inputs.put("refreshToken", "test-refresh-token");
            var delegatedConnector = new TestableDelegatedConnector();
            delegatedConnector.setInputParameters(inputs);
            delegatedConnector.validateInputParameters();
        }
    }

    @Nested
    @DisplayName("Phase 4: Disconnect")
    class Disconnect {

        @Test
        @DisplayName("should_disconnect_gracefully_when_client_present")
        void should_disconnect_gracefully_when_client_present() throws Exception {
            MsTeamsGraphClient mockClient = mock(MsTeamsGraphClient.class);
            connector.setGraphClient(mockClient);
            connector.disconnect();
            verify(mockClient).close();
        }

        @Test
        @DisplayName("should_handle_null_client_on_disconnect")
        void should_handle_null_client_on_disconnect() throws Exception {
            connector.disconnect();
            // no exception
        }

        @Test
        @DisplayName("should_handle_close_error_gracefully")
        void should_handle_close_error_gracefully() throws Exception {
            MsTeamsGraphClient mockClient = mock(MsTeamsGraphClient.class);
            doThrow(new RuntimeException("close error")).when(mockClient).close();
            connector.setGraphClient(mockClient);
            connector.disconnect();
            // no exception propagated
        }
    }

    // Testable concrete implementation for app-only mode
    static class TestableGraphConnector extends AbstractGraphConnector {
        @Override
        protected void validateOperationParameters(List<String> errors) {
            // no-op
        }

        @Override
        protected void executeBusinessLogic() throws ConnectorException {
            setSuccessOutputs();
        }
    }

    // Testable concrete implementation for delegated mode
    static class TestableDelegatedConnector extends AbstractGraphConnector {
        @Override
        protected String getAuthMode() {
            return "delegated";
        }

        @Override
        protected void validateOperationParameters(List<String> errors) {
            // no-op
        }

        @Override
        protected void executeBusinessLogic() throws ConnectorException {
            setSuccessOutputs();
        }
    }
}
