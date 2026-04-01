package com.bonitasoft.connectors.msteams.common;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.Mockito.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.bonitasoft.engine.connector.ConnectorException;
import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

@DisplayName("AbstractWebhookConnector")
class AbstractWebhookConnectorTest {

    private TestableWebhookConnector connector;
    private final Map<String, Object> inputs = new HashMap<>();

    @BeforeEach
    void setUp() {
        connector = new TestableWebhookConnector();
        inputs.clear();
        inputs.put("webhookUrl", "https://outlook.office.com/webhook/test");
    }

    private void setInputs() {
        connector.setInputParameters(inputs);
    }

    @Nested
    @DisplayName("Phase 1: Validation")
    class Validation {

        @Test
        void should_pass_validation_when_webhook_url_valid() throws Exception {
            setInputs();
            connector.validateInputParameters();
        }

        @Test
        void should_fail_validation_when_webhook_url_missing() {
            inputs.remove("webhookUrl");
            setInputs();
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }

        @Test
        void should_fail_validation_when_webhook_url_not_https() {
            inputs.put("webhookUrl", "http://insecure.example.com/webhook");
            setInputs();
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }

        @Test
        void should_fail_validation_when_connectTimeout_negative() {
            inputs.put("connectTimeout", -5);
            setInputs();
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }
    }

    @Nested
    @DisplayName("Phase 2: Connect")
    class Connect {

        @Test
        void should_create_webhook_client_on_connect() throws Exception {
            setInputs();
            connector.connect();
            assertThat(connector.webhookClient).isNotNull();
        }

        @Test
        void should_use_custom_timeout() throws Exception {
            inputs.put("connectTimeout", 5000);
            setInputs();
            connector.connect();
            assertThat(connector.webhookClient).isNotNull();
        }
    }

    @Nested
    @DisplayName("Phase 4: Disconnect")
    class Disconnect {

        @Test
        void should_disconnect_gracefully() throws Exception {
            MsTeamsWebhookClient mockClient = mock(MsTeamsWebhookClient.class);
            connector.setWebhookClient(mockClient);
            connector.disconnect();
            verify(mockClient).close();
        }

        @Test
        void should_handle_null_client_on_disconnect() throws Exception {
            connector.disconnect();
        }

        @Test
        void should_handle_close_error_gracefully() throws Exception {
            MsTeamsWebhookClient mockClient = mock(MsTeamsWebhookClient.class);
            doThrow(new RuntimeException("close error")).when(mockClient).close();
            connector.setWebhookClient(mockClient);
            connector.disconnect();
        }
    }

    static class TestableWebhookConnector extends AbstractWebhookConnector {
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
