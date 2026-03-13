package com.bonitasoft.connectors.msteams.messaging;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.*;

import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import com.bonitasoft.connectors.msteams.common.MsTeamsWebhookClient;
import com.fasterxml.jackson.databind.node.ObjectNode;
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
import java.util.Map;

@ExtendWith(MockitoExtension.class)
class SendChannelMessageConnectorTest {

    @Mock
    private MsTeamsWebhookClient webhookClient;

    private SendChannelMessageConnector connector;
    private Map<String, Object> inputs;

    @BeforeEach
    void setUp() {
        connector = new SendChannelMessageConnector();
        inputs = new HashMap<>();
        inputs.put("webhookUrl", "https://outlook.office.com/webhook/test");
        inputs.put("message", "Hello Teams!");
        connector.setInputParameters(inputs);
    }

    @Nested
    @DisplayName("Phase 1: Validation")
    class ValidationTests {

        @Test
        void should_pass_validation_with_valid_inputs() throws ConnectorValidationException {
            connector.validateInputParameters();
        }

        @Test
        void should_fail_when_webhook_url_is_missing() {
            var c = new SendChannelMessageConnector();
            var incompleteInputs = new HashMap<>(inputs);
            incompleteInputs.remove("webhookUrl");
            c.setInputParameters(incompleteInputs);
            assertThatThrownBy(() -> c.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }

        @Test
        void should_fail_when_webhook_url_is_not_https() {
            inputs.put("webhookUrl", "http://insecure.com/webhook");
            connector.setInputParameters(inputs);
            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }

        @Test
        void should_fail_when_message_is_missing() {
            var c = new SendChannelMessageConnector();
            var incompleteInputs = new HashMap<>(inputs);
            incompleteInputs.remove("message");
            c.setInputParameters(incompleteInputs);
            assertThatThrownBy(() -> c.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class);
        }
    }

    @Nested
    @DisplayName("Phase 3: Execute")
    class ExecutionTests {

        @BeforeEach
        void setUpClient() {
            connector.setWebhookClient(webhookClient);
        }

        @Test
        void should_send_message_successfully() throws ConnectorException {
            connector.executeBusinessLogic();

            verify(webhookClient).send(eq("https://outlook.office.com/webhook/test"), any(ObjectNode.class));
            assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        }

        @Test
        void should_include_title_when_provided() throws ConnectorException {
            inputs.put("title", "Important");
            connector.setInputParameters(inputs);

            connector.executeBusinessLogic();

            verify(webhookClient).send(any(), any(ObjectNode.class));
        }

        @Test
        void should_handle_webhook_error() {
            doThrow(new MsTeamsException.ExecutionException("Webhook error", 400))
                    .when(webhookClient).send(any(), any());

            assertThatThrownBy(() -> connector.executeBusinessLogic())
                    .isInstanceOf(ConnectorException.class);
            assertThat(connector.getOutputs().get("success")).isEqualTo(false);
        }
    }

    @Nested
    @DisplayName("Phase 4: Disconnect")
    class DisconnectTests {

        @Test
        void should_close_webhook_client() throws ConnectorException {
            connector.setWebhookClient(webhookClient);
            connector.disconnect();
            verify(webhookClient).close();
        }

        @Test
        void should_handle_null_client() throws ConnectorException {
            connector.disconnect(); // Should not throw
        }
    }
}
