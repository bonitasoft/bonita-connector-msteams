package com.bonitasoft.connectors.msteams.messaging;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.*;

import com.bonitasoft.connectors.msteams.common.MsTeamsWebhookClient;
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
class SendAdaptiveCardConnectorTest {

    @Mock
    private MsTeamsWebhookClient webhookClient;

    private SendAdaptiveCardConnector connector;
    private Map<String, Object> inputs;

    @BeforeEach
    void setUp() {
        connector = new SendAdaptiveCardConnector();
        inputs = new HashMap<>();
        inputs.put("webhookUrl", "https://outlook.office.com/webhook/test");
        inputs.put("cardTitle", "Test Card");
        inputs.put("cardBody", "Test Body");
        connector.setInputParameters(inputs);
    }

    @Nested
    @DisplayName("Phase 1: Validation")
    class ValidationTests {

        @Test
        void should_pass_with_card_title() throws ConnectorValidationException {
            connector.validateInputParameters();
        }

        @Test
        void should_pass_with_advanced_json() throws ConnectorValidationException {
            inputs.remove("cardTitle");
            inputs.remove("cardBody");
            inputs.put("advancedCardJson", "{\"type\":\"AdaptiveCard\"}");
            connector.setInputParameters(inputs);
            connector.validateInputParameters();
        }

        @Test
        void should_fail_when_no_card_content() {
            var c = new SendAdaptiveCardConnector();
            var incompleteInputs = new HashMap<>(inputs);
            incompleteInputs.remove("cardTitle");
            incompleteInputs.remove("cardBody");
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
        void should_send_card_via_builder() throws ConnectorException {
            connector.executeBusinessLogic();
            verify(webhookClient).send(any(), any());
            assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        }

        @Test
        void should_send_raw_json_card() throws ConnectorException {
            inputs.put("advancedCardJson", "{\"type\":\"AdaptiveCard\",\"version\":\"1.4\",\"body\":[]}");
            connector.setInputParameters(inputs);
            connector.executeBusinessLogic();
            verify(webhookClient).send(any(), any());
            assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        }

        @Test
        void should_send_card_with_facts() throws ConnectorException {
            inputs.put("factsJson", "[{\"key\":\"Status\",\"value\":\"OK\"}]");
            connector.setInputParameters(inputs);
            connector.executeBusinessLogic();
            verify(webhookClient).send(any(), any());
        }
    }
}
