package com.bonitasoft.connectors.msteams.integration;

import static com.github.tomakehurst.wiremock.client.WireMock.*;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import com.bonitasoft.connectors.msteams.common.MsTeamsWebhookClient;
import com.bonitasoft.connectors.msteams.messaging.SendChannelMessageConnector;
import com.bonitasoft.connectors.msteams.messaging.SendAdaptiveCardConnector;
import com.github.tomakehurst.wiremock.junit5.WireMockRuntimeInfo;
import com.github.tomakehurst.wiremock.junit5.WireMockTest;
import org.bonitasoft.engine.connector.ConnectorException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;

@Tag("integration")
@WireMockTest
class WebhookConnectorIT {

    @Nested
    @DisplayName("SendChannelMessage via Webhook")
    class SendChannelMessageTests {

        @Test
        @DisplayName("Should send message successfully when webhook returns 200 + '1'")
        void should_send_message_when_webhook_returns_success(WireMockRuntimeInfo wmInfo) throws Exception {
            stubFor(post(urlEqualTo("/webhook/test"))
                    .willReturn(aResponse().withStatus(200).withBody("1")));

            var connector = new SendChannelMessageConnector();
            Map<String, Object> inputs = new HashMap<>();
            inputs.put("webhookUrl", wmInfo.getHttpBaseUrl() + "/webhook/test");
            inputs.put("message", "Hello from Bonita!");
            inputs.put("title", "Test Message");
            connector.setInputParameters(inputs);

            var client = new MsTeamsWebhookClient(30000, 0);
            connector.setWebhookClient(client);

            connector.executeForTest();

            assertThat(connector.getOutputs().get("success")).isEqualTo(true);
            verify(postRequestedFor(urlEqualTo("/webhook/test"))
                    .withRequestBody(containing("Hello from Bonita!")));
        }

        @Test
        @DisplayName("Should handle webhook error 400")
        void should_handle_webhook_error(WireMockRuntimeInfo wmInfo) {
            stubFor(post(urlEqualTo("/webhook/error"))
                    .willReturn(aResponse().withStatus(400).withBody("Bad Request")));

            var connector = new SendChannelMessageConnector();
            Map<String, Object> inputs = new HashMap<>();
            inputs.put("webhookUrl", wmInfo.getHttpBaseUrl() + "/webhook/error");
            inputs.put("message", "Test");
            connector.setInputParameters(inputs);

            var client = new MsTeamsWebhookClient(30000, 0);
            connector.setWebhookClient(client);

            assertThatThrownBy(() -> connector.executeForTest())
                    .isInstanceOf(ConnectorException.class);
            assertThat(connector.getOutputs().get("success")).isEqualTo(false);
        }

        @Test
        @DisplayName("Should handle webhook unexpected body")
        void should_handle_unexpected_body(WireMockRuntimeInfo wmInfo) {
            stubFor(post(urlEqualTo("/webhook/odd"))
                    .willReturn(aResponse().withStatus(200).withBody("unexpected")));

            var connector = new SendChannelMessageConnector();
            Map<String, Object> inputs = new HashMap<>();
            inputs.put("webhookUrl", wmInfo.getHttpBaseUrl() + "/webhook/odd");
            inputs.put("message", "Test");
            connector.setInputParameters(inputs);

            var client = new MsTeamsWebhookClient(30000, 0);
            connector.setWebhookClient(client);

            assertThatThrownBy(() -> connector.executeForTest())
                    .isInstanceOf(ConnectorException.class);
        }
    }

    @Nested
    @DisplayName("SendAdaptiveCard via Webhook")
    class SendAdaptiveCardTests {

        @Test
        @DisplayName("Should send adaptive card successfully")
        void should_send_adaptive_card(WireMockRuntimeInfo wmInfo) throws Exception {
            stubFor(post(urlEqualTo("/webhook/card"))
                    .willReturn(aResponse().withStatus(200).withBody("1")));

            var connector = new SendAdaptiveCardConnector();
            Map<String, Object> inputs = new HashMap<>();
            inputs.put("webhookUrl", wmInfo.getHttpBaseUrl() + "/webhook/card");
            inputs.put("cardTitle", "Deployment Alert");
            inputs.put("cardBody", "v2.1.0 deployed to production");
            inputs.put("themeColor", "00FF00");
            connector.setInputParameters(inputs);

            var client = new MsTeamsWebhookClient(30000, 0);
            connector.setWebhookClient(client);

            connector.executeForTest();

            assertThat(connector.getOutputs().get("success")).isEqualTo(true);
            verify(postRequestedFor(urlEqualTo("/webhook/card"))
                    .withRequestBody(containing("AdaptiveCard")));
        }

        @Test
        @DisplayName("Should send raw JSON card")
        void should_send_raw_json_card(WireMockRuntimeInfo wmInfo) throws Exception {
            stubFor(post(urlEqualTo("/webhook/raw"))
                    .willReturn(aResponse().withStatus(200).withBody("1")));

            var connector = new SendAdaptiveCardConnector();
            Map<String, Object> inputs = new HashMap<>();
            inputs.put("webhookUrl", wmInfo.getHttpBaseUrl() + "/webhook/raw");
            inputs.put("advancedCardJson", "{\"type\":\"AdaptiveCard\",\"version\":\"1.4\",\"body\":[{\"type\":\"TextBlock\",\"text\":\"Custom\"}]}");
            connector.setInputParameters(inputs);

            var client = new MsTeamsWebhookClient(30000, 0);
            connector.setWebhookClient(client);

            connector.executeForTest();

            assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        }
    }
}
