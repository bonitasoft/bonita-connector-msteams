package com.bonitasoft.connectors.msteams.common;

import static com.github.tomakehurst.wiremock.client.WireMock.*;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.github.tomakehurst.wiremock.junit5.WireMockRuntimeInfo;
import com.github.tomakehurst.wiremock.junit5.WireMockTest;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@WireMockTest
@DisplayName("MsTeamsWebhookClient")
class MsTeamsWebhookClientTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Test
    void should_send_payload_successfully(WireMockRuntimeInfo wm) {
        stubFor(post("/webhook").willReturn(ok("1")));

        MsTeamsWebhookClient client = new MsTeamsWebhookClient(5000, 0);
        ObjectNode payload = MAPPER.createObjectNode().put("text", "Hello");
        client.send(wm.getHttpBaseUrl() + "/webhook", payload);

        verify(postRequestedFor(urlEqualTo("/webhook"))
                .withHeader("Content-Type", equalTo("application/json")));
    }

    @Test
    void should_throw_on_4xx_error(WireMockRuntimeInfo wm) {
        stubFor(post("/webhook").willReturn(badRequest().withBody("Bad request")));

        MsTeamsWebhookClient client = new MsTeamsWebhookClient(5000, 0);
        ObjectNode payload = MAPPER.createObjectNode().put("text", "Hello");

        assertThatThrownBy(() -> client.send(wm.getHttpBaseUrl() + "/webhook", payload))
                .isInstanceOf(MsTeamsException.ExecutionException.class)
                .hasMessageContaining("400");
    }

    @Test
    void should_throw_on_unexpected_200_body(WireMockRuntimeInfo wm) {
        stubFor(post("/webhook").willReturn(ok("unexpected")));

        MsTeamsWebhookClient client = new MsTeamsWebhookClient(5000, 0);
        ObjectNode payload = MAPPER.createObjectNode().put("text", "Hello");

        assertThatThrownBy(() -> client.send(wm.getHttpBaseUrl() + "/webhook", payload))
                .isInstanceOf(MsTeamsException.class)
                .hasMessageContaining("unexpected response");
    }

    @Test
    void should_close_without_error() {
        MsTeamsWebhookClient client = new MsTeamsWebhookClient(5000, 0);
        client.close(); // no exception
    }
}
