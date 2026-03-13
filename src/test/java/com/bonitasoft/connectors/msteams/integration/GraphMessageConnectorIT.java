package com.bonitasoft.connectors.msteams.integration;

import static com.github.tomakehurst.wiremock.client.WireMock.*;
import static org.assertj.core.api.Assertions.assertThat;

import com.bonitasoft.connectors.msteams.common.MsTeamsGraphClient;
import com.bonitasoft.connectors.msteams.messaging.ReplyMessageConnector;
import com.bonitasoft.connectors.msteams.messaging.SendChatMessageConnector;
import com.github.tomakehurst.wiremock.junit5.WireMockRuntimeInfo;
import com.github.tomakehurst.wiremock.junit5.WireMockTest;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;

@Tag("integration")
@WireMockTest
class GraphMessageConnectorIT {

    private MsTeamsGraphClient graphClient;

    @BeforeEach
    void setUp(WireMockRuntimeInfo wmInfo) {
        graphClient = new MsTeamsGraphClient(wmInfo.getHttpBaseUrl(), "fake-token", 30000, 0);
    }

    @Test
    @DisplayName("Should send chat message")
    void should_send_chat_message() throws Exception {
        stubFor(post(urlPathEqualTo("/chats/chat-1/messages"))
                .willReturn(okJson("{\"id\":\"msg-1\",\"createdDateTime\":\"2025-01-15T10:00:00Z\"}")));

        var connector = new SendChatMessageConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("refreshToken", "rt");
        inputs.put("chatId", "chat-1");
        inputs.put("message", "Hello via Graph");
        connector.setInputParameters(inputs);
        connector.setGraphClient(graphClient);

        connector.executeForTest();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("messageId")).isEqualTo("msg-1");
    }

    @Test
    @DisplayName("Should reply to channel message")
    void should_reply_to_message() throws Exception {
        stubFor(post(urlPathEqualTo("/teams/team-1/channels/ch-1/messages/msg-1/replies"))
                .willReturn(okJson("{\"id\":\"reply-1\",\"createdDateTime\":\"2025-01-15T10:05:00Z\"}")));

        var connector = new ReplyMessageConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("refreshToken", "rt");
        inputs.put("teamId", "team-1");
        inputs.put("channelId", "ch-1");
        inputs.put("messageId", "msg-1");
        inputs.put("replyContent", "Thanks!");
        connector.setInputParameters(inputs);
        connector.setGraphClient(graphClient);

        connector.executeForTest();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("replyId")).isEqualTo("reply-1");
    }
}
