package com.bonitasoft.connectors.msteams.common;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.Duration;

/**
 * HTTP client for MS Teams Incoming Webhooks.
 * Webhooks use simple POST with JSON body, no auth header needed.
 * Success: HTTP 200 + body "1".
 */
@Slf4j
public class MsTeamsWebhookClient implements AutoCloseable {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private final HttpClient httpClient;
    private final RetryPolicy retryPolicy;

    public MsTeamsWebhookClient(int connectTimeoutMs, int maxRetries) {
        this.httpClient = HttpClient.newBuilder()
                .connectTimeout(Duration.ofMillis(connectTimeoutMs))
                .build();
        this.retryPolicy = new RetryPolicy(maxRetries);
    }

    /**
     * Send a message card to a webhook URL.
     *
     * @param webhookUrl the incoming webhook URL
     * @param payload    JSON payload (MessageCard or AdaptiveCard)
     */
    public void send(String webhookUrl, ObjectNode payload) {
        retryPolicy.execute(() -> {
            doSend(webhookUrl, payload);
            return null;
        });
    }

    private void doSend(String webhookUrl, ObjectNode payload) throws IOException, InterruptedException {
        String jsonBody = MAPPER.writeValueAsString(payload);
        log.debug("Sending webhook payload to {}: {} chars", webhookUrl, jsonBody.length());

        HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(webhookUrl))
                .header("Content-Type", "application/json")
                .POST(HttpRequest.BodyPublishers.ofString(jsonBody))
                .build();

        HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());
        int statusCode = response.statusCode();
        String body = response.body();

        if (statusCode == 200 && "1".equals(body)) {
            log.debug("Webhook message sent successfully");
            return;
        }

        if (statusCode >= 400) {
            throw new MsTeamsException.ExecutionException(
                    String.format("Webhook error %d: %s", statusCode, body), statusCode);
        }

        // 200 but body != "1" — possible error
        log.warn("Webhook returned HTTP {} with unexpected body: {}", statusCode, body);
        throw new MsTeamsException("Webhook returned unexpected response: " + body);
    }

    @Override
    public void close() {
        log.debug("MsTeamsWebhookClient closed");
    }
}
