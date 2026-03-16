package com.bonitasoft.connectors.msteams;

import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.Duration;
import java.util.concurrent.ThreadLocalRandom;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class MSTeamsClient implements AutoCloseable {

    private static final String TOKEN_URL_TEMPLATE = "https://login.microsoftonline.com/%s/oauth2/v2.0/token";
    private static final String GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
    private static final int MAX_RETRIES = 5;
    private static final long BASE_DELAY_MS = 1_000;
    private static final long MAX_DELAY_MS = 64_000;
    private static final double JITTER_FACTOR = 0.2;
    private static final ObjectMapper MAPPER = new ObjectMapper();

    private final MSTeamsConfiguration configuration;
    private final HttpClient httpClient;
    private String accessToken;

    public MSTeamsClient(MSTeamsConfiguration configuration) {
        this(configuration, HttpClient.newBuilder()
                .connectTimeout(Duration.ofMillis(configuration.connectTimeout()))
                .build());
    }

    /** Package-private constructor for testing with a mock HttpClient. */
    MSTeamsClient(MSTeamsConfiguration configuration, HttpClient httpClient) {
        this.configuration = configuration;
        this.httpClient = httpClient;
    }

    public void authenticate() {
        log.info("Authenticating with Microsoft Graph API (tenant: {})", configuration.tenantId());
        String tokenUrl = String.format(TOKEN_URL_TEMPLATE, configuration.tenantId());
        String body = "grant_type=client_credentials"
                + "&client_id=" + configuration.clientId()
                + "&client_secret=" + configuration.clientSecret()
                + "&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default";
        try {
            HttpRequest request = HttpRequest.newBuilder()
                    .uri(URI.create(tokenUrl))
                    .header("Content-Type", "application/x-www-form-urlencoded")
                    .POST(HttpRequest.BodyPublishers.ofString(body))
                    .timeout(Duration.ofMillis(configuration.readTimeout()))
                    .build();
            HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());
            if (response.statusCode() != 200) {
                throw new MSTeamsException.ConnectionException(
                        "Authentication failed (HTTP " + response.statusCode() + "): " + response.body(), null);
            }
            JsonNode json = MAPPER.readTree(response.body());
            this.accessToken = json.get("access_token").asText();
            log.info("Successfully authenticated with Microsoft Graph API");
        } catch (MSTeamsException e) {
            throw e;
        } catch (Exception e) {
            throw new MSTeamsException.ConnectionException("Failed to authenticate: " + e.getMessage(), e);
        }
    }

    public JsonNode get(String path) {
        return executeWithRetry(() -> doRequest("GET", path, null), "GET " + path);
    }

    public JsonNode post(String path, String body) {
        return executeWithRetry(() -> doRequest("POST", path, body), "POST " + path);
    }

    public JsonNode delete(String path) {
        return executeWithRetry(() -> doRequest("DELETE", path, null), "DELETE " + path);
    }

    public JsonNode put(String path, byte[] body, String contentType) {
        return executeWithRetry(() -> doPutRequest(path, body, contentType), "PUT " + path);
    }

    public ObjectMapper getObjectMapper() { return MAPPER; }

    public MSTeamsConfiguration getConfiguration() { return configuration; }

    @Override
    public void close() { log.debug("Closing MS Teams client"); }

    private JsonNode doPutRequest(String path, byte[] body, String contentType) throws IOException, InterruptedException {
        String url = GRAPH_BASE_URL + path;
        HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + accessToken)
                .header("Content-Type", contentType)
                .timeout(Duration.ofMillis(configuration.readTimeout()))
                .PUT(HttpRequest.BodyPublishers.ofByteArray(body))
                .build();
        HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());
        int statusCode = response.statusCode();
        if (statusCode == 204) return null;
        if (statusCode >= 200 && statusCode < 300) return MAPPER.readTree(response.body());
        String errorCode = "unknown";
        String errorMessage = response.body();
        try {
            JsonNode errorJson = MAPPER.readTree(response.body());
            JsonNode error = errorJson.get("error");
            if (error != null) {
                errorCode = error.has("code") ? error.get("code").asText() : "unknown";
                errorMessage = error.has("message") ? error.get("message").asText() : response.body();
            }
        } catch (Exception ignored) {}
        throw MSTeamsException.fromGraphError(statusCode, errorCode, errorMessage);
    }

    private JsonNode doRequest(String method, String path, String body) throws IOException, InterruptedException {
        String url = GRAPH_BASE_URL + path;
        HttpRequest.Builder builder = HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + accessToken)
                .header("Content-Type", "application/json")
                .timeout(Duration.ofMillis(configuration.readTimeout()));
        switch (method) {
            case "GET" -> builder.GET();
            case "POST" -> builder.POST(HttpRequest.BodyPublishers.ofString(body != null ? body : ""));
            case "DELETE" -> builder.DELETE();
            default -> throw new IllegalArgumentException("Unsupported HTTP method: " + method);
        }
        HttpResponse<String> response = httpClient.send(builder.build(), HttpResponse.BodyHandlers.ofString());
        int statusCode = response.statusCode();
        if (statusCode == 204) return null;
        if (statusCode >= 200 && statusCode < 300) return MAPPER.readTree(response.body());
        String errorCode = "unknown";
        String errorMessage = response.body();
        try {
            JsonNode errorJson = MAPPER.readTree(response.body());
            JsonNode error = errorJson.get("error");
            if (error != null) {
                errorCode = error.has("code") ? error.get("code").asText() : "unknown";
                errorMessage = error.has("message") ? error.get("message").asText() : response.body();
            }
        } catch (Exception ignored) {}
        throw MSTeamsException.fromGraphError(statusCode, errorCode, errorMessage);
    }

    private <T> T executeWithRetry(RetryableOperation<T> op, String desc) {
        int attempt = 0;
        MSTeamsException lastEx = null;
        while (attempt <= MAX_RETRIES) {
            try {
                if (attempt > 0) log.debug("Retry {}/{} for: {}", attempt, MAX_RETRIES, desc);
                return op.call();
            } catch (MSTeamsException e) {
                if (!e.isRetryable()) throw e;
                lastEx = e;
                attempt++;
                if (attempt > MAX_RETRIES) break;
                long delay = calcDelay(attempt);
                log.warn("Retryable error (attempt {}/{}), waiting {}ms: {}", attempt, MAX_RETRIES, delay, e.getMessage());
                doSleep(delay);
            } catch (Exception e) {
                throw new MSTeamsException("Unexpected error during " + desc + ": " + e.getMessage(), e);
            }
        }
        throw new MSTeamsException(String.format("Max retries (%d) exceeded for %s: %s", MAX_RETRIES, desc, lastEx.getMessage()), lastEx);
    }

    private long calcDelay(int attempt) {
        long base = BASE_DELAY_MS * (1L << (attempt - 1));
        long capped = Math.min(base, MAX_DELAY_MS);
        return capped + ThreadLocalRandom.current().nextLong((long) (capped * JITTER_FACTOR));
    }

    /** Package-private for test override to avoid real sleeps. */
    void doSleep(long ms) {
        try { Thread.sleep(ms); }
        catch (InterruptedException e) { Thread.currentThread().interrupt(); throw new MSTeamsException("Interrupted", e); }
    }

    @FunctionalInterface
    interface RetryableOperation<T> { T call() throws Exception; }
}
