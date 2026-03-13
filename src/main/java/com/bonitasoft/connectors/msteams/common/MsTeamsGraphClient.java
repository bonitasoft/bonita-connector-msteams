package com.bonitasoft.connectors.msteams.common;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

/**
 * HTTP client for Microsoft Graph API.
 * Uses java.net.http.HttpClient (zero external HTTP deps).
 */
@Slf4j
public class MsTeamsGraphClient implements AutoCloseable {

    private static final String DEFAULT_GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
    private static final ObjectMapper MAPPER = new ObjectMapper();

    private final String graphBaseUrl;
    private final HttpClient httpClient;
    private final TokenManager tokenManager;
    private final RetryPolicy retryPolicy;
    private String accessToken;

    public MsTeamsGraphClient(TokenManager tokenManager, int connectTimeoutMs, int maxRetries) {
        this(DEFAULT_GRAPH_BASE_URL, tokenManager, connectTimeoutMs, maxRetries);
    }

    public MsTeamsGraphClient(String baseUrl, TokenManager tokenManager, int connectTimeoutMs, int maxRetries) {
        this.graphBaseUrl = baseUrl.endsWith("/") ? baseUrl.substring(0, baseUrl.length() - 1) : baseUrl;
        this.tokenManager = tokenManager;
        this.httpClient = HttpClient.newBuilder()
                .connectTimeout(Duration.ofMillis(connectTimeoutMs))
                .build();
        this.retryPolicy = new RetryPolicy(maxRetries);
    }

    /**
     * Testing constructor: accepts a custom base URL and pre-set access token (no TokenManager needed).
     *
     * @param baseUrl       custom base URL (e.g., WireMock URL)
     * @param accessToken   pre-set access token
     * @param connectTimeoutMs connection timeout in milliseconds
     * @param maxRetries    max retry attempts
     */
    public MsTeamsGraphClient(String baseUrl, String accessToken, int connectTimeoutMs, int maxRetries) {
        this.graphBaseUrl = baseUrl.endsWith("/") ? baseUrl.substring(0, baseUrl.length() - 1) : baseUrl;
        this.tokenManager = null;
        this.accessToken = accessToken;
        this.httpClient = HttpClient.newBuilder()
                .connectTimeout(Duration.ofMillis(connectTimeoutMs))
                .build();
        this.retryPolicy = new RetryPolicy(maxRetries);
    }

    public void authenticateAppOnly() {
        this.accessToken = tokenManager.getAppOnlyToken();
    }

    public void authenticateDelegated(String refreshToken) {
        this.accessToken = tokenManager.getDelegatedToken(refreshToken);
    }

    /** GET request to Graph API, returns parsed JSON. */
    public JsonNode get(String path) {
        return retryPolicy.execute(() -> executeRequest("GET", path, null));
    }

    /** POST request to Graph API with JSON body. */
    public JsonNode post(String path, Object body) {
        return retryPolicy.execute(() -> executeRequest("POST", path, body));
    }

    /** PUT request to Graph API with body bytes. */
    public JsonNode put(String path, byte[] body, String contentType) {
        return retryPolicy.execute(() -> {
            HttpRequest request = HttpRequest.newBuilder()
                    .uri(URI.create(graphBaseUrl + path))
                    .header("Authorization", "Bearer " + accessToken)
                    .header("Content-Type", contentType)
                    .PUT(HttpRequest.BodyPublishers.ofByteArray(body))
                    .build();
            return handleResponse(httpClient.send(request, HttpResponse.BodyHandlers.ofString()));
        });
    }

    /** DELETE request to Graph API. */
    public void delete(String path) {
        retryPolicy.execute(() -> {
            executeRequest("DELETE", path, null);
            return null;
        });
    }

    /** GET with pagination support - collects all pages. */
    public List<JsonNode> getWithPagination(String path) {
        List<JsonNode> allItems = new ArrayList<>();
        String nextLink = graphBaseUrl + path;

        while (nextLink != null) {
            final String url = nextLink;
            JsonNode response = retryPolicy.execute(() -> {
                HttpRequest request = HttpRequest.newBuilder()
                        .uri(URI.create(url))
                        .header("Authorization", "Bearer " + accessToken)
                        .header("Accept", "application/json")
                        .GET()
                        .build();
                return handleResponse(httpClient.send(request, HttpResponse.BodyHandlers.ofString()));
            });

            JsonNode value = response.get("value");
            if (value != null && value.isArray()) {
                value.forEach(allItems::add);
            }

            JsonNode nextLinkNode = response.get("@odata.nextLink");
            nextLink = (nextLinkNode != null && !nextLinkNode.isNull()) ? nextLinkNode.asText() : null;
        }

        return allItems;
    }

    private JsonNode executeRequest(String method, String path, Object body) throws IOException, InterruptedException {
        HttpRequest.Builder builder = HttpRequest.newBuilder()
                .uri(URI.create(graphBaseUrl + path))
                .header("Authorization", "Bearer " + accessToken)
                .header("Accept", "application/json");

        if (body != null) {
            String jsonBody = MAPPER.writeValueAsString(body);
            builder.header("Content-Type", "application/json")
                    .method(method, HttpRequest.BodyPublishers.ofString(jsonBody));
        } else {
            builder.method(method, HttpRequest.BodyPublishers.noBody());
        }

        HttpResponse<String> response = httpClient.send(builder.build(), HttpResponse.BodyHandlers.ofString());
        return handleResponse(response);
    }

    private JsonNode handleResponse(HttpResponse<String> response) {
        int statusCode = response.statusCode();
        String responseBody = response.body();

        if (statusCode >= 200 && statusCode < 300) {
            if (responseBody == null || responseBody.isBlank()) {
                return MAPPER.createObjectNode().put("status", statusCode);
            }
            try {
                return MAPPER.readTree(responseBody);
            } catch (IOException e) {
                throw new MsTeamsException("Failed to parse Graph API response", e);
            }
        }

        // Parse error from Graph API
        String errorMessage = parseGraphError(responseBody, statusCode);
        throw new MsTeamsException.ExecutionException(errorMessage, statusCode);
    }

    private String parseGraphError(String responseBody, int statusCode) {
        try {
            JsonNode errorNode = MAPPER.readTree(responseBody);
            JsonNode error = errorNode.get("error");
            if (error != null) {
                String code = error.has("code") ? error.get("code").asText() : "Unknown";
                String message = error.has("message") ? error.get("message").asText() : "No message";
                return String.format("Graph API error %d [%s]: %s", statusCode, code, message);
            }
        } catch (Exception ignored) {
            // Fall through to default
        }
        return String.format("Graph API error %d: %s", statusCode,
                responseBody != null ? responseBody.substring(0, Math.min(responseBody.length(), 500)) : "empty response");
    }

    @Override
    public void close() {
        log.debug("MsTeamsGraphClient closed");
    }
}
