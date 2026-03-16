package com.bonitasoft.connectors.msteams;

import java.io.IOException;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.any;
import static org.mockito.Mockito.lenient;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.reset;
import static org.mockito.Mockito.times;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

@DisplayName("MSTeamsClient")
@ExtendWith(MockitoExtension.class)
class MSTeamsClientTest {

    private static final MSTeamsConfiguration CONFIG =
            new MSTeamsConfiguration("tenant-123", "client-id", "client-secret", 5000, 10000, true);

    @Mock
    private HttpClient httpClient;

    @Mock
    private HttpResponse<String> httpResponse;

    private MSTeamsClient client;

    @SuppressWarnings("unchecked")
    private void mockResponse(int statusCode, String body) throws Exception {
        lenient().when(httpResponse.statusCode()).thenReturn(statusCode);
        lenient().when(httpResponse.body()).thenReturn(body);
        when(httpClient.send(any(HttpRequest.class), any(HttpResponse.BodyHandler.class)))
                .thenReturn(httpResponse);
    }

    @SuppressWarnings("unchecked")
    private void mockResponseSequence(int firstStatus, String firstBody, int secondStatus, String secondBody) throws Exception {
        HttpResponse<String> firstResponse = mock(HttpResponse.class);
        HttpResponse<String> secondResponse = mock(HttpResponse.class);
        when(firstResponse.statusCode()).thenReturn(firstStatus);
        when(firstResponse.body()).thenReturn(firstBody);
        when(secondResponse.statusCode()).thenReturn(secondStatus);
        when(secondResponse.body()).thenReturn(secondBody);
        when(httpClient.send(any(HttpRequest.class), any(HttpResponse.BodyHandler.class)))
                .thenReturn(firstResponse)
                .thenReturn(secondResponse);
    }

    @Nested
    @DisplayName("Constructor and basic methods")
    class Basic {

        @Test
        void should_create_client_with_valid_configuration() {
            var c = new MSTeamsClient(CONFIG);
            assertThat(c.getConfiguration()).isEqualTo(CONFIG);
            assertThat(c.getObjectMapper()).isNotNull();
        }

        @Test
        void should_create_client_with_injected_http_client() {
            var c = new MSTeamsClient(CONFIG, httpClient);
            assertThat(c.getConfiguration()).isEqualTo(CONFIG);
        }

        @Test
        void should_close_without_error() {
            var c = new MSTeamsClient(CONFIG);
            c.close();
        }

        @Test
        void should_return_shared_object_mapper() {
            var c1 = new MSTeamsClient(CONFIG);
            var c2 = new MSTeamsClient(CONFIG);
            assertThat(c1.getObjectMapper()).isSameAs(c2.getObjectMapper());
        }
    }

    @Nested
    @DisplayName("Authentication")
    class Authentication {

        @BeforeEach
        void setUp() {
            client = new MSTeamsClient(CONFIG, httpClient);
        }

        @Test
        void should_authenticate_successfully_when_token_returned() throws Exception {
            mockResponse(200, "{\"access_token\":\"eyJ0eXAiOiJKV1Q\",\"expires_in\":3600}");

            client.authenticate();

            // Verify auth request was sent — successful auth means we can now make API calls
            verify(httpClient).send(any(HttpRequest.class), any());
        }

        @Test
        void should_throw_connection_exception_when_auth_returns_error() throws Exception {
            mockResponse(401, "{\"error\":\"invalid_client\"}");

            assertThatThrownBy(() -> client.authenticate())
                    .isInstanceOf(MSTeamsException.ConnectionException.class)
                    .hasMessageContaining("HTTP 401");
        }

        @Test
        @SuppressWarnings("unchecked")
        void should_throw_connection_exception_when_network_fails() throws Exception {
            when(httpClient.send(any(HttpRequest.class), any(HttpResponse.BodyHandler.class)))
                    .thenThrow(new IOException("Connection refused"));

            assertThatThrownBy(() -> client.authenticate())
                    .isInstanceOf(MSTeamsException.ConnectionException.class)
                    .hasMessageContaining("Failed to authenticate");
        }

        @Test
        @SuppressWarnings("unchecked")
        void should_throw_connection_exception_when_interrupted() throws Exception {
            when(httpClient.send(any(HttpRequest.class), any(HttpResponse.BodyHandler.class)))
                    .thenThrow(new InterruptedException("interrupted"));

            assertThatThrownBy(() -> client.authenticate())
                    .isInstanceOf(MSTeamsException.ConnectionException.class)
                    .hasMessageContaining("Failed to authenticate");
        }
    }

    @Nested
    @DisplayName("GET requests")
    class GetRequests {

        @BeforeEach
        void setUp() throws Exception {
            client = new NoSleepMSTeamsClient(CONFIG, httpClient);
            // Authenticate first
            mockResponse(200, "{\"access_token\":\"test-token\",\"expires_in\":3600}");
            client.authenticate();
            reset(httpClient);
        }

        @Test
        void should_return_json_when_get_succeeds() throws Exception {
            mockResponse(200, "{\"id\":\"msg-123\",\"body\":{\"content\":\"hello\"}}");

            JsonNode result = client.get("/teams/t1/channels/c1/messages");

            assertThat(result.get("id").asText()).isEqualTo("msg-123");
        }

        @Test
        void should_return_null_when_get_returns_204() throws Exception {
            mockResponse(204, "");

            JsonNode result = client.get("/teams/t1");

            assertThat(result).isNull();
        }

        @Test
        void should_throw_when_get_returns_400_with_graph_error() throws Exception {
            mockResponse(400, "{\"error\":{\"code\":\"BadRequest\",\"message\":\"Invalid filter\"}}");

            assertThatThrownBy(() -> client.get("/teams/t1"))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("Invalid request");
        }

        @Test
        void should_throw_when_get_returns_404() throws Exception {
            mockResponse(404, "{\"error\":{\"code\":\"NotFound\",\"message\":\"Team not found\"}}");

            assertThatThrownBy(() -> client.get("/teams/nonexistent"))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("not found");
        }

        @Test
        void should_throw_when_get_returns_error_with_unparseable_body() throws Exception {
            mockResponse(500, "Internal Server Error - not JSON");

            assertThatThrownBy(() -> client.get("/teams/t1"))
                    .isInstanceOf(MSTeamsException.class);
        }

        @Test
        void should_throw_when_get_returns_error_without_error_object() throws Exception {
            mockResponse(400, "{\"status\":\"error\",\"detail\":\"something\"}");

            assertThatThrownBy(() -> client.get("/teams/t1"))
                    .isInstanceOf(MSTeamsException.class);
        }

        @Test
        void should_throw_when_get_returns_error_with_partial_error_object() throws Exception {
            mockResponse(400, "{\"error\":{}}");

            assertThatThrownBy(() -> client.get("/teams/t1"))
                    .isInstanceOf(MSTeamsException.class);
        }
    }

    @Nested
    @DisplayName("POST requests")
    class PostRequests {

        @BeforeEach
        void setUp() throws Exception {
            client = new NoSleepMSTeamsClient(CONFIG, httpClient);
            mockResponse(200, "{\"access_token\":\"test-token\",\"expires_in\":3600}");
            client.authenticate();
            reset(httpClient);
        }

        @Test
        void should_return_json_when_post_succeeds() throws Exception {
            mockResponse(201, "{\"id\":\"new-msg\",\"createdDateTime\":\"2026-03-14T10:00:00Z\"}");

            JsonNode result = client.post("/teams/t1/channels/c1/messages", "{\"body\":{\"content\":\"hi\"}}");

            assertThat(result.get("id").asText()).isEqualTo("new-msg");
        }

        @Test
        void should_handle_post_with_null_body() throws Exception {
            mockResponse(200, "{\"id\":\"test\"}");

            JsonNode result = client.post("/test/path", null);

            assertThat(result).isNotNull();
        }

        @Test
        void should_throw_when_post_returns_403() throws Exception {
            mockResponse(403, "{\"error\":{\"code\":\"Forbidden\",\"message\":\"Insufficient privileges\"}}");

            assertThatThrownBy(() -> client.post("/teams/t1/channels/c1/messages", "{}"))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("Permission denied");
        }
    }

    @Nested
    @DisplayName("DELETE requests")
    class DeleteRequests {

        @BeforeEach
        void setUp() throws Exception {
            client = new NoSleepMSTeamsClient(CONFIG, httpClient);
            mockResponse(200, "{\"access_token\":\"test-token\",\"expires_in\":3600}");
            client.authenticate();
            reset(httpClient);
        }

        @Test
        void should_return_null_when_delete_returns_204() throws Exception {
            mockResponse(204, "");

            JsonNode result = client.delete("/teams/t1/channels/c1");

            assertThat(result).isNull();
        }

        @Test
        void should_throw_when_delete_returns_409() throws Exception {
            mockResponse(409, "{\"error\":{\"code\":\"Conflict\",\"message\":\"Resource in use\"}}");

            assertThatThrownBy(() -> client.delete("/teams/t1/channels/c1"))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("Conflict");
        }
    }

    @Nested
    @DisplayName("PUT requests (binary upload)")
    class PutRequests {

        @BeforeEach
        void setUp() throws Exception {
            client = new NoSleepMSTeamsClient(CONFIG, httpClient);
            mockResponse(200, "{\"access_token\":\"test-token\",\"expires_in\":3600}");
            client.authenticate();
            reset(httpClient);
        }

        @Test
        void should_return_json_when_put_succeeds() throws Exception {
            mockResponse(200, "{\"id\":\"file-123\",\"name\":\"test.pdf\"}");

            JsonNode result = client.put("/drives/d1/items/root:/test.pdf:/content",
                    new byte[]{1, 2, 3}, "application/pdf");

            assertThat(result.get("id").asText()).isEqualTo("file-123");
        }

        @Test
        void should_return_null_when_put_returns_204() throws Exception {
            mockResponse(204, "");

            JsonNode result = client.put("/drives/d1/items/root:/test.pdf:/content",
                    new byte[]{1, 2, 3}, "application/pdf");

            assertThat(result).isNull();
        }

        @Test
        void should_throw_when_put_returns_401() throws Exception {
            mockResponse(401, "{\"error\":{\"code\":\"InvalidAuthenticationToken\",\"message\":\"Token expired\"}}");

            assertThatThrownBy(() -> client.put("/drives/d1/items/root:/test.pdf:/content",
                    new byte[]{1, 2, 3}, "application/pdf"))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("Authentication failed");
        }

        @Test
        void should_throw_when_put_returns_error_with_unparseable_body() throws Exception {
            mockResponse(500, "not json at all");

            assertThatThrownBy(() -> client.put("/test", new byte[]{1}, "application/octet-stream"))
                    .isInstanceOf(MSTeamsException.class);
        }

        @Test
        void should_throw_when_put_returns_error_without_error_object() throws Exception {
            mockResponse(403, "{\"something\":\"else\"}");

            assertThatThrownBy(() -> client.put("/test", new byte[]{1}, "application/octet-stream"))
                    .isInstanceOf(MSTeamsException.class);
        }

        @Test
        void should_throw_when_put_returns_error_with_partial_error_object() throws Exception {
            mockResponse(400, "{\"error\":{}}");

            assertThatThrownBy(() -> client.put("/test", new byte[]{1}, "application/octet-stream"))
                    .isInstanceOf(MSTeamsException.class);
        }
    }

    @Nested
    @DisplayName("Retry logic")
    class RetryLogic {

        @BeforeEach
        void setUp() throws Exception {
            client = new NoSleepMSTeamsClient(CONFIG, httpClient);
            mockResponse(200, "{\"access_token\":\"test-token\",\"expires_in\":3600}");
            client.authenticate();
            reset(httpClient);
        }

        @Test
        void should_retry_on_429_then_succeed() throws Exception {
            mockResponseSequence(
                    429, "{\"error\":{\"code\":\"TooManyRequests\",\"message\":\"Rate limited\"}}",
                    200, "{\"id\":\"success\"}");

            JsonNode result = client.get("/teams/t1");

            assertThat(result.get("id").asText()).isEqualTo("success");
            verify(httpClient, times(2)).send(any(), any());
        }

        @Test
        void should_retry_on_503_then_succeed() throws Exception {
            mockResponseSequence(
                    503, "{\"error\":{\"code\":\"serviceNotAvailable\",\"message\":\"Service unavailable\"}}",
                    200, "{\"value\":[]}");

            JsonNode result = client.get("/teams");

            assertThat(result.get("value")).isNotNull();
            verify(httpClient, times(2)).send(any(), any());
        }

        @Test
        void should_not_retry_on_400() throws Exception {
            mockResponse(400, "{\"error\":{\"code\":\"BadRequest\",\"message\":\"Bad input\"}}");

            assertThatThrownBy(() -> client.get("/teams/t1"))
                    .isInstanceOf(MSTeamsException.class);
            verify(httpClient, times(1)).send(any(), any());
        }

        @Test
        void should_not_retry_on_401() throws Exception {
            mockResponse(401, "{\"error\":{\"code\":\"InvalidToken\",\"message\":\"Expired\"}}");

            assertThatThrownBy(() -> client.get("/teams/t1"))
                    .isInstanceOf(MSTeamsException.class);
            verify(httpClient, times(1)).send(any(), any());
        }

        @Test
        void should_not_retry_on_404() throws Exception {
            mockResponse(404, "{\"error\":{\"code\":\"NotFound\",\"message\":\"Not found\"}}");

            assertThatThrownBy(() -> client.get("/teams/t1"))
                    .isInstanceOf(MSTeamsException.class);
            verify(httpClient, times(1)).send(any(), any());
        }

        @Test
        @SuppressWarnings("unchecked")
        void should_throw_after_max_retries_exceeded() throws Exception {
            // All responses are 429 (retryable) — should exhaust all retries
            when(httpResponse.statusCode()).thenReturn(429);
            when(httpResponse.body()).thenReturn("{\"error\":{\"code\":\"TooManyRequests\",\"message\":\"Rate limited\"}}");
            when(httpClient.send(any(HttpRequest.class), any(HttpResponse.BodyHandler.class)))
                    .thenReturn(httpResponse);

            assertThatThrownBy(() -> client.get("/teams/t1"))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("Max retries");
        }

        @Test
        @SuppressWarnings("unchecked")
        void should_wrap_unexpected_exception_in_msteams_exception() throws Exception {
            when(httpClient.send(any(HttpRequest.class), any(HttpResponse.BodyHandler.class)))
                    .thenThrow(new RuntimeException("Unexpected socket error"));

            assertThatThrownBy(() -> client.get("/teams/t1"))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("Unexpected error");
        }

        @Test
        @SuppressWarnings("unchecked")
        void should_wrap_io_exception_in_msteams_exception() throws Exception {
            when(httpClient.send(any(HttpRequest.class), any(HttpResponse.BodyHandler.class)))
                    .thenThrow(new IOException("Connection reset"));

            assertThatThrownBy(() -> client.get("/teams/t1"))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("Unexpected error");
        }

        @Test
        void should_retry_on_504_then_succeed() throws Exception {
            mockResponseSequence(
                    504, "{\"error\":{\"code\":\"GatewayTimeout\",\"message\":\"Timeout\"}}",
                    200, "{\"id\":\"ok\"}");

            JsonNode result = client.post("/teams/t1/channels/c1/messages", "{\"body\":{}}");

            assertThat(result.get("id").asText()).isEqualTo("ok");
        }

        @Test
        void should_retry_put_on_429() throws Exception {
            mockResponseSequence(
                    429, "{\"error\":{\"code\":\"TooManyRequests\",\"message\":\"Rate limited\"}}",
                    200, "{\"id\":\"uploaded\"}");

            JsonNode result = client.put("/drives/d1/items/root:/f.txt:/content",
                    "data".getBytes(), "text/plain");

            assertThat(result.get("id").asText()).isEqualTo("uploaded");
        }
    }

    @Nested
    @DisplayName("doRequest edge cases")
    class DoRequestEdgeCases {

        @BeforeEach
        void setUp() throws Exception {
            client = new NoSleepMSTeamsClient(CONFIG, httpClient);
            mockResponse(200, "{\"access_token\":\"test-token\",\"expires_in\":3600}");
            client.authenticate();
            reset(httpClient);
        }

        @Test
        void should_return_parsed_json_for_200_range_response() throws Exception {
            mockResponse(201, "{\"id\":\"created\"}");

            JsonNode result = client.post("/teams/t1/channels", "{\"displayName\":\"test\"}");

            assertThat(result.get("id").asText()).isEqualTo("created");
        }
    }

    @Nested
    @DisplayName("doSleep")
    class SleepBehavior {

        @Test
        void should_throw_when_interrupted_during_sleep() {
            var c = new MSTeamsClient(CONFIG, httpClient);
            Thread.currentThread().interrupt();

            assertThatThrownBy(() -> c.doSleep(100))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("Interrupted");

            // Clear flag if still set
            Thread.interrupted();
        }
    }

    /**
     * Test subclass that overrides doSleep to avoid real delays during retry tests.
     */
    static class NoSleepMSTeamsClient extends MSTeamsClient {
        NoSleepMSTeamsClient(MSTeamsConfiguration config, HttpClient httpClient) {
            super(config, httpClient);
        }

        @Override
        void doSleep(long ms) {
            // no-op for fast tests
        }
    }
}
