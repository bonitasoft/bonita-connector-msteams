package com.bonitasoft.connectors.msteams.common;

import static com.github.tomakehurst.wiremock.client.WireMock.*;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import com.fasterxml.jackson.databind.JsonNode;
import com.github.tomakehurst.wiremock.junit5.WireMockRuntimeInfo;
import com.github.tomakehurst.wiremock.junit5.WireMockTest;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Map;

@WireMockTest
@DisplayName("MsTeamsGraphClient")
class MsTeamsGraphClientTest {

    private MsTeamsGraphClient createClient(String baseUrl) {
        return new MsTeamsGraphClient(baseUrl, "test-token", 5000, 0);
    }

    @Nested
    @DisplayName("GET requests")
    class GetRequests {

        @Test
        void should_get_json_response(WireMockRuntimeInfo wm) {
            stubFor(get("/teams/t1").willReturn(okJson("{\"id\":\"t1\",\"displayName\":\"Team1\"}")));

            MsTeamsGraphClient client = createClient(wm.getHttpBaseUrl());
            JsonNode result = client.get("/teams/t1");

            assertThat(result.get("id").asText()).isEqualTo("t1");
            assertThat(result.get("displayName").asText()).isEqualTo("Team1");
            verify(getRequestedFor(urlEqualTo("/teams/t1"))
                    .withHeader("Authorization", equalTo("Bearer test-token")));
        }

        @Test
        void should_throw_on_error_response(WireMockRuntimeInfo wm) {
            stubFor(get("/teams/bad").willReturn(
                    aResponse().withStatus(404)
                            .withBody("{\"error\":{\"code\":\"NotFound\",\"message\":\"Team not found\"}}")));

            MsTeamsGraphClient client = createClient(wm.getHttpBaseUrl());
            assertThatThrownBy(() -> client.get("/teams/bad"))
                    .isInstanceOf(MsTeamsException.ExecutionException.class)
                    .hasMessageContaining("NotFound");
        }

        @Test
        void should_handle_empty_response_body(WireMockRuntimeInfo wm) {
            stubFor(get("/empty").willReturn(aResponse().withStatus(204).withBody("")));

            MsTeamsGraphClient client = createClient(wm.getHttpBaseUrl());
            JsonNode result = client.get("/empty");
            assertThat(result.get("status").asInt()).isEqualTo(204);
        }
    }

    @Nested
    @DisplayName("POST requests")
    class PostRequests {

        @Test
        void should_post_json_body(WireMockRuntimeInfo wm) {
            stubFor(post("/teams/t1/channels").willReturn(
                    okJson("{\"id\":\"ch1\",\"displayName\":\"General\"}")));

            MsTeamsGraphClient client = createClient(wm.getHttpBaseUrl());
            JsonNode result = client.post("/teams/t1/channels",
                    Map.of("displayName", "General", "description", "Test channel"));

            assertThat(result.get("id").asText()).isEqualTo("ch1");
            verify(postRequestedFor(urlEqualTo("/teams/t1/channels"))
                    .withHeader("Content-Type", equalTo("application/json")));
        }
    }

    @Nested
    @DisplayName("PUT requests")
    class PutRequests {

        @Test
        void should_upload_binary_content(WireMockRuntimeInfo wm) {
            stubFor(put(urlPathMatching("/upload.*")).willReturn(
                    okJson("{\"id\":\"file1\",\"webUrl\":\"https://sp/file1\"}")));

            MsTeamsGraphClient client = createClient(wm.getHttpBaseUrl());
            JsonNode result = client.put("/upload?name=test.txt",
                    "file content".getBytes(), "application/octet-stream");

            assertThat(result.get("id").asText()).isEqualTo("file1");
        }
    }

    @Nested
    @DisplayName("DELETE requests")
    class DeleteRequests {

        @Test
        void should_delete_resource(WireMockRuntimeInfo wm) {
            stubFor(delete("/teams/t1/channels/ch1").willReturn(
                    aResponse().withStatus(204).withBody("")));

            MsTeamsGraphClient client = createClient(wm.getHttpBaseUrl());
            client.delete("/teams/t1/channels/ch1");

            verify(deleteRequestedFor(urlEqualTo("/teams/t1/channels/ch1")));
        }
    }

    @Nested
    @DisplayName("Pagination")
    class Pagination {

        @Test
        void should_collect_all_pages(WireMockRuntimeInfo wm) {
            String page1 = """
                    {"value":[{"id":"1"},{"id":"2"}],"@odata.nextLink":"%s/page2"}
                    """.formatted(wm.getHttpBaseUrl());
            String page2 = """
                    {"value":[{"id":"3"}]}
                    """;

            stubFor(get("/items").willReturn(okJson(page1)));
            stubFor(get("/page2").willReturn(okJson(page2)));

            MsTeamsGraphClient client = createClient(wm.getHttpBaseUrl());
            List<JsonNode> items = client.getWithPagination("/items");

            assertThat(items).hasSize(3);
        }

        @Test
        void should_handle_single_page(WireMockRuntimeInfo wm) {
            stubFor(get("/items").willReturn(okJson("{\"value\":[{\"id\":\"1\"}]}")));

            MsTeamsGraphClient client = createClient(wm.getHttpBaseUrl());
            List<JsonNode> items = client.getWithPagination("/items");

            assertThat(items).hasSize(1);
        }
    }

    @Nested
    @DisplayName("Error handling")
    class ErrorHandling {

        @Test
        void should_parse_graph_api_error(WireMockRuntimeInfo wm) {
            stubFor(get("/error").willReturn(
                    aResponse().withStatus(403)
                            .withBody("{\"error\":{\"code\":\"AccessDenied\",\"message\":\"Insufficient privileges\"}}")));

            MsTeamsGraphClient client = createClient(wm.getHttpBaseUrl());
            assertThatThrownBy(() -> client.get("/error"))
                    .isInstanceOf(MsTeamsException.ExecutionException.class)
                    .hasMessageContaining("AccessDenied")
                    .hasMessageContaining("Insufficient privileges");
        }

        @Test
        void should_handle_non_json_error_body(WireMockRuntimeInfo wm) {
            stubFor(get("/error").willReturn(
                    aResponse().withStatus(500).withBody("Internal Server Error")));

            MsTeamsGraphClient client = createClient(wm.getHttpBaseUrl());
            assertThatThrownBy(() -> client.get("/error"))
                    .isInstanceOf(MsTeamsException.ExecutionException.class)
                    .hasMessageContaining("500");
        }
    }

    @Test
    @DisplayName("should_close_without_error")
    void should_close_without_error() {
        MsTeamsGraphClient client = createClient("http://localhost:1234");
        client.close();
    }

    @Test
    @DisplayName("should_strip_trailing_slash_from_base_url")
    void should_strip_trailing_slash_from_base_url(WireMockRuntimeInfo wm) {
        stubFor(get("/test").willReturn(okJson("{\"ok\":true}")));

        MsTeamsGraphClient client = new MsTeamsGraphClient(
                wm.getHttpBaseUrl() + "/", "test-token", 5000, 0);
        JsonNode result = client.get("/test");
        assertThat(result.get("ok").asBoolean()).isTrue();
    }
}
