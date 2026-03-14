package com.bonitasoft.connectors.msteams;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

@DisplayName("AdaptiveCardBuilder")
class AdaptiveCardBuilderTest {

    @Nested
    @DisplayName("build")
    class Build {

        @Test
        void should_build_card_with_title() {
            ObjectNode result = new AdaptiveCardBuilder()
                    .title("Test Title")
                    .build();

            assertThat(result.has("type")).isTrue();
            assertThat(result.get("type").asText()).isEqualTo("message");
            assertThat(result.has("attachments")).isTrue();

            JsonNode card = result.get("attachments").get(0).get("content");
            assertThat(card.get("type").asText()).isEqualTo("AdaptiveCard");
            assertThat(card.get("version").asText()).isEqualTo("1.4");

            JsonNode body = card.get("body");
            assertThat(body.get(0).get("text").asText()).isEqualTo("Test Title");
            assertThat(body.get(0).get("weight").asText()).isEqualTo("Bolder");
        }

        @Test
        void should_build_card_with_body_text() {
            ObjectNode result = new AdaptiveCardBuilder()
                    .body("Hello World")
                    .build();

            JsonNode card = result.get("attachments").get(0).get("content");
            JsonNode body = card.get("body");
            assertThat(body.get(0).get("text").asText()).isEqualTo("Hello World");
            assertThat(body.get(0).get("wrap").asBoolean()).isTrue();
        }

        @Test
        void should_build_card_with_facts() {
            ObjectNode result = new AdaptiveCardBuilder()
                    .fact("Key1", "Value1")
                    .fact("Key2", "Value2")
                    .build();

            JsonNode card = result.get("attachments").get(0).get("content");
            JsonNode factSet = card.get("body").get(0);
            assertThat(factSet.get("type").asText()).isEqualTo("FactSet");
            assertThat(factSet.get("facts")).hasSize(2);
        }

        @Test
        void should_build_card_with_actions() {
            ObjectNode result = new AdaptiveCardBuilder()
                    .title("Test")
                    .action("Click me", "https://example.com")
                    .build();

            JsonNode card = result.get("attachments").get(0).get("content");
            JsonNode actions = card.get("actions");
            assertThat(actions).hasSize(1);
            assertThat(actions.get(0).get("type").asText()).isEqualTo("Action.OpenUrl");
            assertThat(actions.get(0).get("title").asText()).isEqualTo("Click me");
            assertThat(actions.get(0).get("url").asText()).isEqualTo("https://example.com");
        }

        @Test
        void should_include_theme_color() {
            ObjectNode result = new AdaptiveCardBuilder()
                    .title("Test")
                    .themeColor("FF0000")
                    .build();

            assertThat(result.get("themeColor").asText()).isEqualTo("FF0000");
        }

        @Test
        void should_skip_blank_title() {
            ObjectNode result = new AdaptiveCardBuilder()
                    .title("  ")
                    .body("content")
                    .build();

            JsonNode card = result.get("attachments").get(0).get("content");
            assertThat(card.get("body")).hasSize(1); // only body, no title
        }

        @Test
        void should_skip_blank_body() {
            ObjectNode result = new AdaptiveCardBuilder()
                    .title("Title")
                    .body("")
                    .build();

            JsonNode card = result.get("attachments").get(0).get("content");
            assertThat(card.get("body")).hasSize(1); // only title, no body
        }

        @Test
        void should_skip_blank_theme_color() {
            ObjectNode result = new AdaptiveCardBuilder()
                    .title("Test")
                    .themeColor("")
                    .build();

            assertThat(result.has("themeColor")).isFalse();
        }

        @Test
        void should_build_empty_card() {
            ObjectNode result = new AdaptiveCardBuilder().build();
            assertThat(result.get("type").asText()).isEqualTo("message");
            JsonNode card = result.get("attachments").get(0).get("content");
            assertThat(card.get("body")).isEmpty();
        }

        @Test
        void should_reject_payload_exceeding_28kb() {
            AdaptiveCardBuilder builder = new AdaptiveCardBuilder()
                    .body("x".repeat(30_000));

            assertThatThrownBy(builder::build)
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("28KB limit");
        }

        @Test
        void should_skip_null_title() {
            ObjectNode result = new AdaptiveCardBuilder()
                    .title(null)
                    .body("content")
                    .build();
            JsonNode card = result.get("attachments").get(0).get("content");
            assertThat(card.get("body")).hasSize(1); // only body, no title
        }

        @Test
        void should_skip_null_body() {
            ObjectNode result = new AdaptiveCardBuilder()
                    .title("Title")
                    .body(null)
                    .build();
            JsonNode card = result.get("attachments").get(0).get("content");
            assertThat(card.get("body")).hasSize(1); // only title, no body
        }

        @Test
        void should_skip_null_theme_color() {
            ObjectNode result = new AdaptiveCardBuilder()
                    .title("Test")
                    .themeColor(null)
                    .build();
            assertThat(result.has("themeColor")).isFalse();
        }

        @Test
        void should_build_card_with_title_body_facts_and_actions() {
            ObjectNode result = new AdaptiveCardBuilder()
                    .title("Title")
                    .body("Body")
                    .fact("F1", "V1")
                    .action("Click", "https://example.com")
                    .themeColor("00FF00")
                    .build();

            assertThat(result.get("themeColor").asText()).isEqualTo("00FF00");
            JsonNode card = result.get("attachments").get(0).get("content");
            // title + body + factSet = 3 elements
            assertThat(card.get("body")).hasSize(3);
            assertThat(card.get("actions")).hasSize(1);
        }
    }

    @Nested
    @DisplayName("wrapRawCard")
    class WrapRawCard {

        @Test
        void should_wrap_valid_json() {
            String cardJson = "{\"type\":\"AdaptiveCard\",\"body\":[]}";
            ObjectNode result = AdaptiveCardBuilder.wrapRawCard(cardJson);

            assertThat(result.get("type").asText()).isEqualTo("message");
            assertThat(result.get("attachments").get(0).get("contentType").asText())
                    .isEqualTo("application/vnd.microsoft.card.adaptive");
        }

        @Test
        void should_reject_invalid_json() {
            assertThatThrownBy(() -> AdaptiveCardBuilder.wrapRawCard("not json"))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("Invalid Adaptive Card JSON");
        }

        @Test
        void should_reject_oversized_raw_card() {
            String huge = "{\"data\":\"" + "x".repeat(30_000) + "\"}";
            assertThatThrownBy(() -> AdaptiveCardBuilder.wrapRawCard(huge))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("28KB limit");
        }
    }
}
