package com.bonitasoft.connectors.msteams.common;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.DisplayName;

class AdaptiveCardBuilderTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Test
    @DisplayName("Should build card with title and body")
    void should_build_card_with_title_and_body() throws Exception {
        var result = new AdaptiveCardBuilder()
                .title("Test Title")
                .body("Test Body")
                .build();

        assertThat(result.get("type").asText()).isEqualTo("message");
        JsonNode card = result.get("attachments").get(0).get("content");
        assertThat(card.get("type").asText()).isEqualTo("AdaptiveCard");
        assertThat(card.get("version").asText()).isEqualTo("1.4");

        JsonNode bodyArray = card.get("body");
        assertThat(bodyArray).hasSize(2);
        assertThat(bodyArray.get(0).get("text").asText()).isEqualTo("Test Title");
        assertThat(bodyArray.get(1).get("text").asText()).isEqualTo("Test Body");
    }

    @Test
    @DisplayName("Should build card with facts")
    void should_build_card_with_facts() throws Exception {
        var result = new AdaptiveCardBuilder()
                .title("Status")
                .fact("Environment", "Production")
                .fact("Status", "Running")
                .build();

        JsonNode card = result.get("attachments").get(0).get("content");
        JsonNode factSet = card.get("body").get(1);
        assertThat(factSet.get("type").asText()).isEqualTo("FactSet");
        assertThat(factSet.get("facts")).hasSize(2);
    }

    @Test
    @DisplayName("Should build card with actions")
    void should_build_card_with_actions() throws Exception {
        var result = new AdaptiveCardBuilder()
                .title("Alert")
                .action("View Details", "https://example.com")
                .build();

        JsonNode card = result.get("attachments").get(0).get("content");
        JsonNode actions = card.get("actions");
        assertThat(actions).hasSize(1);
        assertThat(actions.get(0).get("type").asText()).isEqualTo("Action.OpenUrl");
        assertThat(actions.get(0).get("title").asText()).isEqualTo("View Details");
    }

    @Test
    @DisplayName("Should set theme color in envelope")
    void should_set_theme_color() {
        var result = new AdaptiveCardBuilder()
                .title("Test")
                .themeColor("FF0000")
                .build();

        assertThat(result.get("themeColor").asText()).isEqualTo("FF0000");
    }

    @Test
    @DisplayName("Should build empty card with no title or body")
    void should_build_empty_card() {
        var result = new AdaptiveCardBuilder().build();
        JsonNode card = result.get("attachments").get(0).get("content");
        assertThat(card.get("body")).isEmpty();
    }

    @Test
    @DisplayName("Should wrap raw card JSON")
    void should_wrap_raw_card_json() {
        String rawJson = "{\"type\":\"AdaptiveCard\",\"version\":\"1.4\",\"body\":[{\"type\":\"TextBlock\",\"text\":\"Hello\"}]}";
        var result = AdaptiveCardBuilder.wrapRawCard(rawJson);
        assertThat(result.get("type").asText()).isEqualTo("message");
        assertThat(result.get("attachments").get(0).get("contentType").asText())
                .isEqualTo("application/vnd.microsoft.card.adaptive");
    }

    @Test
    @DisplayName("Should reject invalid raw JSON")
    void should_reject_invalid_raw_json() {
        assertThatThrownBy(() -> AdaptiveCardBuilder.wrapRawCard("not json"))
                .isInstanceOf(MsTeamsException.class)
                .hasMessageContaining("Invalid Adaptive Card JSON");
    }

    @Test
    @DisplayName("Should reject oversized card")
    void should_reject_oversized_card() {
        String largeBody = "x".repeat(30000);
        assertThatThrownBy(() -> new AdaptiveCardBuilder().body(largeBody).build())
                .isInstanceOf(MsTeamsException.class)
                .hasMessageContaining("28KB");
    }
}
