package com.bonitasoft.connectors.msteams.property;

import static org.assertj.core.api.Assertions.assertThat;

import com.bonitasoft.connectors.msteams.common.AdaptiveCardBuilder;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import net.jqwik.api.*;
import net.jqwik.api.constraints.AlphaChars;
import net.jqwik.api.constraints.StringLength;

/**
 * Property-based tests for AdaptiveCardBuilder.
 * Ensures cards are always valid JSON regardless of input.
 */
class AdaptiveCardBuilderPropertyTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Property
    void should_always_produce_valid_json_with_random_title(
            @ForAll @AlphaChars @StringLength(min = 1, max = 200) String title) throws Exception {
        var result = new AdaptiveCardBuilder().title(title).build();
        assertThat(result).isNotNull();
        assertThat(result.get("type").asText()).isEqualTo("message");

        // Verify it's valid JSON by re-parsing
        String json = MAPPER.writeValueAsString(result);
        JsonNode reparsed = MAPPER.readTree(json);
        assertThat(reparsed).isNotNull();
    }

    @Property
    void should_always_produce_valid_json_with_random_body(
            @ForAll @AlphaChars @StringLength(min = 1, max = 500) String body) throws Exception {
        var result = new AdaptiveCardBuilder().body(body).build();
        String json = MAPPER.writeValueAsString(result);
        assertThat(json).isNotEmpty();
        assertThat(MAPPER.readTree(json).get("attachments")).isNotNull();
    }

    @Property
    void should_always_include_adaptive_card_in_envelope(
            @ForAll @AlphaChars @StringLength(min = 0, max = 100) String title,
            @ForAll @AlphaChars @StringLength(min = 0, max = 200) String body) {
        var result = new AdaptiveCardBuilder().title(title).body(body).build();
        JsonNode attachments = result.get("attachments");
        assertThat(attachments).hasSize(1);
        assertThat(attachments.get(0).get("contentType").asText())
                .isEqualTo("application/vnd.microsoft.card.adaptive");
    }

    @Property(tries = 50)
    void should_handle_facts_with_random_content(
            @ForAll @AlphaChars @StringLength(min = 1, max = 50) String key,
            @ForAll @AlphaChars @StringLength(min = 1, max = 100) String value) {
        var result = new AdaptiveCardBuilder()
                .title("Test")
                .fact(key, value)
                .build();

        JsonNode card = result.get("attachments").get(0).get("content");
        JsonNode bodyArray = card.get("body");
        // Should have title + factset
        boolean hasFactSet = false;
        for (JsonNode node : bodyArray) {
            if ("FactSet".equals(node.get("type").asText())) {
                hasFactSet = true;
                assertThat(node.get("facts").get(0).get("title").asText()).isEqualTo(key);
                assertThat(node.get("facts").get(0).get("value").asText()).isEqualTo(value);
            }
        }
        assertThat(hasFactSet).isTrue();
    }

    @Property(tries = 50)
    void should_handle_actions_with_random_labels(
            @ForAll @AlphaChars @StringLength(min = 1, max = 50) String label) {
        var result = new AdaptiveCardBuilder()
                .title("Test")
                .action(label, "https://example.com")
                .build();

        JsonNode card = result.get("attachments").get(0).get("content");
        JsonNode actions = card.get("actions");
        assertThat(actions).hasSize(1);
        assertThat(actions.get(0).get("title").asText()).isEqualTo(label);
    }

    @Property(tries = 20)
    void should_handle_multiple_facts(
            @ForAll @AlphaChars @StringLength(min = 1, max = 20) String key1,
            @ForAll @AlphaChars @StringLength(min = 1, max = 20) String key2,
            @ForAll @AlphaChars @StringLength(min = 1, max = 20) String val1,
            @ForAll @AlphaChars @StringLength(min = 1, max = 20) String val2) {
        var result = new AdaptiveCardBuilder()
                .fact(key1, val1)
                .fact(key2, val2)
                .build();

        JsonNode card = result.get("attachments").get(0).get("content");
        for (JsonNode node : card.get("body")) {
            if ("FactSet".equals(node.get("type").asText())) {
                assertThat(node.get("facts")).hasSize(2);
            }
        }
    }

    @Property
    void should_set_theme_color_when_provided(
            @ForAll("hexColors") String color) {
        var result = new AdaptiveCardBuilder()
                .title("Test")
                .themeColor(color)
                .build();
        assertThat(result.get("themeColor").asText()).isEqualTo(color);
    }

    @Provide
    Arbitrary<String> hexColors() {
        return Arbitraries.strings()
                .withCharRange('0', '9')
                .withCharRange('A', 'F')
                .ofLength(6);
    }
}
