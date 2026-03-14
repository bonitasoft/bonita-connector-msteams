package com.bonitasoft.connectors.msteams;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.util.ArrayList;
import java.util.List;

/**
 * Fluent builder for Adaptive Card v1.4 JSON payloads.
 * Produces a Graph API-compatible message body with the card as attachment.
 * <p>
 * Size limit: 28KB for Teams message payloads.
 */
public class AdaptiveCardBuilder {

    private static final ObjectMapper MAPPER = new ObjectMapper();
    private static final int MAX_PAYLOAD_SIZE = 28 * 1024;
    private static final String ADAPTIVE_CARD_SCHEMA = "http://adaptivecards.io/schemas/adaptive-card.json";

    private String title;
    private String body;
    private String themeColor;
    private final List<String[]> facts = new ArrayList<>();
    private final List<String[]> actions = new ArrayList<>();

    public AdaptiveCardBuilder title(String title) {
        this.title = title;
        return this;
    }

    public AdaptiveCardBuilder body(String body) {
        this.body = body;
        return this;
    }

    public AdaptiveCardBuilder themeColor(String themeColor) {
        this.themeColor = themeColor;
        return this;
    }

    public AdaptiveCardBuilder fact(String key, String value) {
        facts.add(new String[]{key, value});
        return this;
    }

    public AdaptiveCardBuilder action(String label, String url) {
        actions.add(new String[]{label, url});
        return this;
    }

    /**
     * Build the Adaptive Card JSON wrapped in a Graph API-compatible message envelope.
     *
     * @return ObjectNode ready to send via Graph API
     * @throws MSTeamsException if payload exceeds 28KB
     */
    public ObjectNode build() {
        ObjectNode card = MAPPER.createObjectNode();
        card.put("$schema", ADAPTIVE_CARD_SCHEMA);
        card.put("type", "AdaptiveCard");
        card.put("version", "1.4");

        ArrayNode cardBody = card.putArray("body");

        // Title
        if (title != null && !title.isBlank()) {
            ObjectNode titleBlock = MAPPER.createObjectNode();
            titleBlock.put("type", "TextBlock");
            titleBlock.put("text", title);
            titleBlock.put("size", "Large");
            titleBlock.put("weight", "Bolder");
            cardBody.add(titleBlock);
        }

        // Body text
        if (body != null && !body.isBlank()) {
            ObjectNode bodyBlock = MAPPER.createObjectNode();
            bodyBlock.put("type", "TextBlock");
            bodyBlock.put("text", body);
            bodyBlock.put("wrap", true);
            cardBody.add(bodyBlock);
        }

        // Facts as FactSet
        if (!facts.isEmpty()) {
            ObjectNode factSet = MAPPER.createObjectNode();
            factSet.put("type", "FactSet");
            ArrayNode factsArray = factSet.putArray("facts");
            for (String[] fact : facts) {
                ObjectNode factNode = MAPPER.createObjectNode();
                factNode.put("title", fact[0]);
                factNode.put("value", fact[1]);
                factsArray.add(factNode);
            }
            cardBody.add(factSet);
        }

        // Actions
        if (!actions.isEmpty()) {
            ArrayNode actionsArray = card.putArray("actions");
            for (String[] action : actions) {
                ObjectNode actionNode = MAPPER.createObjectNode();
                actionNode.put("type", "Action.OpenUrl");
                actionNode.put("title", action[0]);
                actionNode.put("url", action[1]);
                actionsArray.add(actionNode);
            }
        }

        // Wrap in message envelope
        ObjectNode envelope = MAPPER.createObjectNode();
        envelope.put("type", "message");
        ArrayNode attachments = envelope.putArray("attachments");
        ObjectNode attachment = MAPPER.createObjectNode();
        attachment.put("contentType", "application/vnd.microsoft.card.adaptive");
        attachment.set("content", card);
        attachments.add(attachment);

        if (themeColor != null && !themeColor.isBlank()) {
            envelope.put("themeColor", themeColor);
        }

        // Size validation
        String json = envelope.toString();
        if (json.length() > MAX_PAYLOAD_SIZE) {
            throw new MSTeamsException(
                    String.format("Adaptive Card payload exceeds 28KB limit: %d bytes", json.length()));
        }

        return envelope;
    }

    /**
     * Wrap a raw Adaptive Card JSON string in the message envelope.
     *
     * @param rawCardJson raw Adaptive Card JSON
     * @return ObjectNode ready to send via Graph API
     */
    public static ObjectNode wrapRawCard(String rawCardJson) {
        try {
            ObjectNode card = (ObjectNode) MAPPER.readTree(rawCardJson);
            ObjectNode envelope = MAPPER.createObjectNode();
            envelope.put("type", "message");
            ArrayNode attachments = envelope.putArray("attachments");
            ObjectNode attachment = MAPPER.createObjectNode();
            attachment.put("contentType", "application/vnd.microsoft.card.adaptive");
            attachment.set("content", card);
            attachments.add(attachment);

            String json = envelope.toString();
            if (json.length() > MAX_PAYLOAD_SIZE) {
                throw new MSTeamsException(
                        String.format("Adaptive Card payload exceeds 28KB limit: %d bytes", json.length()));
            }
            return envelope;
        } catch (MSTeamsException e) {
            throw e;
        } catch (Exception e) {
            throw new MSTeamsException("Invalid Adaptive Card JSON: " + e.getMessage(), e);
        }
    }
}
