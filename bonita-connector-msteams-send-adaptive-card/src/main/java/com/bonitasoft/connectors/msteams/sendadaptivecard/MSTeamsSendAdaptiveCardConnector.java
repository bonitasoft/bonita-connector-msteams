package com.bonitasoft.connectors.msteams.sendadaptivecard;

import java.util.List;

import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import lombok.extern.slf4j.Slf4j;

/**
 * MS Teams connector to send an Adaptive Card to a channel.
 * <p>
 * Uses Microsoft Graph API:
 * POST /teams/{teamId}/channels/{channelId}/messages
 * <p>
 * The card JSON is sent as an attachment with contentType
 * "application/vnd.microsoft.card.adaptive".
 */
@Slf4j
public class MSTeamsSendAdaptiveCardConnector extends AbstractMSTeamsConnector {

    // Input parameters
    public static final String INPUT_TEAM_ID = "teamId";
    public static final String INPUT_CHANNEL_ID = "channelId";
    public static final String INPUT_CARD_JSON = "cardJson";
    public static final String INPUT_SUBJECT = "subject";
    public static final String INPUT_IMPORTANCE = "importance";

    // Output parameters
    public static final String OUTPUT_MESSAGE_ID = "messageId";
    public static final String OUTPUT_WEB_URL = "webUrl";
    public static final String OUTPUT_CREATED_DATE_TIME = "createdDateTime";

    private static final String ADAPTIVE_CARD_CONTENT_TYPE = "application/vnd.microsoft.card.adaptive";

    @Override
    protected String getConnectorName() {
        return "MSTeams-SendAdaptiveCard";
    }

    @Override
    protected void validateOperationInputs(List<String> errors) {
        String teamId = getStringInput(INPUT_TEAM_ID);
        if (teamId == null || teamId.isBlank()) {
            errors.add("teamId is required");
        }

        String channelId = getStringInput(INPUT_CHANNEL_ID);
        if (channelId == null || channelId.isBlank()) {
            errors.add("channelId is required");
        }

        String cardJson = getStringInput(INPUT_CARD_JSON);
        if (cardJson == null || cardJson.isBlank()) {
            errors.add("cardJson is required");
        } else {
            // Validate that cardJson is valid JSON
            try {
                new ObjectMapper().readTree(cardJson);
            } catch (Exception e) {
                errors.add("cardJson must be valid JSON: " + e.getMessage());
            }
        }
    }

    @Override
    protected void executeOperation(MSTeamsClient client) throws MSTeamsException {
        String teamId = getStringInput(INPUT_TEAM_ID);
        String channelId = getStringInput(INPUT_CHANNEL_ID);
        String cardJson = getStringInput(INPUT_CARD_JSON);
        String subject = getStringInput(INPUT_SUBJECT);
        String importance = getStringInputOrDefault(INPUT_IMPORTANCE, "normal");

        log.info("Sending Adaptive Card to team '{}', channel '{}'", teamId, channelId);

        try {
            ObjectMapper mapper = client.getObjectMapper();
            ObjectNode body = mapper.createObjectNode();

            // Body with empty HTML content (card renders from attachment)
            ObjectNode bodyContent = mapper.createObjectNode();
            bodyContent.put("content", "");
            bodyContent.put("contentType", "html");
            body.set("body", bodyContent);

            // Optional fields
            if (subject != null && !subject.isBlank()) {
                body.put("subject", subject);
            }
            body.put("importance", importance);

            // Build attachment with Adaptive Card
            ArrayNode attachments = mapper.createArrayNode();
            ObjectNode attachment = mapper.createObjectNode();
            attachment.put("id", "adaptive-card-1");
            attachment.put("contentType", ADAPTIVE_CARD_CONTENT_TYPE);
            attachment.set("content", mapper.readTree(cardJson));
            attachments.add(attachment);
            body.set("attachments", attachments);

            String jsonBody = mapper.writeValueAsString(body);
            String path = String.format("/teams/%s/channels/%s/messages", teamId, channelId);

            JsonNode response = client.post(path, jsonBody);

            // Set outputs
            setOutputParameter(OUTPUT_MESSAGE_ID, response.path("id").asText(null));
            setOutputParameter(OUTPUT_WEB_URL, response.path("webUrl").asText(null));
            setOutputParameter(OUTPUT_CREATED_DATE_TIME, response.path("createdDateTime").asText(null));

            log.info("Adaptive Card sent successfully, messageId: {}", response.path("id").asText());
        } catch (MSTeamsException e) {
            throw e;
        } catch (Exception e) {
            throw new MSTeamsException("Failed to send Adaptive Card: " + e.getMessage(), e);
        }
    }
}
