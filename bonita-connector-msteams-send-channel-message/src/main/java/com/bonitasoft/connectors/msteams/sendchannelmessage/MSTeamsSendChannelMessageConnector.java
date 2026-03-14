package com.bonitasoft.connectors.msteams.sendchannelmessage;

import java.util.List;

import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import lombok.extern.slf4j.Slf4j;

/**
 * MS Teams connector to send a message to a channel.
 * <p>
 * Uses Microsoft Graph API:
 * POST /teams/{teamId}/channels/{channelId}/messages
 */
@Slf4j
public class MSTeamsSendChannelMessageConnector extends AbstractMSTeamsConnector {

    private static final int MAX_MESSAGE_LENGTH = 28_000;

    // Input parameters
    public static final String INPUT_TEAM_ID = "teamId";
    public static final String INPUT_CHANNEL_ID = "channelId";
    public static final String INPUT_MESSAGE_CONTENT = "messageContent";
    public static final String INPUT_CONTENT_TYPE = "contentType";
    public static final String INPUT_SUBJECT = "subject";
    public static final String INPUT_IMPORTANCE = "importance";

    // Output parameters
    public static final String OUTPUT_MESSAGE_ID = "messageId";
    public static final String OUTPUT_WEB_URL = "webUrl";
    public static final String OUTPUT_CREATED_DATE_TIME = "createdDateTime";

    @Override
    protected String getConnectorName() {
        return "MSTeams-SendChannelMessage";
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

        String messageContent = getStringInput(INPUT_MESSAGE_CONTENT);
        if (messageContent == null || messageContent.isBlank()) {
            errors.add("messageContent is required");
        } else if (messageContent.length() > MAX_MESSAGE_LENGTH) {
            errors.add("messageContent exceeds maximum length (" + MAX_MESSAGE_LENGTH + " chars)");
        }
    }

    @Override
    protected void executeOperation(MSTeamsClient client) throws MSTeamsException {
        String teamId = getStringInput(INPUT_TEAM_ID);
        String channelId = getStringInput(INPUT_CHANNEL_ID);
        String messageContent = getStringInput(INPUT_MESSAGE_CONTENT);
        String contentType = getStringInputOrDefault(INPUT_CONTENT_TYPE, "text");
        String subject = getStringInput(INPUT_SUBJECT);
        String importance = getStringInputOrDefault(INPUT_IMPORTANCE, "normal");

        log.info("Sending message to team '{}', channel '{}'", teamId, channelId);

        try {
            ObjectMapper mapper = client.getObjectMapper();
            ObjectNode body = mapper.createObjectNode();

            // Build body content
            ObjectNode bodyContent = mapper.createObjectNode();
            bodyContent.put("content", messageContent);
            bodyContent.put("contentType", contentType);
            body.set("body", bodyContent);

            // Optional fields
            if (subject != null && !subject.isBlank()) {
                body.put("subject", subject);
            }
            body.put("importance", importance);

            String jsonBody = mapper.writeValueAsString(body);
            String path = String.format("/teams/%s/channels/%s/messages", teamId, channelId);

            JsonNode response = client.post(path, jsonBody);

            // Set outputs
            setOutputParameter(OUTPUT_MESSAGE_ID, response.path("id").asText(null));
            setOutputParameter(OUTPUT_WEB_URL, response.path("webUrl").asText(null));
            setOutputParameter(OUTPUT_CREATED_DATE_TIME, response.path("createdDateTime").asText(null));

            log.info("Message sent successfully, messageId: {}", response.path("id").asText());
        } catch (MSTeamsException e) {
            throw e;
        } catch (Exception e) {
            throw new MSTeamsException("Failed to send channel message: " + e.getMessage(), e);
        }
    }
}
