package com.bonitasoft.connectors.msteams.replymessage;

import java.util.List;

import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import lombok.extern.slf4j.Slf4j;

/**
 * MS Teams connector to reply to a channel message.
 * <p>
 * Uses Microsoft Graph API:
 * POST /teams/{teamId}/channels/{channelId}/messages/{parentMessageId}/replies
 */
@Slf4j
public class MSTeamsReplyMessageConnector extends AbstractMSTeamsConnector {

    private static final int MAX_MESSAGE_LENGTH = 28_000;

    // Input parameters
    public static final String INPUT_TEAM_ID = "teamId";
    public static final String INPUT_CHANNEL_ID = "channelId";
    public static final String INPUT_PARENT_MESSAGE_ID = "parentMessageId";
    public static final String INPUT_MESSAGE_CONTENT = "messageContent";
    public static final String INPUT_CONTENT_TYPE = "contentType";

    // Output parameters
    public static final String OUTPUT_REPLY_ID = "replyId";
    public static final String OUTPUT_CREATED_DATE_TIME = "createdDateTime";

    @Override
    protected String getConnectorName() {
        return "MSTeams-ReplyMessage";
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

        String parentMessageId = getStringInput(INPUT_PARENT_MESSAGE_ID);
        if (parentMessageId == null || parentMessageId.isBlank()) {
            errors.add("parentMessageId is required");
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
        String parentMessageId = getStringInput(INPUT_PARENT_MESSAGE_ID);
        String messageContent = getStringInput(INPUT_MESSAGE_CONTENT);
        String contentType = getStringInputOrDefault(INPUT_CONTENT_TYPE, "text");

        log.info("Replying to message '{}' in team '{}', channel '{}'", parentMessageId, teamId, channelId);

        try {
            ObjectMapper mapper = client.getObjectMapper();
            ObjectNode body = mapper.createObjectNode();

            // Build body content
            ObjectNode bodyContent = mapper.createObjectNode();
            bodyContent.put("content", messageContent);
            bodyContent.put("contentType", contentType);
            body.set("body", bodyContent);

            String jsonBody = mapper.writeValueAsString(body);
            String path = String.format("/teams/%s/channels/%s/messages/%s/replies",
                    teamId, channelId, parentMessageId);

            JsonNode response = client.post(path, jsonBody);

            // Set outputs
            setOutputParameter(OUTPUT_REPLY_ID, response.path("id").asText(null));
            setOutputParameter(OUTPUT_CREATED_DATE_TIME, response.path("createdDateTime").asText(null));

            log.info("Reply sent successfully, replyId: {}", response.path("id").asText());
        } catch (MSTeamsException e) {
            throw e;
        } catch (Exception e) {
            throw new MSTeamsException("Failed to reply to message: " + e.getMessage(), e);
        }
    }
}
