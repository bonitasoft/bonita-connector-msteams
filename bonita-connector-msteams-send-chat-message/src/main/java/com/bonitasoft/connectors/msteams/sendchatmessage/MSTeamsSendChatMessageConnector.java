package com.bonitasoft.connectors.msteams.sendchatmessage;

import java.util.List;

import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import lombok.extern.slf4j.Slf4j;

/**
 * MS Teams connector to send a message to a chat.
 * <p>
 * Uses Microsoft Graph API:
 * POST /chats/{chatId}/messages
 */
@Slf4j
public class MSTeamsSendChatMessageConnector extends AbstractMSTeamsConnector {

    private static final int MAX_MESSAGE_LENGTH = 28_000;

    // Input parameters
    public static final String INPUT_CHAT_ID = "chatId";
    public static final String INPUT_MESSAGE_CONTENT = "messageContent";
    public static final String INPUT_CONTENT_TYPE = "contentType";
    public static final String INPUT_IMPORTANCE = "importance";

    // Output parameters
    public static final String OUTPUT_MESSAGE_ID = "messageId";
    public static final String OUTPUT_CHAT_ID = "chatId";
    public static final String OUTPUT_CREATED_DATE_TIME = "createdDateTime";

    @Override
    protected String getConnectorName() {
        return "MSTeams-SendChatMessage";
    }

    @Override
    protected void validateOperationInputs(List<String> errors) {
        String chatId = getStringInput(INPUT_CHAT_ID);
        if (chatId == null || chatId.isBlank()) {
            errors.add("chatId is required");
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
        String chatId = getStringInput(INPUT_CHAT_ID);
        String messageContent = getStringInput(INPUT_MESSAGE_CONTENT);
        String contentType = getStringInputOrDefault(INPUT_CONTENT_TYPE, "text");
        String importance = getStringInputOrDefault(INPUT_IMPORTANCE, "normal");

        log.info("Sending message to chat '{}'", chatId);

        try {
            ObjectMapper mapper = client.getObjectMapper();
            ObjectNode body = mapper.createObjectNode();

            // Build body content
            ObjectNode bodyContent = mapper.createObjectNode();
            bodyContent.put("content", messageContent);
            bodyContent.put("contentType", contentType);
            body.set("body", bodyContent);

            body.put("importance", importance);

            String jsonBody = mapper.writeValueAsString(body);
            String path = String.format("/chats/%s/messages", chatId);

            JsonNode response = client.post(path, jsonBody);

            // Set outputs
            setOutputParameter(OUTPUT_MESSAGE_ID, response.path("id").asText(null));
            setOutputParameter(OUTPUT_CHAT_ID, chatId);
            setOutputParameter(OUTPUT_CREATED_DATE_TIME, response.path("createdDateTime").asText(null));

            log.info("Chat message sent successfully, messageId: {}", response.path("id").asText());
        } catch (MSTeamsException e) {
            throw e;
        } catch (Exception e) {
            throw new MSTeamsException("Failed to send chat message: " + e.getMessage(), e);
        }
    }
}
