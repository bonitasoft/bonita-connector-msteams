package com.bonitasoft.connectors.msteams.messaging;

import com.bonitasoft.connectors.msteams.common.AbstractGraphConnector;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import org.bonitasoft.engine.connector.ConnectorException;
import java.util.List;
import java.util.Map;

/**
 * Send a message to a Teams chat via Graph API (delegated permissions).
 * Connector ID: msteams-send-chat-message
 */
public class SendChatMessageConnector extends AbstractGraphConnector {

    public static final String INPUT_CHAT_ID = "chatId";
    public static final String INPUT_MESSAGE = "message";
    public static final String INPUT_CONTENT_TYPE = "contentType";
    public static final String OUTPUT_MESSAGE_ID = "messageId";
    public static final String OUTPUT_CREATED_DATE = "createdDateTime";

    @Override
    protected String getAuthMode() {
        return "delegated";
    }

    @Override
    protected void validateOperationParameters(List<String> errors) {
        if (isNullOrEmpty(getStringInput(INPUT_CHAT_ID))) {
            errors.add("chatId is required");
        }
        if (isNullOrEmpty(getStringInput(INPUT_MESSAGE))) {
            errors.add("message is required");
        }
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            String chatId = getStringInput(INPUT_CHAT_ID);
            String message = getStringInput(INPUT_MESSAGE);
            String contentType = getStringInput(INPUT_CONTENT_TYPE);
            if (isNullOrEmpty(contentType)) contentType = "text";

            Map<String, Object> body = Map.of(
                    "body", Map.of(
                            "contentType", contentType,
                            "content", message
                    )
            );

            JsonNode result = graphClient.post("/chats/" + chatId + "/messages", body);
            setOutputParameter(OUTPUT_MESSAGE_ID, result.path("id").asText());
            setOutputParameter(OUTPUT_CREATED_DATE, result.path("createdDateTime").asText());
            setSuccessOutputs();
        } catch (MsTeamsException e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        }
    }
}
