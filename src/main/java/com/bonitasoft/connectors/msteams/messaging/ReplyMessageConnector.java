package com.bonitasoft.connectors.msteams.messaging;

import com.bonitasoft.connectors.msteams.common.AbstractGraphConnector;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import org.bonitasoft.engine.connector.ConnectorException;
import java.util.List;
import java.util.Map;

/**
 * Reply to a message in a Teams channel via Graph API (delegated).
 * Connector ID: msteams-reply-message
 */
public class ReplyMessageConnector extends AbstractGraphConnector {

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String INPUT_CHANNEL_ID = "channelId";
    public static final String INPUT_MESSAGE_ID = "messageId";
    public static final String INPUT_REPLY_CONTENT = "replyContent";
    public static final String INPUT_CONTENT_TYPE = "contentType";
    public static final String OUTPUT_REPLY_ID = "replyId";
    public static final String OUTPUT_CREATED_DATE = "createdDateTime";

    @Override
    protected String getAuthMode() {
        return "delegated";
    }

    @Override
    protected void validateOperationParameters(List<String> errors) {
        if (isNullOrEmpty(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
        if (isNullOrEmpty(getStringInput(INPUT_CHANNEL_ID))) errors.add("channelId is required");
        if (isNullOrEmpty(getStringInput(INPUT_MESSAGE_ID))) errors.add("messageId is required");
        if (isNullOrEmpty(getStringInput(INPUT_REPLY_CONTENT))) errors.add("replyContent is required");
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            String teamId = getStringInput(INPUT_TEAM_ID);
            String channelId = getStringInput(INPUT_CHANNEL_ID);
            String messageId = getStringInput(INPUT_MESSAGE_ID);
            String content = getStringInput(INPUT_REPLY_CONTENT);
            String contentType = getStringInput(INPUT_CONTENT_TYPE);
            if (isNullOrEmpty(contentType)) contentType = "text";

            String path = String.format("/teams/%s/channels/%s/messages/%s/replies", teamId, channelId, messageId);
            Map<String, Object> body = Map.of("body", Map.of("contentType", contentType, "content", content));

            JsonNode result = graphClient.post(path, body);
            setOutputParameter(OUTPUT_REPLY_ID, result.path("id").asText());
            setOutputParameter(OUTPUT_CREATED_DATE, result.path("createdDateTime").asText());
            setSuccessOutputs();
        } catch (MsTeamsException e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        }
    }
}
