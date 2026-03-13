package com.bonitasoft.connectors.msteams.messaging;

import com.bonitasoft.connectors.msteams.common.AbstractWebhookConnector;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.bonitasoft.engine.connector.ConnectorException;
import java.util.List;

/**
 * Send a plain text or HTML message to a Teams channel via incoming webhook.
 * Connector ID: msteams-send-channel-message
 */
public class SendChannelMessageConnector extends AbstractWebhookConnector {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    public static final String INPUT_MESSAGE = "message";
    public static final String INPUT_TITLE = "title";
    public static final String INPUT_THEME_COLOR = "themeColor";

    @Override
    protected void validateOperationParameters(List<String> errors) {
        if (isNullOrEmpty(getStringInput(INPUT_MESSAGE))) {
            errors.add("message is required");
        }
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            String message = getStringInput(INPUT_MESSAGE);
            String title = getStringInput(INPUT_TITLE);
            String themeColor = getStringInput(INPUT_THEME_COLOR);
            String webhookUrl = getStringInput(INPUT_WEBHOOK_URL);

            ObjectNode payload = MAPPER.createObjectNode();
            payload.put("@type", "MessageCard");
            payload.put("@context", "http://schema.org/extensions");
            if (title != null && !title.isBlank()) {
                payload.put("title", title);
            }
            payload.put("text", message);
            if (themeColor != null && !themeColor.isBlank()) {
                payload.put("themeColor", themeColor);
            }

            webhookClient.send(webhookUrl, payload);
            setSuccessOutputs();
        } catch (MsTeamsException e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        }
    }
}
