package com.bonitasoft.connectors.msteams.common;

import java.util.List;

import org.bonitasoft.engine.connector.ConnectorException;

/**
 * Base class for webhook-based connectors (SendChannelMessage, SendAdaptiveCard).
 * Uses incoming webhook URLs — no Azure AD auth required.
 */
public abstract class AbstractWebhookConnector extends AbstractMsTeamsConnector {

    public static final String INPUT_WEBHOOK_URL = "webhookUrl";

    protected MsTeamsWebhookClient webhookClient;

    @Override
    protected void validateConnectionParameters(List<String> errors) {
        String webhookUrl = getStringInput(INPUT_WEBHOOK_URL);
        if (isNullOrEmpty(webhookUrl)) {
            errors.add("webhookUrl is required");
        } else if (!webhookUrl.startsWith("https://")) {
            errors.add("webhookUrl must use HTTPS");
        }

        Integer connectTimeout = getIntegerInput(INPUT_CONNECT_TIMEOUT);
        if (connectTimeout != null && connectTimeout < 0) {
            errors.add("connectTimeout must be a positive number");
        }
    }

    @Override
    public void connect() throws ConnectorException {
        try {
            int connectTimeout = getIntegerInputOrDefault(INPUT_CONNECT_TIMEOUT, 30000);
            webhookClient = new MsTeamsWebhookClient(connectTimeout, 3);
        } catch (Exception e) {
            throw new ConnectorException("Failed to create webhook client: " + e.getMessage(), e);
        }
    }

    @Override
    public void disconnect() throws ConnectorException {
        if (webhookClient != null) {
            try {
                webhookClient.close();
            } catch (Exception e) {
                LOGGER.warning("Error closing webhook client: " + e.getMessage());
            }
        }
    }

    public void setWebhookClient(MsTeamsWebhookClient client) {
        this.webhookClient = client;
    }
}
