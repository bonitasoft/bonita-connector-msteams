package com.bonitasoft.connectors.msteams.messaging;

import com.bonitasoft.connectors.msteams.common.AbstractWebhookConnector;
import com.bonitasoft.connectors.msteams.common.AdaptiveCardBuilder;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.bonitasoft.engine.connector.ConnectorException;
import java.util.List;

/**
 * Send an Adaptive Card to a Teams channel via incoming webhook.
 * Supports both fluent builder mode and raw JSON mode.
 * Connector ID: msteams-send-adaptive-card
 */
public class SendAdaptiveCardConnector extends AbstractWebhookConnector {

    public static final String INPUT_CARD_TITLE = "cardTitle";
    public static final String INPUT_CARD_BODY = "cardBody";
    public static final String INPUT_THEME_COLOR = "themeColor";
    public static final String INPUT_FACTS_JSON = "factsJson";
    public static final String INPUT_ACTIONS_JSON = "actionsJson";
    public static final String INPUT_ADVANCED_CARD_JSON = "advancedCardJson";

    @Override
    protected void validateOperationParameters(List<String> errors) {
        String advancedJson = getStringInput(INPUT_ADVANCED_CARD_JSON);
        if (isNullOrEmpty(advancedJson)) {
            // Builder mode: at least title or body required
            if (isNullOrEmpty(getStringInput(INPUT_CARD_TITLE)) && isNullOrEmpty(getStringInput(INPUT_CARD_BODY))) {
                errors.add("Either cardTitle, cardBody, or advancedCardJson is required");
            }
        }
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            String webhookUrl = getStringInput(INPUT_WEBHOOK_URL);
            String advancedJson = getStringInput(INPUT_ADVANCED_CARD_JSON);
            ObjectNode payload;

            if (!isNullOrEmpty(advancedJson)) {
                payload = AdaptiveCardBuilder.wrapRawCard(advancedJson);
            } else {
                AdaptiveCardBuilder builder = new AdaptiveCardBuilder();
                String cardTitle = getStringInput(INPUT_CARD_TITLE);
                String cardBody = getStringInput(INPUT_CARD_BODY);
                String themeColor = getStringInput(INPUT_THEME_COLOR);

                if (cardTitle != null) builder.title(cardTitle);
                if (cardBody != null) builder.body(cardBody);
                if (themeColor != null) builder.themeColor(themeColor);

                // Parse facts JSON array: [{"key":"k","value":"v"},...]
                String factsJson = getStringInput(INPUT_FACTS_JSON);
                if (!isNullOrEmpty(factsJson)) {
                    try {
                        var mapper = new com.fasterxml.jackson.databind.ObjectMapper();
                        var factsArray = mapper.readTree(factsJson);
                        if (factsArray.isArray()) {
                            for (var factNode : factsArray) {
                                builder.fact(factNode.get("key").asText(), factNode.get("value").asText());
                            }
                        }
                    } catch (Exception e) {
                        throw new MsTeamsException("Invalid factsJson format: " + e.getMessage(), e);
                    }
                }

                // Parse actions JSON array: [{"label":"l","url":"u"},...]
                String actionsJson = getStringInput(INPUT_ACTIONS_JSON);
                if (!isNullOrEmpty(actionsJson)) {
                    try {
                        var mapper = new com.fasterxml.jackson.databind.ObjectMapper();
                        var actionsArray = mapper.readTree(actionsJson);
                        if (actionsArray.isArray()) {
                            for (var actionNode : actionsArray) {
                                builder.action(actionNode.get("label").asText(), actionNode.get("url").asText());
                            }
                        }
                    } catch (Exception e) {
                        throw new MsTeamsException("Invalid actionsJson format: " + e.getMessage(), e);
                    }
                }

                payload = builder.build();
            }

            webhookClient.send(webhookUrl, payload);
            setSuccessOutputs();
        } catch (MsTeamsException e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        }
    }
}
