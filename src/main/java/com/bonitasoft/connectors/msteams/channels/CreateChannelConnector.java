package com.bonitasoft.connectors.msteams.channels;

import com.bonitasoft.connectors.msteams.common.AbstractGraphConnector;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import org.bonitasoft.engine.connector.ConnectorException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Create a channel in a Teams team via Graph API (app-only).
 * Connector ID: msteams-create-channel
 */
public class CreateChannelConnector extends AbstractGraphConnector {

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String INPUT_DISPLAY_NAME = "displayName";
    public static final String INPUT_DESCRIPTION = "description";
    public static final String INPUT_MEMBERSHIP_TYPE = "membershipType";
    public static final String OUTPUT_CHANNEL_ID = "channelId";
    public static final String OUTPUT_WEB_URL = "webUrl";

    @Override
    protected void validateOperationParameters(List<String> errors) {
        if (isNullOrEmpty(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
        if (isNullOrEmpty(getStringInput(INPUT_DISPLAY_NAME))) errors.add("displayName is required");
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            String teamId = getStringInput(INPUT_TEAM_ID);
            String displayName = getStringInput(INPUT_DISPLAY_NAME);
            String description = getStringInput(INPUT_DESCRIPTION);
            String membershipType = getStringInput(INPUT_MEMBERSHIP_TYPE);

            Map<String, Object> body = new HashMap<>();
            body.put("displayName", displayName);
            if (description != null) body.put("description", description);
            if (membershipType != null) body.put("membershipType", membershipType);
            else body.put("membershipType", "standard");

            JsonNode result = graphClient.post("/teams/" + teamId + "/channels", body);
            setOutputParameter(OUTPUT_CHANNEL_ID, result.path("id").asText());
            setOutputParameter(OUTPUT_WEB_URL, result.path("webUrl").asText());
            setSuccessOutputs();
        } catch (MsTeamsException e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        }
    }
}
