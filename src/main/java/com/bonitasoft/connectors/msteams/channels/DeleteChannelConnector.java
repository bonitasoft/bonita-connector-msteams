package com.bonitasoft.connectors.msteams.channels;

import com.bonitasoft.connectors.msteams.common.AbstractGraphConnector;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import org.bonitasoft.engine.connector.ConnectorException;
import java.util.List;

/**
 * Delete a channel from a Teams team via Graph API (app-only).
 * Connector ID: msteams-delete-channel
 */
public class DeleteChannelConnector extends AbstractGraphConnector {

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String INPUT_CHANNEL_ID = "channelId";

    @Override
    protected void validateOperationParameters(List<String> errors) {
        if (isNullOrEmpty(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
        if (isNullOrEmpty(getStringInput(INPUT_CHANNEL_ID))) errors.add("channelId is required");
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            String teamId = getStringInput(INPUT_TEAM_ID);
            String channelId = getStringInput(INPUT_CHANNEL_ID);
            graphClient.delete("/teams/" + teamId + "/channels/" + channelId);
            setSuccessOutputs();
        } catch (MsTeamsException e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        }
    }
}
