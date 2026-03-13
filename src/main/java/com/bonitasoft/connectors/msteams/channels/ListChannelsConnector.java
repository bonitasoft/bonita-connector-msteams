package com.bonitasoft.connectors.msteams.channels;

import com.bonitasoft.connectors.msteams.common.AbstractGraphConnector;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import org.bonitasoft.engine.connector.ConnectorException;
import java.util.List;

/**
 * List channels in a Teams team via Graph API (app-only).
 * Connector ID: msteams-list-channels
 */
public class ListChannelsConnector extends AbstractGraphConnector {

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String OUTPUT_CHANNELS_JSON = "channelsJson";
    public static final String OUTPUT_CHANNEL_COUNT = "channelCount";

    @Override
    protected void validateOperationParameters(List<String> errors) {
        if (isNullOrEmpty(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            String teamId = getStringInput(INPUT_TEAM_ID);
            List<JsonNode> channels = graphClient.getWithPagination("/teams/" + teamId + "/channels");
            String json = new com.fasterxml.jackson.databind.ObjectMapper().writeValueAsString(channels);
            setOutputParameter(OUTPUT_CHANNELS_JSON, json);
            setOutputParameter(OUTPUT_CHANNEL_COUNT, channels.size());
            setSuccessOutputs();
        } catch (MsTeamsException e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        } catch (Exception e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        }
    }
}
