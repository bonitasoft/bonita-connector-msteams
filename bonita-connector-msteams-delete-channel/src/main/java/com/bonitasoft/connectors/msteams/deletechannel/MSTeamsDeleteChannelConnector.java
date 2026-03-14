package com.bonitasoft.connectors.msteams.deletechannel;

import java.util.List;
import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import lombok.extern.slf4j.Slf4j;

@Slf4j
public class MSTeamsDeleteChannelConnector extends AbstractMSTeamsConnector {

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String INPUT_CHANNEL_ID = "channelId";

    @Override protected String getConnectorName() { return "MSTeams-DeleteChannel"; }

    @Override
    protected void validateOperationInputs(List<String> errors) {
        if (isBlank(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
        if (isBlank(getStringInput(INPUT_CHANNEL_ID))) errors.add("channelId is required");
    }

    @Override
    protected void executeOperation(MSTeamsClient client) throws MSTeamsException {
        String teamId = getStringInput(INPUT_TEAM_ID);
        String channelId = getStringInput(INPUT_CHANNEL_ID);
        log.info("Deleting channel {} from team {}", channelId, teamId);
        client.delete("/teams/" + teamId + "/channels/" + channelId);
        log.info("Channel {} deleted from team {}", channelId, teamId);
    }
}
