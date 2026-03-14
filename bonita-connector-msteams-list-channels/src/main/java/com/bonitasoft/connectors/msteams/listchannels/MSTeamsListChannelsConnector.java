package com.bonitasoft.connectors.msteams.listchannels;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import lombok.extern.slf4j.Slf4j;

@Slf4j
public class MSTeamsListChannelsConnector extends AbstractMSTeamsConnector {

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String OUTPUT_CHANNELS = "channels";
    public static final String OUTPUT_TOTAL_COUNT = "totalCount";

    @Override protected String getConnectorName() { return "MSTeams-ListChannels"; }

    @Override
    protected void validateOperationInputs(List<String> errors) {
        if (isBlank(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
    }

    @Override
    protected void executeOperation(MSTeamsClient client) throws MSTeamsException {
        String teamId = getStringInput(INPUT_TEAM_ID);
        log.info("Listing channels for team {}", teamId);
        JsonNode response = client.get("/teams/" + teamId + "/channels");
        List<Map<String, Object>> list = new ArrayList<>();
        JsonNode value = response.get("value");
        if (value != null && value.isArray()) {
            for (JsonNode n : value) {
                Map<String, Object> ch = new HashMap<>();
                ch.put("id", n.has("id") ? n.get("id").asText() : null);
                ch.put("displayName", n.has("displayName") ? n.get("displayName").asText() : null);
                ch.put("description", n.has("description") ? n.get("description").asText() : null);
                ch.put("membershipType", n.has("membershipType") ? n.get("membershipType").asText() : null);
                list.add(ch);
            }
        }
        setOutputParameter(OUTPUT_CHANNELS, list);
        setOutputParameter(OUTPUT_TOTAL_COUNT, list.size());
        log.info("Found {} channels", list.size());
    }
}
