package com.bonitasoft.connectors.msteams.getteam;

import java.util.List;
import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import lombok.extern.slf4j.Slf4j;

@Slf4j
public class MSTeamsGetTeamConnector extends AbstractMSTeamsConnector {

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String OUTPUT_TEAM_ID = "teamId";
    public static final String OUTPUT_DISPLAY_NAME = "displayName";
    public static final String OUTPUT_DESCRIPTION = "description";
    public static final String OUTPUT_VISIBILITY = "visibility";
    public static final String OUTPUT_IS_ARCHIVED = "isArchived";
    public static final String OUTPUT_WEB_URL = "webUrl";

    @Override protected String getConnectorName() { return "MSTeams-GetTeam"; }

    @Override
    protected void validateOperationInputs(List<String> errors) {
        if (isBlank(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
    }

    @Override
    protected void executeOperation(MSTeamsClient client) throws MSTeamsException {
        String teamId = getStringInput(INPUT_TEAM_ID);
        log.info("Getting team details for {}", teamId);
        JsonNode r = client.get("/teams/" + teamId);
        setOutputParameter(OUTPUT_TEAM_ID, r.has("id") ? r.get("id").asText() : null);
        setOutputParameter(OUTPUT_DISPLAY_NAME, r.has("displayName") ? r.get("displayName").asText() : null);
        setOutputParameter(OUTPUT_DESCRIPTION, r.has("description") && !r.get("description").isNull() ? r.get("description").asText() : null);
        setOutputParameter(OUTPUT_VISIBILITY, r.has("visibility") ? r.get("visibility").asText() : null);
        setOutputParameter(OUTPUT_IS_ARCHIVED, r.has("isArchived") ? r.get("isArchived").asBoolean() : false);
        setOutputParameter(OUTPUT_WEB_URL, r.has("webUrl") ? r.get("webUrl").asText() : null);
        log.info("Team retrieved: {}", teamId);
    }
}
