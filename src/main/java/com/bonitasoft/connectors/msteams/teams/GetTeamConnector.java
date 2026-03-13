package com.bonitasoft.connectors.msteams.teams;

import com.bonitasoft.connectors.msteams.common.AbstractGraphConnector;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import org.bonitasoft.engine.connector.ConnectorException;
import java.util.List;

/**
 * Get team details via Graph API (app-only).
 * Connector ID: msteams-get-team
 */
public class GetTeamConnector extends AbstractGraphConnector {

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String OUTPUT_TEAM_JSON = "teamJson";
    public static final String OUTPUT_DISPLAY_NAME = "displayName";
    public static final String OUTPUT_DESCRIPTION = "description";

    @Override
    protected void validateOperationParameters(List<String> errors) {
        if (isNullOrEmpty(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            String teamId = getStringInput(INPUT_TEAM_ID);
            JsonNode result = graphClient.get("/teams/" + teamId);
            setOutputParameter(OUTPUT_TEAM_JSON, result.toString());
            setOutputParameter(OUTPUT_DISPLAY_NAME, result.path("displayName").asText());
            setOutputParameter(OUTPUT_DESCRIPTION, result.path("description").asText());
            setSuccessOutputs();
        } catch (MsTeamsException e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        }
    }
}
