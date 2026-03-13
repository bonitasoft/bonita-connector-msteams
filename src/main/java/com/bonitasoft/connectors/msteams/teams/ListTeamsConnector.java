package com.bonitasoft.connectors.msteams.teams;

import com.bonitasoft.connectors.msteams.common.AbstractGraphConnector;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import org.bonitasoft.engine.connector.ConnectorException;
import java.util.List;

/**
 * List all teams in the organization via Graph API (app-only).
 * Connector ID: msteams-list-teams
 */
public class ListTeamsConnector extends AbstractGraphConnector {

    public static final String OUTPUT_TEAMS_JSON = "teamsJson";
    public static final String OUTPUT_TEAM_COUNT = "teamCount";

    @Override
    protected void validateOperationParameters(List<String> errors) {
        // No operation-specific parameters
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            List<JsonNode> teams = graphClient.getWithPagination("/groups?$filter=resourceProvisioningOptions/Any(x:x%20eq%20'Team')");
            String json = new com.fasterxml.jackson.databind.ObjectMapper().writeValueAsString(teams);
            setOutputParameter(OUTPUT_TEAMS_JSON, json);
            setOutputParameter(OUTPUT_TEAM_COUNT, teams.size());
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
