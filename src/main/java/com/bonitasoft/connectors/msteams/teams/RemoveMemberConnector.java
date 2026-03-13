package com.bonitasoft.connectors.msteams.teams;

import com.bonitasoft.connectors.msteams.common.AbstractGraphConnector;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import org.bonitasoft.engine.connector.ConnectorException;
import java.util.List;

/**
 * Remove a member from a Teams team via Graph API (app-only).
 * Connector ID: msteams-remove-member
 */
public class RemoveMemberConnector extends AbstractGraphConnector {

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String INPUT_MEMBERSHIP_ID = "membershipId";

    @Override
    protected void validateOperationParameters(List<String> errors) {
        if (isNullOrEmpty(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
        if (isNullOrEmpty(getStringInput(INPUT_MEMBERSHIP_ID))) errors.add("membershipId is required");
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            String teamId = getStringInput(INPUT_TEAM_ID);
            String membershipId = getStringInput(INPUT_MEMBERSHIP_ID);
            graphClient.delete("/teams/" + teamId + "/members/" + membershipId);
            setSuccessOutputs();
        } catch (MsTeamsException e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        }
    }
}
