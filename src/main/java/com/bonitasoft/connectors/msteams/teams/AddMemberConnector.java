package com.bonitasoft.connectors.msteams.teams;

import com.bonitasoft.connectors.msteams.common.AbstractGraphConnector;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import org.bonitasoft.engine.connector.ConnectorException;
import java.util.List;
import java.util.Map;

/**
 * Add a member to a Teams team via Graph API (app-only).
 * Connector ID: msteams-add-member
 */
public class AddMemberConnector extends AbstractGraphConnector {

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String INPUT_USER_ID = "userId";
    public static final String INPUT_ROLE = "role";
    public static final String OUTPUT_MEMBERSHIP_ID = "membershipId";

    @Override
    protected void validateOperationParameters(List<String> errors) {
        if (isNullOrEmpty(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
        if (isNullOrEmpty(getStringInput(INPUT_USER_ID))) errors.add("userId is required");
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            String teamId = getStringInput(INPUT_TEAM_ID);
            String userId = getStringInput(INPUT_USER_ID);
            String role = getStringInput(INPUT_ROLE);

            Map<String, Object> body = Map.of(
                    "@odata.type", "#microsoft.graph.aadUserConversationMember",
                    "roles", role != null ? List.of(role) : List.of(),
                    "user@odata.bind", "https://graph.microsoft.com/v1.0/users('" + userId + "')"
            );

            JsonNode result = graphClient.post("/teams/" + teamId + "/members", body);
            setOutputParameter(OUTPUT_MEMBERSHIP_ID, result.path("id").asText());
            setSuccessOutputs();
        } catch (MsTeamsException e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        }
    }
}
