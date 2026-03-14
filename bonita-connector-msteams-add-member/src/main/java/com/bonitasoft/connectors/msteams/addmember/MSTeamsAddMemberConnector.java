package com.bonitasoft.connectors.msteams.addmember;

import java.util.List;
import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import lombok.extern.slf4j.Slf4j;

@Slf4j
public class MSTeamsAddMemberConnector extends AbstractMSTeamsConnector {

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String INPUT_USER_PRINCIPAL_NAME = "userPrincipalName";
    public static final String INPUT_ROLE = "role";
    public static final String OUTPUT_MEMBERSHIP_ID = "membershipId";
    public static final String OUTPUT_DISPLAY_NAME = "displayName";

    @Override protected String getConnectorName() { return "MSTeams-AddMember"; }

    @Override
    protected void validateOperationInputs(List<String> errors) {
        if (isBlank(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
        if (isBlank(getStringInput(INPUT_USER_PRINCIPAL_NAME))) errors.add("userPrincipalName is required");
    }

    @Override
    protected void executeOperation(MSTeamsClient client) throws MSTeamsException {
        String teamId = getStringInput(INPUT_TEAM_ID);
        String upn = getStringInput(INPUT_USER_PRINCIPAL_NAME);
        String role = getStringInputOrDefault(INPUT_ROLE, "member");
        log.info("Adding {} as {} to team {}", upn, role, teamId);

        String rolesArray = "owner".equalsIgnoreCase(role) ? "[\"owner\"]" : "[]";
        String body = "{\"@odata.type\":\"#microsoft.graph.aadUserConversationMember\","
                + "\"roles\":" + rolesArray + ","
                + "\"user@odata.bind\":\"https://graph.microsoft.com/v1.0/users('" + esc(upn) + "')\"}";

        JsonNode r = client.post("/teams/" + teamId + "/members", body);
        setOutputParameter(OUTPUT_MEMBERSHIP_ID, r.has("id") ? r.get("id").asText() : null);
        setOutputParameter(OUTPUT_DISPLAY_NAME, r.has("displayName") ? r.get("displayName").asText() : null);
        log.info("Member added successfully");
    }

    private String esc(String v) {
        if (v == null) return "";
        return v.replace("\\", "\\\\").replace("\"", "\\\"").replace("'", "\\'");
    }
}
