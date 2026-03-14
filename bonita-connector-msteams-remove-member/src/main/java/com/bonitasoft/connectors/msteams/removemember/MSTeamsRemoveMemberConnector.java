package com.bonitasoft.connectors.msteams.removemember;

import java.util.List;
import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import lombok.extern.slf4j.Slf4j;

@Slf4j
public class MSTeamsRemoveMemberConnector extends AbstractMSTeamsConnector {

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String INPUT_MEMBERSHIP_ID = "membershipId";

    @Override protected String getConnectorName() { return "MSTeams-RemoveMember"; }

    @Override
    protected void validateOperationInputs(List<String> errors) {
        if (isBlank(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
        if (isBlank(getStringInput(INPUT_MEMBERSHIP_ID))) errors.add("membershipId is required");
    }

    @Override
    protected void executeOperation(MSTeamsClient client) throws MSTeamsException {
        String teamId = getStringInput(INPUT_TEAM_ID);
        String membershipId = getStringInput(INPUT_MEMBERSHIP_ID);
        log.info("Removing member {} from team {}", membershipId, teamId);
        client.delete("/teams/" + teamId + "/members/" + membershipId);
        log.info("Member {} removed from team {}", membershipId, teamId);
    }
}
