package com.bonitasoft.connectors.msteams.createchannel;

import java.util.List;
import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import lombok.extern.slf4j.Slf4j;

@Slf4j
public class MSTeamsCreateChannelConnector extends AbstractMSTeamsConnector {

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String INPUT_DISPLAY_NAME = "displayName";
    public static final String INPUT_DESCRIPTION = "description";
    public static final String INPUT_MEMBERSHIP_TYPE = "membershipType";
    public static final String OUTPUT_CHANNEL_ID = "channelId";
    public static final String OUTPUT_DISPLAY_NAME = "displayName";
    public static final String OUTPUT_WEB_URL = "webUrl";
    private static final int MAX_DISPLAY_NAME_LENGTH = 50;

    @Override
    protected String getConnectorName() { return "MSTeams-CreateChannel"; }

    @Override
    protected void validateOperationInputs(List<String> errors) {
        if (isBlank(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
        String dn = getStringInput(INPUT_DISPLAY_NAME);
        if (isBlank(dn)) errors.add("displayName is required");
        else if (dn.length() > MAX_DISPLAY_NAME_LENGTH)
            errors.add("displayName exceeds maximum length (" + MAX_DISPLAY_NAME_LENGTH + " chars)");
    }

    @Override
    protected void executeOperation(MSTeamsClient client) throws MSTeamsException {
        String teamId = getStringInput(INPUT_TEAM_ID);
        String displayName = getStringInput(INPUT_DISPLAY_NAME);
        String description = getStringInput(INPUT_DESCRIPTION);
        String membershipType = getStringInputOrDefault(INPUT_MEMBERSHIP_TYPE, "standard");

        log.info("Creating channel '{}' in team {}", displayName, teamId);

        StringBuilder body = new StringBuilder();
        body.append("{\"displayName\":\"").append(esc(displayName)).append("\"");
        if (description != null && !description.isBlank())
            body.append(",\"description\":\"").append(esc(description)).append("\"");
        body.append(",\"membershipType\":\"").append(esc(membershipType)).append("\"}");

        JsonNode r = client.post("/teams/" + teamId + "/channels", body.toString());
        setOutputParameter(OUTPUT_CHANNEL_ID, r.has("id") ? r.get("id").asText() : null);
        setOutputParameter(OUTPUT_DISPLAY_NAME, r.has("displayName") ? r.get("displayName").asText() : null);
        setOutputParameter(OUTPUT_WEB_URL, r.has("webUrl") ? r.get("webUrl").asText() : null);
        log.info("Channel created successfully");
    }

    private String esc(String v) {
        if (v == null) return "";
        return v.replace("\\", "\\\\").replace("\"", "\\\"").replace("\n", "\\n").replace("\r", "\\r");
    }
}
