package com.bonitasoft.connectors.msteams.createmeeting;

import java.util.List;

import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class MSTeamsCreateMeetingConnector extends AbstractMSTeamsConnector {

    public static final String INPUT_USER_ID = "userId";
    public static final String INPUT_SUBJECT = "subject";
    public static final String INPUT_START_DATE_TIME = "startDateTime";
    public static final String INPUT_END_DATE_TIME = "endDateTime";
    public static final String INPUT_ALLOWED_PRESENTERS = "allowedPresenters";
    public static final String INPUT_LOBBY_BYPASS_SCOPE = "lobbyBypassScope";
    public static final String INPUT_RECORD_AUTOMATICALLY = "recordAutomatically";

    public static final String OUTPUT_MEETING_ID = "meetingId";
    public static final String OUTPUT_JOIN_WEB_URL = "joinWebUrl";
    public static final String OUTPUT_MEETING_CODE = "meetingCode";
    public static final String OUTPUT_SUBJECT = "subject";
    public static final String OUTPUT_START_DATE_TIME = "startDateTime";
    public static final String OUTPUT_END_DATE_TIME = "endDateTime";

    @Override
    protected String getConnectorName() { return "MSTeams-CreateMeeting"; }

    @Override
    protected void validateOperationInputs(List<String> errors) {
        if (isBlank(getStringInput(INPUT_USER_ID))) errors.add("userId is required");
        if (isBlank(getStringInput(INPUT_SUBJECT))) errors.add("subject is required");
        if (isBlank(getStringInput(INPUT_START_DATE_TIME))) errors.add("startDateTime is required (ISO 8601 format)");
        if (isBlank(getStringInput(INPUT_END_DATE_TIME))) errors.add("endDateTime is required (ISO 8601 format)");
    }

    @Override
    public void executeOperation(MSTeamsClient client) throws MSTeamsException {
        String userId = getStringInput(INPUT_USER_ID);
        String subject = getStringInput(INPUT_SUBJECT);
        String startDateTime = getStringInput(INPUT_START_DATE_TIME);
        String endDateTime = getStringInput(INPUT_END_DATE_TIME);
        String allowedPresenters = getStringInputOrDefault(INPUT_ALLOWED_PRESENTERS, "everyone");
        String lobbyBypassScope = getStringInputOrDefault(INPUT_LOBBY_BYPASS_SCOPE, "organization");
        boolean recordAutomatically = getBooleanInputOrDefault(INPUT_RECORD_AUTOMATICALLY, false);

        log.info("Creating online meeting '{}' for user {}", subject, userId);

        try {
            ObjectMapper mapper = client.getObjectMapper();
            ObjectNode body = mapper.createObjectNode();
            body.put("subject", subject);
            body.put("startDateTime", startDateTime);
            body.put("endDateTime", endDateTime);
            body.put("allowedPresenters", allowedPresenters);
            body.put("recordAutomatically", recordAutomatically);

            ObjectNode lobbyBypass = mapper.createObjectNode();
            lobbyBypass.put("scope", lobbyBypassScope);
            body.set("lobbyBypassSettings", lobbyBypass);

            String path = "/users/" + userId + "/onlineMeetings";
            JsonNode response = client.post(path, mapper.writeValueAsString(body));

            setOutputParameter(OUTPUT_MEETING_ID, response.path("id").asText(null));
            setOutputParameter(OUTPUT_JOIN_WEB_URL, response.path("joinWebUrl").asText(null));
            setOutputParameter(OUTPUT_MEETING_CODE, response.path("joinMeetingIdSettings").path("joinMeetingId").asText(null));
            setOutputParameter(OUTPUT_SUBJECT, response.path("subject").asText(null));
            setOutputParameter(OUTPUT_START_DATE_TIME, response.path("startDateTime").asText(null));
            setOutputParameter(OUTPUT_END_DATE_TIME, response.path("endDateTime").asText(null));

            log.info("Meeting created successfully: {}", response.path("id").asText());
        } catch (MSTeamsException e) {
            throw e;
        } catch (Exception e) {
            throw new MSTeamsException("Failed to create meeting: " + e.getMessage(), e);
        }
    }
}
