package com.bonitasoft.connectors.msteams.getmeeting;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.JsonNode;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class MSTeamsGetMeetingConnector extends AbstractMSTeamsConnector {

    public static final String INPUT_USER_ID = "userId";
    public static final String INPUT_MEETING_ID = "meetingId";

    public static final String OUTPUT_MEETING_ID = "meetingId";
    public static final String OUTPUT_JOIN_WEB_URL = "joinWebUrl";
    public static final String OUTPUT_SUBJECT = "subject";
    public static final String OUTPUT_START_DATE_TIME = "startDateTime";
    public static final String OUTPUT_END_DATE_TIME = "endDateTime";
    public static final String OUTPUT_PARTICIPANTS = "participants";

    @Override
    protected String getConnectorName() { return "MSTeams-GetMeeting"; }

    @Override
    protected void validateOperationInputs(List<String> errors) {
        if (isBlank(getStringInput(INPUT_USER_ID))) errors.add("userId is required");
        if (isBlank(getStringInput(INPUT_MEETING_ID))) errors.add("meetingId is required");
    }

    @Override
    @SuppressWarnings("unchecked")
    public void executeOperation(MSTeamsClient client) throws MSTeamsException {
        String userId = getStringInput(INPUT_USER_ID);
        String meetingId = getStringInput(INPUT_MEETING_ID);

        log.info("Retrieving meeting {} for user {}", meetingId, userId);

        String path = "/users/" + userId + "/onlineMeetings/" + meetingId;
        JsonNode response = client.get(path);

        setOutputParameter(OUTPUT_MEETING_ID, response.path("id").asText(null));
        setOutputParameter(OUTPUT_JOIN_WEB_URL, response.path("joinWebUrl").asText(null));
        setOutputParameter(OUTPUT_SUBJECT, response.path("subject").asText(null));
        setOutputParameter(OUTPUT_START_DATE_TIME, response.path("startDateTime").asText(null));
        setOutputParameter(OUTPUT_END_DATE_TIME, response.path("endDateTime").asText(null));

        Map<String, Object> participants = new HashMap<>();
        JsonNode participantsNode = response.path("participants");
        if (!participantsNode.isMissingNode()) {
            participants = client.getObjectMapper().convertValue(participantsNode, Map.class);
        }
        setOutputParameter(OUTPUT_PARTICIPANTS, participants);

        log.info("Meeting retrieved: {}", meetingId);
    }
}
