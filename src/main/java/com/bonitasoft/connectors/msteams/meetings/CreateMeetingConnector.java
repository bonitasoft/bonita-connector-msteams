package com.bonitasoft.connectors.msteams.meetings;

import com.bonitasoft.connectors.msteams.common.AbstractGraphConnector;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import org.bonitasoft.engine.connector.ConnectorException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Create an online meeting via Graph API (delegated permissions).
 * Connector ID: msteams-create-meeting
 */
public class CreateMeetingConnector extends AbstractGraphConnector {

    public static final String INPUT_USER_ID = "userId";
    public static final String INPUT_SUBJECT = "subject";
    public static final String INPUT_START_DATE_TIME = "startDateTime";
    public static final String INPUT_END_DATE_TIME = "endDateTime";
    public static final String OUTPUT_MEETING_ID = "meetingId";
    public static final String OUTPUT_JOIN_URL = "joinUrl";
    public static final String OUTPUT_MEETING_JSON = "meetingJson";

    @Override
    protected String getAuthMode() {
        return "delegated";
    }

    @Override
    protected void validateOperationParameters(List<String> errors) {
        if (isNullOrEmpty(getStringInput(INPUT_USER_ID))) errors.add("userId is required");
        if (isNullOrEmpty(getStringInput(INPUT_SUBJECT))) errors.add("subject is required");
        if (isNullOrEmpty(getStringInput(INPUT_START_DATE_TIME))) errors.add("startDateTime is required");
        if (isNullOrEmpty(getStringInput(INPUT_END_DATE_TIME))) errors.add("endDateTime is required");
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            String userId = getStringInput(INPUT_USER_ID);
            String subject = getStringInput(INPUT_SUBJECT);
            String startDateTime = getStringInput(INPUT_START_DATE_TIME);
            String endDateTime = getStringInput(INPUT_END_DATE_TIME);

            Map<String, Object> body = new HashMap<>();
            body.put("subject", subject);
            body.put("startDateTime", startDateTime);
            body.put("endDateTime", endDateTime);

            JsonNode result = graphClient.post("/users/" + userId + "/onlineMeetings", body);
            setOutputParameter(OUTPUT_MEETING_ID, result.path("id").asText());
            setOutputParameter(OUTPUT_JOIN_URL, result.path("joinWebUrl").asText());
            setOutputParameter(OUTPUT_MEETING_JSON, result.toString());
            setSuccessOutputs();
        } catch (MsTeamsException e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        }
    }
}
