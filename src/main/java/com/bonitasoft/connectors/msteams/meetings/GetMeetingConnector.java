package com.bonitasoft.connectors.msteams.meetings;

import com.bonitasoft.connectors.msteams.common.AbstractGraphConnector;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import org.bonitasoft.engine.connector.ConnectorException;
import java.util.List;

/**
 * Get online meeting details via Graph API (delegated permissions).
 * Connector ID: msteams-get-meeting
 */
public class GetMeetingConnector extends AbstractGraphConnector {

    public static final String INPUT_USER_ID = "userId";
    public static final String INPUT_MEETING_ID = "meetingId";
    public static final String OUTPUT_MEETING_JSON = "meetingJson";
    public static final String OUTPUT_JOIN_URL = "joinUrl";
    public static final String OUTPUT_SUBJECT = "subject";

    @Override
    protected String getAuthMode() {
        return "delegated";
    }

    @Override
    protected void validateOperationParameters(List<String> errors) {
        if (isNullOrEmpty(getStringInput(INPUT_USER_ID))) errors.add("userId is required");
        if (isNullOrEmpty(getStringInput(INPUT_MEETING_ID))) errors.add("meetingId is required");
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            String userId = getStringInput(INPUT_USER_ID);
            String meetingId = getStringInput(INPUT_MEETING_ID);

            JsonNode result = graphClient.get("/users/" + userId + "/onlineMeetings/" + meetingId);
            setOutputParameter(OUTPUT_MEETING_JSON, result.toString());
            setOutputParameter(OUTPUT_JOIN_URL, result.path("joinWebUrl").asText());
            setOutputParameter(OUTPUT_SUBJECT, result.path("subject").asText());
            setSuccessOutputs();
        } catch (MsTeamsException e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        }
    }
}
