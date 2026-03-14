package com.bonitasoft.connectors.msteams.e2e;

import java.util.HashMap;
import java.util.Map;

import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;

import org.bonitasoft.engine.connector.ConnectorException;

/**
 * Helper class for end-to-end connector testing.
 * <p>
 * Provides utilities to configure, execute, and verify MS Teams connectors
 * with test data. Designed to work with both unit tests (mocked client)
 * and integration tests (real MS Teams tenant).
 */
public final class ConnectorTestToolkit {

    private ConnectorTestToolkit() {
        // utility class
    }

    /**
     * Creates a map of common authentication inputs required by all connectors.
     *
     * @param tenantId     Azure AD tenant ID
     * @param clientId     Azure AD application (client) ID
     * @param clientSecret Azure AD client secret value
     * @return map with tenantId, clientId, clientSecret keys
     */
    public static Map<String, Object> authInputs(String tenantId, String clientId, String clientSecret) {
        var inputs = new HashMap<String, Object>();
        inputs.put("tenantId", tenantId);
        inputs.put("clientId", clientId);
        inputs.put("clientSecret", clientSecret);
        return inputs;
    }

    /**
     * Creates authentication inputs including delegated auth parameters.
     *
     * @param tenantId     Azure AD tenant ID
     * @param clientId     Azure AD application (client) ID
     * @param clientSecret Azure AD client secret value
     * @param userId       User ID for delegated operations
     * @return map with auth keys including userId
     */
    public static Map<String, Object> delegatedAuthInputs(String tenantId, String clientId,
                                                           String clientSecret, String userId) {
        var inputs = authInputs(tenantId, clientId, clientSecret);
        inputs.put("userId", userId);
        return inputs;
    }

    /**
     * Merges multiple input maps into a single map.
     * Later maps override earlier ones for duplicate keys.
     *
     * @param maps input maps to merge
     * @return merged map
     */
    @SafeVarargs
    public static Map<String, Object> mergeInputs(Map<String, Object>... maps) {
        var merged = new HashMap<String, Object>();
        for (var map : maps) {
            merged.putAll(map);
        }
        return merged;
    }

    /**
     * Creates operation-specific inputs for Send Channel Message.
     */
    public static Map<String, Object> sendChannelMessageInputs(String teamId, String channelId,
                                                                 String message) {
        return Map.of(
                "teamId", teamId,
                "channelId", channelId,
                "messageContent", message,
                "contentType", "text"
        );
    }

    /**
     * Creates operation-specific inputs for Send Chat Message.
     */
    public static Map<String, Object> sendChatMessageInputs(String chatId, String message) {
        return Map.of(
                "chatId", chatId,
                "messageContent", message
        );
    }

    /**
     * Creates operation-specific inputs for Reply Message.
     */
    public static Map<String, Object> replyMessageInputs(String teamId, String channelId,
                                                           String parentMessageId, String message) {
        return Map.of(
                "teamId", teamId,
                "channelId", channelId,
                "parentMessageId", parentMessageId,
                "messageContent", message
        );
    }

    /**
     * Creates operation-specific inputs for Send Adaptive Card.
     */
    public static Map<String, Object> sendAdaptiveCardInputs(String teamId, String channelId,
                                                               String cardJson) {
        return Map.of(
                "teamId", teamId,
                "channelId", channelId,
                "cardJson", cardJson
        );
    }

    /**
     * Creates operation-specific inputs for Create Channel.
     */
    public static Map<String, Object> createChannelInputs(String teamId, String displayName,
                                                            String description) {
        return Map.of(
                "teamId", teamId,
                "displayName", displayName,
                "description", description
        );
    }

    /**
     * Creates operation-specific inputs for List Channels.
     */
    public static Map<String, Object> listChannelsInputs(String teamId) {
        return Map.of("teamId", teamId);
    }

    /**
     * Creates operation-specific inputs for Delete Channel.
     */
    public static Map<String, Object> deleteChannelInputs(String teamId, String channelId) {
        return Map.of(
                "teamId", teamId,
                "channelId", channelId
        );
    }

    /**
     * Creates operation-specific inputs for List Teams (no additional inputs required).
     */
    public static Map<String, Object> listTeamsInputs() {
        return Map.of();
    }

    /**
     * Creates operation-specific inputs for Get Team.
     */
    public static Map<String, Object> getTeamInputs(String teamId) {
        return Map.of("teamId", teamId);
    }

    /**
     * Creates operation-specific inputs for Add Member.
     */
    public static Map<String, Object> addMemberInputs(String teamId, String userPrincipalName) {
        return Map.of(
                "teamId", teamId,
                "userPrincipalName", userPrincipalName
        );
    }

    /**
     * Creates operation-specific inputs for Remove Member.
     */
    public static Map<String, Object> removeMemberInputs(String teamId, String membershipId) {
        return Map.of(
                "teamId", teamId,
                "membershipId", membershipId
        );
    }

    /**
     * Creates operation-specific inputs for Create Meeting.
     */
    public static Map<String, Object> createMeetingInputs(String subject, String startDateTime,
                                                            String endDateTime) {
        return Map.of(
                "subject", subject,
                "startDateTime", startDateTime,
                "endDateTime", endDateTime
        );
    }

    /**
     * Creates operation-specific inputs for Get Meeting.
     */
    public static Map<String, Object> getMeetingInputs(String meetingId) {
        return Map.of("meetingId", meetingId);
    }

    /**
     * Creates operation-specific inputs for Upload File.
     */
    public static Map<String, Object> uploadFileInputs(String teamId, String channelId,
                                                         String fileName, String fileContentBase64) {
        return Map.of(
                "teamId", teamId,
                "channelId", channelId,
                "fileName", fileName,
                "fileContent", fileContentBase64
        );
    }

    /**
     * Configures a connector with the given inputs and executes the full Bonita lifecycle:
     * validate -> connect -> executeBusinessLogic -> disconnect.
     * <p>
     * Uses {@link AbstractMSTeamsConnector#execute()} which is {@code public final}
     * on {@code AbstractConnector} and returns the output parameters directly.
     *
     * @param connector the connector instance to execute
     * @param inputs    map of input parameter name to value
     * @return the connector's output parameters
     * @throws ConnectorException if connector execution fails
     */
    public static Map<String, Object> executeConnector(AbstractMSTeamsConnector connector,
                                                        Map<String, Object> inputs) throws ConnectorException {
        connector.setInputParameters(inputs);
        // execute() is public final on AbstractConnector:
        // validates inputs, connects, runs business logic, disconnects, returns outputs
        return connector.execute();
    }

    /**
     * Checks if the connector output indicates success.
     *
     * @param outputs connector output map
     * @return true if the "success" output is Boolean.TRUE
     */
    public static boolean isSuccess(Map<String, Object> outputs) {
        Object success = outputs.get("success");
        return Boolean.TRUE.equals(success);
    }

    /**
     * Gets the error message from connector outputs.
     *
     * @param outputs connector output map
     * @return the errorMessage value, or null
     */
    public static String getErrorMessage(Map<String, Object> outputs) {
        Object msg = outputs.get("errorMessage");
        return msg != null ? msg.toString() : null;
    }
}
