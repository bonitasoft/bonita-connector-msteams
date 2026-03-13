package com.bonitasoft.connectors.msteams.files;

import com.bonitasoft.connectors.msteams.common.AbstractGraphConnector;
import com.bonitasoft.connectors.msteams.common.MsTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import org.bonitasoft.engine.connector.ConnectorException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.net.URLEncoder;
import java.util.List;

/**
 * Upload a file to a Teams channel's SharePoint folder via Graph API (app-only).
 * Connector ID: msteams-upload-file
 */
public class UploadFileConnector extends AbstractGraphConnector {

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String INPUT_CHANNEL_ID = "channelId";
    public static final String INPUT_FILE_PATH = "filePath";
    public static final String INPUT_FILE_NAME = "fileName";
    public static final String INPUT_FILE_CONTENT = "fileContent";
    public static final String OUTPUT_FILE_ID = "fileId";
    public static final String OUTPUT_WEB_URL = "webUrl";

    @Override
    protected void validateOperationParameters(List<String> errors) {
        if (isNullOrEmpty(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
        if (isNullOrEmpty(getStringInput(INPUT_CHANNEL_ID))) errors.add("channelId is required");
        if (isNullOrEmpty(getStringInput(INPUT_FILE_NAME))) errors.add("fileName is required");
        // Either filePath or fileContent must be provided
        if (isNullOrEmpty(getStringInput(INPUT_FILE_PATH)) && getInputParameter(INPUT_FILE_CONTENT) == null) {
            errors.add("Either filePath or fileContent is required");
        }
    }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        try {
            String teamId = getStringInput(INPUT_TEAM_ID);
            String channelId = getStringInput(INPUT_CHANNEL_ID);
            String fileName = getStringInput(INPUT_FILE_NAME);

            byte[] content;
            String filePath = getStringInput(INPUT_FILE_PATH);
            if (!isNullOrEmpty(filePath)) {
                content = Files.readAllBytes(Path.of(filePath));
            } else {
                Object fileContent = getInputParameter(INPUT_FILE_CONTENT);
                if (fileContent instanceof byte[] bytes) {
                    content = bytes;
                } else if (fileContent instanceof String s) {
                    content = s.getBytes(StandardCharsets.UTF_8);
                } else {
                    throw new MsTeamsException("fileContent must be byte[] or String");
                }
            }

            String path = String.format("/teams/%s/channels/%s/filesFolder/content?fileName=%s",
                    teamId, channelId, URLEncoder.encode(fileName, StandardCharsets.UTF_8));

            JsonNode result = graphClient.put(path, content, "application/octet-stream");
            setOutputParameter(OUTPUT_FILE_ID, result.path("id").asText());
            setOutputParameter(OUTPUT_WEB_URL, result.path("webUrl").asText());
            setSuccessOutputs();
        } catch (MsTeamsException e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        } catch (Exception e) {
            setErrorOutputs(e.getMessage());
            throw new ConnectorException(e.getMessage(), e);
        }
    }
}
