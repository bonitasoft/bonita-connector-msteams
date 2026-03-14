package com.bonitasoft.connectors.msteams.uploadfile;

import java.util.Base64;
import java.util.List;

import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.JsonNode;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class MSTeamsUploadFileConnector extends AbstractMSTeamsConnector {

    static final int MAX_FILE_NAME_LENGTH = 255;

    public static final String INPUT_TEAM_ID = "teamId";
    public static final String INPUT_CHANNEL_ID = "channelId";
    public static final String INPUT_FILE_NAME = "fileName";
    public static final String INPUT_FILE_CONTENT = "fileContent";
    public static final String INPUT_CONTENT_TYPE = "contentType";

    public static final String OUTPUT_FILE_ID = "fileId";
    public static final String OUTPUT_FILE_NAME = "fileName";
    public static final String OUTPUT_WEB_URL = "webUrl";
    public static final String OUTPUT_SIZE = "size";

    @Override
    protected String getConnectorName() { return "MSTeams-UploadFile"; }

    @Override
    protected void validateOperationInputs(List<String> errors) {
        if (isBlank(getStringInput(INPUT_TEAM_ID))) errors.add("teamId is required");
        if (isBlank(getStringInput(INPUT_CHANNEL_ID))) errors.add("channelId is required");

        String fileName = getStringInput(INPUT_FILE_NAME);
        if (isBlank(fileName)) {
            errors.add("fileName is required");
        } else if (fileName.length() > MAX_FILE_NAME_LENGTH) {
            errors.add("fileName exceeds maximum length (" + MAX_FILE_NAME_LENGTH + " chars)");
        }

        String fileContent = getStringInput(INPUT_FILE_CONTENT);
        if (isBlank(fileContent)) {
            errors.add("fileContent is required (Base64 encoded)");
        } else {
            try {
                Base64.getDecoder().decode(fileContent);
            } catch (IllegalArgumentException e) {
                errors.add("fileContent must be valid Base64 encoded data");
            }
        }
    }

    @Override
    public void executeOperation(MSTeamsClient client) throws MSTeamsException {
        String teamId = getStringInput(INPUT_TEAM_ID);
        String channelId = getStringInput(INPUT_CHANNEL_ID);
        String fileName = getStringInput(INPUT_FILE_NAME);
        String fileContent = getStringInput(INPUT_FILE_CONTENT);
        String contentType = getStringInputOrDefault(INPUT_CONTENT_TYPE, "application/octet-stream");

        log.info("Uploading file '{}' to team {} channel {}", fileName, teamId, channelId);

        // Step 1: Get the channel's filesFolder
        String filesFolderPath = "/teams/" + teamId + "/channels/" + channelId + "/filesFolder";
        JsonNode folderJson = client.get(filesFolderPath);
        String driveId = folderJson.path("parentReference").path("driveId").asText(null);
        String folderId = folderJson.path("id").asText(null);

        if (driveId == null || folderId == null) {
            throw new MSTeamsException("Could not resolve driveId or folderId from filesFolder response");
        }
        log.debug("Resolved driveId={}, folderId={}", driveId, folderId);

        // Step 2: Decode Base64 content
        byte[] contentBytes = Base64.getDecoder().decode(fileContent);
        log.debug("File size: {} bytes", contentBytes.length);

        // Step 3: Upload file via PUT
        String uploadPath = "/drives/" + driveId + "/items/" + folderId + ":/" + fileName + ":/content";
        JsonNode uploadJson = client.put(uploadPath, contentBytes, contentType);

        setOutputParameter(OUTPUT_FILE_ID, uploadJson.path("id").asText(null));
        setOutputParameter(OUTPUT_FILE_NAME, uploadJson.path("name").asText(null));
        setOutputParameter(OUTPUT_WEB_URL, uploadJson.path("webUrl").asText(null));
        setOutputParameter(OUTPUT_SIZE, uploadJson.path("size").asLong(0L));

        log.info("File uploaded: {} (ID: {})", uploadJson.path("name").asText(), uploadJson.path("id").asText());
    }
}
