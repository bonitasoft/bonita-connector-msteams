package com.bonitasoft.connectors.msteams.uploadfile;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.when;

import java.util.Base64;
import java.util.HashMap;
import java.util.Map;

import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

@DisplayName("MSTeamsUploadFileConnector")
@ExtendWith(MockitoExtension.class)
class MSTeamsUploadFileConnectorTest {

    private MSTeamsUploadFileConnector connector;

    @Mock
    private MSTeamsClient client;

    private final ObjectMapper objectMapper = new ObjectMapper();

    @BeforeEach
    void setUp() {
        connector = new MSTeamsUploadFileConnector();
    }

    @Nested
    @DisplayName("Validation")
    class Validation {

        @Test
        @DisplayName("should fail when teamId is missing")
        void should_fail_when_teamId_missing() {
            Map<String, Object> params = validParams();
            params.remove("teamId");

            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("teamId");
        }

        @Test
        @DisplayName("should fail when channelId is missing")
        void should_fail_when_channelId_missing() {
            Map<String, Object> params = validParams();
            params.remove("channelId");

            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("channelId");
        }

        @Test
        @DisplayName("should fail when fileName is missing")
        void should_fail_when_fileName_missing() {
            Map<String, Object> params = validParams();
            params.remove("fileName");

            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("fileName");
        }

        @Test
        @DisplayName("should fail when fileName exceeds max length")
        void should_fail_when_fileName_too_long() {
            Map<String, Object> params = validParams();
            params.put("fileName", "a".repeat(256));

            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("maximum length");
        }

        @Test
        @DisplayName("should fail when fileContent is invalid Base64")
        void should_fail_when_fileContent_invalid_base64() {
            Map<String, Object> params = validParams();
            params.put("fileContent", "not-valid-base64!!!");

            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("Base64");
        }

        @Test
        @DisplayName("should fail when fileContent is missing")
        void should_fail_when_fileContent_missing() {
            Map<String, Object> params = validParams();
            params.remove("fileContent");

            connector.setInputParameters(params);

            assertThatThrownBy(() -> connector.validateInputParameters())
                    .isInstanceOf(ConnectorValidationException.class)
                    .hasMessageContaining("fileContent");
        }

        @Test
        @DisplayName("should pass with all required parameters")
        void should_pass_when_all_required_params_present() throws ConnectorValidationException {
            connector.setInputParameters(validParams());

            connector.validateInputParameters(); // Should not throw
        }

        private Map<String, Object> validParams() {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("teamId", "team-123");
            params.put("channelId", "channel-456");
            params.put("fileName", "report.pdf");
            params.put("fileContent", Base64.getEncoder().encodeToString("test content".getBytes()));
            return params;
        }
    }

    @Nested
    @DisplayName("Execution")
    class Execution {

        @Test
        @DisplayName("should upload file and set outputs")
        void should_upload_file_when_valid_inputs() throws Exception {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("teamId", "team-123");
            params.put("channelId", "channel-456");
            params.put("fileName", "report.pdf");
            params.put("fileContent", Base64.getEncoder().encodeToString("test file content".getBytes()));
            params.put("contentType", "application/pdf");

            connector.setInputParameters(params);

            String filesFolderJson = """
                    {
                        "id": "folder-789",
                        "parentReference": {
                            "driveId": "drive-abc"
                        }
                    }
                    """;
            JsonNode filesFolderNode = objectMapper.readTree(filesFolderJson);

            String uploadJson = """
                    {
                        "id": "file-xyz-123",
                        "name": "report.pdf",
                        "webUrl": "https://teams.sharepoint.com/sites/team/report.pdf",
                        "size": 17
                    }
                    """;
            JsonNode uploadNode = objectMapper.readTree(uploadJson);

            when(client.get(anyString())).thenReturn(filesFolderNode);
            when(client.put(anyString(), any(byte[].class), anyString())).thenReturn(uploadNode);

            connector.executeOperation(client);

            assertThat(connector.getOutputs().get("fileId")).isEqualTo("file-xyz-123");
            assertThat(connector.getOutputs().get("fileName")).isEqualTo("report.pdf");
            assertThat(connector.getOutputs().get("webUrl")).isEqualTo("https://teams.sharepoint.com/sites/team/report.pdf");
            assertThat(connector.getOutputs().get("size")).isEqualTo(17L);
        }

        @Test
        @DisplayName("should handle API error on filesFolder")
        void should_handle_api_error_on_files_folder() {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("teamId", "team-123");
            params.put("channelId", "channel-456");
            params.put("fileName", "report.pdf");
            params.put("fileContent", Base64.getEncoder().encodeToString("test".getBytes()));
            connector.setInputParameters(params);

            when(client.get(anyString())).thenThrow(new MSTeamsException("error", 404, "NotFound", false, null));
            assertThatThrownBy(() -> connector.executeOperation(client)).isInstanceOf(MSTeamsException.class);
        }

        @Test
        @DisplayName("should handle missing driveId in response")
        void should_handle_missing_driveId() throws Exception {
            Map<String, Object> params = new HashMap<>();
            params.put("tenantId", "test-tenant");
            params.put("clientId", "test-client-id");
            params.put("clientSecret", "test-secret");
            params.put("teamId", "team-123");
            params.put("channelId", "channel-456");
            params.put("fileName", "report.pdf");
            params.put("fileContent", Base64.getEncoder().encodeToString("test".getBytes()));
            connector.setInputParameters(params);

            when(client.get(anyString())).thenReturn(objectMapper.readTree("{\"id\":\"f\",\"parentReference\":{}}"));
            assertThatThrownBy(() -> connector.executeOperation(client)).isInstanceOf(MSTeamsException.class).hasMessageContaining("driveId");
        }

        @Test
        @DisplayName("should return correct connector name")
        void should_return_correct_connector_name() {
            assertThat(connector.getConnectorName()).isEqualTo("MSTeams-UploadFile");
        }
    }
}
