package com.bonitasoft.connectors.msteams.files;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

import com.bonitasoft.connectors.msteams.common.MsTeamsGraphClient;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.bonitasoft.engine.connector.ConnectorException;
import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.HashMap;
import java.util.Map;

@ExtendWith(MockitoExtension.class)
class UploadFileConnectorTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Mock
    private MsTeamsGraphClient graphClient;

    private UploadFileConnector connector;
    private Map<String, Object> inputs;

    @BeforeEach
    void setUp() {
        connector = new UploadFileConnector();
        inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("teamId", "team-1");
        inputs.put("channelId", "ch-1");
        inputs.put("fileName", "report.txt");
        inputs.put("fileContent", "Hello World");
        connector.setInputParameters(inputs);
    }

    @Test
    void should_pass_validation() throws ConnectorValidationException {
        connector.validateInputParameters();
    }

    @Test
    void should_fail_when_no_file_source() {
        var c = new UploadFileConnector();
        var incompleteInputs = new HashMap<>(inputs);
        incompleteInputs.remove("fileContent");
        c.setInputParameters(incompleteInputs);
        assertThatThrownBy(() -> c.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class);
    }

    @Test
    void should_upload_file_from_content() throws Exception {
        connector.setGraphClient(graphClient);
        when(graphClient.put(any(), any(byte[].class), eq("application/octet-stream")))
                .thenReturn(MAPPER.readTree("{\"id\":\"file-1\",\"webUrl\":\"https://sharepoint.com/file\"}"));

        connector.executeBusinessLogic();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("fileId")).isEqualTo("file-1");
    }
}
