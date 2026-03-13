package com.bonitasoft.connectors.msteams.integration;

import static com.github.tomakehurst.wiremock.client.WireMock.*;
import static org.assertj.core.api.Assertions.assertThat;

import com.bonitasoft.connectors.msteams.common.MsTeamsGraphClient;
import com.bonitasoft.connectors.msteams.files.UploadFileConnector;
import com.github.tomakehurst.wiremock.junit5.WireMockRuntimeInfo;
import com.github.tomakehurst.wiremock.junit5.WireMockTest;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;

@Tag("integration")
@WireMockTest
class GraphFileConnectorIT {

    private MsTeamsGraphClient graphClient;

    @BeforeEach
    void setUp(WireMockRuntimeInfo wmInfo) {
        graphClient = new MsTeamsGraphClient(wmInfo.getHttpBaseUrl(), "fake-token", 30000, 0);
    }

    @Test
    @DisplayName("Should upload file to channel")
    void should_upload_file() throws Exception {
        stubFor(put(urlPathMatching("/teams/team-1/channels/ch-1/filesFolder/content.*"))
                .willReturn(okJson("{\"id\":\"file-1\",\"webUrl\":\"https://sharepoint.com/sites/team/file.txt\"}")));

        var connector = new UploadFileConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("teamId", "team-1");
        inputs.put("channelId", "ch-1");
        inputs.put("fileName", "report.txt");
        inputs.put("fileContent", "File content here");
        connector.setInputParameters(inputs);
        connector.setGraphClient(graphClient);

        connector.executeForTest();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("fileId")).isEqualTo("file-1");
    }
}
