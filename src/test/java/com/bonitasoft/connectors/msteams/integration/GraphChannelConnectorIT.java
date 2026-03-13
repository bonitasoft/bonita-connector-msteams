package com.bonitasoft.connectors.msteams.integration;

import static com.github.tomakehurst.wiremock.client.WireMock.*;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import com.bonitasoft.connectors.msteams.channels.CreateChannelConnector;
import com.bonitasoft.connectors.msteams.channels.DeleteChannelConnector;
import com.bonitasoft.connectors.msteams.channels.ListChannelsConnector;
import com.bonitasoft.connectors.msteams.common.MsTeamsGraphClient;
import com.github.tomakehurst.wiremock.junit5.WireMockRuntimeInfo;
import com.github.tomakehurst.wiremock.junit5.WireMockTest;
import org.bonitasoft.engine.connector.ConnectorException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;

@Tag("integration")
@WireMockTest
class GraphChannelConnectorIT {

    private MsTeamsGraphClient graphClient;

    @BeforeEach
    void setUp(WireMockRuntimeInfo wmInfo) {
        graphClient = new MsTeamsGraphClient(wmInfo.getHttpBaseUrl(), "fake-token", 30000, 0);
    }

    @Test
    @DisplayName("Should list channels with pagination")
    void should_list_channels_with_pagination(WireMockRuntimeInfo wmInfo) throws Exception {
        // Page 1 — nextLink points to WireMock for page 2
        String nextLink = wmInfo.getHttpBaseUrl() + "/teams/team-1/channels?$skip=1";
        stubFor(get(urlPathEqualTo("/teams/team-1/channels"))
                .withQueryParam("$skip", absent())
                .willReturn(okJson("{\"value\":[{\"id\":\"ch-1\",\"displayName\":\"General\"}],\"@odata.nextLink\":\"" + nextLink + "\"}")));
        // Page 2
        stubFor(get(urlPathEqualTo("/teams/team-1/channels"))
                .withQueryParam("$skip", equalTo("1"))
                .willReturn(okJson("{\"value\":[{\"id\":\"ch-2\",\"displayName\":\"Dev\"}]}")));

        var connector = new ListChannelsConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("teamId", "team-1");
        connector.setInputParameters(inputs);
        connector.setGraphClient(graphClient);

        connector.executeForTest();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat((String) connector.getOutputs().get("channelsJson")).contains("General");
        assertThat(connector.getOutputs().get("channelCount")).isEqualTo(2);
    }

    @Test
    @DisplayName("Should create channel successfully")
    void should_create_channel() throws Exception {
        stubFor(post(urlPathEqualTo("/teams/team-1/channels"))
                .willReturn(okJson("{\"id\":\"ch-new\",\"webUrl\":\"https://teams.microsoft.com/channel/ch-new\"}")));

        var connector = new CreateChannelConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("teamId", "team-1");
        inputs.put("displayName", "New Channel");
        inputs.put("description", "Test channel");
        connector.setInputParameters(inputs);
        connector.setGraphClient(graphClient);

        connector.executeForTest();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("channelId")).isEqualTo("ch-new");
        verify(postRequestedFor(urlPathEqualTo("/teams/team-1/channels"))
                .withRequestBody(containing("New Channel")));
    }

    @Test
    @DisplayName("Should delete channel successfully")
    void should_delete_channel() throws Exception {
        stubFor(delete(urlPathEqualTo("/teams/team-1/channels/ch-del"))
                .willReturn(aResponse().withStatus(204)));

        var connector = new DeleteChannelConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("teamId", "team-1");
        inputs.put("channelId", "ch-del");
        connector.setInputParameters(inputs);
        connector.setGraphClient(graphClient);

        connector.executeForTest();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
    }

    @Test
    @DisplayName("Should handle 404 error from Graph API")
    void should_handle_not_found() {
        stubFor(get(urlPathEqualTo("/teams/bad-team/channels"))
                .willReturn(aResponse().withStatus(404)
                        .withHeader("Content-Type", "application/json")
                        .withBody("{\"error\":{\"code\":\"NotFound\",\"message\":\"Team not found\"}}")));

        var connector = new ListChannelsConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("teamId", "bad-team");
        connector.setInputParameters(inputs);
        connector.setGraphClient(graphClient);

        assertThatThrownBy(() -> connector.executeForTest())
                .isInstanceOf(ConnectorException.class);
        assertThat(connector.getOutputs().get("success")).isEqualTo(false);
        assertThat((String) connector.getOutputs().get("errorMessage")).contains("NotFound");
    }
}
