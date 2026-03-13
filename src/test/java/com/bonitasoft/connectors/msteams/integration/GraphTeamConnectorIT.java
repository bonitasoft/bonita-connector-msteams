package com.bonitasoft.connectors.msteams.integration;

import static com.github.tomakehurst.wiremock.client.WireMock.*;
import static org.assertj.core.api.Assertions.assertThat;

import com.bonitasoft.connectors.msteams.common.MsTeamsGraphClient;
import com.bonitasoft.connectors.msteams.teams.AddMemberConnector;
import com.bonitasoft.connectors.msteams.teams.GetTeamConnector;
import com.bonitasoft.connectors.msteams.teams.ListTeamsConnector;
import com.bonitasoft.connectors.msteams.teams.RemoveMemberConnector;
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
class GraphTeamConnectorIT {

    private MsTeamsGraphClient graphClient;

    @BeforeEach
    void setUp(WireMockRuntimeInfo wmInfo) {
        graphClient = new MsTeamsGraphClient(wmInfo.getHttpBaseUrl(), "fake-token", 30000, 0);
    }

    @Test
    @DisplayName("Should list teams")
    void should_list_teams() throws Exception {
        stubFor(get(urlPathEqualTo("/groups"))
                .withQueryParam("$filter", containing("resourceProvisioningOptions"))
                .willReturn(okJson("{\"value\":[{\"id\":\"t-1\",\"displayName\":\"Engineering\"}]}")));

        var connector = new ListTeamsConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        connector.setInputParameters(inputs);
        connector.setGraphClient(graphClient);

        connector.executeForTest();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("teamCount")).isEqualTo(1);
    }

    @Test
    @DisplayName("Should get team details")
    void should_get_team() throws Exception {
        stubFor(get(urlPathEqualTo("/teams/team-1"))
                .willReturn(okJson("{\"displayName\":\"Engineering\",\"description\":\"Dev team\"}")));

        var connector = new GetTeamConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("teamId", "team-1");
        connector.setInputParameters(inputs);
        connector.setGraphClient(graphClient);

        connector.executeForTest();

        assertThat(connector.getOutputs().get("displayName")).isEqualTo("Engineering");
    }

    @Test
    @DisplayName("Should add member to team")
    void should_add_member() throws Exception {
        stubFor(post(urlPathEqualTo("/teams/team-1/members"))
                .willReturn(okJson("{\"id\":\"mem-1\"}")));

        var connector = new AddMemberConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("teamId", "team-1");
        inputs.put("userId", "user-1");
        inputs.put("role", "owner");
        connector.setInputParameters(inputs);
        connector.setGraphClient(graphClient);

        connector.executeForTest();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("membershipId")).isEqualTo("mem-1");
    }

    @Test
    @DisplayName("Should remove member from team")
    void should_remove_member() throws Exception {
        stubFor(delete(urlPathEqualTo("/teams/team-1/members/mem-1"))
                .willReturn(aResponse().withStatus(204)));

        var connector = new RemoveMemberConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("teamId", "team-1");
        inputs.put("membershipId", "mem-1");
        connector.setInputParameters(inputs);
        connector.setGraphClient(graphClient);

        connector.executeForTest();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
    }
}
