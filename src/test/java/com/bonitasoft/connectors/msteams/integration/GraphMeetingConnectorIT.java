package com.bonitasoft.connectors.msteams.integration;

import static com.github.tomakehurst.wiremock.client.WireMock.*;
import static org.assertj.core.api.Assertions.assertThat;

import com.bonitasoft.connectors.msteams.common.MsTeamsGraphClient;
import com.bonitasoft.connectors.msteams.meetings.CreateMeetingConnector;
import com.bonitasoft.connectors.msteams.meetings.GetMeetingConnector;
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
class GraphMeetingConnectorIT {

    private MsTeamsGraphClient graphClient;

    @BeforeEach
    void setUp(WireMockRuntimeInfo wmInfo) {
        graphClient = new MsTeamsGraphClient(wmInfo.getHttpBaseUrl(), "fake-token", 30000, 0);
    }

    @Test
    @DisplayName("Should create online meeting")
    void should_create_meeting() throws Exception {
        stubFor(post(urlPathEqualTo("/users/user-1/onlineMeetings"))
                .willReturn(okJson("{\"id\":\"meet-1\",\"joinWebUrl\":\"https://teams.microsoft.com/l/meetup-join/123\",\"subject\":\"Sprint Review\"}")));

        var connector = new CreateMeetingConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("refreshToken", "rt");
        inputs.put("userId", "user-1");
        inputs.put("subject", "Sprint Review");
        inputs.put("startDateTime", "2025-01-20T14:00:00Z");
        inputs.put("endDateTime", "2025-01-20T15:00:00Z");
        connector.setInputParameters(inputs);
        connector.setGraphClient(graphClient);

        connector.executeForTest();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("meetingId")).isEqualTo("meet-1");
        assertThat((String) connector.getOutputs().get("joinUrl")).contains("meetup-join");
    }

    @Test
    @DisplayName("Should get meeting details")
    void should_get_meeting() throws Exception {
        stubFor(get(urlPathEqualTo("/users/user-1/onlineMeetings/meet-1"))
                .willReturn(okJson("{\"subject\":\"Sprint Review\",\"joinWebUrl\":\"https://teams.microsoft.com/l/meetup-join/123\"}")));

        var connector = new GetMeetingConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("refreshToken", "rt");
        inputs.put("userId", "user-1");
        inputs.put("meetingId", "meet-1");
        connector.setInputParameters(inputs);
        connector.setGraphClient(graphClient);

        connector.executeForTest();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("subject")).isEqualTo("Sprint Review");
    }
}
