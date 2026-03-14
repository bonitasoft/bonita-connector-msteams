package com.bonitasoft.connectors.msteams.integration;

import com.bonitasoft.connectors.msteams.sendchannelmessage.MSTeamsSendChannelMessageConnector;
import com.bonitasoft.connectors.msteams.sendchatmessage.MSTeamsSendChatMessageConnector;
import com.bonitasoft.connectors.msteams.replymessage.MSTeamsReplyMessageConnector;
import com.bonitasoft.connectors.msteams.sendadaptivecard.MSTeamsSendAdaptiveCardConnector;
import com.bonitasoft.connectors.msteams.createchannel.MSTeamsCreateChannelConnector;
import com.bonitasoft.connectors.msteams.listchannels.MSTeamsListChannelsConnector;
import com.bonitasoft.connectors.msteams.deletechannel.MSTeamsDeleteChannelConnector;
import com.bonitasoft.connectors.msteams.listteams.MSTeamsListTeamsConnector;
import com.bonitasoft.connectors.msteams.getteam.MSTeamsGetTeamConnector;
import com.bonitasoft.connectors.msteams.addmember.MSTeamsAddMemberConnector;
import com.bonitasoft.connectors.msteams.removemember.MSTeamsRemoveMemberConnector;
import com.bonitasoft.connectors.msteams.createmeeting.MSTeamsCreateMeetingConnector;
import com.bonitasoft.connectors.msteams.getmeeting.MSTeamsGetMeetingConnector;
import com.bonitasoft.connectors.msteams.uploadfile.MSTeamsUploadFileConnector;
import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsConfiguration;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.bonitasoft.connectors.msteams.RetryPolicy;
import com.bonitasoft.connectors.msteams.AdaptiveCardBuilder;

import org.junit.jupiter.api.*;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Integration test that verifies all 14 connector classes are accessible
 * from the -all module and can be instantiated correctly.
 * This validates the module assembly and class loading.
 */
@Tag("integration")
@DisplayName("All Module Assembly Integration Test")
class GraphConnectorIT {

    @Test
    @DisplayName("should instantiate all 14 connector classes")
    void should_instantiate_all_connectors() {
        assertThat(new MSTeamsSendChannelMessageConnector()).isNotNull();
        assertThat(new MSTeamsSendChatMessageConnector()).isNotNull();
        assertThat(new MSTeamsReplyMessageConnector()).isNotNull();
        assertThat(new MSTeamsSendAdaptiveCardConnector()).isNotNull();
        assertThat(new MSTeamsCreateChannelConnector()).isNotNull();
        assertThat(new MSTeamsListChannelsConnector()).isNotNull();
        assertThat(new MSTeamsDeleteChannelConnector()).isNotNull();
        assertThat(new MSTeamsListTeamsConnector()).isNotNull();
        assertThat(new MSTeamsGetTeamConnector()).isNotNull();
        assertThat(new MSTeamsAddMemberConnector()).isNotNull();
        assertThat(new MSTeamsRemoveMemberConnector()).isNotNull();
        assertThat(new MSTeamsCreateMeetingConnector()).isNotNull();
        assertThat(new MSTeamsGetMeetingConnector()).isNotNull();
        assertThat(new MSTeamsUploadFileConnector()).isNotNull();
    }

    @Test
    @DisplayName("should access common classes from all module")
    void should_access_common_classes() {
        assertThat(MSTeamsConfiguration.class).isNotNull();
        assertThat(MSTeamsException.class).isNotNull();
        assertThat(RetryPolicy.class).isNotNull();
        assertThat(AdaptiveCardBuilder.class).isNotNull();
    }

    @Test
    @DisplayName("all 14 connector classes should be distinct types")
    void should_have_14_distinct_connector_types() {
        var connectors = new AbstractMSTeamsConnector[] {
            new MSTeamsSendChannelMessageConnector(),
            new MSTeamsSendChatMessageConnector(),
            new MSTeamsReplyMessageConnector(),
            new MSTeamsSendAdaptiveCardConnector(),
            new MSTeamsCreateChannelConnector(),
            new MSTeamsListChannelsConnector(),
            new MSTeamsDeleteChannelConnector(),
            new MSTeamsListTeamsConnector(),
            new MSTeamsGetTeamConnector(),
            new MSTeamsAddMemberConnector(),
            new MSTeamsRemoveMemberConnector(),
            new MSTeamsCreateMeetingConnector(),
            new MSTeamsGetMeetingConnector(),
            new MSTeamsUploadFileConnector()
        };

        var types = java.util.Arrays.stream(connectors)
                .map(Object::getClass)
                .collect(java.util.stream.Collectors.toSet());

        assertThat(types).hasSize(14); // all unique types
    }

    @Test
    @DisplayName("should create MSTeamsConfiguration")
    void should_create_configuration() {
        var config = new MSTeamsConfiguration("tenant", "client", "secret", 5000, 10000, true);
        assertThat(config.tenantId()).isEqualTo("tenant");
        assertThat(config.isAppOnly()).isTrue();
    }

    @Test
    @DisplayName("should build adaptive card")
    void should_build_adaptive_card() {
        var card = new AdaptiveCardBuilder()
                .title("Test")
                .body("Body text")
                .fact("Key", "Value")
                .action("Click", "https://example.com")
                .build();
        assertThat(card.get("type").asText()).isEqualTo("message");
    }

    @Test
    @DisplayName("should validate retry policy")
    void should_validate_retry_policy() {
        assertThat(RetryPolicy.isRetryableStatusCode(429)).isTrue();
        assertThat(RetryPolicy.isRetryableStatusCode(200)).isFalse();
    }
}
