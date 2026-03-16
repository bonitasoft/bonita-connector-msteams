package com.bonitasoft.connectors.msteams.sendchannelmessage;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import java.util.HashMap;
import java.util.Map;

import net.jqwik.api.*;
import net.jqwik.api.constraints.*;

import org.bonitasoft.engine.connector.ConnectorValidationException;

/**
 * Property-based tests for MSTeamsSendChannelMessageConnector using jqwik.
 * Tests input validation with randomly generated data to find edge cases.
 */
class MSTeamsSendChannelMessagePropertyTest {

    private Map<String, Object> validAuthInputs() {
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("tenantId", "test-tenant-id");
        inputs.put("clientId", "test-client-id");
        inputs.put("clientSecret", "test-client-secret");
        return inputs;
    }

    @Property(tries = 100)
    @Label("Valid inputs should always pass validation")
    void should_pass_validation_when_all_inputs_valid(
            @ForAll @AlphaChars @StringLength(min = 1, max = 100) String teamId,
            @ForAll @AlphaChars @StringLength(min = 1, max = 100) String channelId,
            @ForAll @AlphaChars @StringLength(min = 1, max = 1000) String messageContent)
            throws ConnectorValidationException {
        var connector = new MSTeamsSendChannelMessageConnector();
        Map<String, Object> inputs = validAuthInputs();
        inputs.put("teamId", teamId);
        inputs.put("channelId", channelId);
        inputs.put("messageContent", messageContent);

        connector.setInputParameters(inputs);
        connector.validateInputParameters();
    }

    @Property(tries = 50)
    @Label("Null teamId should always fail validation")
    void should_fail_validation_when_teamId_null(
            @ForAll @AlphaChars @StringLength(min = 1, max = 100) String channelId,
            @ForAll @AlphaChars @StringLength(min = 1, max = 100) String messageContent) {
        var connector = new MSTeamsSendChannelMessageConnector();
        Map<String, Object> inputs = validAuthInputs();
        inputs.put("channelId", channelId);
        inputs.put("messageContent", messageContent);

        connector.setInputParameters(inputs);
        assertThatThrownBy(connector::validateInputParameters)
                .isInstanceOf(ConnectorValidationException.class);
    }

    @Property(tries = 50)
    @Label("Empty channelId should always fail validation")
    void should_fail_validation_when_channelId_empty(
            @ForAll @AlphaChars @StringLength(min = 1, max = 100) String teamId,
            @ForAll @AlphaChars @StringLength(min = 1, max = 100) String messageContent) {
        var connector = new MSTeamsSendChannelMessageConnector();
        Map<String, Object> inputs = validAuthInputs();
        inputs.put("teamId", teamId);
        inputs.put("channelId", "");
        inputs.put("messageContent", messageContent);

        connector.setInputParameters(inputs);
        assertThatThrownBy(connector::validateInputParameters)
                .isInstanceOf(ConnectorValidationException.class);
    }

    @Property(tries = 50)
    @Label("Null messageContent should always fail validation")
    void should_fail_validation_when_messageContent_null(
            @ForAll @AlphaChars @StringLength(min = 1, max = 100) String teamId,
            @ForAll @AlphaChars @StringLength(min = 1, max = 100) String channelId) {
        var connector = new MSTeamsSendChannelMessageConnector();
        Map<String, Object> inputs = validAuthInputs();
        inputs.put("teamId", teamId);
        inputs.put("channelId", channelId);

        connector.setInputParameters(inputs);
        assertThatThrownBy(connector::validateInputParameters)
                .isInstanceOf(ConnectorValidationException.class);
    }

    @Property(tries = 50)
    @Label("Empty messageContent should always fail validation")
    void should_fail_validation_when_messageContent_empty(
            @ForAll @AlphaChars @StringLength(min = 1, max = 100) String teamId,
            @ForAll @AlphaChars @StringLength(min = 1, max = 100) String channelId) {
        var connector = new MSTeamsSendChannelMessageConnector();
        Map<String, Object> inputs = validAuthInputs();
        inputs.put("teamId", teamId);
        inputs.put("channelId", channelId);
        inputs.put("messageContent", "   ");

        connector.setInputParameters(inputs);
        assertThatThrownBy(connector::validateInputParameters)
                .isInstanceOf(ConnectorValidationException.class);
    }

    @Property(tries = 30)
    @Label("Message exceeding max length should always fail validation")
    void should_fail_validation_when_messageContent_exceeds_max_length(
            @ForAll @AlphaChars @StringLength(min = 1, max = 100) String teamId,
            @ForAll @AlphaChars @StringLength(min = 1, max = 100) String channelId) {
        var connector = new MSTeamsSendChannelMessageConnector();
        Map<String, Object> inputs = validAuthInputs();
        inputs.put("teamId", teamId);
        inputs.put("channelId", channelId);
        inputs.put("messageContent", "x".repeat(28_001));

        connector.setInputParameters(inputs);
        assertThatThrownBy(connector::validateInputParameters)
                .isInstanceOf(ConnectorValidationException.class);
    }
}
