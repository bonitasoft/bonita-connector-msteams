package com.bonitasoft.connectors.msteams.property;

import static org.assertj.core.api.Assertions.assertThatThrownBy;

import com.bonitasoft.connectors.msteams.messaging.SendChannelMessageConnector;
import com.bonitasoft.connectors.msteams.channels.CreateChannelConnector;
import com.bonitasoft.connectors.msteams.meetings.CreateMeetingConnector;
import net.jqwik.api.*;
import net.jqwik.api.constraints.StringLength;
import org.bonitasoft.engine.connector.ConnectorValidationException;

import java.util.HashMap;
import java.util.Map;

/**
 * Property-based tests for connector validation.
 * Ensures validators catch all invalid inputs.
 */
class ConnectorValidationPropertyTest {

    @Property
    void should_reject_non_https_webhook_urls(
            @ForAll("httpUrls") String url) {
        var connector = new SendChannelMessageConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("webhookUrl", url);
        inputs.put("message", "test");
        connector.setInputParameters(inputs);

        assertThatThrownBy(() -> connector.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class);
    }

    @Provide
    Arbitrary<String> httpUrls() {
        return Arbitraries.of(
                "http://example.com/webhook",
                "ftp://example.com/webhook",
                "not-a-url",
                "",
                "htt://bad"
        );
    }

    @Property
    void should_accept_any_https_webhook_url_with_message(
            @ForAll("httpsUrls") String url,
            @ForAll("nonBlankMessages") String message) throws ConnectorValidationException {
        var connector = new SendChannelMessageConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("webhookUrl", url);
        inputs.put("message", message);
        connector.setInputParameters(inputs);

        connector.validateInputParameters(); // Should not throw
    }

    @Provide
    Arbitrary<String> nonBlankMessages() {
        return Arbitraries.strings()
                .alpha().ofMinLength(1).ofMaxLength(100)
                .filter(s -> !s.isBlank());
    }

    @Provide
    Arbitrary<String> httpsUrls() {
        return Arbitraries.strings()
                .alpha().ofMinLength(5).ofMaxLength(50)
                .map(s -> "https://" + s + ".com/webhook");
    }

    @Property(tries = 50)
    void should_reject_graph_connector_without_credentials(
            @ForAll @StringLength(min = 1, max = 50) String teamId,
            @ForAll @StringLength(min = 1, max = 50) String displayName) {
        var connector = new CreateChannelConnector();
        Map<String, Object> inputs = new HashMap<>();
        // Missing tenantId, clientId, clientSecret
        inputs.put("teamId", teamId);
        inputs.put("displayName", displayName);
        connector.setInputParameters(inputs);

        assertThatThrownBy(() -> connector.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class);
    }

    @Property(tries = 30)
    void should_reject_meeting_without_required_fields(
            @ForAll("missingMeetingInputs") Map<String, Object> inputs) {
        var connector = new CreateMeetingConnector();
        connector.setInputParameters(inputs);

        assertThatThrownBy(() -> connector.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class);
    }

    @Provide
    Arbitrary<Map<String, Object>> missingMeetingInputs() {
        return Arbitraries.of(
                // Missing userId
                Map.of("tenantId", "t", "clientId", "c", "clientSecret", "s",
                        "refreshToken", "rt", "subject", "Test",
                        "startDateTime", "2025-01-01T00:00:00Z", "endDateTime", "2025-01-01T01:00:00Z"),
                // Missing subject
                Map.of("tenantId", "t", "clientId", "c", "clientSecret", "s",
                        "refreshToken", "rt", "userId", "u",
                        "startDateTime", "2025-01-01T00:00:00Z", "endDateTime", "2025-01-01T01:00:00Z"),
                // Missing startDateTime
                Map.of("tenantId", "t", "clientId", "c", "clientSecret", "s",
                        "refreshToken", "rt", "userId", "u", "subject", "Test",
                        "endDateTime", "2025-01-01T01:00:00Z"),
                // Missing endDateTime
                Map.of("tenantId", "t", "clientId", "c", "clientSecret", "s",
                        "refreshToken", "rt", "userId", "u", "subject", "Test",
                        "startDateTime", "2025-01-01T00:00:00Z")
        );
    }

    @Property
    void should_reject_negative_timeout(
            @ForAll @net.jqwik.api.constraints.IntRange(min = -10000, max = -1) int timeout) {
        var connector = new SendChannelMessageConnector();
        Map<String, Object> inputs = new HashMap<>();
        inputs.put("webhookUrl", "https://test.com/webhook");
        inputs.put("message", "test");
        inputs.put("connectTimeout", timeout);
        connector.setInputParameters(inputs);

        assertThatThrownBy(() -> connector.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class);
    }
}
