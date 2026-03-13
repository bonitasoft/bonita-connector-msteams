package com.bonitasoft.connectors.msteams.e2e;

import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;

/**
 * End-to-end tests for MS Teams connectors.
 * Requires Docker (Bonita container) and WireMock for Teams API simulation.
 *
 * Run with: mvn verify -PE2E
 */
@Tag("e2e")
@Disabled("Requires Docker and Bonita container - run manually with -PE2E")
class MsTeamsConnectorE2EIT {

    @Test
    @DisplayName("Should deploy and execute SendChannelMessage connector in Bonita")
    void should_execute_send_channel_message_in_bonita() {
        // TODO: Implement with Testcontainers + Bonita Docker
        // 1. Start Bonita container
        // 2. Build BusinessArchive with ConnectorTestToolkit
        // 3. Deploy process
        // 4. Start process instance
        // 5. Verify connector outputs
    }

    @Test
    @DisplayName("Should deploy and execute SendAdaptiveCard connector in Bonita")
    void should_execute_send_adaptive_card_in_bonita() {
        // TODO: Implement with Testcontainers
    }

    @Test
    @DisplayName("Should deploy and execute ListChannels connector in Bonita")
    void should_execute_list_channels_in_bonita() {
        // TODO: Implement with Testcontainers
    }
}
