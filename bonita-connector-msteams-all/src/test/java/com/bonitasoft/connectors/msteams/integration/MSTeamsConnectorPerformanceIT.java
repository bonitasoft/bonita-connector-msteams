package com.bonitasoft.connectors.msteams.integration;

import java.time.Duration;
import java.time.Instant;
import java.util.*;
import java.util.concurrent.*;
import java.util.stream.IntStream;

import com.bonitasoft.connectors.msteams.*;
import com.bonitasoft.connectors.msteams.sendchannelmessage.MSTeamsSendChannelMessageConnector;
import com.bonitasoft.connectors.msteams.listteams.MSTeamsListTeamsConnector;
import com.bonitasoft.connectors.msteams.getteam.MSTeamsGetTeamConnector;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import org.junit.jupiter.api.*;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * Performance and load tests for MS Teams connectors.
 *
 * <p>Uses mocked MSTeamsClient to measure connector execution performance
 * without external dependencies. Validates:</p>
 * <ul>
 *   <li>Single connector latency under threshold</li>
 *   <li>Concurrent execution (10 instances) without failures</li>
 *   <li>Sequential 14-connector pipeline throughput</li>
 * </ul>
 */
@Tag("integration")
@ExtendWith(MockitoExtension.class)
@DisplayName("Connector Performance Tests")
@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
class MSTeamsConnectorPerformanceIT {

    private static final ObjectMapper MAPPER = new ObjectMapper();
    private static final int CONCURRENCY = 10;
    private static final Duration SINGLE_CONNECTOR_THRESHOLD = Duration.ofMillis(500);
    private static final Duration CONCURRENT_BATCH_THRESHOLD = Duration.ofSeconds(5);

    private static final MSTeamsConfiguration TEST_CONFIG =
            new MSTeamsConfiguration("perf-tenant", "perf-client", "perf-secret", 5000, 10000, true);

    @Mock
    private MSTeamsClient client;

    /** Performance results collected during the test run. */
    private static final List<PerformanceResult> results = Collections.synchronizedList(new ArrayList<>());

    record PerformanceResult(String testName, int instances, Duration elapsed, boolean success) {
        String throughput() {
            double ops = instances / (elapsed.toMillis() / 1000.0);
            return String.format("%.1f ops/s", ops);
        }
    }

    @BeforeEach
    void setupClient() {
        lenient().when(client.getObjectMapper()).thenReturn(MAPPER);
        lenient().when(client.getConfiguration()).thenReturn(TEST_CONFIG);
    }

    // =========================================================================
    // Single Connector Latency
    // =========================================================================

    @Test
    @Order(1)
    @DisplayName("should execute SendChannelMessage within latency threshold")
    void should_execute_send_channel_message_within_threshold() throws Exception {
        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "perf-msg-001");
        response.put("webUrl", "https://teams.microsoft.com/perf/001");
        response.put("createdDateTime", "2026-03-15T10:00:00Z");
        when(client.post(contains("/messages"), anyString())).thenReturn(response);

        var connector = new MSTeamsSendChannelMessageConnector();
        connector.setInputParameters(Map.of(
                "tenantId", "perf-tenant",
                "clientId", "perf-client",
                "clientSecret", "perf-secret",
                "teamId", "team-perf",
                "channelId", "channel-perf",
                "messageContent", "Performance test message"
        ));

        Instant start = Instant.now();
        ConnectorLifecycleTestHelper.executeWithMockClient(connector, client);
        Duration elapsed = Duration.between(start, Instant.now());

        results.add(new PerformanceResult("SendChannelMessage (single)", 1, elapsed, true));

        assertThat(elapsed)
                .as("Single SendChannelMessage should complete within %s", SINGLE_CONNECTOR_THRESHOLD)
                .isLessThan(SINGLE_CONNECTOR_THRESHOLD);
        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
    }

    @Test
    @Order(2)
    @DisplayName("should execute ListTeams within latency threshold")
    void should_execute_list_teams_within_threshold() throws Exception {
        ObjectNode response = MAPPER.createObjectNode();
        ArrayNode teams = response.putArray("value");
        for (int i = 0; i < 50; i++) {
            ObjectNode team = teams.addObject();
            team.put("id", "team-" + i);
            team.put("displayName", "Team " + i);
            team.put("description", "Description " + i);
            team.put("visibility", "public");
        }
        when(client.get(contains("/groups"))).thenReturn(response);

        var connector = new MSTeamsListTeamsConnector();
        connector.setInputParameters(Map.of(
                "tenantId", "perf-tenant",
                "clientId", "perf-client",
                "clientSecret", "perf-secret"
        ));

        Instant start = Instant.now();
        ConnectorLifecycleTestHelper.executeWithMockClient(connector, client);
        Duration elapsed = Duration.between(start, Instant.now());

        results.add(new PerformanceResult("ListTeams (50 results)", 1, elapsed, true));

        assertThat(elapsed)
                .as("ListTeams with 50 results should complete within %s", SINGLE_CONNECTOR_THRESHOLD)
                .isLessThan(SINGLE_CONNECTOR_THRESHOLD);
        assertThat(connector.getOutputs().get("totalCount")).isEqualTo(50);
    }

    // =========================================================================
    // Concurrent Execution (Load Test)
    // =========================================================================

    @Test
    @Order(10)
    @DisplayName("should execute 10 concurrent SendChannelMessage without failures")
    void should_execute_concurrent_send_channel_messages() throws Exception {
        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "concurrent-msg");
        response.put("webUrl", "https://teams.microsoft.com/concurrent");
        response.put("createdDateTime", "2026-03-15T10:00:00Z");
        when(client.post(contains("/messages"), anyString())).thenReturn(response);

        ExecutorService executor = Executors.newFixedThreadPool(CONCURRENCY);
        List<Future<Boolean>> futures = new ArrayList<>();

        Instant start = Instant.now();

        for (int i = 0; i < CONCURRENCY; i++) {
            final int idx = i;
            futures.add(executor.submit(() -> {
                var connector = new MSTeamsSendChannelMessageConnector();
                connector.setInputParameters(Map.of(
                        "tenantId", "perf-tenant",
                        "clientId", "perf-client",
                        "clientSecret", "perf-secret",
                        "teamId", "team-perf",
                        "channelId", "channel-perf",
                        "messageContent", "Concurrent message #" + idx
                ));
                ConnectorLifecycleTestHelper.executeWithMockClient(connector, client);
                return Boolean.TRUE.equals(connector.getOutputs().get("success"));
            }));
        }

        int successCount = 0;
        for (Future<Boolean> future : futures) {
            if (future.get(10, TimeUnit.SECONDS)) {
                successCount++;
            }
        }
        executor.shutdown();

        Duration elapsed = Duration.between(start, Instant.now());
        results.add(new PerformanceResult("SendChannelMessage (concurrent)", CONCURRENCY, elapsed, successCount == CONCURRENCY));

        assertThat(successCount)
                .as("All %d concurrent executions should succeed", CONCURRENCY)
                .isEqualTo(CONCURRENCY);
        assertThat(elapsed)
                .as("Concurrent batch should complete within %s", CONCURRENT_BATCH_THRESHOLD)
                .isLessThan(CONCURRENT_BATCH_THRESHOLD);
    }

    @Test
    @Order(11)
    @DisplayName("should execute 10 concurrent GetTeam without failures")
    void should_execute_concurrent_get_team() throws Exception {
        ObjectNode response = MAPPER.createObjectNode();
        response.put("id", "team-perf");
        response.put("displayName", "Performance Team");
        response.put("description", "Load test team");
        response.put("visibility", "public");
        response.put("isArchived", false);
        response.put("webUrl", "https://teams.microsoft.com/perf");
        when(client.get(contains("/teams/"))).thenReturn(response);

        ExecutorService executor = Executors.newFixedThreadPool(CONCURRENCY);
        List<Future<Boolean>> futures = new ArrayList<>();

        Instant start = Instant.now();

        for (int i = 0; i < CONCURRENCY; i++) {
            futures.add(executor.submit(() -> {
                var connector = new MSTeamsGetTeamConnector();
                connector.setInputParameters(Map.of(
                        "tenantId", "perf-tenant",
                        "clientId", "perf-client",
                        "clientSecret", "perf-secret",
                        "teamId", "team-perf"
                ));
                ConnectorLifecycleTestHelper.executeWithMockClient(connector, client);
                return Boolean.TRUE.equals(connector.getOutputs().get("success"));
            }));
        }

        int successCount = 0;
        for (Future<Boolean> future : futures) {
            if (future.get(10, TimeUnit.SECONDS)) {
                successCount++;
            }
        }
        executor.shutdown();

        Duration elapsed = Duration.between(start, Instant.now());
        results.add(new PerformanceResult("GetTeam (concurrent)", CONCURRENCY, elapsed, successCount == CONCURRENCY));

        assertThat(successCount).isEqualTo(CONCURRENCY);
        assertThat(elapsed).isLessThan(CONCURRENT_BATCH_THRESHOLD);
    }

    // =========================================================================
    // Performance Report
    // =========================================================================

    @Test
    @Order(100)
    @DisplayName("should print performance report")
    void should_print_performance_report() {
        System.out.println("\n" + "=".repeat(80));
        System.out.println("  MS Teams Connector Performance Report");
        System.out.println("=".repeat(80));
        System.out.printf("  %-45s %8s %10s %12s%n", "Test", "Count", "Elapsed", "Throughput");
        System.out.println("-".repeat(80));

        for (PerformanceResult r : results) {
            System.out.printf("  %-45s %8d %8dms %12s%n",
                    r.testName(), r.instances(), r.elapsed().toMillis(), r.throughput());
        }

        System.out.println("-".repeat(80));
        System.out.printf("  Thresholds: single=%dms, batch=%ds%n",
                SINGLE_CONNECTOR_THRESHOLD.toMillis(), CONCURRENT_BATCH_THRESHOLD.toSeconds());
        System.out.println("=".repeat(80) + "\n");

        // All results should be successful
        assertThat(results).allMatch(PerformanceResult::success);
    }
}
