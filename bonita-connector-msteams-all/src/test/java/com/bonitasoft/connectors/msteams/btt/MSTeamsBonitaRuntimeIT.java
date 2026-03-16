package com.bonitasoft.connectors.msteams.btt;

import java.io.File;
import java.time.Duration;

import com.bonitasoft.test.toolkit.model.ProcessDefinition;
import com.bonitasoft.test.toolkit.model.ProcessInstance;
import com.bonitasoft.test.toolkit.model.User;

import org.junit.jupiter.api.*;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Bonita Test Toolkit (BTT) integration tests for MS Teams Connector.
 *
 * <p>Validates connector deployment and execution inside a real Bonita runtime
 * (Docker via Testcontainers). The process is built programmatically as a .bar
 * and deployed via REST API before tests run.</p>
 *
 * <h3>Setup Steps:</h3>
 * <ol>
 *   <li>Build: {@code mvn clean install -DskipTests}</li>
 *   <li>Run: {@code mvn verify -Pbtt -pl bonita-connector-msteams-all -am}</li>
 * </ol>
 */
@Tag("btt")
@DisplayName("Bonita Runtime Integration Tests (BTT)")
@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
class MSTeamsBonitaRuntimeIT extends AbstractProcessTest {

    private static final String PROCESS_NAME = "MSTeamsConnectorTestProcess";
    private static final String PROCESS_VERSION = "1.0";

    private ProcessDefinition process;

    @BeforeAll
    void setUp() throws Exception {
        // 1. Build .bar programmatically (ZIP-based, no Bonita API dependency)
        File barFile = new File("target/MSTeamsConnectorTestProcess--1.0.bar");
        MSTeamsProcessBarBuilder.writeToFile(barFile);

        // 2. Deploy to running Bonita via REST API (Testcontainers URL)
        var deployer = new BonitaRestApiDeployer(getBonitaUrl(), "install", "install");
        deployer.deploy(barFile);

        // 3. Get process via BTT
        process = getProcess(PROCESS_NAME, PROCESS_VERSION);
    }

    // =========================================================================
    // Infrastructure
    // =========================================================================

    @Test
    @Order(0)
    @DisplayName("should deploy process to Bonita runtime")
    void should_deploy_process_to_bonita_runtime() {
        assertThat(process).isNotNull();
        assertThat(process.getName()).isEqualTo(PROCESS_NAME);
    }

    @Test
    @Order(1)
    @DisplayName("should start process instance when deployed")
    void should_start_process_instance_when_deployed() {
        User initiator = toolkit.getUser("install");
        ProcessInstance instance = process.startProcessFor(initiator);
        assertThat(instance).isNotNull();
        assertThat(instance.getId()).isNotNull();
    }

    @Test
    @Order(2)
    @DisplayName("should have failed flow nodes when connectors cannot authenticate")
    void should_have_failed_flow_nodes_when_connectors_cannot_authenticate() {
        User initiator = toolkit.getUser("install");
        ProcessInstance instance = process.startProcessFor(initiator);

        // Since tasks are automatic and connectors will fail (no real MS Teams credentials),
        // we expect at least one failed flow node at the first connector
        try {
            Thread.sleep(5000); // Wait for automatic execution attempt
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
        }

        assertThat(instance.getNumberOfFailedFlowNodes())
                .as("First connector should fail because it cannot authenticate with MS Teams")
                .isGreaterThanOrEqualTo(1);
    }
}
