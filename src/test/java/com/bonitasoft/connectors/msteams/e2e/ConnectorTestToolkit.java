package com.bonitasoft.connectors.msteams.e2e;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.List;
import java.util.Map;
import java.util.function.Predicate;
import java.util.jar.JarEntry;
import java.util.jar.JarFile;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import org.bonitasoft.engine.bpm.bar.BarResource;
import org.bonitasoft.engine.bpm.bar.BusinessArchive;
import org.bonitasoft.engine.bpm.bar.BusinessArchiveBuilder;
import org.bonitasoft.engine.bpm.bar.BusinessArchiveFactory;
import org.bonitasoft.engine.bpm.bar.actorMapping.Actor;
import org.bonitasoft.engine.bpm.bar.actorMapping.ActorMapping;
import org.bonitasoft.engine.bpm.connector.ConnectorEvent;
import org.bonitasoft.engine.bpm.process.DesignProcessDefinition;
import org.bonitasoft.engine.bpm.process.impl.ProcessDefinitionBuilder;
import org.bonitasoft.engine.expression.ExpressionBuilder;
import org.bonitasoft.engine.expression.InvalidExpressionException;
import org.bonitasoft.engine.operation.OperationBuilder;

/**
 * Helper for testing connectors in a Docker instance of Bonita.
 * Adapted from bonita-connector-rest ConnectorTestToolkit.
 */
public class ConnectorTestToolkit {

    /**
     * Build a business archive containing a process with a connector to test.
     *
     * @param connectorId  connector definition ID
     * @param versionId    connector version
     * @param inputs       map of input name to constant string value
     * @param outputs      map of output variable name to Output descriptor
     * @param artifactId   Maven artifact ID for finding the JAR
     * @return BusinessArchive ready to deploy
     */
    public static BusinessArchive buildConnectorToTest(
            String connectorId, String versionId,
            Map<String, String> inputs, Map<String, Output> outputs,
            String artifactId) throws Exception {
        var process = buildConnectorInProcess(connectorId, versionId, inputs, outputs);
        return buildBusinessArchive(process, connectorId, artifactId);
    }

    private static BusinessArchive buildBusinessArchive(
            DesignProcessDefinition process, String connectorId,
            String artifactId) throws Exception {
        var barBuilder = new BusinessArchiveBuilder();
        barBuilder.createNewBusinessArchive();
        barBuilder.setProcessDefinition(process);

        var foundFiles = new File("").getAbsoluteFile().toPath()
                .resolve("target")
                .toFile()
                .listFiles((dir, name) ->
                        Pattern.matches(artifactId + "-.*.jar", name)
                                && !name.endsWith("-sources.jar")
                                && !name.endsWith("-javadoc.jar"));

        assertThat(foundFiles).hasSize(1);
        var connectorJar = foundFiles[0];
        assertThat(connectorJar).exists();

        List<JarEntry> jarEntries = findJarEntries(connectorJar,
                entry -> entry.getName().equals(connectorId + ".impl"));
        assertThat(jarEntries).hasSize(1);

        byte[] content;
        try (JarFile jarFile = new JarFile(connectorJar)) {
            InputStream inputStream = jarFile.getInputStream(jarEntries.get(0));
            content = inputStream.readAllBytes();
        }

        barBuilder.addConnectorImplementation(
                new BarResource(connectorId + ".impl", content));
        barBuilder.addClasspathResource(
                new BarResource(connectorJar.getName(), Files.readAllBytes(connectorJar.toPath())));

        ActorMapping actorMapping = new ActorMapping();
        var systemActor = new Actor("system");
        systemActor.addRole("member");
        actorMapping.addActor(systemActor);
        barBuilder.setActorMapping(actorMapping);

        return barBuilder.done();
    }

    private static DesignProcessDefinition buildConnectorInProcess(
            String connectorId, String versionId,
            Map<String, String> inputs, Map<String, Output> outputs) throws Exception {
        var processBuilder = new ProcessDefinitionBuilder();
        var expBuilder = new ExpressionBuilder();
        processBuilder.createNewInstance("MSTEAMS_CONNECTOR_TEST", "1.0");
        processBuilder.addActor("system");

        var connectorBuilder = processBuilder.addConnector(
                "connector-under-test", connectorId, versionId, ConnectorEvent.ON_ENTER);

        inputs.forEach((name, value) -> {
            try {
                connectorBuilder.addInput(name, expBuilder.createConstantStringExpression(value));
            } catch (InvalidExpressionException e) {
                throw new RuntimeException(e);
            }
        });

        if (outputs != null) {
            outputs.forEach((name, output) -> {
                try {
                    processBuilder.addData(name, output.type(), null);
                    connectorBuilder.addOutput(new OperationBuilder().createSetDataOperation(name,
                            new ExpressionBuilder().createDataExpression(output.name(), output.type())));
                } catch (InvalidExpressionException e) {
                    throw new RuntimeException(e);
                }
            });
        }

        processBuilder.addUserTask("waiting task", "system");
        return processBuilder.done();
    }

    private static List<JarEntry> findJarEntries(File file, Predicate<? super JarEntry> predicate)
            throws IOException {
        try (JarFile jarFile = new JarFile(file)) {
            return jarFile.stream()
                    .filter(predicate)
                    .collect(Collectors.toList());
        }
    }

    /**
     * Output descriptor for process variables.
     */
    public record Output(String name, String type) {
        public static Output create(String name, String type) {
            return new Output(name, type);
        }
    }
}
