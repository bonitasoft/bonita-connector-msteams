package com.bonitasoft.connectors.msteams.integration;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;

import javax.xml.parsers.DocumentBuilderFactory;

import org.junit.jupiter.api.*;
import org.w3c.dom.*;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Validates the .proc BPMN definition against the connector .def files.
 *
 * <p>Cross-references:</p>
 * <ul>
 *   <li>Each connector task in .proc references a valid connectorId+version</li>
 *   <li>All mandatory .def inputs have a mapping in .proc</li>
 *   <li>All .proc output mappings reference existing .def outputs</li>
 *   <li>All .proc output target variables exist as process variables</li>
 *   <li>All .proc input source variables exist as process variables</li>
 * </ul>
 */
@Tag("integration")
@DisplayName(".proc ↔ .def Cross-Reference Validation")
@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
class ProcDefinitionValidationIT {

    private static final String BPMN_NS = "http://www.omg.org/spec/BPMN/20100524/MODEL";
    private static final String BONITA_NS = "http://bonitasoft.com/ns/process/client/7.x";
    private static final String DEF_NS = "http://www.bonitasoft.org/ns/connector/definition/6.1";

    private static Document procDoc;
    private static Set<String> processVariables;
    private static Map<String, ConnectorTaskMapping> taskMappings;

    record ConnectorTaskMapping(
            String taskId,
            String taskName,
            String connectorId,
            String version,
            Map<String, String> inputMappings,   // connectorInput → expression
            Map<String, String> outputMappings    // connectorOutput → targetVariable
    ) {}

    record DefInput(String name, String type, boolean mandatory) {}
    record DefOutput(String name, String type) {}

    @BeforeAll
    static void parseProcFile() throws Exception {
        var factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);

        try (InputStream is = ProcDefinitionValidationIT.class
                .getClassLoader().getResourceAsStream("MSTeamsConnectorTestProcess.bpmn")) {
            assertThat(is).as(".proc file should be on classpath").isNotNull();
            procDoc = factory.newDocumentBuilder().parse(is);
        }

        // Extract process variables
        processVariables = new LinkedHashSet<>();
        var dataNodes = procDoc.getElementsByTagNameNS(BONITA_NS, "data");
        for (int i = 0; i < dataNodes.getLength(); i++) {
            processVariables.add(((Element) dataNodes.item(i)).getAttribute("name"));
        }

        // Extract connector task mappings
        taskMappings = new LinkedHashMap<>();
        var serviceTasks = procDoc.getElementsByTagNameNS(BPMN_NS, "serviceTask");
        for (int i = 0; i < serviceTasks.getLength(); i++) {
            Element task = (Element) serviceTasks.item(i);
            String taskId = task.getAttribute("id");
            String taskName = task.getAttribute("name");

            var connectors = task.getElementsByTagNameNS(BONITA_NS, "connector");
            if (connectors.getLength() == 0) continue;

            Element connector = (Element) connectors.item(0);
            String connectorId = connector.getAttribute("connectorId");
            String version = connector.getAttribute("version");

            var inputMap = new LinkedHashMap<String, String>();
            var inputs = connector.getElementsByTagNameNS(BONITA_NS, "input");
            for (int j = 0; j < inputs.getLength(); j++) {
                Element input = (Element) inputs.item(j);
                inputMap.put(input.getAttribute("name"), input.getAttribute("expression"));
            }

            var outputMap = new LinkedHashMap<String, String>();
            var outputs = connector.getElementsByTagNameNS(BONITA_NS, "output");
            for (int j = 0; j < outputs.getLength(); j++) {
                Element output = (Element) outputs.item(j);
                outputMap.put(output.getAttribute("name"), output.getAttribute("expression"));
            }

            taskMappings.put(taskId, new ConnectorTaskMapping(
                    taskId, taskName, connectorId, version, inputMap, outputMap));
        }
    }

    @Test
    @Order(1)
    @DisplayName("should have 14 service tasks with connector definitions")
    void should_have_14_service_tasks() {
        assertThat(taskMappings).hasSize(14);
    }

    @Test
    @Order(2)
    @DisplayName("should reference valid connectorId matching existing .def files")
    void should_reference_valid_connector_ids() throws Exception {
        for (var entry : taskMappings.entrySet()) {
            var mapping = entry.getValue();

            // .def file should exist in the JAR or as a resource on classpath
            String defResourceName = mapping.connectorId() + ".def";
            // Look for the .def in the module resources
            Path defPath = findDefFile(mapping.connectorId());
            assertThat(defPath)
                    .as("Task '%s' references connectorId '%s' — .def file should exist",
                            mapping.taskName(), mapping.connectorId())
                    .isNotNull();
        }
    }

    @Test
    @Order(3)
    @DisplayName("should map all mandatory .def inputs in .proc")
    void should_map_all_mandatory_inputs() throws Exception {
        var errors = new ArrayList<String>();

        for (var entry : taskMappings.entrySet()) {
            var mapping = entry.getValue();
            List<DefInput> defInputs = parseDefInputs(mapping.connectorId());

            for (DefInput defInput : defInputs) {
                if (defInput.mandatory() && !mapping.inputMappings().containsKey(defInput.name())) {
                    errors.add(String.format("Task '%s' (%s): mandatory input '%s' is NOT mapped in .proc",
                            mapping.taskName(), mapping.connectorId(), defInput.name()));
                }
            }
        }

        assertThat(errors)
                .as("All mandatory .def inputs should be mapped in the .proc. Unmapped:\n%s",
                        String.join("\n", errors))
                .isEmpty();
    }

    @Test
    @Order(4)
    @DisplayName("should only map outputs that exist in .def")
    void should_only_map_valid_outputs() throws Exception {
        var errors = new ArrayList<String>();

        for (var entry : taskMappings.entrySet()) {
            var mapping = entry.getValue();
            Set<String> defOutputNames = parseDefOutputs(mapping.connectorId()).stream()
                    .map(DefOutput::name)
                    .collect(Collectors.toSet());

            for (String outputName : mapping.outputMappings().keySet()) {
                if (!defOutputNames.contains(outputName)) {
                    errors.add(String.format("Task '%s' (%s): output '%s' is mapped in .proc but NOT defined in .def",
                            mapping.taskName(), mapping.connectorId(), outputName));
                }
            }
        }

        assertThat(errors)
                .as("All .proc output mappings should reference valid .def outputs. Invalid:\n%s",
                        String.join("\n", errors))
                .isEmpty();
    }

    @Test
    @Order(5)
    @DisplayName("should map output targets to existing process variables")
    void should_map_output_targets_to_existing_variables() {
        var errors = new ArrayList<String>();

        for (var entry : taskMappings.entrySet()) {
            var mapping = entry.getValue();
            for (var output : mapping.outputMappings().entrySet()) {
                String targetVar = output.getValue();
                if (!processVariables.contains(targetVar)) {
                    errors.add(String.format("Task '%s' (%s): output '%s' → variable '%s' does NOT exist as process variable",
                            mapping.taskName(), mapping.connectorId(), output.getKey(), targetVar));
                }
            }
        }

        assertThat(errors)
                .as("All output target variables should exist in the process definition. Missing:\n%s",
                        String.join("\n", errors))
                .isEmpty();
    }

    @Test
    @Order(6)
    @DisplayName("should map input expressions to existing process variables or literals")
    void should_map_input_expressions_to_valid_sources() {
        // Known literal values (not process variables)
        Set<String> knownLiterals = Set.of("text", "html", "standard", "private");
        var warnings = new ArrayList<String>();

        for (var entry : taskMappings.entrySet()) {
            var mapping = entry.getValue();
            for (var input : mapping.inputMappings().entrySet()) {
                String expression = input.getValue();
                if (!processVariables.contains(expression) && !knownLiterals.contains(expression)) {
                    warnings.add(String.format("Task '%s' (%s): input '%s' expression '%s' is neither a process variable nor a known literal",
                            mapping.taskName(), mapping.connectorId(), input.getKey(), expression));
                }
            }
        }

        assertThat(warnings)
                .as("All input expressions should reference existing process variables or known literals. Issues:\n%s",
                        String.join("\n", warnings))
                .isEmpty();
    }

    @Test
    @Order(7)
    @DisplayName("should have consistent connector versions across all tasks")
    void should_have_consistent_connector_versions() {
        Set<String> versions = taskMappings.values().stream()
                .map(ConnectorTaskMapping::version)
                .collect(Collectors.toSet());

        assertThat(versions)
                .as("All connector tasks should use the same version")
                .hasSize(1)
                .containsExactly("1.0.0");
    }

    @Test
    @Order(8)
    @DisplayName("should have process variables for success and errorMessage")
    void should_have_shared_output_variables() {
        assertThat(processVariables)
                .as("Process should define 'success' and 'errorMessage' variables for connector outputs")
                .contains("success", "errorMessage");
    }

    // ========================================================================
    // Helpers
    // ========================================================================

    private static Path findDefFile(String connectorId) throws Exception {
        Path projectRoot = Paths.get("").toAbsolutePath().getParent();
        String moduleName = "bonita-connector-msteams-" + connectorId.replace("msteams-", "");
        Path defPath = projectRoot.resolve(moduleName)
                .resolve("src/main/resources")
                .resolve(connectorId + ".def");
        return Files.exists(defPath) ? defPath : null;
    }

    private static List<DefInput> parseDefInputs(String connectorId) throws Exception {
        Path defPath = findDefFile(connectorId);
        if (defPath == null) return List.of();

        var factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        Document defDoc = factory.newDocumentBuilder().parse(defPath.toFile());

        var inputs = new ArrayList<DefInput>();
        var inputNodes = defDoc.getElementsByTagNameNS(DEF_NS, "input");
        // Fallback: try without namespace
        if (inputNodes.getLength() == 0) {
            inputNodes = defDoc.getElementsByTagName("input");
        }

        for (int i = 0; i < inputNodes.getLength(); i++) {
            Element el = (Element) inputNodes.item(i);
            inputs.add(new DefInput(
                    el.getAttribute("name"),
                    el.getAttribute("type"),
                    "true".equals(el.getAttribute("mandatory"))));
        }
        return inputs;
    }

    private static List<DefOutput> parseDefOutputs(String connectorId) throws Exception {
        Path defPath = findDefFile(connectorId);
        if (defPath == null) return List.of();

        var factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        Document defDoc = factory.newDocumentBuilder().parse(defPath.toFile());

        var outputs = new ArrayList<DefOutput>();
        var outputNodes = defDoc.getElementsByTagNameNS(DEF_NS, "output");
        if (outputNodes.getLength() == 0) {
            outputNodes = defDoc.getElementsByTagName("output");
        }

        for (int i = 0; i < outputNodes.getLength(); i++) {
            Element el = (Element) outputNodes.item(i);
            outputs.add(new DefOutput(
                    el.getAttribute("name"),
                    el.getAttribute("type")));
        }
        return outputs;
    }
}
