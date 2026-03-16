package com.bonitasoft.connectors.msteams.integration;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Set;
import java.util.jar.JarEntry;
import java.util.jar.JarFile;
import java.util.stream.Collectors;

import org.junit.jupiter.api.*;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Integration test that validates the bonita-connector-msteams-all-*-bonita.jar
 * is properly structured for Bonita Studio import.
 *
 * <p>Checks:</p>
 * <ul>
 *   <li>14 .def files at JAR root</li>
 *   <li>14 .impl files at JAR root with resolved dependencies</li>
 *   <li>14 .properties files at JAR root</li>
 *   <li>msteams.png icon at JAR root</li>
 *   <li>.def XML has no &lt;label&gt; or &lt;description&gt; child elements</li>
 *   <li>.impl has resolved &lt;jarDependencies&gt; (no ${connector-dependencies})</li>
 *   <li>Implementation classes exist as .class files in JAR</li>
 * </ul>
 */
@Tag("integration")
@DisplayName("Bonita JAR Studio Compatibility Validation")
@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
class BonitoJarValidationIT {

    private static final List<String> CONNECTOR_IDS = List.of(
            "msteams-send-channel-message",
            "msteams-send-chat-message",
            "msteams-reply-message",
            "msteams-send-adaptive-card",
            "msteams-create-channel",
            "msteams-list-channels",
            "msteams-delete-channel",
            "msteams-list-teams",
            "msteams-get-team",
            "msteams-add-member",
            "msteams-remove-member",
            "msteams-create-meeting",
            "msteams-get-meeting",
            "msteams-upload-file"
    );

    private static Path jarPath;
    private static JarFile jarFile;
    private static Set<String> rootEntryNames;

    @BeforeAll
    static void findAndOpenJar() throws Exception {
        Path targetDir = Paths.get("target");
        assertThat(targetDir).exists();

        jarPath = Files.list(targetDir)
                .filter(p -> p.getFileName().toString().matches("bonita-connector-msteams-all-.*-bonita\\.jar"))
                .findFirst()
                .orElseThrow(() -> new AssertionError(
                        "bonita JAR not found in target/. Run 'mvn package -DskipTests' first."));

        jarFile = new JarFile(jarPath.toFile());

        rootEntryNames = jarFile.stream()
                .map(JarEntry::getName)
                .filter(name -> !name.contains("/"))
                .collect(Collectors.toSet());
    }

    @AfterAll
    static void closeJar() throws Exception {
        if (jarFile != null) jarFile.close();
    }

    @Test
    @Order(1)
    @DisplayName("should contain 14 .def files at JAR root")
    void should_contain_14_def_files_at_root() {
        var defFiles = rootEntryNames.stream()
                .filter(n -> n.endsWith(".def"))
                .collect(Collectors.toSet());

        assertThat(defFiles).hasSize(14);
        for (String id : CONNECTOR_IDS) {
            assertThat(defFiles).contains(id + ".def");
        }
    }

    @Test
    @Order(2)
    @DisplayName("should contain 14 .impl files at JAR root")
    void should_contain_14_impl_files_at_root() {
        var implFiles = rootEntryNames.stream()
                .filter(n -> n.endsWith(".impl"))
                .collect(Collectors.toSet());

        assertThat(implFiles).hasSize(14);
        for (String id : CONNECTOR_IDS) {
            assertThat(implFiles).contains(id + ".impl");
        }
    }

    @Test
    @Order(3)
    @DisplayName("should contain 14 .properties files at JAR root")
    void should_contain_14_properties_files_at_root() {
        var propsFiles = rootEntryNames.stream()
                .filter(n -> n.endsWith(".properties"))
                .collect(Collectors.toSet());

        assertThat(propsFiles).hasSize(14);
        for (String id : CONNECTOR_IDS) {
            assertThat(propsFiles).contains(id + ".properties");
        }
    }

    @Test
    @Order(4)
    @DisplayName("should contain msteams.png icon at JAR root")
    void should_contain_icon_at_root() {
        assertThat(rootEntryNames).contains("msteams.png");
    }

    @Test
    @Order(5)
    @DisplayName("should have .def files with no <label> or <description> child elements")
    void should_have_def_files_without_label_or_description_elements() throws Exception {
        for (String id : CONNECTOR_IDS) {
            JarEntry entry = jarFile.getJarEntry(id + ".def");
            assertThat(entry).as("Missing .def for " + id).isNotNull();

            String content;
            try (InputStream is = jarFile.getInputStream(entry)) {
                content = new String(is.readAllBytes());
            }

            assertThat(content)
                    .as(".def for %s should not contain <label> child element", id)
                    .doesNotContainPattern("<label>[^<]*</label>");
            assertThat(content)
                    .as(".def for %s should not contain <description> child element", id)
                    .doesNotContainPattern("<description>[^<]*</description>");
        }
    }

    @Test
    @Order(6)
    @DisplayName("should have .impl files with resolved jarDependencies (no Maven placeholders)")
    void should_have_impl_files_with_resolved_dependencies() throws Exception {
        for (String id : CONNECTOR_IDS) {
            JarEntry entry = jarFile.getJarEntry(id + ".impl");
            assertThat(entry).as("Missing .impl for " + id).isNotNull();

            String content;
            try (InputStream is = jarFile.getInputStream(entry)) {
                content = new String(is.readAllBytes());
            }

            assertThat(content)
                    .as(".impl for %s should not contain unresolved ${connector-dependencies}", id)
                    .doesNotContain("${connector-dependencies}");

            assertThat(content)
                    .as(".impl for %s should contain <jarDependencies>", id)
                    .contains("<jarDependencies>");
        }
    }

    @Test
    @Order(7)
    @DisplayName("should have .impl definitionId matching .def id")
    void should_have_matching_definition_ids() throws Exception {
        for (String id : CONNECTOR_IDS) {
            String implContent;
            try (InputStream is = jarFile.getInputStream(jarFile.getJarEntry(id + ".impl"))) {
                implContent = new String(is.readAllBytes());
            }

            assertThat(implContent)
                    .as(".impl for %s should reference definitionId=%s", id, id)
                    .contains("<definitionId>" + id + "</definitionId>");
        }
    }

    @Test
    @Order(8)
    @DisplayName("should have implementation classes present as .class in JAR")
    void should_have_implementation_classes_in_jar() throws Exception {
        for (String id : CONNECTOR_IDS) {
            String implContent;
            try (InputStream is = jarFile.getInputStream(jarFile.getJarEntry(id + ".impl"))) {
                implContent = new String(is.readAllBytes());
            }

            // Extract implementationClassname from XML
            var matcher = java.util.regex.Pattern
                    .compile("<implementationClassname>([^<]+)</implementationClassname>")
                    .matcher(implContent);
            assertThat(matcher.find())
                    .as("Could not find implementationClassname in %s.impl", id)
                    .isTrue();

            String className = matcher.group(1);
            String classPath = className.replace('.', '/') + ".class";

            assertThat(jarFile.getJarEntry(classPath))
                    .as("Class %s not found in JAR for connector %s", className, id)
                    .isNotNull();
        }
    }

    @Test
    @Order(9)
    @DisplayName("should have <category> as self-closing tag in all .def files")
    void should_have_self_closing_category_in_def_files() throws Exception {
        for (String id : CONNECTOR_IDS) {
            String content;
            try (InputStream is = jarFile.getInputStream(jarFile.getJarEntry(id + ".def"))) {
                content = new String(is.readAllBytes());
            }

            assertThat(content)
                    .as(".def for %s should have self-closing <category/> tag", id)
                    .containsPattern("<category[^>]*/>");
            assertThat(content)
                    .as(".def for %s should not have </category> closing tag", id)
                    .doesNotContain("</category>");
        }
    }

    @Test
    @Order(10)
    @DisplayName("should have widget inputName as attribute in all .def files")
    void should_have_widget_input_name_as_attribute() throws Exception {
        for (String id : CONNECTOR_IDS) {
            String content;
            try (InputStream is = jarFile.getInputStream(jarFile.getJarEntry(id + ".def"))) {
                content = new String(is.readAllBytes());
            }

            // Every <widget> should have inputName="..." as attribute
            var widgetPattern = java.util.regex.Pattern.compile("<widget[^>]*>");
            var widgetMatcher = widgetPattern.matcher(content);
            while (widgetMatcher.find()) {
                String widgetTag = widgetMatcher.group();
                assertThat(widgetTag)
                        .as("Widget in %s.def should have inputName attribute", id)
                        .containsPattern("inputName=\"[^\"]+\"");
            }

            // Should NOT have <inputName> as child element
            assertThat(content)
                    .as(".def for %s should not have <inputName> child element", id)
                    .doesNotContain("<inputName>");
        }
    }

    @Test
    @Order(11)
    @DisplayName("should use correct EMF namespace in .def files")
    void should_use_correct_emf_namespace_in_def() throws Exception {
        for (String id : CONNECTOR_IDS) {
            String content;
            try (InputStream is = jarFile.getInputStream(jarFile.getJarEntry(id + ".def"))) {
                content = new String(is.readAllBytes());
            }

            assertThat(content)
                    .as(".def for %s should use namespace 6.1", id)
                    .contains("http://www.bonitasoft.org/ns/connector/definition/6.1");
        }
    }

    @Test
    @Order(12)
    @DisplayName("should use correct EMF namespace in .impl files")
    void should_use_correct_emf_namespace_in_impl() throws Exception {
        for (String id : CONNECTOR_IDS) {
            String content;
            try (InputStream is = jarFile.getInputStream(jarFile.getJarEntry(id + ".impl"))) {
                content = new String(is.readAllBytes());
            }

            assertThat(content)
                    .as(".impl for %s should use namespace 6.0", id)
                    .contains("http://www.bonitasoft.org/ns/connector/implementation/6.0");
        }
    }
}
