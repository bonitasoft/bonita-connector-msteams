package com.bonitasoft.connectors.msteams.common;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Properties;
import java.util.Set;
import java.util.stream.Stream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * Validates all 14 connector .def files for structural correctness.
 * Adapted from PR#4 DefinitionTest pattern, consolidated for single-module.
 */
@DisplayName("Connector Definition Validation (all 14 connectors)")
class ConnectorDefinitionTest {

    static Stream<String> connectorIds() {
        return Stream.of(
                "msteams-send-channel-message",
                "msteams-send-adaptive-card",
                "msteams-send-chat-message",
                "msteams-reply-message",
                "msteams-list-channels",
                "msteams-create-channel",
                "msteams-delete-channel",
                "msteams-list-teams",
                "msteams-get-team",
                "msteams-add-member",
                "msteams-remove-member",
                "msteams-create-meeting",
                "msteams-get-meeting",
                "msteams-upload-file"
        );
    }

    private record DefInfo(Document doc, List<String> inputNames, List<String> outputNames,
                           List<String> mandatoryInputs, String connectorId) {}

    private DefInfo loadDef(String connectorId) throws Exception {
        String defPath = connectorId + ".def";
        InputStream is = getClass().getClassLoader().getResourceAsStream(defPath);
        assertThat(is).as("Definition file '%s' should exist on classpath", defPath).isNotNull();

        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document doc = builder.parse(is);

        List<String> inputNames = new ArrayList<>();
        List<String> mandatoryInputs = new ArrayList<>();
        NodeList inputs = doc.getElementsByTagNameNS("*", "input");
        for (int i = 0; i < inputs.getLength(); i++) {
            Element input = (Element) inputs.item(i);
            String name = input.getAttribute("name");
            inputNames.add(name);
            if ("true".equals(input.getAttribute("mandatory"))) {
                mandatoryInputs.add(name);
            }
        }

        List<String> outputNames = new ArrayList<>();
        NodeList outputs = doc.getElementsByTagNameNS("*", "output");
        for (int i = 0; i < outputs.getLength(); i++) {
            outputNames.add(((Element) outputs.item(i)).getAttribute("name"));
        }

        return new DefInfo(doc, inputNames, outputNames, mandatoryInputs, connectorId);
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("connectorIds")
    @DisplayName("Definition file should exist and be parseable")
    void should_have_parseable_def_file(String connectorId) throws Exception {
        DefInfo def = loadDef(connectorId);
        assertThat(def.doc()).isNotNull();
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("connectorIds")
    @DisplayName("Definition should have correct connector ID")
    void should_have_correct_connector_id(String connectorId) throws Exception {
        DefInfo def = loadDef(connectorId);
        NodeList ids = def.doc().getElementsByTagNameNS("*", "id");
        assertThat(ids.getLength()).isGreaterThan(0);
        assertThat(ids.item(0).getTextContent()).isEqualTo(connectorId);
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("connectorIds")
    @DisplayName("Definition should have a version")
    void should_have_version(String connectorId) throws Exception {
        DefInfo def = loadDef(connectorId);
        NodeList versions = def.doc().getElementsByTagNameNS("*", "version");
        assertThat(versions.getLength()).isGreaterThan(0);
        assertThat(versions.item(0).getTextContent()).isNotBlank();
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("connectorIds")
    @DisplayName("Definition should reference msteams.png icon")
    void should_have_icon(String connectorId) throws Exception {
        DefInfo def = loadDef(connectorId);
        NodeList icons = def.doc().getElementsByTagNameNS("*", "icon");
        assertThat(icons.getLength()).isGreaterThan(0);
        assertThat(icons.item(0).getTextContent()).isEqualTo("msteams.png");
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("connectorIds")
    @DisplayName("Definition should have standard success and errorMessage outputs")
    void should_have_standard_outputs(String connectorId) throws Exception {
        DefInfo def = loadDef(connectorId);
        assertThat(def.outputNames()).contains("success", "errorMessage");
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("connectorIds")
    @DisplayName("Definition should have at least one UI page")
    void should_have_ui_pages(String connectorId) throws Exception {
        DefInfo def = loadDef(connectorId);
        NodeList pages = def.doc().getElementsByTagNameNS("*", "page");
        assertThat(pages.getLength()).isGreaterThanOrEqualTo(1);
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("connectorIds")
    @DisplayName("All inputs should be assigned to a widget")
    void should_assign_all_inputs_to_widgets(String connectorId) throws Exception {
        DefInfo def = loadDef(connectorId);
        Set<String> widgetInputNames = new HashSet<>();
        NodeList widgets = def.doc().getElementsByTagNameNS("*", "widget");
        for (int i = 0; i < widgets.getLength(); i++) {
            Element widget = (Element) widgets.item(i);
            String inputName = widget.getAttribute("inputName");
            if (inputName != null && !inputName.isBlank()) {
                widgetInputNames.add(inputName);
            }
        }
        for (String inputName : def.inputNames()) {
            assertThat(widgetInputNames)
                    .as("Input '%s' in %s should be assigned to a widget", inputName, connectorId)
                    .contains(inputName);
        }
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("connectorIds")
    @DisplayName("Definition should not have label/description child elements (EMF rule)")
    void should_not_have_label_description_children(String connectorId) throws Exception {
        DefInfo def = loadDef(connectorId);
        NodeList inputs = def.doc().getElementsByTagNameNS("*", "input");
        for (int i = 0; i < inputs.getLength(); i++) {
            Element input = (Element) inputs.item(i);
            NodeList children = input.getChildNodes();
            for (int j = 0; j < children.getLength(); j++) {
                Node child = children.item(j);
                if (child.getNodeType() == Node.ELEMENT_NODE) {
                    assertThat(child.getLocalName())
                            .as("Input '%s' in %s should not have <%s> child element",
                                    input.getAttribute("name"), connectorId, child.getLocalName())
                            .isNotIn("label", "description");
                }
            }
        }
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("connectorIds")
    @DisplayName("Category should be self-closing with icon and id attributes")
    void should_have_self_closing_category(String connectorId) throws Exception {
        DefInfo def = loadDef(connectorId);
        NodeList categories = def.doc().getElementsByTagNameNS("*", "category");
        assertThat(categories.getLength()).isGreaterThan(0);
        Element category = (Element) categories.item(0);
        assertThat(category.getAttribute("id")).isNotBlank();
        assertThat(category.getAttribute("icon")).isNotBlank();
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("connectorIds")
    @DisplayName("Properties file should exist alongside .def")
    void should_have_properties_file(String connectorId) throws Exception {
        String propsPath = connectorId + ".properties";
        InputStream is = getClass().getClassLoader().getResourceAsStream(propsPath);
        assertThat(is).as("Properties file '%s' should exist", propsPath).isNotNull();

        Properties props = new Properties();
        props.load(is);
        assertThat(props).isNotEmpty();
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("connectorIds")
    @DisplayName("Implementation file should exist alongside .def")
    void should_have_impl_file(String connectorId) throws Exception {
        String implPath = connectorId + ".impl";
        InputStream is = getClass().getClassLoader().getResourceAsStream(implPath);
        assertThat(is).as("Implementation file '%s' should exist", implPath).isNotNull();
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("connectorIds")
    @DisplayName("Icon file should exist on classpath")
    void should_have_icon_file(String connectorId) throws Exception {
        InputStream is = getClass().getClassLoader().getResourceAsStream("msteams.png");
        assertThat(is).as("msteams.png icon should exist on classpath").isNotNull();
    }
}
