package com.bonitasoft.connectors.msteams.createmeeting;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

/**
 * Tests that the connector definition XML (.def) is valid and consistent.
 */
@DisplayName("MS Teams CreateMeeting Connector Definition Tests")
class MSTeamsCreateMeetingDefinitionTest {

    private static Document defDocument;
    private static List<String> defInputNames;
    private static List<String> defOutputNames;
    private static List<String> defMandatoryInputs;

    @BeforeAll
    static void loadDefinition() throws Exception {
        String defPath = "msteams-create-meeting.def";
        InputStream defIs = MSTeamsCreateMeetingDefinitionTest.class
                .getClassLoader()
                .getResourceAsStream(defPath);
        assertThat(defIs).as("Definition file '%s' should exist", defPath).isNotNull();

        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();
        defDocument = builder.parse(defIs);

        defInputNames = new ArrayList<>();
        defMandatoryInputs = new ArrayList<>();
        NodeList inputs = defDocument.getElementsByTagNameNS("*", "input");
        for (int i = 0; i < inputs.getLength(); i++) {
            Element input = (Element) inputs.item(i);
            String name = input.getAttribute("name");
            defInputNames.add(name);
            if ("true".equals(input.getAttribute("mandatory"))) {
                defMandatoryInputs.add(name);
            }
        }

        defOutputNames = new ArrayList<>();
        NodeList outputs = defDocument.getElementsByTagNameNS("*", "output");
        for (int i = 0; i < outputs.getLength(); i++) {
            Element output = (Element) outputs.item(i);
            defOutputNames.add(output.getAttribute("name"));
        }
    }

    @Test
    @DisplayName("Definition should have a valid connector ID")
    void should_have_valid_connector_id() {
        NodeList ids = defDocument.getElementsByTagNameNS("*", "id");
        assertThat(ids.getLength()).isGreaterThan(0);
        String connectorId = ids.item(0).getTextContent();
        assertThat(connectorId).isEqualTo("msteams-create-meeting");
    }

    @Test
    @DisplayName("Definition should have standard output parameters")
    void should_have_standard_outputs() {
        assertThat(defOutputNames).contains("success", "errorMessage");
    }

    @Test
    @DisplayName("Definition should have at least one operation-specific output or standard outputs")
    void should_have_outputs() {
        assertThat(defOutputNames).isNotEmpty();
        assertThat(defOutputNames).contains("success", "errorMessage");
    }

    @Test
    @DisplayName("Definition should have auth input parameters")
    void should_have_auth_inputs() {
        assertThat(defInputNames).contains("tenantId", "clientId", "clientSecret");
    }

    @Test
    @DisplayName("Definition should have operation input parameters")
    void should_have_operation_inputs() {
        // All connectors should have at least auth inputs plus some operation inputs
        assertThat(defInputNames.size())
                .as("Should have at least some inputs")
                .isGreaterThan(3);
    }

    @Test
    @DisplayName("Mandatory inputs should be marked correctly")
    void should_have_correct_mandatory_inputs() {
        assertThat(defMandatoryInputs)
                .as("Should have at least one mandatory input")
                .isNotEmpty();
    }

    @Test
    @DisplayName("Optional inputs should not be mandatory")
    void should_have_correct_optional_inputs() {
        // connectTimeout and readTimeout should always be optional
        assertThat(defMandatoryInputs).doesNotContain("connectTimeout", "readTimeout");
    }

    @Test
    @DisplayName("Definition should have at least 2 UI pages")
    void should_have_ui_pages() {
        NodeList pages = defDocument.getElementsByTagNameNS("*", "page");
        assertThat(pages.getLength())
                .as("Should have auth page and operation page")
                .isGreaterThanOrEqualTo(2);
    }

    @Test
    @DisplayName("Widget inputs should reference valid input parameters")
    void should_reference_valid_inputs_in_widgets() {
        Set<String> inputNameSet = new HashSet<>(defInputNames);
        NodeList widgets = defDocument.getElementsByTagNameNS("*", "widget");
        for (int i = 0; i < widgets.getLength(); i++) {
            Element widget = (Element) widgets.item(i);
            String inputName = widget.getAttribute("inputName");
            if (inputName != null && !inputName.isBlank()) {
                assertThat(inputNameSet)
                        .as("Widget references input '%s' which should exist", inputName)
                        .contains(inputName);
            }
        }
    }

    @Test
    @DisplayName("Widgets should have labels in properties file")
    void should_have_widget_labels_in_properties() {
        InputStream propsIs =
                getClass()
                        .getClassLoader()
                        .getResourceAsStream("msteams-create-meeting.properties");
        assertThat(propsIs).as("Properties file should exist").isNotNull();

        java.util.Properties props = new java.util.Properties();
        try {
            props.load(propsIs);
        } catch (Exception e) {
            assertThat(false)
                    .as("Properties file should be loadable: " + e.getMessage())
                    .isTrue();
        }

        NodeList widgets = defDocument.getElementsByTagNameNS("*", "widget");
        int labelCount = 0;
        for (int i = 0; i < widgets.getLength(); i++) {
            Element widget = (Element) widgets.item(i);
            String widgetId = widget.getAttribute("id");
            if (widgetId != null && !widgetId.isBlank()) {
                String labelKey = widgetId + ".label";
                if (props.containsKey(labelKey)) {
                    labelCount++;
                }
            }
        }
        assertThat(labelCount)
                .as("Widgets should have descriptive labels in .properties file")
                .isGreaterThanOrEqualTo(5);
    }
}
