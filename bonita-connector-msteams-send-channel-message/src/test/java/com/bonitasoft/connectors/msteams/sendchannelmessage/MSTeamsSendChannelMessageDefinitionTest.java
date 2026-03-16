package com.bonitasoft.connectors.msteams.sendchannelmessage;

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
@DisplayName("MS Teams Send Channel Message Connector Definition Tests")
class MSTeamsSendChannelMessageDefinitionTest {

    private static Document defDocument;
    private static List<String> defInputNames;
    private static List<String> defOutputNames;
    private static List<String> defMandatoryInputs;

    @BeforeAll
    static void loadDefinition() throws Exception {
        String defPath = "msteams-send-channel-message.def";
        InputStream defIs = MSTeamsSendChannelMessageDefinitionTest.class
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
        assertThat(connectorId).isEqualTo("msteams-send-channel-message");
    }

    @Test
    @DisplayName("Definition should have standard output parameters")
    void should_have_standard_outputs() {
        assertThat(defOutputNames).contains("success", "errorMessage");
    }

    @Test
    @DisplayName("Definition should have operation-specific outputs")
    void should_have_operation_specific_outputs() {
        assertThat(defOutputNames).contains("messageId", "webUrl", "createdDateTime");
    }

    @Test
    @DisplayName("Definition should have auth input parameters")
    void should_have_auth_inputs() {
        assertThat(defInputNames).contains("tenantId", "clientId", "clientSecret");
    }

    @Test
    @DisplayName("Definition should have operation input parameters")
    void should_have_operation_inputs() {
        assertThat(defInputNames).contains("teamId", "channelId", "messageContent");
    }

    @Test
    @DisplayName("Mandatory inputs should be marked correctly")
    void should_have_correct_mandatory_inputs() {
        assertThat(defMandatoryInputs)
                .contains(
                        "tenantId",
                        "clientId",
                        "clientSecret",
                        "teamId",
                        "channelId",
                        "messageContent");
    }

    @Test
    @DisplayName("Optional inputs should not be mandatory")
    void should_have_correct_optional_inputs() {
        assertThat(defMandatoryInputs)
                .doesNotContain(
                        "contentType",
                        "subject",
                        "importance",
                        "authMode",
                        "refreshToken",
                        "userId",
                        "connectTimeout",
                        "readTimeout");
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
    @DisplayName("All inputs should be assigned to a UI widget")
    void should_assign_all_inputs_to_widgets() {
        Set<String> widgetInputNames = new HashSet<>();
        NodeList widgets = defDocument.getElementsByTagNameNS("*", "widget");
        for (int i = 0; i < widgets.getLength(); i++) {
            Element widget = (Element) widgets.item(i);
            String inputName = widget.getAttribute("inputName");
            if (inputName != null && !inputName.isBlank()) {
                widgetInputNames.add(inputName);
            }
        }

        for (String inputName : defInputNames) {
            assertThat(widgetInputNames)
                    .as("Input '%s' should be assigned to a UI widget", inputName)
                    .contains(inputName);
        }
    }

    @Test
    @DisplayName("Widgets should have labels in properties file")
    void should_have_widget_labels_in_properties() {
        InputStream propsIs =
                getClass()
                        .getClassLoader()
                        .getResourceAsStream("msteams-send-channel-message.properties");
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
                .isGreaterThanOrEqualTo(10);
    }
}
