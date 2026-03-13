package com.bonitasoft.connectors.msteams.messaging;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.*;

import com.bonitasoft.connectors.msteams.common.MsTeamsGraphClient;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.bonitasoft.engine.connector.ConnectorException;
import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.HashMap;
import java.util.Map;

@ExtendWith(MockitoExtension.class)
class SendChatMessageConnectorTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Mock
    private MsTeamsGraphClient graphClient;

    private SendChatMessageConnector connector;
    private Map<String, Object> inputs;

    @BeforeEach
    void setUp() {
        connector = new SendChatMessageConnector();
        inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("refreshToken", "rt");
        inputs.put("chatId", "chat-1");
        inputs.put("message", "Hello");
        connector.setInputParameters(inputs);
    }

    @Test
    void should_pass_validation() throws ConnectorValidationException {
        connector.validateInputParameters();
    }

    @Test
    void should_fail_when_chat_id_missing() {
        var c = new SendChatMessageConnector();
        var incompleteInputs = new HashMap<>(inputs);
        incompleteInputs.remove("chatId");
        c.setInputParameters(incompleteInputs);
        assertThatThrownBy(() -> c.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class);
    }

    @Test
    void should_use_delegated_auth_mode() {
        assertThat(connector.getAuthMode()).isEqualTo("delegated");
    }

    @Test
    void should_send_chat_message_successfully() throws Exception {
        connector.setGraphClient(graphClient);
        when(graphClient.post(eq("/chats/chat-1/messages"), any()))
                .thenReturn(MAPPER.readTree("{\"id\":\"msg-1\",\"createdDateTime\":\"2025-01-01T00:00:00Z\"}"));

        connector.executeBusinessLogic();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("messageId")).isEqualTo("msg-1");
    }
}
