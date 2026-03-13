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
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.HashMap;
import java.util.Map;

@ExtendWith(MockitoExtension.class)
class ReplyMessageConnectorTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Mock
    private MsTeamsGraphClient graphClient;

    private ReplyMessageConnector connector;
    private Map<String, Object> inputs;

    @BeforeEach
    void setUp() {
        connector = new ReplyMessageConnector();
        inputs = new HashMap<>();
        inputs.put("tenantId", "t");
        inputs.put("clientId", "c");
        inputs.put("clientSecret", "s");
        inputs.put("refreshToken", "rt");
        inputs.put("teamId", "team-1");
        inputs.put("channelId", "ch-1");
        inputs.put("messageId", "msg-1");
        inputs.put("replyContent", "Reply text");
        connector.setInputParameters(inputs);
    }

    @Test
    void should_pass_validation() throws ConnectorValidationException {
        connector.validateInputParameters();
    }

    @Test
    void should_fail_when_reply_content_missing() {
        var c = new ReplyMessageConnector();
        var incompleteInputs = new HashMap<>(inputs);
        incompleteInputs.remove("replyContent");
        c.setInputParameters(incompleteInputs);
        assertThatThrownBy(() -> c.validateInputParameters())
                .isInstanceOf(ConnectorValidationException.class);
    }

    @Test
    void should_reply_successfully() throws Exception {
        connector.setGraphClient(graphClient);
        when(graphClient.post(eq("/teams/team-1/channels/ch-1/messages/msg-1/replies"), any()))
                .thenReturn(MAPPER.readTree("{\"id\":\"reply-1\",\"createdDateTime\":\"2025-01-01T00:00:00Z\"}"));

        connector.executeBusinessLogic();

        assertThat(connector.getOutputs().get("success")).isEqualTo(true);
        assertThat(connector.getOutputs().get("replyId")).isEqualTo("reply-1");
    }
}
