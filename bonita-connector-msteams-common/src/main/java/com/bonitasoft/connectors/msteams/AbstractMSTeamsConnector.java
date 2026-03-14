package com.bonitasoft.connectors.msteams;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.bonitasoft.engine.connector.AbstractConnector;
import org.bonitasoft.engine.connector.ConnectorException;
import org.bonitasoft.engine.connector.ConnectorValidationException;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public abstract class AbstractMSTeamsConnector extends AbstractConnector {

    public static final String INPUT_TENANT_ID = "tenantId";
    public static final String INPUT_CLIENT_ID = "clientId";
    public static final String INPUT_CLIENT_SECRET = "clientSecret";
    public static final String INPUT_CONNECT_TIMEOUT = "connectTimeout";
    public static final String INPUT_READ_TIMEOUT = "readTimeout";
    public static final String OUTPUT_SUCCESS = "success";
    public static final String OUTPUT_ERROR_MESSAGE = "errorMessage";
    protected static final int MAX_MESSAGE_LENGTH = 28_000;

    protected MSTeamsConfiguration configuration;
    protected MSTeamsClient client;

    @Override
    public void validateInputParameters() throws ConnectorValidationException {
        log.debug("Validating input parameters for {}", getConnectorName());
        List<String> errors = new ArrayList<>();
        validateAuthenticationInputs(errors);
        validateOperationInputs(errors);
        if (!errors.isEmpty()) {
            log.warn("Validation failed for {}: {}", getConnectorName(), errors);
            throw new ConnectorValidationException(this, errors);
        }
        log.debug("Input validation passed for {}", getConnectorName());
    }

    protected void validateAuthenticationInputs(List<String> errors) {
        String tenantId = getStringInput(INPUT_TENANT_ID);
        String clientId = getStringInput(INPUT_CLIENT_ID);
        String clientSecret = getStringInput(INPUT_CLIENT_SECRET);
        if (isBlank(tenantId) || isBlank(clientId) || isBlank(clientSecret)) {
            errors.add("Authentication required: tenantId, clientId, and clientSecret must all be provided.");
        }
        Integer ct = getIntInput(INPUT_CONNECT_TIMEOUT);
        if (ct != null && ct < 0) errors.add("connectTimeout must be a positive number");
        Integer rt = getIntInput(INPUT_READ_TIMEOUT);
        if (rt != null && rt < 0) errors.add("readTimeout must be a positive number");
    }

    protected abstract void validateOperationInputs(List<String> errors);
    protected abstract String getConnectorName();

    @Override
    public void connect() throws ConnectorException {
        log.info("Connecting to Microsoft Teams for {}", getConnectorName());
        try {
            configuration = buildConfiguration();
            client = new MSTeamsClient(configuration);
            client.authenticate();
        } catch (MSTeamsException e) {
            log.error("Failed to connect: {}", e.getMessage());
            throw new ConnectorException(MSTeamsException.truncateMessage("Failed to connect: " + e.getMessage()), e);
        }
    }

    protected MSTeamsConfiguration buildConfiguration() {
        return new MSTeamsConfiguration(
                getStringInput(INPUT_TENANT_ID), getStringInput(INPUT_CLIENT_ID), getStringInput(INPUT_CLIENT_SECRET),
                getIntInputOrDefault(INPUT_CONNECT_TIMEOUT, 30_000), getIntInputOrDefault(INPUT_READ_TIMEOUT, 60_000),
                getBooleanInputOrDefault("appOnly", false));
    }

    public void setClient(MSTeamsClient client) { this.client = client; }

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        log.info("Executing {} operation", getConnectorName());
        try {
            executeOperation(client);
            setSuccessOutputs();
            log.info("{} completed successfully", getConnectorName());
        } catch (MSTeamsException e) {
            log.error("{} failed: {}", getConnectorName(), e.getMessage());
            setErrorOutputs(e.toOutputMessage());
            throw new ConnectorException(MSTeamsException.truncateMessage(e.getMessage()), e);
        } catch (Exception e) {
            log.error("{} failed unexpectedly: {}", getConnectorName(), e.getMessage(), e);
            setErrorOutputs("Unexpected error: " + e.getMessage());
            throw new ConnectorException(MSTeamsException.truncateMessage("Unexpected error: " + e.getMessage()), e);
        }
    }

    protected abstract void executeOperation(MSTeamsClient client) throws MSTeamsException;

    @Override
    public void disconnect() throws ConnectorException {
        log.debug("Disconnecting from Microsoft Teams for {}", getConnectorName());
        if (client != null) { try { client.close(); } catch (Exception e) { log.warn("Error closing client: {}", e.getMessage()); } }
    }

    protected void setSuccessOutputs() {
        setOutputParameter(OUTPUT_SUCCESS, true);
        setOutputParameter(OUTPUT_ERROR_MESSAGE, "");
    }

    protected void setErrorOutputs(String message) {
        setOutputParameter(OUTPUT_SUCCESS, false);
        setOutputParameter(OUTPUT_ERROR_MESSAGE, MSTeamsException.truncateMessage(message));
    }

    public Map<String, Object> getOutputs() { return getOutputParameters(); }

    public String getStringInput(String name) {
        Object v = getInputParameter(name); return v != null ? v.toString() : null;
    }

    public String getStringInputOrDefault(String name, String def) {
        String v = getStringInput(name); return (v != null && !v.isBlank()) ? v : def;
    }

    public Integer getIntInput(String name) {
        Object v = getInputParameter(name);
        if (v == null) return null;
        if (v instanceof Integer i) return i;
        if (v instanceof Number n) return n.intValue();
        try { return Integer.parseInt(v.toString()); } catch (NumberFormatException e) { return null; }
    }

    public int getIntInputOrDefault(String name, int def) {
        Integer v = getIntInput(name); return v != null ? v : def;
    }

    public Boolean getBooleanInput(String name) {
        Object v = getInputParameter(name);
        if (v == null) return null;
        if (v instanceof Boolean b) return b;
        return Boolean.parseBoolean(v.toString());
    }

    public boolean getBooleanInputOrDefault(String name, boolean def) {
        Boolean v = getBooleanInput(name); return v != null ? v : def;
    }

    @SuppressWarnings("unchecked")
    public <T> List<T> getListInput(String name) {
        Object v = getInputParameter(name);
        if (v instanceof List) return (List<T>) v;
        return null;
    }

    protected static boolean isBlank(String s) { return s == null || s.isBlank(); }
}
