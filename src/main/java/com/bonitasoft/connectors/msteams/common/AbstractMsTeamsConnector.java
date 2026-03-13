package com.bonitasoft.connectors.msteams.common;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.bonitasoft.engine.connector.AbstractConnector;
import org.bonitasoft.engine.connector.ConnectorException;
import org.bonitasoft.engine.connector.ConnectorValidationException;

/**
 * Abstract base connector for MS Teams operations.
 * <p>
 * Implements the Bonita 4-phase lifecycle:
 * <ol>
 *   <li>VALIDATE — check inputs, fail fast</li>
 *   <li>CONNECT — create client, authenticate</li>
 *   <li>EXECUTE — call API, set outputs</li>
 *   <li>DISCONNECT — close client, release resources</li>
 * </ol>
 */
public abstract class AbstractMsTeamsConnector extends AbstractConnector {

    protected static final Logger LOGGER = Logger.getLogger(AbstractMsTeamsConnector.class.getName());

    // Standard output parameters
    public static final String OUTPUT_SUCCESS = "success";
    public static final String OUTPUT_ERROR_MESSAGE = "errorMessage";

    // Common input parameters
    public static final String INPUT_CONNECT_TIMEOUT = "connectTimeout";
    public static final String INPUT_READ_TIMEOUT = "readTimeout";

    // --- Phase 1: VALIDATE ---

    @Override
    public void validateInputParameters() throws ConnectorValidationException {
        List<String> errors = new ArrayList<>();
        validateConnectionParameters(errors);
        validateOperationParameters(errors);
        if (!errors.isEmpty()) {
            throw new ConnectorValidationException(this, errors);
        }
    }

    /**
     * Validate connection-level parameters (auth, timeouts).
     * Subclasses (AbstractWebhookConnector, AbstractGraphConnector) override this.
     */
    protected abstract void validateConnectionParameters(List<String> errors);

    /**
     * Validate operation-specific parameters.
     * Each concrete connector overrides this.
     */
    protected abstract void validateOperationParameters(List<String> errors);

    // --- Output accessors (for testing) ---

    /**
     * Public accessor for output parameters map.
     * Delegates to the protected method from AbstractConnector.
     */
    public Map<String, Object> getOutputs() {
        return getOutputParameters();
    }

    // --- Test support ---

    /**
     * Public accessor to execute business logic from integration tests
     * in a different package. Bypasses validate/connect/disconnect phases.
     */
    public void executeForTest() throws ConnectorException {
        executeBusinessLogic();
    }

    // --- Output helpers ---

    protected void setSuccessOutputs() {
        setOutputParameter(OUTPUT_SUCCESS, true);
        setOutputParameter(OUTPUT_ERROR_MESSAGE, "");
    }

    protected void setErrorOutputs(String message) {
        setOutputParameter(OUTPUT_SUCCESS, false);
        String truncated = message != null && message.length() > 1000
                ? message.substring(0, 1000) : message;
        setOutputParameter(OUTPUT_ERROR_MESSAGE, truncated);
    }

    // --- Input helpers ---

    protected String getStringInput(String name) {
        Object value = getInputParameter(name);
        return value != null ? value.toString() : null;
    }

    protected Integer getIntegerInput(String name) {
        Object value = getInputParameter(name);
        if (value == null) return null;
        if (value instanceof Integer i) return i;
        return Integer.valueOf(value.toString());
    }

    protected int getIntegerInputOrDefault(String name, int defaultValue) {
        Integer value = getIntegerInput(name);
        return value != null ? value : defaultValue;
    }

    protected Boolean getBooleanInput(String name) {
        Object value = getInputParameter(name);
        if (value == null) return null;
        if (value instanceof Boolean b) return b;
        return Boolean.valueOf(value.toString());
    }

    protected boolean getBooleanInputOrDefault(String name, boolean defaultValue) {
        Boolean value = getBooleanInput(name);
        return value != null ? value : defaultValue;
    }

    protected boolean isNullOrEmpty(String value) {
        return value == null || value.trim().isEmpty();
    }
}
