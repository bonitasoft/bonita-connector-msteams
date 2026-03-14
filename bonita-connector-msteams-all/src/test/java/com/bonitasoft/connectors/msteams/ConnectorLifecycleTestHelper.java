package com.bonitasoft.connectors.msteams;

import org.bonitasoft.engine.connector.ConnectorException;
import org.bonitasoft.engine.connector.ConnectorValidationException;

/**
 * Test helper in the same package as {@link AbstractMSTeamsConnector}
 * to access the protected {@code executeBusinessLogic()} method.
 * <p>
 * Simulates the Bonita engine lifecycle with a mocked MSTeamsClient:
 * validate → inject client → executeBusinessLogic → disconnect.
 */
public final class ConnectorLifecycleTestHelper {

    private ConnectorLifecycleTestHelper() {}

    /**
     * Executes the full connector lifecycle using a pre-configured mock client
     * instead of authenticating against the real Microsoft Graph API.
     *
     * @param connector  connector with input parameters already set
     * @param mockClient mocked MSTeamsClient
     * @throws ConnectorValidationException if input validation fails
     * @throws ConnectorException           if execution or disconnect fails
     */
    public static void executeWithMockClient(AbstractMSTeamsConnector connector,
                                              MSTeamsClient mockClient)
            throws ConnectorValidationException, ConnectorException {
        connector.validateInputParameters();
        connector.setClient(mockClient);
        connector.executeBusinessLogic();
        connector.disconnect();
    }
}
