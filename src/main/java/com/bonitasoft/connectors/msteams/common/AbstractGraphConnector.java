package com.bonitasoft.connectors.msteams.common;

import java.util.List;

import org.bonitasoft.engine.connector.ConnectorException;

/**
 * Base class for Graph API-based connectors.
 * Supports app-only (client credentials) and delegated (refresh token) auth.
 */
public abstract class AbstractGraphConnector extends AbstractMsTeamsConnector {

    public static final String INPUT_TENANT_ID = "tenantId";
    public static final String INPUT_CLIENT_ID = "clientId";
    public static final String INPUT_CLIENT_SECRET = "clientSecret";
    public static final String INPUT_REFRESH_TOKEN = "refreshToken";

    protected MsTeamsGraphClient graphClient;

    /**
     * Auth mode: "app-only" (default) or "delegated".
     * Subclasses override to change.
     */
    protected String getAuthMode() {
        return "app-only";
    }

    @Override
    protected void validateConnectionParameters(List<String> errors) {
        if (isNullOrEmpty(getStringInput(INPUT_TENANT_ID))) {
            errors.add("tenantId is required");
        }
        if (isNullOrEmpty(getStringInput(INPUT_CLIENT_ID))) {
            errors.add("clientId is required");
        }
        if (isNullOrEmpty(getStringInput(INPUT_CLIENT_SECRET))) {
            errors.add("clientSecret is required");
        }
        if ("delegated".equals(getAuthMode()) && isNullOrEmpty(getStringInput(INPUT_REFRESH_TOKEN))) {
            errors.add("refreshToken is required for delegated auth");
        }

        Integer connectTimeout = getIntegerInput(INPUT_CONNECT_TIMEOUT);
        if (connectTimeout != null && connectTimeout < 0) {
            errors.add("connectTimeout must be a positive number");
        }
    }

    @Override
    public void connect() throws ConnectorException {
        try {
            String tenantId = getStringInput(INPUT_TENANT_ID);
            String clientId = getStringInput(INPUT_CLIENT_ID);
            String clientSecret = getStringInput(INPUT_CLIENT_SECRET);
            int connectTimeout = getIntegerInputOrDefault(INPUT_CONNECT_TIMEOUT, 30000);

            TokenManager tokenManager = new TokenManager(tenantId, clientId, clientSecret);
            graphClient = new MsTeamsGraphClient(tokenManager, connectTimeout, 3);

            if ("delegated".equals(getAuthMode())) {
                graphClient.authenticateDelegated(getStringInput(INPUT_REFRESH_TOKEN));
            } else {
                graphClient.authenticateAppOnly();
            }
        } catch (MsTeamsException e) {
            throw new ConnectorException("Failed to connect to Graph API: " + e.getMessage(), e);
        }
    }

    @Override
    public void disconnect() throws ConnectorException {
        if (graphClient != null) {
            try {
                graphClient.close();
            } catch (Exception e) {
                LOGGER.warning("Error closing Graph client: " + e.getMessage());
            }
        }
    }

    public void setGraphClient(MsTeamsGraphClient client) {
        this.graphClient = client;
    }
}
