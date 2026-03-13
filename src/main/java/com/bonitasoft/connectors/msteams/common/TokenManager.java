package com.bonitasoft.connectors.msteams.common;

import com.microsoft.aad.msal4j.ClientCredentialFactory;
import com.microsoft.aad.msal4j.ClientCredentialParameters;
import com.microsoft.aad.msal4j.ConfidentialClientApplication;
import com.microsoft.aad.msal4j.IAuthenticationResult;
import lombok.extern.slf4j.Slf4j;

import java.util.Collections;
import java.util.concurrent.CompletableFuture;

/**
 * MSAL4J token management with caching.
 * Supports app-only (client credentials) and delegated (refresh token) flows.
 */
@Slf4j
public class TokenManager {

    private static final String GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default";

    private final String tenantId;
    private final String clientId;
    private final String clientSecret;
    private ConfidentialClientApplication app;
    private String cachedToken;
    private long tokenExpiresAt;

    public TokenManager(String tenantId, String clientId, String clientSecret) {
        this.tenantId = tenantId;
        this.clientId = clientId;
        this.clientSecret = clientSecret;
    }

    /**
     * Get an access token using client credentials (app-only).
     */
    public String getAppOnlyToken() {
        if (cachedToken != null && System.currentTimeMillis() < tokenExpiresAt - 60_000) {
            return cachedToken;
        }
        try {
            if (app == null) {
                app = ConfidentialClientApplication.builder(
                                clientId, ClientCredentialFactory.createFromSecret(clientSecret))
                        .authority("https://login.microsoftonline.com/" + tenantId)
                        .build();
            }
            ClientCredentialParameters parameters = ClientCredentialParameters
                    .builder(Collections.singleton(GRAPH_DEFAULT_SCOPE))
                    .build();
            CompletableFuture<IAuthenticationResult> future = app.acquireToken(parameters);
            IAuthenticationResult result = future.join();
            cachedToken = result.accessToken();
            tokenExpiresAt = result.expiresOnDate().getTime();
            log.debug("Acquired app-only token, expires at {}", result.expiresOnDate());
            return cachedToken;
        } catch (Exception e) {
            throw new MsTeamsException.ConnectionException("Failed to acquire app-only token", e);
        }
    }

    /**
     * Get an access token using delegated permissions (refresh token flow).
     */
    public String getDelegatedToken(String refreshToken) {
        // For delegated flow, tokens should be managed per-request
        // as refresh tokens are user-specific
        try {
            if (app == null) {
                app = ConfidentialClientApplication.builder(
                                clientId, ClientCredentialFactory.createFromSecret(clientSecret))
                        .authority("https://login.microsoftonline.com/" + tenantId)
                        .build();
            }
            // Use refresh token to get new access token
            var parameters = com.microsoft.aad.msal4j.RefreshTokenParameters
                    .builder(Collections.singleton(GRAPH_DEFAULT_SCOPE), refreshToken)
                    .build();
            IAuthenticationResult result = app.acquireToken(parameters).join();
            log.debug("Acquired delegated token, expires at {}", result.expiresOnDate());
            return result.accessToken();
        } catch (Exception e) {
            throw new MsTeamsException.ConnectionException("Failed to acquire delegated token", e);
        }
    }

    public void invalidateCache() {
        cachedToken = null;
        tokenExpiresAt = 0;
    }
}
