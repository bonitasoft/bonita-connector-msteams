package com.bonitasoft.connectors.msteams.common;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("TokenManager")
class TokenManagerTest {

    @Test
    void should_create_token_manager_with_credentials() {
        TokenManager tm = new TokenManager("tenant-1", "client-1", "secret-1");
        assertThat(tm).isNotNull();
    }

    @Test
    void should_throw_connection_exception_on_app_only_failure() {
        TokenManager tm = new TokenManager("invalid-tenant", "invalid-client", "invalid-secret");
        assertThatThrownBy(tm::getAppOnlyToken)
                .isInstanceOf(MsTeamsException.ConnectionException.class)
                .hasMessageContaining("Failed to acquire app-only token");
    }

    @Test
    void should_throw_connection_exception_on_delegated_failure() {
        TokenManager tm = new TokenManager("invalid-tenant", "invalid-client", "invalid-secret");
        assertThatThrownBy(() -> tm.getDelegatedToken("invalid-refresh-token"))
                .isInstanceOf(MsTeamsException.ConnectionException.class)
                .hasMessageContaining("Failed to acquire delegated token");
    }

    @Test
    void should_invalidate_cache() {
        TokenManager tm = new TokenManager("tenant", "client", "secret");
        tm.invalidateCache();
        // After invalidation, next call should attempt new token acquisition
        assertThatThrownBy(tm::getAppOnlyToken)
                .isInstanceOf(MsTeamsException.ConnectionException.class);
    }
}
