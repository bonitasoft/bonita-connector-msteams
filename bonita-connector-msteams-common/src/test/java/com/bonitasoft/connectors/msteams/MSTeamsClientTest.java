package com.bonitasoft.connectors.msteams;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

@DisplayName("MSTeamsClient")
class MSTeamsClientTest {

    @Nested
    @DisplayName("Constructor and basic methods")
    class Basic {

        @Test
        void should_create_client_with_valid_configuration() {
            var config = new MSTeamsConfiguration("t", "c", "s", 5000, 10000, true);
            var client = new MSTeamsClient(config);
            assertThat(client.getConfiguration()).isEqualTo(config);
            assertThat(client.getObjectMapper()).isNotNull();
        }

        @Test
        void should_close_without_error() {
            var config = new MSTeamsConfiguration("t", "c", "s", 5000, 10000, true);
            var client = new MSTeamsClient(config);
            client.close(); // should not throw
        }
    }

    @Nested
    @DisplayName("Authentication")
    class Authentication {

        @Test
        void should_throw_on_invalid_tenant() {
            var config = new MSTeamsConfiguration("invalid-tenant", "c", "s", 1000, 1000, true);
            var client = new MSTeamsClient(config);
            assertThatThrownBy(client::authenticate)
                    .isInstanceOf(MSTeamsException.ConnectionException.class);
        }
    }

    @Nested
    @DisplayName("API calls without authentication")
    class UnauthenticatedCalls {

        @Test
        void should_throw_on_get_without_auth() {
            var config = new MSTeamsConfiguration("t", "c", "s", 1000, 1000, true);
            var client = new MSTeamsClient(config);
            // accessToken is null, should fail
            assertThatThrownBy(() -> client.get("/teams/test"))
                    .isInstanceOf(MSTeamsException.class);
        }

        @Test
        void should_throw_on_post_without_auth() {
            var config = new MSTeamsConfiguration("t", "c", "s", 1000, 1000, true);
            var client = new MSTeamsClient(config);
            assertThatThrownBy(() -> client.post("/teams/test", "{}"))
                    .isInstanceOf(MSTeamsException.class);
        }

        @Test
        void should_throw_on_delete_without_auth() {
            var config = new MSTeamsConfiguration("t", "c", "s", 1000, 1000, true);
            var client = new MSTeamsClient(config);
            assertThatThrownBy(() -> client.delete("/teams/test"))
                    .isInstanceOf(MSTeamsException.class);
        }

        @Test
        void should_throw_on_put_without_auth() {
            var config = new MSTeamsConfiguration("t", "c", "s", 1000, 1000, true);
            var client = new MSTeamsClient(config);
            assertThatThrownBy(() -> client.put("/test", new byte[]{1, 2, 3}, "application/octet-stream"))
                    .isInstanceOf(MSTeamsException.class);
        }
    }
}
