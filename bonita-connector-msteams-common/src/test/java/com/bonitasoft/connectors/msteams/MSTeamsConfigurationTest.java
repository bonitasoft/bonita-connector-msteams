package com.bonitasoft.connectors.msteams;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

@DisplayName("MSTeamsConfiguration")
class MSTeamsConfigurationTest {

    @Nested
    @DisplayName("Record constructor")
    class Constructor {

        @Test
        void should_create_with_valid_params() {
            var config = new MSTeamsConfiguration("t", "c", "s", 5000, 10000, true);
            assertThat(config.tenantId()).isEqualTo("t");
            assertThat(config.clientId()).isEqualTo("c");
            assertThat(config.clientSecret()).isEqualTo("s");
            assertThat(config.connectTimeout()).isEqualTo(5000);
            assertThat(config.readTimeout()).isEqualTo(10000);
            assertThat(config.isAppOnly()).isTrue();
        }

        @Test
        void should_use_default_connect_timeout_when_zero() {
            var config = new MSTeamsConfiguration("t", "c", "s", 0, 10000, false);
            assertThat(config.connectTimeout()).isEqualTo(30_000);
        }

        @Test
        void should_use_default_read_timeout_when_zero() {
            var config = new MSTeamsConfiguration("t", "c", "s", 5000, 0, false);
            assertThat(config.readTimeout()).isEqualTo(60_000);
        }

        @Test
        void should_reject_negative_connect_timeout() {
            assertThatThrownBy(() -> new MSTeamsConfiguration("t", "c", "s", -1, 10000, false))
                    .isInstanceOf(IllegalArgumentException.class)
                    .hasMessageContaining("connectTimeout");
        }

        @Test
        void should_reject_negative_read_timeout() {
            assertThatThrownBy(() -> new MSTeamsConfiguration("t", "c", "s", 5000, -1, false))
                    .isInstanceOf(IllegalArgumentException.class)
                    .hasMessageContaining("readTimeout");
        }
    }

    @Nested
    @DisplayName("from(Map)")
    class FromMap {

        @Test
        void should_create_from_map() {
            Map<String, Object> inputs = new HashMap<>();
            inputs.put("tenantId", "t1");
            inputs.put("clientId", "c1");
            inputs.put("clientSecret", "s1");
            inputs.put("connectTimeout", 15000);
            inputs.put("readTimeout", 30000);
            inputs.put("appOnly", true);

            var config = MSTeamsConfiguration.from(inputs);
            assertThat(config.tenantId()).isEqualTo("t1");
            assertThat(config.clientId()).isEqualTo("c1");
            assertThat(config.connectTimeout()).isEqualTo(15000);
            assertThat(config.isAppOnly()).isTrue();
        }

        @Test
        void should_use_defaults_for_missing_values() {
            Map<String, Object> inputs = new HashMap<>();
            inputs.put("tenantId", "t");
            inputs.put("clientId", "c");
            inputs.put("clientSecret", "s");

            var config = MSTeamsConfiguration.from(inputs);
            assertThat(config.connectTimeout()).isEqualTo(30_000);
            assertThat(config.readTimeout()).isEqualTo(60_000);
            assertThat(config.isAppOnly()).isFalse();
        }

        @Test
        void should_handle_string_integer_values() {
            Map<String, Object> inputs = new HashMap<>();
            inputs.put("tenantId", "t");
            inputs.put("clientId", "c");
            inputs.put("clientSecret", "s");
            inputs.put("connectTimeout", "5000");
            inputs.put("readTimeout", "10000");

            var config = MSTeamsConfiguration.from(inputs);
            assertThat(config.connectTimeout()).isEqualTo(5000);
            assertThat(config.readTimeout()).isEqualTo(10000);
        }

        @Test
        void should_use_default_for_invalid_integer() {
            Map<String, Object> inputs = new HashMap<>();
            inputs.put("tenantId", "t");
            inputs.put("clientId", "c");
            inputs.put("clientSecret", "s");
            inputs.put("connectTimeout", "not-a-number");

            var config = MSTeamsConfiguration.from(inputs);
            assertThat(config.connectTimeout()).isEqualTo(30_000);
        }

        @Test
        void should_handle_null_map_values() {
            Map<String, Object> inputs = new HashMap<>();
            var config = MSTeamsConfiguration.from(inputs);
            assertThat(config.tenantId()).isNull();
        }

        @Test
        void should_handle_string_boolean_values() {
            Map<String, Object> inputs = new HashMap<>();
            inputs.put("tenantId", "t");
            inputs.put("clientId", "c");
            inputs.put("clientSecret", "s");
            inputs.put("appOnly", "true");
            var config = MSTeamsConfiguration.from(inputs);
            assertThat(config.isAppOnly()).isTrue();
        }

        @Test
        void should_handle_boolean_value_directly() {
            Map<String, Object> inputs = new HashMap<>();
            inputs.put("tenantId", "t");
            inputs.put("clientId", "c");
            inputs.put("clientSecret", "s");
            inputs.put("appOnly", Boolean.TRUE);
            var config = MSTeamsConfiguration.from(inputs);
            assertThat(config.isAppOnly()).isTrue();
        }

        @Test
        void should_handle_integer_value_directly_in_map() {
            Map<String, Object> inputs = new HashMap<>();
            inputs.put("tenantId", "t");
            inputs.put("clientId", "c");
            inputs.put("clientSecret", "s");
            inputs.put("connectTimeout", Integer.valueOf(7000));
            var config = MSTeamsConfiguration.from(inputs);
            assertThat(config.connectTimeout()).isEqualTo(7000);
        }

        @Test
        void should_use_default_for_null_boolean() {
            Map<String, Object> inputs = new HashMap<>();
            inputs.put("tenantId", "t");
            inputs.put("clientId", "c");
            inputs.put("clientSecret", "s");
            // appOnly not set at all
            var config = MSTeamsConfiguration.from(inputs);
            assertThat(config.isAppOnly()).isFalse();
        }

        @Test
        void should_use_default_for_null_timeout() {
            Map<String, Object> inputs = new HashMap<>();
            inputs.put("tenantId", "t");
            inputs.put("clientId", "c");
            inputs.put("clientSecret", "s");
            inputs.put("connectTimeout", null);
            inputs.put("readTimeout", null);
            var config = MSTeamsConfiguration.from(inputs);
            assertThat(config.connectTimeout()).isEqualTo(30_000);
            assertThat(config.readTimeout()).isEqualTo(60_000);
        }
    }

    @Nested
    @DisplayName("validate")
    class Validate {

        @Test
        void should_pass_with_all_auth_fields() {
            var config = new MSTeamsConfiguration("t", "c", "s", 5000, 10000, false);
            config.validate(); // should not throw
        }

        @Test
        void should_fail_when_tenantId_blank() {
            var config = new MSTeamsConfiguration("", "c", "s", 5000, 10000, false);
            assertThatThrownBy(config::validate)
                    .isInstanceOf(MSTeamsException.ValidationException.class)
                    .hasMessageContaining("Authentication required");
        }

        @Test
        void should_fail_when_clientId_null() {
            var config = new MSTeamsConfiguration("t", null, "s", 5000, 10000, false);
            assertThatThrownBy(config::validate)
                    .isInstanceOf(MSTeamsException.ValidationException.class);
        }

        @Test
        void should_fail_when_clientSecret_blank() {
            var config = new MSTeamsConfiguration("t", "c", "  ", 5000, 10000, false);
            assertThatThrownBy(config::validate)
                    .isInstanceOf(MSTeamsException.ValidationException.class);
        }
    }
}
