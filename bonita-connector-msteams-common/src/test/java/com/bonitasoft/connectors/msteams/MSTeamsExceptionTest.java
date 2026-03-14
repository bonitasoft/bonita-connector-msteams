package com.bonitasoft.connectors.msteams;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

@DisplayName("MSTeamsException")
class MSTeamsExceptionTest {

    @Nested
    @DisplayName("truncateMessage")
    class TruncateMessage {

        @Test
        void should_return_null_when_null() {
            assertThat(MSTeamsException.truncateMessage(null)).isNull();
        }

        @Test
        void should_return_short_message_unchanged() {
            assertThat(MSTeamsException.truncateMessage("short")).isEqualTo("short");
        }

        @Test
        void should_truncate_long_message() {
            String longMsg = "x".repeat(1100);
            String result = MSTeamsException.truncateMessage(longMsg);
            assertThat(result).hasSize(MSTeamsException.MAX_ERROR_MESSAGE_LENGTH);
            assertThat(result).endsWith("...");
        }

        @Test
        void should_not_truncate_message_at_exact_limit() {
            String exactMsg = "x".repeat(MSTeamsException.MAX_ERROR_MESSAGE_LENGTH);
            assertThat(MSTeamsException.truncateMessage(exactMsg)).isEqualTo(exactMsg);
        }
    }

    @Nested
    @DisplayName("toOutputMessage")
    class ToOutputMessage {

        @Test
        void should_format_with_status_code() {
            MSTeamsException ex = new MSTeamsException("test error", 404, "NotFound", false, null);
            String output = ex.toOutputMessage();
            assertThat(output).contains("HTTP 404").contains("NotFound").contains("test error");
        }

        @Test
        void should_return_plain_message_when_no_status_code() {
            MSTeamsException ex = new MSTeamsException("plain error");
            assertThat(ex.toOutputMessage()).isEqualTo("plain error");
        }
    }

    @Nested
    @DisplayName("fromGraphError")
    class FromGraphError {

        @Test
        void should_create_retryable_for_429() {
            MSTeamsException ex = MSTeamsException.fromGraphError(429, "TooManyRequests", "slow down");
            assertThat(ex.isRetryable()).isTrue();
            assertThat(ex.getStatusCode()).isEqualTo(429);
        }

        @Test
        void should_create_retryable_for_503() {
            MSTeamsException ex = MSTeamsException.fromGraphError(503, "serviceNotAvailable", "try later");
            assertThat(ex.isRetryable()).isTrue();
        }

        @Test
        void should_create_retryable_for_504() {
            MSTeamsException ex = MSTeamsException.fromGraphError(504, null, "timeout");
            assertThat(ex.isRetryable()).isTrue();
        }

        @Test
        void should_create_retryable_for_activityLimitReached() {
            MSTeamsException ex = MSTeamsException.fromGraphError(403, "activityLimitReached", "limit");
            assertThat(ex.isRetryable()).isTrue();
        }

        @Test
        void should_not_be_retryable_for_400() {
            MSTeamsException ex = MSTeamsException.fromGraphError(400, "BadRequest", "bad");
            assertThat(ex.isRetryable()).isFalse();
        }

        @Test
        void should_not_be_retryable_for_401() {
            MSTeamsException ex = MSTeamsException.fromGraphError(401, null, "auth");
            assertThat(ex.isRetryable()).isFalse();
        }

        @Test
        void should_not_be_retryable_for_404() {
            MSTeamsException ex = MSTeamsException.fromGraphError(404, "NotFound", "missing");
            assertThat(ex.isRetryable()).isFalse();
        }

        @Test
        void should_handle_null_error_code() {
            MSTeamsException ex = MSTeamsException.fromGraphError(500, null, "server error");
            assertThat(ex.getReason()).isEqualTo("unknown");
        }

        @Test
        void should_create_retryable_for_serviceNotAvailable() {
            MSTeamsException ex = MSTeamsException.fromGraphError(500, "serviceNotAvailable", "unavailable");
            assertThat(ex.isRetryable()).isTrue();
        }

        @Test
        void should_not_be_retryable_for_403_generic() {
            MSTeamsException ex = MSTeamsException.fromGraphError(403, "Forbidden", "denied");
            assertThat(ex.isRetryable()).isFalse();
        }

        @Test
        void should_create_with_409_status() {
            MSTeamsException ex = MSTeamsException.fromGraphError(409, "Conflict", "conflict");
            assertThat(ex.isRetryable()).isFalse();
            assertThat(ex.getStatusCode()).isEqualTo(409);
        }
    }

    @Nested
    @DisplayName("createUserFriendlyMessage")
    class UserFriendlyMessage {

        @Test
        void should_format_400() {
            assertThat(MSTeamsException.createUserFriendlyMessage(400, null, "bad input"))
                    .startsWith("Invalid request:");
        }

        @Test
        void should_format_401() {
            assertThat(MSTeamsException.createUserFriendlyMessage(401, null, "x"))
                    .contains("Authentication failed");
        }

        @Test
        void should_format_403_with_activityLimit() {
            assertThat(MSTeamsException.createUserFriendlyMessage(403, "activityLimitReached", "x"))
                    .contains("Activity limit reached");
        }

        @Test
        void should_format_403_generic() {
            assertThat(MSTeamsException.createUserFriendlyMessage(403, "other", "denied"))
                    .startsWith("Permission denied:");
        }

        @Test
        void should_format_404() {
            assertThat(MSTeamsException.createUserFriendlyMessage(404, null, "x"))
                    .contains("Resource not found");
        }

        @Test
        void should_format_409() {
            assertThat(MSTeamsException.createUserFriendlyMessage(409, null, "conflict"))
                    .startsWith("Conflict:");
        }

        @Test
        void should_format_429() {
            assertThat(MSTeamsException.createUserFriendlyMessage(429, null, "x"))
                    .contains("Too many requests");
        }

        @Test
        void should_format_503_serviceNotAvailable() {
            assertThat(MSTeamsException.createUserFriendlyMessage(503, "serviceNotAvailable", "x"))
                    .contains("temporarily unavailable");
        }

        @Test
        void should_format_503_generic() {
            assertThat(MSTeamsException.createUserFriendlyMessage(503, "other", "x"))
                    .contains("Service unavailable");
        }

        @Test
        void should_format_504() {
            assertThat(MSTeamsException.createUserFriendlyMessage(504, null, "x"))
                    .contains("Gateway timeout");
        }

        @Test
        void should_return_original_for_unknown_code() {
            assertThat(MSTeamsException.createUserFriendlyMessage(418, null, "teapot"))
                    .isEqualTo("teapot");
        }
    }

    @Nested
    @DisplayName("Subclasses")
    class Subclasses {

        @Test
        void should_create_validation_exception() {
            MSTeamsException.ValidationException ex = new MSTeamsException.ValidationException("bad input");
            assertThat(ex).isInstanceOf(MSTeamsException.class);
            assertThat(ex.getMessage()).isEqualTo("bad input");
        }

        @Test
        void should_create_connection_exception() {
            MSTeamsException.ConnectionException ex = new MSTeamsException.ConnectionException("no conn", new RuntimeException("cause"));
            assertThat(ex).isInstanceOf(MSTeamsException.class);
            assertThat(ex.getCause()).isNotNull();
        }

        @Test
        void should_create_execution_exception_retryable() {
            MSTeamsException.ExecutionException ex = new MSTeamsException.ExecutionException("rate limit", 429);
            assertThat(ex.isRetryable()).isTrue();
            assertThat(ex.getStatusCode()).isEqualTo(429);
        }

        @Test
        void should_create_execution_exception_not_retryable() {
            MSTeamsException.ExecutionException ex = new MSTeamsException.ExecutionException("not found", 404);
            assertThat(ex.isRetryable()).isFalse();
        }

        @Test
        void should_create_execution_exception_with_cause() {
            Throwable cause = new RuntimeException("root");
            MSTeamsException.ExecutionException ex = new MSTeamsException.ExecutionException("error", 500, cause);
            assertThat(ex.getCause()).isEqualTo(cause);
        }
    }

    @Nested
    @DisplayName("Constructors")
    class Constructors {

        @Test
        void should_create_with_message_only() {
            MSTeamsException ex = new MSTeamsException("msg");
            assertThat(ex.getStatusCode()).isEqualTo(0);
            assertThat(ex.getReason()).isEqualTo("unknown");
            assertThat(ex.isRetryable()).isFalse();
        }

        @Test
        void should_create_with_message_and_cause() {
            Throwable cause = new RuntimeException("c");
            MSTeamsException ex = new MSTeamsException("msg", cause);
            assertThat(ex.getCause()).isEqualTo(cause);
            assertThat(ex.getStatusCode()).isEqualTo(0);
        }

        @Test
        void should_create_with_all_params() {
            MSTeamsException ex = new MSTeamsException("msg", 503, "ServiceUnavailable", true, null);
            assertThat(ex.getStatusCode()).isEqualTo(503);
            assertThat(ex.getReason()).isEqualTo("ServiceUnavailable");
            assertThat(ex.isRetryable()).isTrue();
        }

        @Test
        void should_default_reason_to_unknown_when_null() {
            MSTeamsException ex = new MSTeamsException("msg", 500, null, false, null);
            assertThat(ex.getReason()).isEqualTo("unknown");
        }
    }
}
