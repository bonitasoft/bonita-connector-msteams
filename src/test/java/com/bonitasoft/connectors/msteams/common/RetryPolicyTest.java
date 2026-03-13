package com.bonitasoft.connectors.msteams.common;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.DisplayName;

class RetryPolicyTest {

    @Test
    @DisplayName("Should execute successfully on first attempt")
    void should_execute_successfully_on_first_attempt() {
        var policy = new RetryPolicy(3);
        String result = policy.execute(() -> "success");
        assertThat(result).isEqualTo("success");
    }

    @Test
    @DisplayName("Should retry on retryable exception")
    void should_retry_on_retryable_exception() {
        var policy = new TestRetryPolicy(3);
        var counter = new int[]{0};

        String result = policy.execute(() -> {
            counter[0]++;
            if (counter[0] < 2) {
                throw new MsTeamsException.ExecutionException("Rate limited", 429);
            }
            return "success";
        });

        assertThat(result).isEqualTo("success");
        assertThat(counter[0]).isEqualTo(2);
    }

    @Test
    @DisplayName("Should not retry on non-retryable exception")
    void should_not_retry_on_non_retryable_exception() {
        var policy = new RetryPolicy(3);
        assertThatThrownBy(() -> policy.execute(() -> {
            throw new MsTeamsException.ExecutionException("Not found", 404);
        }))
                .isInstanceOf(MsTeamsException.ExecutionException.class)
                .hasMessageContaining("Not found");
    }

    @Test
    @DisplayName("Should throw after max retries")
    void should_throw_after_max_retries() {
        var policy = new TestRetryPolicy(2);
        assertThatThrownBy(() -> policy.execute(() -> {
            throw new MsTeamsException.ExecutionException("Server error", 500);
        }))
                .isInstanceOf(MsTeamsException.ExecutionException.class);
    }

    @Test
    @DisplayName("Should calculate exponential wait with cap")
    void should_calculate_exponential_wait() {
        var policy = new RetryPolicy(3);
        long wait0 = policy.calculateWait(0);
        long wait5 = policy.calculateWait(5);
        long wait10 = policy.calculateWait(10);

        assertThat(wait0).isBetween(1000L, 1500L);
        assertThat(wait5).isBetween(32000L, 48000L);
        assertThat(wait10).isBetween(64000L, 96000L); // capped at MAX_WAIT_MS
    }

    @Test
    @DisplayName("Should identify retryable status codes")
    void should_identify_retryable_status_codes() {
        assertThat(RetryPolicy.isRetryableStatusCode(429)).isTrue();
        assertThat(RetryPolicy.isRetryableStatusCode(500)).isTrue();
        assertThat(RetryPolicy.isRetryableStatusCode(502)).isTrue();
        assertThat(RetryPolicy.isRetryableStatusCode(503)).isTrue();
        assertThat(RetryPolicy.isRetryableStatusCode(400)).isFalse();
        assertThat(RetryPolicy.isRetryableStatusCode(404)).isFalse();
        assertThat(RetryPolicy.isRetryableStatusCode(200)).isFalse();
    }

    /** Test subclass that skips actual sleep */
    static class TestRetryPolicy extends RetryPolicy {
        TestRetryPolicy(int maxRetries) {
            super(maxRetries);
        }

        @Override
        void sleep(long millis) {
            // No-op for tests
        }
    }
}
