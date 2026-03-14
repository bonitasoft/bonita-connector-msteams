package com.bonitasoft.connectors.msteams;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.concurrent.atomic.AtomicInteger;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

@DisplayName("RetryPolicy")
class RetryPolicyTest {

    @Nested
    @DisplayName("execute")
    class Execute {

        @Test
        void should_return_result_on_first_attempt() throws Exception {
            RetryPolicy policy = new RetryPolicy(3);
            String result = policy.execute(() -> "success");
            assertThat(result).isEqualTo("success");
        }

        @Test
        void should_retry_on_retryable_exception() throws Exception {
            RetryPolicy policy = new TestableRetryPolicy(3);
            AtomicInteger attempts = new AtomicInteger(0);

            String result = policy.execute(() -> {
                if (attempts.incrementAndGet() < 3) {
                    throw new MSTeamsException("rate limited", 429, "TooManyRequests", true, null);
                }
                return "recovered";
            });

            assertThat(result).isEqualTo("recovered");
            assertThat(attempts.get()).isEqualTo(3);
        }

        @Test
        void should_throw_immediately_on_non_retryable_exception() {
            RetryPolicy policy = new RetryPolicy(3);
            AtomicInteger attempts = new AtomicInteger(0);

            assertThatThrownBy(() -> policy.execute(() -> {
                attempts.incrementAndGet();
                throw new MSTeamsException("not found", 404, "NotFound", false, null);
            }))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("not found");

            assertThat(attempts.get()).isEqualTo(1);
        }

        @Test
        void should_throw_after_max_retries_exceeded() {
            TestableRetryPolicy policy = new TestableRetryPolicy(2);
            AtomicInteger attempts = new AtomicInteger(0);

            assertThatThrownBy(() -> policy.execute(() -> {
                attempts.incrementAndGet();
                throw new MSTeamsException("rate limited", 429, "TooManyRequests", true, null);
            }))
                    .isInstanceOf(MSTeamsException.class);

            assertThat(attempts.get()).isEqualTo(3); // initial + 2 retries
        }

        @Test
        void should_wrap_unexpected_exceptions() {
            RetryPolicy policy = new RetryPolicy(3);

            assertThatThrownBy(() -> policy.execute(() -> {
                throw new RuntimeException("unexpected");
            }))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("Unexpected error");
        }
    }

    @Nested
    @DisplayName("calculateWait")
    class CalculateWait {

        @Test
        void should_increase_exponentially() {
            RetryPolicy policy = new RetryPolicy(5);
            long wait0 = policy.calculateWait(0);
            long wait1 = policy.calculateWait(1);
            long wait2 = policy.calculateWait(2);

            // Each wait should be roughly double (with jitter)
            assertThat(wait0).isGreaterThanOrEqualTo(1000L);
            assertThat(wait1).isGreaterThanOrEqualTo(2000L);
            assertThat(wait2).isGreaterThanOrEqualTo(4000L);
        }

        @Test
        void should_cap_at_max_wait() {
            RetryPolicy policy = new RetryPolicy(20);
            long wait = policy.calculateWait(10); // 2^10 * 1000 = 1024000 > 64000
            assertThat(wait).isLessThanOrEqualTo(64000L + 32000L); // max + half jitter
        }
    }

    @Nested
    @DisplayName("isRetryableStatusCode")
    class IsRetryable {

        @Test
        void should_return_true_for_429() {
            assertThat(RetryPolicy.isRetryableStatusCode(429)).isTrue();
        }

        @Test
        void should_return_true_for_500() {
            assertThat(RetryPolicy.isRetryableStatusCode(500)).isTrue();
        }

        @Test
        void should_return_true_for_502() {
            assertThat(RetryPolicy.isRetryableStatusCode(502)).isTrue();
        }

        @Test
        void should_return_true_for_503() {
            assertThat(RetryPolicy.isRetryableStatusCode(503)).isTrue();
        }

        @Test
        void should_return_false_for_404() {
            assertThat(RetryPolicy.isRetryableStatusCode(404)).isFalse();
        }

        @Test
        void should_return_false_for_200() {
            assertThat(RetryPolicy.isRetryableStatusCode(200)).isFalse();
        }
    }

    @Nested
    @DisplayName("sleep")
    class Sleep {

        @Test
        void should_throw_on_interrupt() {
            RetryPolicy policy = new RetryPolicy(1);
            Thread.currentThread().interrupt();
            assertThatThrownBy(() -> policy.sleep(1))
                    .isInstanceOf(MSTeamsException.class)
                    .hasMessageContaining("interrupted");
            // Clear interrupted status
            Thread.interrupted();
        }
    }

    /**
     * Test-friendly subclass that overrides sleep to avoid actual delays.
     */
    static class TestableRetryPolicy extends RetryPolicy {
        TestableRetryPolicy(int maxRetries) {
            super(maxRetries);
        }

        @Override
        void sleep(long millis) {
            // No-op for tests
        }
    }
}
