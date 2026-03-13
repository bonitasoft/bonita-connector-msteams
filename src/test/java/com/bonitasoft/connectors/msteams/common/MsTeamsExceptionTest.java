package com.bonitasoft.connectors.msteams.common;

import static org.assertj.core.api.Assertions.assertThat;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.DisplayName;

class MsTeamsExceptionTest {

    @Test
    @DisplayName("Should create base exception as non-retryable")
    void should_create_base_exception_as_non_retryable() {
        var ex = new MsTeamsException("test");
        assertThat(ex.isRetryable()).isFalse();
        assertThat(ex.getMessage()).isEqualTo("test");
    }

    @Test
    @DisplayName("Should create retryable exception")
    void should_create_retryable_exception() {
        var ex = new MsTeamsException("retry", null, true);
        assertThat(ex.isRetryable()).isTrue();
    }

    @Test
    @DisplayName("Should create validation exception")
    void should_create_validation_exception() {
        var ex = new MsTeamsException.ValidationException("bad input");
        assertThat(ex).isInstanceOf(MsTeamsException.class);
        assertThat(ex.isRetryable()).isFalse();
    }

    @Test
    @DisplayName("Should create connection exception")
    void should_create_connection_exception() {
        var cause = new RuntimeException("timeout");
        var ex = new MsTeamsException.ConnectionException("conn failed", cause);
        assertThat(ex.getCause()).isEqualTo(cause);
        assertThat(ex.isRetryable()).isFalse();
    }

    @Test
    @DisplayName("Should mark 429 as retryable")
    void should_mark_429_as_retryable() {
        var ex = new MsTeamsException.ExecutionException("rate limited", 429);
        assertThat(ex.isRetryable()).isTrue();
        assertThat(ex.getHttpStatusCode()).isEqualTo(429);
    }

    @Test
    @DisplayName("Should mark 5xx as retryable")
    void should_mark_5xx_as_retryable() {
        assertThat(new MsTeamsException.ExecutionException("err", 500).isRetryable()).isTrue();
        assertThat(new MsTeamsException.ExecutionException("err", 502).isRetryable()).isTrue();
        assertThat(new MsTeamsException.ExecutionException("err", 503).isRetryable()).isTrue();
    }

    @Test
    @DisplayName("Should mark 4xx (non-429) as non-retryable")
    void should_mark_4xx_as_non_retryable() {
        assertThat(new MsTeamsException.ExecutionException("err", 400).isRetryable()).isFalse();
        assertThat(new MsTeamsException.ExecutionException("err", 404).isRetryable()).isFalse();
    }
}
