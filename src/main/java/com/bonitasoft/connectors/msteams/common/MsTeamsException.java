package com.bonitasoft.connectors.msteams.common;

/**
 * Exception hierarchy for MS Teams connector.
 *
 * <pre>
 * MsTeamsException (base — extends RuntimeException)
 * ├── ValidationException  (Phase 1: bad inputs)
 * ├── ConnectionException  (Phase 2: connection failures)
 * └── ExecutionException   (Phase 3: API errors)
 * </pre>
 */
public class MsTeamsException extends RuntimeException {

    private final boolean retryable;

    public MsTeamsException(String message) {
        super(message);
        this.retryable = false;
    }

    public MsTeamsException(String message, Throwable cause) {
        super(message, cause);
        this.retryable = false;
    }

    public MsTeamsException(String message, Throwable cause, boolean retryable) {
        super(message, cause);
        this.retryable = retryable;
    }

    public boolean isRetryable() {
        return retryable;
    }

    public static class ValidationException extends MsTeamsException {
        public ValidationException(String message) {
            super(message);
        }
    }

    public static class ConnectionException extends MsTeamsException {
        public ConnectionException(String message, Throwable cause) {
            super(message, cause);
        }
    }

    public static class ExecutionException extends MsTeamsException {
        private final int httpStatusCode;

        public ExecutionException(String message, int httpStatusCode) {
            super(message, null, isRetryableStatus(httpStatusCode));
            this.httpStatusCode = httpStatusCode;
        }

        public ExecutionException(String message, int httpStatusCode, Throwable cause) {
            super(message, cause, isRetryableStatus(httpStatusCode));
            this.httpStatusCode = httpStatusCode;
        }

        public int getHttpStatusCode() {
            return httpStatusCode;
        }

        private static boolean isRetryableStatus(int code) {
            return code == 429 || (code >= 500 && code <= 503);
        }
    }
}
