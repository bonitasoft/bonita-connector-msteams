package com.bonitasoft.connectors.msteams;

/**
 * Exception hierarchy for MS Teams connector.
 * <p>
 * Hierarchy:
 * <pre>
 * MSTeamsException (base — extends RuntimeException)
 * ├── ValidationException  (Phase 1: bad inputs)
 * ├── ConnectionException  (Phase 2: connection failures)
 * └── ExecutionException   (Phase 3: API errors)
 * </pre>
 */
public class MSTeamsException extends RuntimeException {

    /**
     * Maximum error message length for Bonita's bpm_failure table.
     * H2 column ERRORMESSAGE is VARCHAR_IGNORECASE(1024).
     */
    static final int MAX_ERROR_MESSAGE_LENGTH = 1000;

    private final int statusCode;
    private final String reason;
    private final boolean retryable;

    public MSTeamsException(String message) {
        super(message);
        this.statusCode = 0;
        this.reason = "unknown";
        this.retryable = false;
    }

    public MSTeamsException(String message, Throwable cause) {
        super(message, cause);
        this.statusCode = 0;
        this.reason = "unknown";
        this.retryable = false;
    }

    public MSTeamsException(String message, int statusCode, String reason, boolean retryable, Throwable cause) {
        super(message, cause);
        this.statusCode = statusCode;
        this.reason = reason != null ? reason : "unknown";
        this.retryable = retryable;
    }

    /**
     * HTTP status code (0 if not from HTTP response).
     */
    public int getStatusCode() {
        return statusCode;
    }

    /**
     * Error reason code from Microsoft Graph API.
     */
    public String getReason() {
        return reason;
    }

    /**
     * Whether this error can be retried (rate-limit, transient server error).
     */
    public boolean isRetryable() {
        return retryable;
    }

    /**
     * Truncate a message to fit Bonita's ERRORMESSAGE column (VARCHAR 1024).
     * Leaves a 24-char margin for toOutputMessage() prefix formatting.
     *
     * @param message the message to truncate
     * @return the truncated message, or null if input was null
     */
    static String truncateMessage(String message) {
        if (message == null || message.length() <= MAX_ERROR_MESSAGE_LENGTH) {
            return message;
        }
        return message.substring(0, MAX_ERROR_MESSAGE_LENGTH - 3) + "...";
    }

    /**
     * Format error message for Bonita output parameter.
     */
    public String toOutputMessage() {
        String msg = truncateMessage(getMessage());
        if (statusCode > 0) {
            String formatted = String.format("[HTTP %d - %s] %s", statusCode, reason, msg);
            return truncateMessage(formatted);
        }
        return msg;
    }

    /**
     * Create MSTeamsException from Microsoft Graph API error response.
     *
     * @param statusCode HTTP status code
     * @param errorCode  Graph API error code (e.g. "activityLimitReached")
     * @param message    error message from the response
     * @return a properly classified MSTeamsException
     */
    public static MSTeamsException fromGraphError(int statusCode, String errorCode, String message) {
        boolean retryable = isRetryableError(statusCode, errorCode);
        String userMessage = createUserFriendlyMessage(statusCode, errorCode, message);
        return new MSTeamsException(userMessage, statusCode, errorCode != null ? errorCode : "unknown", retryable, null);
    }

    /**
     * Determine if an error is retryable based on status code and error code.
     */
    static boolean isRetryableError(int code, String errorCode) {
        // Rate limit — always retryable
        if (code == 429) return true;
        // Server errors
        if (code == 503 || code == 504) return true;
        // Graph-specific retryable error codes
        if ("activityLimitReached".equals(errorCode)) return true;
        if ("serviceNotAvailable".equals(errorCode)) return true;

        return false;
    }

    /**
     * Create a user-friendly message based on HTTP status and error code.
     */
    static String createUserFriendlyMessage(int code, String errorCode, String originalMessage) {
        return switch (code) {
            case 400 -> "Invalid request: " + originalMessage;
            case 401 -> "Authentication failed. Check your tenant ID, client ID, and client secret.";
            case 403 -> {
                if ("activityLimitReached".equals(errorCode)) {
                    yield "Activity limit reached. The request will be retried automatically.";
                }
                yield "Permission denied: " + originalMessage;
            }
            case 404 -> "Resource not found. Verify the team, channel, or message ID is correct and you have access.";
            case 409 -> "Conflict: " + originalMessage;
            case 429 -> "Too many requests. The request will be retried automatically.";
            case 503 -> {
                if ("serviceNotAvailable".equals(errorCode)) {
                    yield "Microsoft Teams service temporarily unavailable. Retrying...";
                }
                yield "Service unavailable. Retrying...";
            }
            case 504 -> "Gateway timeout. Retrying...";
            default -> originalMessage;
        };
    }

    // --- Specialized subclasses ---

    /**
     * Phase 1 error: invalid input parameters.
     */
    public static class ValidationException extends MSTeamsException {
        public ValidationException(String message) {
            super(message);
        }
    }

    /**
     * Phase 2 error: cannot establish connection.
     */
    public static class ConnectionException extends MSTeamsException {
        public ConnectionException(String message, Throwable cause) {
            super(message, cause);
        }
    }

    /**
     * Phase 3 error: API returned an error during execution.
     */
    public static class ExecutionException extends MSTeamsException {
        public ExecutionException(String message, int httpStatusCode) {
            super(message, httpStatusCode, "execution_error", isRetryableStatus(httpStatusCode), null);
        }

        public ExecutionException(String message, int httpStatusCode, Throwable cause) {
            super(message, httpStatusCode, "execution_error", isRetryableStatus(httpStatusCode), cause);
        }

        private static boolean isRetryableStatus(int code) {
            return code == 429 || code == 503 || code == 504;
        }
    }
}
