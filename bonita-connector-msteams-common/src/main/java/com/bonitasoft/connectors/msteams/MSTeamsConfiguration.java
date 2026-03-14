package com.bonitasoft.connectors.msteams;

import java.util.Map;

public record MSTeamsConfiguration(
        String tenantId,
        String clientId,
        String clientSecret,
        int connectTimeout,
        int readTimeout,
        boolean appOnly
) {
    private static final int DEFAULT_CONNECT_TIMEOUT = 30_000;
    private static final int DEFAULT_READ_TIMEOUT = 60_000;

    public MSTeamsConfiguration {
        if (connectTimeout < 0) throw new IllegalArgumentException("connectTimeout must be non-negative");
        if (readTimeout < 0) throw new IllegalArgumentException("readTimeout must be non-negative");
        if (connectTimeout == 0) connectTimeout = DEFAULT_CONNECT_TIMEOUT;
        if (readTimeout == 0) readTimeout = DEFAULT_READ_TIMEOUT;
    }

    public static MSTeamsConfiguration from(Map<String, Object> inputs) {
        return new MSTeamsConfiguration(
                str(inputs, "tenantId"), str(inputs, "clientId"), str(inputs, "clientSecret"),
                intOrDef(inputs, "connectTimeout", DEFAULT_CONNECT_TIMEOUT),
                intOrDef(inputs, "readTimeout", DEFAULT_READ_TIMEOUT),
                boolOrDef(inputs, "appOnly", false));
    }

    public void validate() {
        if (blank(tenantId) || blank(clientId) || blank(clientSecret))
            throw new MSTeamsException.ValidationException(
                    "Authentication required: tenantId, clientId, and clientSecret must all be provided.");
    }

    public boolean isAppOnly() { return appOnly; }

    private static String str(Map<String, Object> m, String k) {
        Object v = m.get(k); return v != null ? v.toString() : null;
    }
    private static int intOrDef(Map<String, Object> m, String k, int d) {
        Object v = m.get(k); if (v == null) return d;
        if (v instanceof Integer i) return i;
        try { return Integer.parseInt(v.toString()); } catch (NumberFormatException e) { return d; }
    }
    private static boolean boolOrDef(Map<String, Object> m, String k, boolean d) {
        Object v = m.get(k); if (v == null) return d;
        if (v instanceof Boolean b) return b;
        return Boolean.parseBoolean(v.toString());
    }
    private static boolean blank(String s) { return s == null || s.isBlank(); }
}
