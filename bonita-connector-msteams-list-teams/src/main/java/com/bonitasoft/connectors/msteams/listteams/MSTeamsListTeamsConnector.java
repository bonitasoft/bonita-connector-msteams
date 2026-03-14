package com.bonitasoft.connectors.msteams.listteams;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import com.bonitasoft.connectors.msteams.AbstractMSTeamsConnector;
import com.bonitasoft.connectors.msteams.MSTeamsClient;
import com.bonitasoft.connectors.msteams.MSTeamsException;
import com.fasterxml.jackson.databind.JsonNode;
import lombok.extern.slf4j.Slf4j;

@Slf4j
public class MSTeamsListTeamsConnector extends AbstractMSTeamsConnector {

    public static final String INPUT_FILTER = "filter";
    public static final String INPUT_MAX_RESULTS = "maxResults";
    public static final String OUTPUT_TEAMS = "teams";
    public static final String OUTPUT_TOTAL_COUNT = "totalCount";

    @Override protected String getConnectorName() { return "MSTeams-ListTeams"; }

    @Override
    protected void validateOperationInputs(List<String> errors) { /* No mandatory op inputs */ }

    @Override
    protected void executeOperation(MSTeamsClient client) throws MSTeamsException {
        String filter = getStringInput(INPUT_FILTER);
        int maxResults = getIntInputOrDefault(INPUT_MAX_RESULTS, 100);
        boolean appOnly = client.getConfiguration().isAppOnly();
        log.info("Listing teams (appOnly={}, maxResults={})", appOnly, maxResults);

        String path;
        if (appOnly) {
            String base = "resourceProvisioningOptions/Any(x:x eq 'Team')";
            String combined = (filter != null && !filter.isBlank()) ? base + " and " + filter : base;
            path = "/groups?$filter=" + enc(combined) + "&$top=" + maxResults + "&$select=id,displayName,description,visibility";
        } else {
            path = "/me/joinedTeams?$top=" + maxResults;
        }

        JsonNode response = client.get(path);
        List<Map<String, Object>> list = new ArrayList<>();
        JsonNode value = response.get("value");
        if (value != null && value.isArray()) {
            for (JsonNode n : value) {
                Map<String, Object> t = new HashMap<>();
                t.put("id", n.has("id") ? n.get("id").asText() : null);
                t.put("displayName", n.has("displayName") ? n.get("displayName").asText() : null);
                t.put("description", n.has("description") && !n.get("description").isNull() ? n.get("description").asText() : null);
                t.put("visibility", n.has("visibility") ? n.get("visibility").asText() : null);
                list.add(t);
            }
        }
        setOutputParameter(OUTPUT_TEAMS, list);
        setOutputParameter(OUTPUT_TOTAL_COUNT, list.size());
        log.info("Found {} teams", list.size());
    }

    private String enc(String v) { return v.replace(" ", "%20").replace("'", "%27"); }
}
