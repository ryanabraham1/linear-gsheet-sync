import fs from "node:fs/promises";

const source = await fs.readFile("Code.local.gs", "utf8");
const apiKey = source.match(/LINEAR_API_KEY:\s*'([^']+)'/)[1];
const teamId = source.match(/LINEAR_TEAM_ID:\s*'([^']+)'/)[1];
const projectId = source.match(/LINEAR_PROJECT_ID:\s*'([^']+)'/)[1];
const labelId = source.match(/LINEAR_LABEL_IDS:\s*\['([^']+)'\]/)[1];

const query = `
  query SyncContext($teamId: String!, $projectId: String!, $labelId: String!) {
    viewer { name email }
    team(id: $teamId) { id name key }
    project(id: $projectId) { id name }
    issueLabel(id: $labelId) { id name }
  }
`;

const response = await fetch("https://api.linear.app/graphql", {
  method: "POST",
  headers: {
    Authorization: apiKey,
    "Content-Type": "application/json",
  },
  body: JSON.stringify({ query, variables: { teamId, projectId, labelId } }),
});
const body = await response.json();
if (!response.ok || body.errors) {
  throw new Error(JSON.stringify(body.errors || { status: response.status }));
}
console.log(
  JSON.stringify({
    viewer: body.data.viewer.name,
    team: body.data.team.key,
    project: body.data.project.name,
    label: body.data.issueLabel.name,
  })
);
