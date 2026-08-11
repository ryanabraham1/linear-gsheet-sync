import fs from "node:fs/promises";

const source = await fs.readFile("Code.local.gs", "utf8");
const apiKey = source.match(/LINEAR_API_KEY:\s*'([^']+)'/)[1];
const teamId = source.match(/LINEAR_TEAM_ID:\s*'([^']+)'/)[1];
const labelId = source.match(/LINEAR_LABEL_IDS:\s*\['([^']+)'\]/)[1];
const botProjectsSource = source.match(/BOT_PROJECTS:\s*\{([\s\S]*?)\n\s*\},\n\s*LINEAR_LABEL_IDS:/)[1];
const botProjects = Array.from(
  botProjectsSource.matchAll(
    /botName:\s*'([^']+)'[\s\S]*?projectId:\s*'([^']+)'[\s\S]*?projectName:\s*'([^']+)'[\s\S]*?milestoneId:\s*'([^']+)'[\s\S]*?milestoneName:\s*'([^']+)'/g
  ),
  (match) => ({
    botName: match[1],
    projectId: match[2],
    projectName: match[3],
    milestoneId: match[4],
    milestoneName: match[5],
  })
);
if (botProjects.length !== 3) throw new Error("Expected three configured bot projects.");

const query = `
  query SyncContext($teamId: String!, $labelId: String!) {
    viewer { name email }
    team(id: $teamId) { id name key }
    issueLabel(id: $labelId) { id name }
  }
`;

const response = await fetch("https://api.linear.app/graphql", {
  method: "POST",
  headers: {
    Authorization: apiKey,
    "Content-Type": "application/json",
  },
  body: JSON.stringify({ query, variables: { teamId, labelId } }),
});
const body = await response.json();
if (!response.ok || body.errors) {
  throw new Error(JSON.stringify(body.errors || { status: response.status }));
}
const checkedProjects = [];
for (const configured of botProjects) {
  const projectQuery = `
    query BotProject($projectId: String!, $milestoneId: String!) {
      project(id: $projectId) { id name }
      projectMilestone(id: $milestoneId) { id name project { id } }
    }
  `;
  const projectResponse = await fetch("https://api.linear.app/graphql", {
    method: "POST",
    headers: { Authorization: apiKey, "Content-Type": "application/json" },
    body: JSON.stringify({
      query: projectQuery,
      variables: {
        projectId: configured.projectId,
        milestoneId: configured.milestoneId,
      },
    }),
  });
  const projectBody = await projectResponse.json();
  if (!projectResponse.ok || projectBody.errors) {
    throw new Error(JSON.stringify(projectBody.errors || { status: projectResponse.status }));
  }
  const { project, projectMilestone } = projectBody.data;
  if (
    !project ||
    !projectMilestone ||
    project.name !== configured.projectName ||
    projectMilestone.name !== configured.milestoneName ||
    projectMilestone.project?.id !== project.id
  ) {
    throw new Error(`Configured Linear mapping did not match for ${configured.botName}.`);
  }
  checkedProjects.push(`${configured.botName} -> ${project.name} / ${projectMilestone.name}`);
}

console.log(JSON.stringify({
  viewer: body.data.viewer.name,
  team: body.data.team.key,
  label: body.data.issueLabel.name,
  botProjects: checkedProjects,
}));
