# @elisra-devops/docgen-data-provider

[![npm](https://img.shields.io/npm/v/@elisra-devops/docgen-data-provider)](https://www.npmjs.com/package/@elisra-devops/docgen-data-provider)
[![license: MIT](https://img.shields.io/badge/license-MIT-yellow.svg)](./LICENSE)

Azure DevOps data provider used by DocGen to fetch work items, Git, Pipelines and Test data via the Azure DevOps REST APIs.

## Installation

```bash
npm i @elisra-devops/docgen-data-provider
```

## Authentication

Pass a token string to the constructor:

- **Azure DevOps PAT** (most common): pass the PAT as-is.
- **Bearer token** (e.g. AAD/OIDC): pass as `bearer:<token>` or `bearer <token>` to send `Authorization: Bearer …`.

`orgUrl` should be your organization base URL, typically `https://dev.azure.com/<org>/` (note the trailing slash).

## Usage

```ts
import DgDataProviderAzureDevOps from '@elisra-devops/docgen-data-provider';

const orgUrl = 'https://dev.azure.com/<org>/';
const token = process.env.AZDO_TOKEN!; // PAT, or: `bearer:<access-token>`

const provider = new DgDataProviderAzureDevOps(orgUrl, token, undefined, process.env.JFROG_TOKEN);

const mgmt = await provider.getMangementDataProvider();
const projects = await mgmt.GetProjects();

const tickets = await provider.getTicketsDataProvider();
const workItem = await tickets.GetWorkItem('<project>', '123');
```

## Modules

The default export is `DgDataProviderAzureDevOps`, which creates module-specific providers:

- `getMangementDataProvider()` – org/project helpers (projects, profile, connection data).
- `getTicketsDataProvider()` – work items + WIQL queries + attachments/images + shared-query helpers.
- `getGitDataProvider()` – repos/branches/tags/files/commits/PRs + linked work items in ranges.
- `getPipelinesDataProvider()` – pipeline runs, artifacts, releases, “previous run” lookup, trigger builds.
- `getTestDataProvider()` – test plans/suites/cases/points/runs + parses test steps (including shared steps).
- `getResultDataProvider()` – test result summaries (group/summary/detailed) and “test reporter” output.
- `getJfrogDataProvider()` – JFrog build URL lookup (requires `jfrogToken` in the constructor).

## Notable APIs (by module)

- `MangementDataProvider`: `GetProjects()`, `GetProjectByName()`, `CheckOrgUrlValidity()`
- `TicketsDataProvider`: `GetWorkItem()`, `GetQueryResultsFromWiql()`, `GetSharedQueries()`, `CreateNewWorkItem()`, `UpdateWorkItem()`
- `GitDataProvider`: `GetTeamProjectGitReposList()`, `GetFileFromGitRepo()`, `GetCommitsInCommitRange()`, `CreatePullRequestComment()`
- `PipelinesDataProvider`: `GetPipelineRunHistory()`, `getPipelineRunDetails()`, `GetArtifactByBuildId()`, `TriggerBuildById()`
- `TestDataProvider`: `GetTestPlans()`, `GetTestSuitesByPlan()`, `GetTestCasesBySuites()`, `CreateTestRun()`, `UploadTestAttachment()`
- `ResultDataProvider`: `getCombinedResultsSummary()`, `getTestReporterResults()`

## Notes

- This library uses `axios` and retries some transient failures (timeouts/429/5xx).
- `TicketsDataProvider.GetSharedQueries()` supports doc-type specific query layouts (e.g. `std`, `str`, `svd`, `srs`, `test-reporter`) and falls back to the provided root path when a dedicated folder is missing.
- In Node.js, the HTTP client is configured with `rejectUnauthorized: false` in `src/helpers/tfs.ts`, which may be required for some internal setups but is a security tradeoff.
