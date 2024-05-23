# dg-data-provider-azuredevops

[![npm version](https://badge.fury.io/js/@doc-gen%2Fdg-data-provider-azuredevops.svg)](https://badge.fury.io/js/@doc-gen%2Fdg-data-provider-azuredevops)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

[![Chat on Discord](https://badgen.net/badge/icon/discord?icon=discord&label)](https://discord.com/channels/904432165901709333/904432165901709336)

An azuredevops data provider using the document generator interface.
supported modules:

- **managment** - a module for retriving general purpose data, for example: number of projects and projects detailes.
- **git** - a module for retriving git related data, for example: commit detailes, branches etc.
- **pipelines** - a module for retriving pipelines data, for example: pipline history, pipline details etc.
- **tests** - a module for retriving tests data, for example:testplans list, test steps etc.
- **tickets** - a module for retriving tickets data, for example: open tickets, tickets list by query results etc.

### Installation:

```
npm i @doc-gen/dg-data-provider-azuredevops
```

### Usage:

```
import DgDataProviderAzureDevOps from "@doc-gen/dg-data-provider-azuredevops";

let dgDataProviderAzureDevOps = new DgDataProviderAzureDevOps(
      orgUrl,
      token
    );
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();



```

### Contributors:
