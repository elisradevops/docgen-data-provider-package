import DgDataProviderAzureDevOps from "../..";

require("dotenv").config();
jest.setTimeout(60000);
const orgUrl = process.env.ORG_URL;
const token = process.env.PAT;

const dgDataProviderAzureDevOps = new DgDataProviderAzureDevOps(orgUrl,token);


describe("git module - tests", () => {
  test("should return repo list for teamProject", async () => {
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let json: any = await gitDataProvider.GetTeamProjectGitReposList(
      "tests"
    );
    expect(json).toBeDefined();
  });
  test("should return repo from repoid", async () => {
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let json: any = await gitDataProvider.GetGitRepoFromRepoId(
      "68f2aee7-0864-458e-93ce-320303a080ed"
    );
    expect(json).toBeDefined();
  });
  test("should return repo by pullrequest id", async () => {
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let json: any = await gitDataProvider.GetGitRepoFromPrId(
      73
    );
    expect(json.repository).toBeDefined();
  });
  test("should return commits by pullrequest & repo id", async () => {
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let json: any = await gitDataProvider.GetGitRepoFromPrId(73);
    let commitJson: any = await gitDataProvider.GetPullRequestCommits(
      json.repository.id,
      73
    );
    expect(commitJson.value.length).toBeGreaterThan(0);
  });
  test("should return pullrequest threads", async () => {
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let json = await gitDataProvider.GetPullRequestComments(
      "tests",
      "68f2aee7-0864-458e-93ce-320303a080ed",
      73
    );
    expect(json.count).toBeDefined;
  });
  test("should create pullrequest thread", async () => {
    let data = {
      comments: [
        {
          parentCommentId: 0,
          content: "Should we add a comment about what this value means?",
          commentType: 1,
        },
      ],
      status: 1,
      threadContext: {
        filePath: "/assaf.txt",
        leftFileEnd: null,
        leftFileStart: null,
        rightFileEnd: {
          line: 2,
          offset: 1,
        },
        rightFileStart: {
          line: 2,
          offset: 5,
        },
      },
      pullRequestThreadContext: {
        changeTrackingId: 1,
        iterationContext: {
          firstComparingIteration: 1,
          secondComparingIteration: 1,
        },
      },
    };
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let json = await gitDataProvider.CreatePullRequestComment(
      "tests",
      "68f2aee7-0864-458e-93ce-320303a080ed",
      73,
      data
    );
    expect(json.comment).toBeDefined;
  });
  test("should return commits with linked items in date range", async () => {
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let commitRange = await gitDataProvider.GetCommitsInDateRange(
      "Azuretraining",
      "b77bccb5-38de-4dca-9dcb-b0f8046bc045",
      "2016-10-21T12:51:51Z",
      "2024-10-24T12:51:51Z",
      "main"
    );
    let items = await gitDataProvider.GetItemsInCommitRange(
      "Azuretraining",
      "b77bccb5-38de-4dca-9dcb-b0f8046bc045",
      commitRange
    );
    expect(items[0].workItem).toBeDefined();
  });
  test("should return source trigger commit for pipline", async () => {
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let json = await gitDataProvider.GetCommitForPipeline(
      "tests",
      248
    );
    expect(json).toBe("59d59691ee002815e7aa774f0a90ef28a6e4708f");
  });
  test("should return commits in commit range", async () => {
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let json = await gitDataProvider.GetCommitsInCommitRange(
      "tests",
      "68f2aee7-0864-458e-93ce-320303a080ed",
      "4ce7f96f74f10bb60d27d7180a8d1bd44da1ffac",
      "e46f8023be49db94b5cf188b41f7ba9db6fd8274"
    );
    expect(json.count).toBeGreaterThan(0);
  });
  test("should return items linked in build range", async () => {
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let json = await gitDataProvider.GetItemsForPipelinesRange(
      "tests",
      244,
      244
    );
    expect(json.length).toBeGreaterThan(0);
  });
  test("should return commits range between dates ", async () => {
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let json = await gitDataProvider.GetCommitsInDateRange(
      "tests",
      "68f2aee7-0864-458e-93ce-320303a080ed",
      "2009-11-30T12:51:51Z",
      "2021-11-30T12:51:51Z"
    );
    expect(json.count).toBeGreaterThanOrEqual(14);
  });
  test("should return all commits for repo ", async ()=>{
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let json = await gitDataProvider.GetCommitsForRepo(
      "tests",  
      "68f2aee7-0864-458e-93ce-320303a080ed",
      "master"
    );
    expect(json.count).toBeGreaterThanOrEqual(0);
  })
  test("should return all pullrequests for repo ", async ()=>{
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let json = await gitDataProvider.GetPullRequestsForRepo(
      "tests",
      "68f2aee7-0864-458e-93ce-320303a080ed"
    );
    expect(json.count).toBeGreaterThanOrEqual(0);
  })

  test("should return all pullrequests for repo with Given PrId ", async ()=>{
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let json = await gitDataProvider.GetItemsInPullRequestRange(
      "tests",
      "68f2aee7-0864-458e-93ce-320303a080ed",
      [73,74]
      );
    expect(json.count).toBeGreaterThanOrEqual(0);
  })

  test("should return branches of repository ", async ()=>{
    let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();
    let json = await gitDataProvider.GetRepoBranches(
      "tests",
      "68f2aee7-0864-458e-93ce-320303a080ed",
      );
    expect(json.count).toBeGreaterThanOrEqual(0);
  })

}); //describe

test("should return pull requests in commit range without linked items", async () => {
  let gitDataProvider = await dgDataProviderAzureDevOps.getGitDataProvider();

  // Get the commit range using the date range
  let commitRange = await gitDataProvider.GetCommitsInDateRange(
    "Azuretraining",
    "b77bccb5-38de-4dca-9dcb-b0f8046bc045",
    "2023-7-21T12:51:51Z",
    "2024-8-21T12:51:51Z",
    "main"
  );

  // Use the new function to get pull requests without linked items
  let pullRequests = await gitDataProvider.GetPullRequestsInCommitRangeWithoutLinkedItems(
    "Azuretraining",
    "b77bccb5-38de-4dca-9dcb-b0f8046bc045",
    commitRange
  );

  // Print the pull requests
  console.log("Pull Requests:", pullRequests);

  expect(pullRequests).toBeDefined();
  expect(pullRequests.length).toBeGreaterThan(0);
});
