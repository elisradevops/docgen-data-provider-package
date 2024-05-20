import DgDataProviderAzureDevOps from "../..";

require("dotenv").config();
jest.setTimeout(60000);


const orgUrl = process.env.ORG_URL;
const token = process.env.PAT;
const dgDataProviderAzureDevOps = new DgDataProviderAzureDevOps(orgUrl,token);

const wiql =
  "SELECT [System.Id],[System.WorkItemType],[System.Title],[System.AssignedTo],[System.State],[System.Tags] FROM workitems WHERE [System.TeamProject]=@project";

describe("ticket module - tests", () => {
  test("should create a new work item", async () => { 
    let TicketDataProvider = await dgDataProviderAzureDevOps.getTicketsDataProvider();
    let body =
      [{ 
        op: "add",
        path: "/fields/System.IterationPath",
        value: "tests"
      },
      { 
        op: "add",
        path: "/fields/System.State",
        value: "New"
      },
      { 
        op: "add",
        path: "/fields/System.AreaPath",
        value: "tests"
      },
      { 
        op: "add",  
        path: "/fields/System.Title",
        value: "new-test-title2"
      },
      {
        op: "add",
        path: "/fields/System.Tags",
        value: "tag-test"
      },
      { 
        op: "add",
        path: "/fields/System.AssignedTo",
        value: "denispankove"
      }];
    let res = await TicketDataProvider.CreateNewWorkItem(
      "tests",
      body,
      "Epic",
      true
    );
    expect(typeof res.id).toBe("number");
    body[3].value = "edited";
    let updatedWI = await TicketDataProvider.UpdateWorkItem(
      "tests",
      body,
      res.id,
      true
    );
    expect(updatedWI.fields["System.Title"]).toBe("edited");
  });
  test("should return shared queires", async () => {
    let TicketDataProvider = await dgDataProviderAzureDevOps.getTicketsDataProvider();
    let json = await TicketDataProvider.GetSharedQueries(
      "tests",
      "",
    );
    expect(json.length).toBeGreaterThan(1);
  });
  test("should return query results", async () => {
    let TicketDataProvider = await dgDataProviderAzureDevOps.getTicketsDataProvider();
    let json = await TicketDataProvider.GetSharedQueries(
      "tests",
      ""
    );
    let query = json.find((o: { wiql: undefined }) => o.wiql != undefined);
    let result = await TicketDataProvider.GetQueryResultsByWiqlHref(
      query.wiql.href,
      "tests"
    );
    expect(result.length).toBeGreaterThanOrEqual(1);
  });
  test("should return query results by query id", async () => {
    let TicketDataProvider = await dgDataProviderAzureDevOps.getTicketsDataProvider();
    let result = await TicketDataProvider.GetQueryResultById(
      "08e044be-b9bc-4962-99c9-ffebb47ff95a",
      "tests"
    );
    expect(result.length).toBeGreaterThanOrEqual(1);
  });
  test("should return wi base on wiql string", async () => {
    let TicketDataProvider = await dgDataProviderAzureDevOps.getTicketsDataProvider();
    let result = await TicketDataProvider.GetQueryResultsByWiqlString(
      wiql,
      "tests"
    );
    expect(result.workItems.length).toBeGreaterThanOrEqual(1);
  });
  test("should return populated work items array", async () => {
    let TicketDataProvider = await dgDataProviderAzureDevOps.getTicketsDataProvider();
    let result = await TicketDataProvider.GetQueryResultsByWiqlString(
      wiql,
      "tests"
    );
    let wiarray = result.workItems.map((o: any) => o.id);
    let res = await TicketDataProvider.PopulateWorkItemsByIds(
      wiarray,
      "tests"
    );
    expect(res.length).toBeGreaterThanOrEqual(1);
  });
  test("should return list of attachments", async () => {
    let TicketDataProvider = await dgDataProviderAzureDevOps.getTicketsDataProvider();
    let attachList = await TicketDataProvider.GetWorkitemAttachments(
      "tests",
      "538"
    );
    expect(attachList.length > 0).toBeDefined();
  });
  test("should return Json data of the attachment", async () => {
    let TicketDataProvider = await dgDataProviderAzureDevOps.getTicketsDataProvider();
    let attachedData = await TicketDataProvider.GetWorkitemAttachmentsJSONData(
      "tests",
      "14933c55-6d84-499c-88db-55202f16dd46"
    );
    expect(JSON.stringify(attachedData).length > 0).toBeDefined();
  });
  test("should return list of id & link object", async () => {
    let TicketDataProvider = await dgDataProviderAzureDevOps.getTicketsDataProvider();
    let attachList = await TicketDataProvider.GetLinksByIds(
      "tests",
      [535]
    );
    expect(attachList.length).toBeGreaterThan(0);
  });
}); //describe
