import { Query, Workitem } from '../models/tfs-data';

export class suiteData {
  name: string;
  id: string;
  parent: string;
  level: number;
  url: string;
  constructor(name: string, id: string, parent: string, level: number) {
    this.name = name;
    this.id = id;
    this.parent = parent;
    this.level = level;
  }
}
export class Relations {
  id: string;
  rels: Array<string> = new Array<string>();
}

export class Links {
  id: string;
  title: string;
  description: string;
  url: string;
  type: string;
  customerId: string;
}
export class Trace {
  id: string;
  title: string;
  url: string;
  customerId: string;
  links: Array<Links>;
}

export class Helper {
  /**
   * Finds test suites recursively starting from a given suite ID
   * @param planId - The test plan ID
   * @param url - Base organization URL
   * @param project - Project name
   * @param suits - Array of all test suites
   * @param foundId - Starting suite ID to search from
   * @param recursive - Whether to search recursively or just direct children
   * @param flatSuiteTestCases - If true and there's only one level 1 suite with children, flatten the hierarchy by one level
   * @returns Array of suiteData objects representing the hierarchy
   */
  public static findSuitesRecursive(
    planId: string,
    url: string,
    project: string,
    suits: any[],
    foundId: string,
    recursive: boolean
  ): Array<suiteData> {
    // Create a map for faster lookups
    const suiteMap = new Map<string, any>();
    suits.forEach((suite) => suiteMap.set(suite.id.toString(), suite));

    const result: Array<suiteData> = [];
    const visited = new Set<string>();

    // Find the starting suite
    const startingSuite = suiteMap.get(foundId);
    if (!startingSuite) {
      return result;
    }

    // Skip root suites (parentSuiteId = 0) - we don't want them in results
    // Just mark them as visited so we can process their children
    if (startingSuite.parentSuiteId === 0) {
      visited.add(foundId);
    }

    // Build the hierarchy
    const startingLevel = startingSuite.parentSuiteId === 0 ? 1 : 0;

    this.buildSuiteHierarchy(
      suiteMap,
      foundId,
      planId,
      url,
      project,
      result,
      visited,
      startingLevel,
      recursive
    );

    return result;
  }

  /**
   * Recursively builds the suite hierarchy
   */
  private static buildSuiteHierarchy(
    suiteMap: Map<string, any>,
    parentId: string,
    planId: string,
    url: string,
    project: string,
    result: Array<suiteData>,
    visited: Set<string>,
    currentLevel: number,
    recursive: boolean
  ): void {
    // Find all direct children of the current parent
    const children = Array.from(suiteMap.values()).filter(
      (suite) => suite.parentSuiteId.toString() === parentId && !visited.has(suite.id.toString())
    );

    for (const child of children) {
      const childId = child.id.toString();

      if (visited.has(childId)) {
        continue; // Skip if already processed
      }

      // Create suite data for this child
      const childSuiteData = this.createSuiteData(child, planId, url, project, currentLevel);
      result.push(childSuiteData);
      visited.add(childId);

      // If recursive is false, only get direct children
      if (!recursive && currentLevel > 0) {
        continue;
      }

      // Recursively process children of this suite
      if (recursive) {
        this.buildSuiteHierarchy(
          suiteMap,
          childId,
          planId,
          url,
          project,
          result,
          visited,
          currentLevel + 1,
          recursive
        );
      }
    }
  }

  /**
   * Creates a suiteData object from a suite
   */
  private static createSuiteData(
    suiteInfo: any,
    planId: string,
    url: string,
    project: string,
    level: number
  ): suiteData {
    const suite = new suiteData(
      suiteInfo.title,
      suiteInfo.id.toString(),
      suiteInfo.parentSuiteId.toString(),
      level
    );

    // Generate appropriate URL based on suite type
    if (suiteInfo.parentSuiteId === 0) {
      // Root suite - link to work item
      suite.url = `${url}/${project}/_workitems/edit/${suiteInfo.id}`;
    } else {
      // Child suite - link to test management
      suite.url = `${url}/${project}/_testManagement?planId=${planId}&suiteId=${suiteInfo.id}&_a=tests`;
    }

    return suite;
  }

  public static levelList: Array<Workitem> = new Array<Workitem>();

  public static LevelBuilder(results: Query, foundId: string): Array<Workitem> {
    // Reset the level list for each call
    this.levelList = [];

    // Build the hierarchy starting from level 0
    this.buildWorkItemHierarchy(results, foundId, 0);

    return this.levelList;
  }

  private static buildWorkItemHierarchy(results: Query, foundId: string, currentLevel: number): void {
    for (let i = 0; i < results.workItems.length; i++) {
      const workItem = results.workItems[i];

      if (workItem.Source == 0) {
        workItem.level = 0;
        if (!this.levelList.includes(workItem)) {
          this.levelList.push(workItem);
        }
      } else if (workItem.Source.toString() == foundId) {
        workItem.level = currentLevel;
        this.levelList.push(workItem);

        // Recursively build hierarchy for children
        this.buildWorkItemHierarchy(results, workItem.fields[0].value, currentLevel + 1);
      }
    }
  }
}
