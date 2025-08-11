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
   * @param flatTreeByOneLevel - If true and there's only one level 1 suite with children, flatten the hierarchy by one level
   * @returns Array of suiteData objects representing the hierarchy
   */
  public static findSuitesRecursive(
    planId: string,
    url: string,
    project: string,
    suits: any[],
    foundId: string,
    recursive: boolean,
    flatTreeByOneLevel: boolean = false
  ): Array<suiteData> {
    console.log(`[findSuitesRecursive] Starting with foundId: ${foundId}, flatTreeByOneLevel: ${flatTreeByOneLevel}`);
    console.log(`[findSuitesRecursive] Total suits provided: ${suits.length}`);
    
    // Create a map for faster lookups
    const suiteMap = new Map<string, any>();
    suits.forEach((suite) => suiteMap.set(suite.id.toString(), suite));

    const result: Array<suiteData> = [];
    const visited = new Set<string>();

    // Find the starting suite
    const startingSuite = suiteMap.get(foundId);
    if (!startingSuite) {
      console.log(`[findSuitesRecursive] ERROR: Starting suite ${foundId} not found in suite map`);
      return result;
    }

    console.log(`[findSuitesRecursive] Starting suite found: ${startingSuite.title} (ID: ${startingSuite.id}, parentSuiteId: ${startingSuite.parentSuiteId})`);

    // Skip root suites (parentSuiteId = 0) - we don't want them in results
    // Just mark them as visited so we can process their children
    if (startingSuite.parentSuiteId === 0) {
      console.log(`[findSuitesRecursive] Starting suite is root (parentSuiteId = 0), marking as visited`);
      visited.add(foundId);
    }

    // Build the hierarchy
    const startingLevel = startingSuite.parentSuiteId === 0 ? 1 : 0;
    console.log(`[findSuitesRecursive] Building hierarchy starting at level: ${startingLevel}`);
    
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

    console.log(`[findSuitesRecursive] Hierarchy built, result count: ${result.length}`);
    console.log(`[findSuitesRecursive] Result suites:`, result.map(s => `${s.name} (ID: ${s.id}, level: ${s.level}, parent: ${s.parent})`));

    // Apply flattening if conditions are met
    if (flatTreeByOneLevel) {
      console.log(`[findSuitesRecursive] Checking flattening conditions...`);
      if (this.shouldFlattenHierarchy(result, suiteMap, foundId)) {
        console.log(`[findSuitesRecursive] Flattening conditions met, applying flattening...`);
        const flattened = this.flattenHierarchyByOneLevel(result, foundId);
        console.log(`[findSuitesRecursive] After flattening:`, flattened.map(s => `${s.name} (ID: ${s.id}, level: ${s.level}, parent: ${s.parent})`));
        return flattened;
      } else {
        console.log(`[findSuitesRecursive] Flattening conditions NOT met`);
      }
    }

    console.log(`[findSuitesRecursive] Returning ${result.length} suites (no flattening applied)`);
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

  /**
   * Determines if hierarchy should be flattened based on conditions:
   * - There is exactly one level 1 suite (direct child of root)
   * - That suite has children (level 2+ suites)
   */
  private static shouldFlattenHierarchy(
    result: Array<suiteData>,
    suiteMap: Map<string, any>,
    rootId: string
  ): boolean {
    console.log(`[shouldFlattenHierarchy] Checking flattening conditions for rootId: ${rootId}`);
    console.log(`[shouldFlattenHierarchy] Total result suites: ${result.length}`);
    
    // Find all level 1 suites (direct children of root) - they have level 1
    const level1Suites = result.filter((suite) => suite.level === 1);
    console.log(`[shouldFlattenHierarchy] Found ${level1Suites.length} level 1 suites:`, level1Suites.map(s => `${s.name} (ID: ${s.id})`));

    // Must have exactly one level 1 suite
    if (level1Suites.length !== 1) {
      console.log(`[shouldFlattenHierarchy] CONDITION FAILED: Not exactly one level 1 suite (found ${level1Suites.length})`);
      return false;
    }

    const singleLevel1Suite = level1Suites[0];
    console.log(`[shouldFlattenHierarchy] Single level 1 suite: ${singleLevel1Suite.name} (ID: ${singleLevel1Suite.id})`);

    // Check if this suite has children by looking for suites with this suite as parent
    const childrenInResult = result.filter((suite) => suite.parent === singleLevel1Suite.id);
    console.log(`[shouldFlattenHierarchy] Found ${childrenInResult.length} children of level 1 suite:`, childrenInResult.map(s => `${s.name} (ID: ${s.id}, level: ${s.level})`));

    // The suite must have children
    const hasChildren = childrenInResult.length > 0;
    console.log(`[shouldFlattenHierarchy] Suite has children: ${hasChildren}`);
    console.log(`[shouldFlattenHierarchy] FLATTENING CONDITIONS ${hasChildren ? 'MET' : 'NOT MET'}`);
    
    return hasChildren;
  }

  /**
   * Flattens the hierarchy by one level:
   * - Removes the single level 1 suite
   * - Promotes all level 2+ suites up by one level
   * - Updates parentSuiteId of new level 1 suites to point to root
   */
  private static flattenHierarchyByOneLevel(result: Array<suiteData>, rootId: string): Array<suiteData> {
    console.log(`[flattenHierarchyByOneLevel] Starting flattening process for rootId: ${rootId}`);
    console.log(`[flattenHierarchyByOneLevel] Input suites:`, result.map(s => `${s.name} (ID: ${s.id}, level: ${s.level}, parent: ${s.parent})`));
    
    // Find the single level 1 suite to remove (level = 1)
    const level1Suite = result.find((suite) => suite.level === 1);
    if (!level1Suite) {
      console.log(`[flattenHierarchyByOneLevel] ERROR: No level 1 suite found to remove`);
      return result;
    }

    const level1SuiteId = level1Suite.id;
    console.log(`[flattenHierarchyByOneLevel] Removing level 1 suite: ${level1Suite.name} (ID: ${level1SuiteId})`);

    // Filter out the level 1 suite and adjust levels/parents for remaining suites
    const flattenedResult = result
      .filter((suite) => suite.id !== level1SuiteId) // Remove the level 1 suite
      .map((suite) => {
        const newSuite = new suiteData(suite.name, suite.id, suite.parent, suite.level - 1);
        newSuite.url = suite.url;

        // If this was a level 2 suite (child of the removed level 1 suite),
        // update its parent to point to the root
        if (suite.parent === level1SuiteId) {
          console.log(`[flattenHierarchyByOneLevel] Updating parent of ${suite.name} from ${level1SuiteId} to ${rootId}`);
          newSuite.parent = rootId;
        }

        console.log(`[flattenHierarchyByOneLevel] Processed suite: ${newSuite.name} (ID: ${newSuite.id}, level: ${suite.level} -> ${newSuite.level}, parent: ${suite.parent} -> ${newSuite.parent})`);
        return newSuite;
      });

    console.log(`[flattenHierarchyByOneLevel] Flattening complete. Result:`, flattenedResult.map(s => `${s.name} (ID: ${s.id}, level: ${s.level}, parent: ${s.parent})`));
    return flattenedResult;
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
