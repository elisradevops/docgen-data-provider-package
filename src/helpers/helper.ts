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
  static level: number = 1;
  static first: boolean = true;
  public static suitList: Array<suiteData> = new Array<suiteData>();

  static buildSuiteslevel(dataSuites: any): any {}

  /**
   * Find suites recursively - O(n) complexity optimization with parent-child lookup map
   * @param planId - The plan identifier
   * @param url - Base URL
   * @param project - Project name
   * @param suits - Array of suite objects
   * @param foundId - ID to search for
   * @param recursive - Whether to search recursively
   * @returns Array of suite data
   */
  public static findSuitesRecursive(
    planId: string,
    url: string,
    project: string,
    suits: any,
    foundId: string,
    recursive: boolean
  ): Array<suiteData> {
    // Early return if no suits provided
    if (!suits || suits.length === 0) {
      return this.suitList;
    }

    // Build parent-child lookup map once - O(n) complexity
    const parentChildMap = new Map<string, any[]>();
    const suiteById = new Map<string, any>();

    for (const suit of suits) {
      // Skip if suit is null/undefined or missing required properties
      if (!suit || suit.id == null || suit.parentSuiteId == null) {
        continue;
      }

      // Store suite by ID for quick lookup
      suiteById.set(suit.id.toString(), suit);

      // Group by parent ID
      const parentId = suit.parentSuiteId.toString();
      if (!parentChildMap.has(parentId)) {
        parentChildMap.set(parentId, []);
      }
      parentChildMap.get(parentId)!.push(suit);
    }

    // Check for single child optimization at the top level
    const directChildren = parentChildMap.get(foundId.toString()) || [];
    const shouldSkipSingleChild = directChildren.length === 1;

    if (shouldSkipSingleChild) {
      // Skip the single child and promote its children to level 1
      const singleChild = directChildren[0];
      const grandChildren = parentChildMap.get(singleChild.id.toString()) || [];

      // Add each grandchild as a level 1 suite (promoted from level 2)
      for (const grandChild of grandChildren) {
        const suite: suiteData = new suiteData(grandChild.title || '', grandChild.id, foundId, this.level++);
        suite.url = `${url}${project}/_testManagement?planId=${planId}&suiteId=${grandChild.id}&_a=tests`;
        this.suitList.push(suite);

        if (!recursive) {
          this.level--; // Restore level before returning
          return this.suitList;
        }

        // Now recursively process this grandchild's children
        this.findSuitesRecursiveOptimized(
          planId,
          url,
          project,
          grandChild.id,
          recursive,
          parentChildMap,
          suiteById
        );
        this.level--;
      }
    } else {
      // Normal processing - start recursion from the found ID
      this.findSuitesRecursiveOptimized(planId, url, project, foundId, recursive, parentChildMap, suiteById);
    }

    return this.suitList;
  }

  /**
   * Optimized recursive helper using lookup maps - O(d) where d is depth
   */
  private static findSuitesRecursiveOptimized(
    planId: string,
    url: string,
    project: string,
    foundId: string,
    recursive: boolean,
    parentChildMap: Map<string, any[]>,
    suiteById: Map<string, any>
  ): void {
    const targetId = foundId.toString();

    // Handle root suite (parentSuiteId === 0) - check if foundId is a root suite
    const rootSuites = parentChildMap.get('0') || [];
    const rootSuite = rootSuites.find((s) => s.id.toString() === targetId);

    if (rootSuite && Helper.first) {
      const suite: suiteData = new suiteData(rootSuite.title || '', rootSuite.id, foundId, this.level);
      suite.url = `${url}${project}/_workitems/edit/${rootSuite.id}`;
      Helper.first = false;

      if (!recursive) {
        return;
      }
    }

    // Process child suites using the lookup map - O(children count)
    const childSuites = parentChildMap.get(targetId) || [];

    // Normal processing - add all children
    for (const childSuite of childSuites) {
      const suite: suiteData = new suiteData(childSuite.title || '', childSuite.id, foundId, this.level++);
      suite.url = `${url}${project}/_testManagement?planId=${planId}&suiteId=${childSuite.id}&_a=tests`;
      this.suitList.push(suite);

      if (!recursive) {
        this.level--; // Restore level before returning
        return;
      }

      // Recursively process children
      this.findSuitesRecursiveOptimized(planId, url, project, childSuite.id, true, parentChildMap, suiteById);

      this.level--;
    }
  }

  /**
   * Optimized level builder without static state
   * @param results - Query results containing work items
   * @param foundId - ID to start building from
   * @returns Array of work items with levels assigned
   */
  public static LevelBuilder(results: Query, foundId: string): Array<Workitem> {
    const levelList: Array<Workitem> = [];
    const processedIds = new Set<string>();
    const workItemMap = new Map<string, Workitem>();

    // Create lookup map for better performance
    for (const workItem of results.workItems) {
      workItemMap.set(workItem.fields[0]?.value || workItem.id?.toString() || '', workItem);
    }

    this.buildLevelsRecursive(results, foundId, 0, levelList, processedIds, workItemMap);
    return levelList;
  }

  /**
   * Internal recursive method for building levels
   */
  private static buildLevelsRecursive(
    results: Query,
    foundId: string,
    currentLevel: number,
    levelList: Array<Workitem>,
    processedIds: Set<string>,
    workItemMap: Map<string, Workitem>
  ): void {
    for (const workItem of results.workItems) {
      const workItemId = workItem.fields[0]?.value || workItem.id?.toString() || '';

      // Skip if already processed
      if (processedIds.has(workItemId)) {
        continue;
      }

      // Handle root items (Source === 0)
      if (workItem.Source === 0) {
        workItem.level = 0;
        levelList.push(workItem);
        processedIds.add(workItemId);
      }
      // Handle items with matching source
      else if (workItem.Source?.toString() === foundId) {
        workItem.level = currentLevel;
        levelList.push(workItem);
        processedIds.add(workItemId);

        // Recursively process children
        this.buildLevelsRecursive(
          results,
          workItemId,
          currentLevel + 1,
          levelList,
          processedIds,
          workItemMap
        );
      }
    }
  }
}
