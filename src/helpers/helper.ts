import { Query, Workitem } from '../models/tfs-data';

export class suiteData {
  name: string;
  id: string;
  parent: string;
  level: number;
  url: string;
  description: string;
  constructor(name: string, id: string, parent: string, level: number, description: string = '') {
    this.name = name;
    this.id = id;
    this.parent = parent;
    this.level = level;
    this.description = description;
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
   * Finds test suites recursively starting from a given suite ID.
   *
   * Sibling order is taken from `suits`' array order (callers are expected to
   * hand this a tree-ordered array — see TestDataProvider.normalizeAndEnrichSuitesResponse
   * / treeOrder), not recomputed here. Uses local state only, so concurrent calls
   * cannot interleave.
   *
   * @param planId - The test plan ID
   * @param url - Base organization URL
   * @param project - Project name
   * @param suits - Array of all test suites, in display order
   * @param foundId - Starting suite ID to search from
   * @param recursive - Whether to search recursively or just direct children
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
    const childrenByParent = new Map<string, any[]>();
    for (const suite of suits) {
      if (suite.parentSuiteId != 0) {
        const key = String(suite.parentSuiteId);
        const list = childrenByParent.get(key);
        if (list) {
          list.push(suite);
        } else {
          childrenByParent.set(key, [suite]);
        }
      }
    }

    const selfSuite = suits.find((suite) => suite.id == foundId);
    const result: suiteData[] = [];
    if (!selfSuite) {
      return result;
    }

    const buildUrl = (suiteId: any) =>
      url + project + '/_testManagement?planId=' + planId + '&suiteId=' + suiteId + '&_a=tests';

    const visitChildren = (parentId: any, level: number) => {
      const children = childrenByParent.get(String(parentId)) || [];
      for (const child of children) {
        const suit = new suiteData(child.title, child.id, parentId, level, child.description || '');
        suit.url = buildUrl(child.id);
        result.push(suit);
        visitChildren(child.id, level + 1);
      }
    };

    if (selfSuite.parentSuiteId == 0) {
      // Root suite match: the root itself is never emitted, only its descendants.
      if (!recursive) {
        return result;
      }
      visitChildren(selfSuite.id, 1);
    } else {
      // Nested suite match: emitted first, at level 1; children start at level 2.
      const suit = new suiteData(
        selfSuite.title,
        selfSuite.id,
        selfSuite.parentSuiteId,
        1,
        selfSuite.description || ''
      );
      suit.url = buildUrl(selfSuite.id);
      result.push(suit);
      if (!recursive) {
        return result;
      }
      visitChildren(selfSuite.id, 2);
    }

    return result;
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
