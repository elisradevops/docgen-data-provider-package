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
    recursive: boolean,
    isTopLevel: boolean = true
  ): Array<suiteData> {

    
    // Only reset static state on the top-level call
    if (isTopLevel) {
      this.suitList = new Array<suiteData>();
      this.level = 0; // Start at 0 so top-level suites get level 1
      this.first = true;
    }

    for (let i = 0; i < suits.length; i++) {

      if (suits[i].parentSuiteId != 0) {
        // Child suites (parentSuiteId != 0)
        if (suits[i].parentSuiteId == foundId) {
          // Find the parent suite's level in our current results
          let parentLevel = this.level + 1; // Default fallback
          for (let j = 0; j < this.suitList.length; j++) {
            if (this.suitList[j].id == foundId) {
              parentLevel = this.suitList[j].level + 1;
              break;
            }
          }
          


          // Found children of the selected suite - add them to results
          let suit: suiteData = new suiteData(suits[i].title, suits[i].id, foundId, parentLevel);
          suit.url =
            url + project + '/_testManagement?planId=' + planId + '&suiteId=' + suits[i].id + '&_a=tests';
          this.suitList.push(suit);
          if (recursive == false) {
            return this.suitList;
          }
          this.level++; // Increment level before recursive call
          this.findSuitesRecursive(planId, url, project, suits, suits[i].id, true, false);
          this.level--; // Decrement level after recursive call
        } else if (suits[i].id == foundId && this.first) {


          // Found the selected nested suite itself - add it to results
          let suit: suiteData = new suiteData(
            suits[i].title,
            suits[i].id,
            suits[i].parentSuiteId,
            this.level + 1
          );
          suit.url =
            url + project + '/_testManagement?planId=' + planId + '&suiteId=' + suits[i].id + '&_a=tests';
          this.suitList.push(suit);
          this.first = false;
          if (recursive == false) {
            return this.suitList;
          }
        }
      } else {
        // Root suites (parentSuiteId = 0) - these do NOT get added to results
        if (suits[i].id == foundId && Helper.first) {

          let suit: suiteData = new suiteData(suits[i].title, suits[i].id, foundId, this.level);
          suit.url = url + project + '/_workitems/edit/' + suits[i].id;
          Helper.first = false;
          if (recursive == false) {
            return this.suitList;
          }
        }
      }
    }

    return this.suitList;
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
