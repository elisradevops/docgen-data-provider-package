import { Query, Workitem } from "../models/tfs-data";

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
  static buildSuiteslevel(dataSuites: any): any {}
  public static suitList: Array<suiteData> = new Array<suiteData>();
  public static findSuitesRecursive(
    planId: string,
    url: string,
    project: string,
    suits: any,
    foundId: string,
    recursive: boolean
  ): Array<suiteData> {
    for (let i = 0; i < suits.length; i++) {
      if (suits[i].parentSuiteId != 0) {
        if (suits[i].parentSuiteId == foundId) {
          let suit: suiteData = new suiteData(
            suits[i].title,
            suits[i].id,
            foundId,
            this.level++
          );
          suit.url =
            url +
            project +
            "/_testManagement?planId=" +
            planId +
            "&suiteId=" +
            suits[i].id +
            "&_a=tests";
          this.suitList.push(suit);
          if (recursive == false) {
            return this.suitList;
          }
          this.findSuitesRecursive(
            planId,
            url,
            project,
            suits,
            suits[i].id,
            true
            );
          this.level--;
        }
      } else {
          if (suits[i].id == foundId && Helper.first) {
            let suit: suiteData = new suiteData(
              suits[i].title,
              suits[i].id,
              foundId,
              this.level
            );
            suit.url = url + project + "/_workitems/edit/" + suits[i].id;
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
    for (let i = 0; i < results.workItems.length; i++) {
      if (results.workItems[i].Source == 0) {
        results.workItems[i].level = 0;
        if (!this.levelList.includes(results.workItems[i]))
          this.levelList.push(results.workItems[i]);
      } else if (results.workItems[i].Source.toString() == foundId) {
        results.workItems[i].level = this.level++;
        this.levelList.push(results.workItems[i]);

        this.LevelBuilder(results, results.workItems[i].fields[0].value);
        this.level--;
      }
    }

    return this.levelList;
  }
}
