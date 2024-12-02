import { TFSServices } from '../helpers/tfs';
import { TestSteps, Workitem } from '../models/tfs-data';
import * as xml2js from 'xml2js';
import logger from './logger';

export default class TestStepParserHelper {
  private orgUrl: string;
  private token: string;

  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
  }

  private extractStepData(stepNode: any, stepPosition: string, parentStepId: string = ''): TestSteps {
    const step = new TestSteps();
    step.stepId = parentStepId === '' ? stepNode.$.id : `${parentStepId};${stepNode.$.id}`;
    step.stepPosition = stepPosition;
    step.action = stepNode.parameterizedString?.[0]?._ || '';
    step.expected = stepNode.parameterizedString?.[1]?._ || '';
    step.isSharedStepTitle = false;
    return step;
  }

  private async fetchSharedSteps(
    sharedStepId: string,
    sharedStepIdToRevisionLookupMap: Map<number, number>,
    parentStepPosition: string,
    parentStepId: string = ''
  ): Promise<TestSteps[]> {
    try {
      let sharedStepsList: TestSteps[] = [];
      const revision = sharedStepIdToRevisionLookupMap.get(Number(sharedStepId));

      const wiUrl = revision
        ? `${this.orgUrl}/_apis/wit/workitems/${sharedStepId}/revisions/${revision}`
        : `${this.orgUrl}/_apis/wit/workitems/${sharedStepId}`;
      const sharedStepWI = await TFSServices.getItemContent(wiUrl, this.token);
      const stepsXML = sharedStepWI?.fields['Microsoft.VSTS.TCM.Steps'] || null;
      const sharedStepTitle = sharedStepWI?.fields['System.Title'] || null;
      if (stepsXML && sharedStepTitle) {
        const stepsList = await this.parseTestSteps(
          stepsXML,
          sharedStepIdToRevisionLookupMap,
          `${parentStepPosition}.`,
          parentStepId
        );
        const titleObj = {
          stepId: parentStepId,
          stepPosition: parentStepPosition,
          action: `<b>${sharedStepTitle}<b/>`,
          expected: '',
          isSharedStepTitle: true,
        };
        sharedStepsList = [titleObj, ...stepsList];
      }
      return sharedStepsList;
    } catch (err: any) {
      const errorMsg = `failed to fetch shared step WI: ${err.message}`;
      logger.error(errorMsg);
      throw new Error(errorMsg);
    }
  }

  private async processCompref(
    comprefNode: any,
    sharedStepIdToRevisionLookupMap: Map<number, number>,
    stepPosition: string,
    parentStepId: string = ''
  ): Promise<TestSteps[]> {
    const stepList: TestSteps[] = [];
    if (comprefNode.$ && comprefNode.$.ref) {
      //Nested Steps
      const sharedStepId = comprefNode.$.ref;
      const comprefStepId = comprefNode.$.id;
      //Fetch the shared step data using the ID
      const sharedSteps = await this.fetchSharedSteps(
        sharedStepId,
        sharedStepIdToRevisionLookupMap,
        stepPosition,
        comprefStepId
      );
      stepList.push(...sharedSteps);
    }
    // If 'compref' contains nested steps
    if (comprefNode.children) {
      for (const child of comprefNode.children) {
        const nodeName = child['#name'];
        const currentPosition = comprefNode.children.indexOf(child) + 1;
        let currentPositionStr = '';
        if (stepPosition.includes('.')) {
          const positions = stepPosition.split('.');
          const lastPosition = Number(positions.pop());
          positions.push(`${currentPosition + lastPosition}`);
          currentPositionStr = positions.join('.');
        } else {
          currentPositionStr = `${currentPosition + Number(stepPosition)}`;
        }
        if (nodeName === 'step') {
          stepList.push(this.extractStepData(child, currentPositionStr, parentStepId));
        } else if (nodeName === 'compref') {
          // Handle nested 'compref' elements recursively
          const nestedSteps = await this.processCompref(
            child,
            sharedStepIdToRevisionLookupMap,
            currentPositionStr,
            parentStepId
          );
          stepList.push(...nestedSteps);
        }
      }
    }
    return stepList;
  }

  private async processSteps(
    parsedResultStepsXml: any,
    sharedStepIdToRevisionLookupMap: Map<number, number>,
    parentStepPosition: string,
    parentStepId: string = ''
  ): Promise<TestSteps[]> {
    const stepsList: TestSteps[] = [];
    const root = parsedResultStepsXml;

    const children: any[] = root.children || [];
    for (const child of children) {
      const nodeName = child['#name'];
      const currentStepPosition = children.indexOf(child) + 1;
      if (nodeName === 'step') {
        //Process a regular step
        stepsList.push(
          this.extractStepData(child, `${parentStepPosition}${currentStepPosition}`, parentStepId)
        );
      } else if (nodeName === 'compref') {
        // Process shared steps
        const sharedSteps = await this.processCompref(
          child,
          sharedStepIdToRevisionLookupMap,
          `${parentStepPosition}${currentStepPosition}`,
          parentStepId
        );
        stepsList.push(...sharedSteps);
      }
    }
    return stepsList;
  }

  public async parseTestSteps(
    xmlSteps: string,
    sharedStepIdToRevisionLookupMap: Map<number, number>,
    level: string = '',
    parentStepId: string = ''
  ): Promise<TestSteps[]> {
    try {
      const result = await xml2js.parseStringPromise(xmlSteps, {
        explicitChildren: true,
        preserveChildrenOrder: true,
        explicitArray: false, // Prevents unnecessary arrays
        childkey: 'children', // Key to store child elements
      });
      const stepsList = await this.processSteps(
        result.steps,
        sharedStepIdToRevisionLookupMap,
        level,
        parentStepId
      );
      return stepsList;
    } catch (err) {
      logger.error('Failed to parse XML test steps.', err);
      return [];
    }
  }
}
