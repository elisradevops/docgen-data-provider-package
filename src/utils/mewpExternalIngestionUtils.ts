import type {
  MewpBugLink,
  MewpExternalFileRef,
  MewpL3L4Link,
} from '../models/mewp-reporting';
import logger from './logger';
import MewpExternalTableUtils from './mewpExternalTableUtils';

export interface MewpExternalIngestionAdapters {
  toComparableText: (value: any) => string;
  toRequirementKey: (value: string) => string;
  resolveBugResponsibility: (fields: Record<string, any>) => string;
  isExternalStateInScope: (value: string, itemType: 'bug' | 'requirement') => boolean;
  isExcludedL3L4BySapWbs: (value: string) => boolean;
  resolveRequirementSapWbsByBaseKey?: (baseKey: string) => string;
}

export default class MewpExternalIngestionUtils {
  private externalTableUtils: MewpExternalTableUtils;

  constructor(externalTableUtils: MewpExternalTableUtils) {
    this.externalTableUtils = externalTableUtils;
  }

  public async loadExternalBugsByTestCase(
    externalBugsFile: MewpExternalFileRef | null | undefined,
    adapters: MewpExternalIngestionAdapters
  ): Promise<Map<number, MewpBugLink[]>> {
    const rows = await this.externalTableUtils.loadExternalTableRows(externalBugsFile, 'bugs');
    if (rows.length === 0) return new Map<number, MewpBugLink[]>();

    const map = new Map<number, MewpBugLink[]>();
    let parsedRows = 0;
    for (const row of rows) {
      const testCaseId = this.toPositiveNumber(
        this.externalTableUtils.readExternalCell(row, [
          'Elisra_SortIndex',
          'Elisra SortIndex',
          'ElisraSortIndex',
        ])
      );
      if (!testCaseId) continue;
      const requirementBaseKey = adapters.toRequirementKey(
        adapters.toComparableText(this.externalTableUtils.readExternalCell(row, ['SR']))
      );
      if (!requirementBaseKey) continue;

      const bugId =
        this.toPositiveNumber(
          this.externalTableUtils.readExternalCell(row, [
            'TargetWorkItemId',
            'Bug ID',
            'BugId',
            'Links.TargetWorkItem.WorkItemId',
            'Links TargetWorkItem WorkItemId',
          ])
        ) || 0;
      if (!bugId) continue;

      const bugState = adapters.toComparableText(
        this.externalTableUtils.readExternalCell(row, ['TargetState', 'State'])
      );
      if (!adapters.isExternalStateInScope(bugState, 'bug')) continue;

      const bugTitle = adapters.toComparableText(
        this.externalTableUtils.readExternalCell(row, ['Bug Title', 'Title', 'Links.TargetWorkItem.Title'])
      );
      const bugResponsibilityRaw = adapters.toComparableText(
        this.externalTableUtils.readExternalCell(row, [
          'Responsibility',
          'Division',
          'SAPWBS',
          'TargetSapWbs',
        ])
      );

      const bug: MewpBugLink = {
        id: bugId,
        title: bugTitle,
        responsibility: adapters.resolveBugResponsibility({
          'Custom.SAPWBS': bugResponsibilityRaw,
        }),
        requirementBaseKey,
      };

      if (!map.has(testCaseId)) map.set(testCaseId, []);
      map.get(testCaseId)!.push(bug);
      parsedRows += 1;
    }

    const deduped = new Map<number, MewpBugLink[]>();
    for (const [testCaseId, bugs] of map.entries()) {
      const byId = new Map<string, MewpBugLink>();
      for (const bug of bugs || []) {
        const bugId = Number(bug?.id || 0);
        if (!Number.isFinite(bugId) || bugId <= 0) continue;
        const baseKey = String(bug?.requirementBaseKey || '').trim();
        const compositeKey = `${bugId}|${baseKey}`;
        const existing = byId.get(compositeKey);
        if (!existing) {
          byId.set(compositeKey, bug);
          continue;
        }

        byId.set(compositeKey, {
          id: bugId,
          requirementBaseKey: baseKey,
          title: String(existing?.title || bug?.title || '').trim(),
          responsibility: String(existing?.responsibility || bug?.responsibility || '').trim(),
        });
      }
      deduped.set(
        testCaseId,
        [...byId.values()].sort((a, b) => {
          const idCompare = Number(a.id) - Number(b.id);
          if (idCompare !== 0) return idCompare;
          return String(a.requirementBaseKey || '').localeCompare(String(b.requirementBaseKey || ''));
        })
      );
    }

    if (parsedRows === 0) {
      logger.warn(
        `External bugs source was loaded but no valid rows were parsed. ` +
          `Expected columns include Elisra_SortIndex, SR and bug ID fields.`
      );
    }

    return deduped;
  }

  public async loadExternalL3L4ByBaseKey(
    externalL3L4File: MewpExternalFileRef | null | undefined,
    adapters: MewpExternalIngestionAdapters
  ): Promise<Map<string, MewpL3L4Link[]>> {
    const rows = await this.externalTableUtils.loadExternalTableRows(externalL3L4File, 'l3l4');
    if (rows.length === 0) return new Map<string, MewpL3L4Link[]>();

    const linksByBaseKey = new Map<string, Map<string, MewpL3L4Link>>();
    const addLink = (baseKey: string, level: 'L3' | 'L4', id: number, title: string) => {
      if (!baseKey || !id) return;
      if (!linksByBaseKey.has(baseKey)) {
        linksByBaseKey.set(baseKey, new Map<string, MewpL3L4Link>());
      }
      const idKey = String(id).trim();
      const dedupeKey = `${level}:${idKey}`;
      linksByBaseKey.get(baseKey)!.set(dedupeKey, {
        id: idKey,
        title: String(title || '').trim(),
        level,
      });
    };

    let parsedRows = 0;
    for (const row of rows) {
      const srRaw = adapters.toComparableText(this.externalTableUtils.readExternalCell(row, ['SR']));
      const baseKey = adapters.toRequirementKey(srRaw);
      if (!baseKey) continue;
      const requirementSapWbsFallback = adapters.resolveRequirementSapWbsByBaseKey?.(baseKey) || '';

      const area = adapters
        .toComparableText(this.externalTableUtils.readExternalCell(row, ['AREA 34', 'AREA34']))
        .toLowerCase();
      const targetIdLevel3 = this.toPositiveNumber(
        this.externalTableUtils.readExternalCell(row, [
          'TargetWorkItemId Level 3',
          'TargetWorkItemIdLevel 3',
          'TargetWorkItemIdLevel3',
        ])
      );
      const targetTitleLevel3 = adapters.toComparableText(
        this.externalTableUtils.readExternalCell(row, [
          'TargetTitleLevel3',
          'TargetTitleLevel 3',
          'TargetTitle Level 3',
        ])
      );
      const targetStateLevel3 = adapters.toComparableText(
        this.externalTableUtils.readExternalCell(row, ['TargetStateLevel 3', 'TargetStateLevel3'])
      );
      const targetSapWbsLevel3 = adapters.toComparableText(
        this.externalTableUtils.readExternalCell(row, [
          'TargetSapWbsLevel 3',
          'TargetSapWbsLevel3',
          'TargetSapWbs Level 3',
        ])
      );
      const targetIdLevel4 = this.toPositiveNumber(
        this.externalTableUtils.readExternalCell(row, [
          'TargetWorkItemIdLevel 4',
          'TargetWorkItemId Level 4',
          'TargetWorkItemIdLevel4',
        ])
      );
      const targetTitleLevel4 = adapters.toComparableText(
        this.externalTableUtils.readExternalCell(row, [
          'TargetTitleLevel4',
          'TargetTitleLevel 4',
          'TargetTitle Level 4',
        ])
      );
      const targetStateLevel4 = adapters.toComparableText(
        this.externalTableUtils.readExternalCell(row, ['TargetStateLevel 4', 'TargetStateLevel4'])
      );
      const targetSapWbsLevel4 = adapters.toComparableText(
        this.externalTableUtils.readExternalCell(row, [
          'TargetSapWbsLevel 4',
          'TargetSapWbsLevel4',
          'TargetSapWbs Level 4',
        ])
      );

      if (area.includes('level 4')) {
        const effectiveSapWbsLevel3 = targetSapWbsLevel3 || requirementSapWbsFallback;
        if (
          targetIdLevel3 &&
          adapters.isExternalStateInScope(targetStateLevel3, 'requirement') &&
          !adapters.isExcludedL3L4BySapWbs(effectiveSapWbsLevel3)
        ) {
          addLink(baseKey, 'L4', targetIdLevel3, targetTitleLevel3);
        }
        parsedRows += 1;
        continue;
      }

      const effectiveSapWbsLevel3 = targetSapWbsLevel3 || requirementSapWbsFallback;
      if (
        targetIdLevel3 &&
        adapters.isExternalStateInScope(targetStateLevel3, 'requirement') &&
        !adapters.isExcludedL3L4BySapWbs(effectiveSapWbsLevel3)
      ) {
        addLink(baseKey, 'L3', targetIdLevel3, targetTitleLevel3);
      }
      const effectiveSapWbsLevel4 = targetSapWbsLevel4 || requirementSapWbsFallback;
      if (
        targetIdLevel4 &&
        adapters.isExternalStateInScope(targetStateLevel4, 'requirement') &&
        !adapters.isExcludedL3L4BySapWbs(effectiveSapWbsLevel4)
      ) {
        addLink(baseKey, 'L4', targetIdLevel4, targetTitleLevel4);
      }
      parsedRows += 1;
    }

    const out = new Map<string, MewpL3L4Link[]>();
    for (const [baseKey, linksById] of linksByBaseKey.entries()) {
      out.set(
        baseKey,
        [...linksById.values()].sort((a, b) => {
          if (a.level !== b.level) return a.level === 'L3' ? -1 : 1;
          return a.id.localeCompare(b.id);
        })
      );
    }

    if (parsedRows === 0) {
      logger.warn(
        `External L3/L4 source was loaded but no valid rows were parsed. ` +
          `Expected columns include SR, AREA 34, target IDs/titles/states.`
      );
    }

    return out;
  }

  private toPositiveNumber(value: any): number {
    const parsed = Number(String(value || '').replace(/[^\d]/g, ''));
    return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
  }
}
