import type {
  MewpBugLink,
  MewpExternalFileRef,
  MewpL3L4Pair,
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
    const sourceName = String(
      externalBugsFile?.name || externalBugsFile?.objectName || externalBugsFile?.text || externalBugsFile?.url || ''
    ).trim();
    if (rows.length === 0) {
      if (sourceName) {
        logger.warn(`MEWP external bugs ingestion: source '${sourceName}' loaded with 0 data rows.`);
      }
      return new Map<number, MewpBugLink[]>();
    }
    logger.info(`MEWP external bugs ingestion: start source='${sourceName || 'unknown'}' rows=${rows.length}`);

    let skippedMissingTestCaseId = 0;
    let skippedMissingRequirement = 0;
    let skippedMissingBugId = 0;
    let skippedOutOfScopeState = 0;

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
      if (!testCaseId) {
        skippedMissingTestCaseId += 1;
        continue;
      }
      const requirementBaseKey = adapters.toRequirementKey(
        adapters.toComparableText(this.externalTableUtils.readExternalCell(row, ['SR']))
      );
      if (!requirementBaseKey) {
        skippedMissingRequirement += 1;
        continue;
      }

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
      if (!bugId) {
        skippedMissingBugId += 1;
        continue;
      }

      const bugState = adapters.toComparableText(
        this.externalTableUtils.readExternalCell(row, ['TargetState', 'State'])
      );
      if (!adapters.isExternalStateInScope(bugState, 'bug')) {
        skippedOutOfScopeState += 1;
        continue;
      }

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
    let dedupedRows = 0;
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
      dedupedRows += deduped.get(testCaseId)?.length || 0;
    }

    if (parsedRows === 0) {
      logger.warn(
        `External bugs source was loaded but no valid rows were parsed. ` +
          `Expected columns include Elisra_SortIndex, SR and bug ID fields.`
      );
    }
    logger.info(
      `MEWP external bugs ingestion: done source='${sourceName || 'unknown'}' ` +
        `rows=${rows.length} parsed=${parsedRows} deduped=${dedupedRows} ` +
        `testCases=${deduped.size} ` +
        `skippedMissingTestCaseId=${skippedMissingTestCaseId} ` +
        `skippedMissingRequirement=${skippedMissingRequirement} ` +
        `skippedMissingBugId=${skippedMissingBugId} ` +
        `skippedOutOfScopeState=${skippedOutOfScopeState}`
    );

    return deduped;
  }

  public async loadExternalL3L4ByBaseKey(
    externalL3L4File: MewpExternalFileRef | null | undefined,
    adapters: MewpExternalIngestionAdapters
  ): Promise<Map<string, MewpL3L4Pair[]>> {
    const rows = await this.externalTableUtils.loadExternalTableRows(externalL3L4File, 'l3l4');
    const sourceName = String(
      externalL3L4File?.name || externalL3L4File?.objectName || externalL3L4File?.text || externalL3L4File?.url || ''
    ).trim();
    if (rows.length === 0) {
      if (sourceName) {
        logger.warn(`MEWP external L3/L4 ingestion: source '${sourceName}' loaded with 0 data rows.`);
      }
      return new Map<string, MewpL3L4Pair[]>();
    }
    logger.info(`MEWP external L3/L4 ingestion: start source='${sourceName || 'unknown'}' rows=${rows.length}`);

    const pairsByBaseKey = new Map<string, Map<string, MewpL3L4Pair>>();
    const addPair = (baseKey: string, pair: MewpL3L4Pair) => {
      const l3Id = String(pair?.l3Id || '').trim();
      const l4Id = String(pair?.l4Id || '').trim();
      if (!baseKey || (!l3Id && !l4Id)) return;
      if (!pairsByBaseKey.has(baseKey)) {
        pairsByBaseKey.set(baseKey, new Map<string, MewpL3L4Pair>());
      }
      const key = `${l3Id}|${l4Id}`;
      const byPair = pairsByBaseKey.get(baseKey)!;
      const existing = byPair.get(key);
      if (!existing) {
        byPair.set(key, {
          l3Id,
          l3Title: String(pair?.l3Title || '').trim(),
          l4Id,
          l4Title: String(pair?.l4Title || '').trim(),
        });
        return;
      }

      byPair.set(key, {
        l3Id: existing.l3Id || l3Id,
        l3Title: existing.l3Title || String(pair?.l3Title || '').trim(),
        l4Id: existing.l4Id || l4Id,
        l4Title: existing.l4Title || String(pair?.l4Title || '').trim(),
      });
    };

    let parsedRows = 0;
    let skippedMissingRequirement = 0;
    let acceptedL3 = 0;
    let acceptedL4 = 0;
    let filteredByState = 0;
    let filteredBySapWbs = 0;
    for (const row of rows) {
      const srRaw = adapters.toComparableText(this.externalTableUtils.readExternalCell(row, ['SR']));
      const baseKey = adapters.toRequirementKey(srRaw);
      if (!baseKey) {
        skippedMissingRequirement += 1;
        continue;
      }
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
        // AREA 34 = Level 4 supports two modes:
        // 1) Direct L2->L4: only Level-3 target column is populated; treat it as L4.
        // 2) Paired L3+L4 in same row: Level-3 and Level-4 columns are both populated.
        const effectiveSapWbsLevel3 = targetSapWbsLevel3 || requirementSapWbsFallback;
        const effectiveSapWbsLevel4 = targetSapWbsLevel4 || requirementSapWbsFallback;
        const allowLevel3State = adapters.isExternalStateInScope(targetStateLevel3, 'requirement');
        const allowLevel3SapWbs = !adapters.isExcludedL3L4BySapWbs(effectiveSapWbsLevel3);
        const allowLevel4State = adapters.isExternalStateInScope(targetStateLevel4, 'requirement');
        const allowLevel4SapWbs = !adapters.isExcludedL3L4BySapWbs(effectiveSapWbsLevel4);

        const hasLevel4Columns = !!targetIdLevel4;
        if (hasLevel4Columns) {
          if (!allowLevel3State) filteredByState += 1;
          if (!allowLevel3SapWbs) filteredBySapWbs += 1;
          if (!allowLevel4State) filteredByState += 1;
          if (!allowLevel4SapWbs) filteredBySapWbs += 1;

          const includeL3 = !!(targetIdLevel3 && allowLevel3State && allowLevel3SapWbs);
          const includeL4 = !!(targetIdLevel4 && allowLevel4State && allowLevel4SapWbs);
          if (includeL3 || includeL4) {
            addPair(baseKey, {
              l3Id: includeL3 ? String(targetIdLevel3) : '',
              l3Title: includeL3 ? targetTitleLevel3 : '',
              l4Id: includeL4 ? String(targetIdLevel4) : '',
              l4Title: includeL4 ? targetTitleLevel4 : '',
            });
          }
          if (includeL3) acceptedL3 += 1;
          if (includeL4) acceptedL4 += 1;
        } else {
          // Direct L2->L4 mode (legacy file semantics): Level-3 target column carries L4 ID/title.
          const allowDirectL4State = adapters.isExternalStateInScope(targetStateLevel3, 'requirement');
          const allowDirectL4SapWbs = !adapters.isExcludedL3L4BySapWbs(effectiveSapWbsLevel3);
          if (!allowDirectL4State) filteredByState += 1;
          if (!allowDirectL4SapWbs) filteredBySapWbs += 1;
          if (targetIdLevel3 && allowDirectL4State && allowDirectL4SapWbs) {
            addPair(baseKey, {
              l3Id: '',
              l3Title: '',
              l4Id: String(targetIdLevel3),
              l4Title: targetTitleLevel3,
            });
            acceptedL4 += 1;
          }
        }
        parsedRows += 1;
        continue;
      }

      const effectiveSapWbsLevel3 = targetSapWbsLevel3 || requirementSapWbsFallback;
      const allowLevel3State = adapters.isExternalStateInScope(targetStateLevel3, 'requirement');
      const allowLevel3SapWbs = !adapters.isExcludedL3L4BySapWbs(effectiveSapWbsLevel3);
      if (!allowLevel3State) filteredByState += 1;
      if (!allowLevel3SapWbs) filteredBySapWbs += 1;

      const effectiveSapWbsLevel4 = targetSapWbsLevel4 || requirementSapWbsFallback;
      const allowLevel4State = adapters.isExternalStateInScope(targetStateLevel4, 'requirement');
      const allowLevel4SapWbs = !adapters.isExcludedL3L4BySapWbs(effectiveSapWbsLevel4);
      if (!allowLevel4State) filteredByState += 1;
      if (!allowLevel4SapWbs) filteredBySapWbs += 1;

      const includeL3 = !!(targetIdLevel3 && allowLevel3State && allowLevel3SapWbs);
      const includeL4 = !!(targetIdLevel4 && allowLevel4State && allowLevel4SapWbs);
      if (includeL3 || includeL4) {
        addPair(baseKey, {
          l3Id: includeL3 ? String(targetIdLevel3) : '',
          l3Title: includeL3 ? targetTitleLevel3 : '',
          l4Id: includeL4 ? String(targetIdLevel4) : '',
          l4Title: includeL4 ? targetTitleLevel4 : '',
        });
      }
      if (includeL3) acceptedL3 += 1;
      if (includeL4) acceptedL4 += 1;
      parsedRows += 1;
    }

    const out = new Map<string, MewpL3L4Pair[]>();
    for (const [baseKey, byPair] of pairsByBaseKey.entries()) {
      out.set(
        baseKey,
        [...byPair.values()].sort((a, b) => {
          const aL3 = String(a?.l3Id || '');
          const bL3 = String(b?.l3Id || '');
          const l3Compare = aL3.localeCompare(bL3);
          if (l3Compare !== 0) return l3Compare;

          const aL4 = String(a?.l4Id || '');
          const bL4 = String(b?.l4Id || '');
          return aL4.localeCompare(bL4);
        })
      );
    }

    if (parsedRows === 0) {
      logger.warn(
        `External L3/L4 source was loaded but no valid rows were parsed. ` +
          `Expected columns include SR, AREA 34, target IDs/titles/states.`
      );
    }
    const totalLinks = [...out.values()].reduce((sum, items) => sum + (items?.length || 0), 0);
    logger.info(
      `MEWP external L3/L4 ingestion: done source='${sourceName || 'unknown'}' ` +
        `rows=${rows.length} parsed=${parsedRows} baseKeys=${out.size} links=${totalLinks} ` +
        `acceptedL3=${acceptedL3} acceptedL4=${acceptedL4} ` +
        `skippedMissingRequirement=${skippedMissingRequirement} ` +
        `filteredByState=${filteredByState} filteredBySapWbs=${filteredBySapWbs}`
    );

    return out;
  }

  private toPositiveNumber(value: any): number {
    const parsed = Number(String(value || '').replace(/[^\d]/g, ''));
    return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
  }
}
