import axios from 'axios';
import * as XLSX from 'xlsx';
import type { MewpExternalFileRef, MewpExternalTableValidationResult } from '../models/mewp-reporting';
import logger from './logger';

export type MewpExternalTableType = 'bugs' | 'l3l4';

interface MewpRequiredColumn {
  label: string;
  aliases: string[];
}

export interface MewpExternalRowsWithMeta {
  rows: Array<Record<string, any>>;
  meta: {
    sourceName: string;
    headerRow: 'A3' | 'A1' | '';
    matchedRequiredColumns: number;
    totalRequiredColumns: number;
  };
}

export class MewpExternalFileValidationError extends Error {
  public readonly statusCode: number;
  public readonly code: string;
  public readonly details: MewpExternalTableValidationResult;

  constructor(message: string, details: MewpExternalTableValidationResult) {
    super(message);
    this.name = 'MewpExternalFileValidationError';
    this.statusCode = 422;
    this.code = 'MEWP_EXTERNAL_FILE_VALIDATION_FAILED';
    this.details = details;
  }
}

export default class MewpExternalTableUtils {
  private static readonly EXTERNAL_BUGS_REQUIRED_COLUMNS: MewpRequiredColumn[] = [
    { label: 'Elisra_SortIndex', aliases: ['elisrasortindex', 'elisra sortindex', 'elisra_sortindex'] },
    { label: 'SR', aliases: ['sr'] },
    { label: 'TargetWorkItemId', aliases: ['targetworkitemid', 'bugid', 'bug id'] },
    { label: 'Title', aliases: ['title', 'bugtitle', 'bug title'] },
    { label: 'TargetState', aliases: ['targetstate', 'state'] },
  ];

  private static readonly EXTERNAL_L3L4_REQUIRED_COLUMNS: MewpRequiredColumn[] = [
    { label: 'SR', aliases: ['sr'] },
    { label: 'AREA 34', aliases: ['area34', 'area 34'] },
    {
      label: 'TargetWorkItemId Level 3',
      aliases: ['targetworkitemidlevel3', 'targetworkitemid level 3', 'targetworkitemidlevel 3'],
    },
    {
      label: 'TargetTitleLevel3',
      aliases: ['targettitlelevel3', 'targettitlelevel 3', 'targettitle level 3'],
    },
    {
      label: 'TargetStateLevel 3',
      aliases: ['targetstatelevel3', 'targetstatelevel 3', 'targetstate level 3'],
    },
    {
      label: 'TargetWorkItemIdLevel 4',
      aliases: ['targetworkitemidlevel4', 'targetworkitemid level 4', 'targetworkitemidlevel 4'],
    },
    {
      label: 'TargetTitleLevel4',
      aliases: ['targettitlelevel4', 'targettitlelevel 4', 'targettitle level 4'],
    },
    {
      label: 'TargetStateLevel 4',
      aliases: ['targetstatelevel4', 'targetstatelevel 4', 'targetstate level 4'],
    },
  ];

  private static readonly ALLOWED_EXTERNAL_FILE_EXTENSIONS = new Set<string>(['.xlsx', '.xls', '.csv']);
  private static readonly DEFAULT_EXTERNAL_FILE_MAX_BYTES = 20 * 1024 * 1024;

  public getRequiredColumnLabels(tableType: MewpExternalTableType): string[] {
    return this.getRequiredColumns(tableType).map((item) => item.label);
  }

  public getRequiredColumnCount(tableType: MewpExternalTableType): number {
    return this.getRequiredColumns(tableType).length;
  }

  public readExternalCell(row: Record<string, any>, aliases: string[]): any {
    if (!row || typeof row !== 'object') return '';
    const byNormalizedKey = new Map<string, any>();
    for (const [key, value] of Object.entries(row)) {
      byNormalizedKey.set(this.normalizeExternalColumnKey(key), value);
    }

    for (const alias of aliases) {
      const value = byNormalizedKey.get(this.normalizeExternalColumnKey(alias));
      if (value !== undefined && value !== null && String(value).trim() !== '') {
        return value;
      }
    }
    return '';
  }

  public async loadExternalTableRows(
    source: MewpExternalFileRef | null | undefined,
    tableType: MewpExternalTableType
  ): Promise<Array<Record<string, any>>> {
    const { rows } = await this.loadExternalTableRowsWithMeta(source, tableType);
    return rows;
  }

  public async loadExternalTableRowsWithMeta(
    source: MewpExternalFileRef | null | undefined,
    tableType: MewpExternalTableType
  ): Promise<MewpExternalRowsWithMeta> {
    const sourceName = String(source?.name || source?.objectName || source?.text || source?.url || '').trim();
    const sourceUrl = this.resolveMewpExternalSourceUrl(source);
    const extension = this.resolveMewpExternalSourceExtension(source);
    const sourceIsolationCheck = this.validateMewpExternalSourceIsolation(source, sourceUrl);
    const requiredColumns = this.getRequiredColumns(tableType);

    if (!sourceName && !sourceUrl) {
      return {
        rows: [],
        meta: {
          sourceName: '',
          headerRow: '',
          matchedRequiredColumns: 0,
          totalRequiredColumns: requiredColumns.length,
        },
      };
    }

    if (!sourceUrl) {
      throw this.createMewpExternalValidationError(
        tableType,
        sourceName || 'unknown',
        '',
        0,
        requiredColumns.length,
        requiredColumns.map((item) => item.label),
        `Missing file URL/object reference for '${sourceName || tableType}'`
      );
    }
    if (!sourceIsolationCheck.valid) {
      throw this.createMewpExternalValidationError(
        tableType,
        sourceName || sourceUrl,
        '',
        0,
        requiredColumns.length,
        requiredColumns.map((item) => item.label),
        sourceIsolationCheck.message
      );
    }

    if (!MewpExternalTableUtils.ALLOWED_EXTERNAL_FILE_EXTENSIONS.has(extension)) {
      throw this.createMewpExternalValidationError(
        tableType,
        sourceName || sourceUrl,
        '',
        0,
        requiredColumns.length,
        requiredColumns.map((item) => item.label),
        `Unsupported file type '${extension || 'unknown'}'. Allowed: .xlsx, .xls, .csv`
      );
    }

    const parseRows = (sheet: XLSX.WorkSheet, startRowZeroBased: number) =>
      XLSX.utils.sheet_to_json(sheet, {
        defval: '',
        raw: false,
        range: startRowZeroBased,
      }) as Array<Record<string, any>>;

    const parseHeaderKeys = (sheet: XLSX.WorkSheet, startRowZeroBased: number): Set<string> => {
      const matrix = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
        range: startRowZeroBased,
      }) as any[][];
      const firstRow = Array.isArray(matrix?.[0]) ? matrix[0] : [];
      const normalized = firstRow
        .map((item) => this.normalizeExternalColumnKey(item))
        .filter((item) => !!item);
      return new Set(normalized);
    };

    const validateHeaders = (headerKeys: Set<string>) => {
      const missingLabels: string[] = [];
      let matched = 0;
      for (const requiredColumn of requiredColumns) {
        const hasMatch = requiredColumn.aliases
          .map((alias) => this.normalizeExternalColumnKey(alias))
          .some((alias) => headerKeys.has(alias));
        if (hasMatch) {
          matched += 1;
        } else {
          missingLabels.push(requiredColumn.label);
        }
      }
      return {
        matched,
        total: requiredColumns.length,
        missingLabels,
        isValid: missingLabels.length === 0,
      };
    };

    try {
      const response = await axios.get(sourceUrl, {
        responseType: 'arraybuffer',
        timeout: 45000,
      });
      const maxBytes = this.resolveMewpExternalMaxFileSize();
      const declaredLength = Number(response?.headers?.['content-length'] || 0);
      if (Number.isFinite(declaredLength) && declaredLength > maxBytes) {
        throw this.createMewpExternalValidationError(
          tableType,
          sourceName || sourceUrl,
          '',
          0,
          requiredColumns.length,
          requiredColumns.map((item) => item.label),
          `File exceeds maximum allowed size (${maxBytes} bytes).`
        );
      }

      const buffer = Buffer.from(response?.data || []);
      if (!buffer.length) {
        throw this.createMewpExternalValidationError(
          tableType,
          sourceName || sourceUrl,
          '',
          0,
          requiredColumns.length,
          requiredColumns.map((item) => item.label),
          'File is empty'
        );
      }
      if (buffer.length > maxBytes) {
        throw this.createMewpExternalValidationError(
          tableType,
          sourceName || sourceUrl,
          '',
          0,
          requiredColumns.length,
          requiredColumns.map((item) => item.label),
          `File exceeds maximum allowed size (${maxBytes} bytes).`
        );
      }

      const workbook = XLSX.read(buffer, { type: 'buffer' });
      const firstSheetName = workbook?.SheetNames?.[0];
      if (!firstSheetName) {
        throw this.createMewpExternalValidationError(
          tableType,
          sourceName || sourceUrl,
          '',
          0,
          requiredColumns.length,
          requiredColumns.map((item) => item.label),
          'No worksheet was found in the uploaded file'
        );
      }
      const sheet = workbook.Sheets[firstSheetName];
      if (!sheet) {
        throw this.createMewpExternalValidationError(
          tableType,
          sourceName || sourceUrl,
          '',
          0,
          requiredColumns.length,
          requiredColumns.map((item) => item.label),
          'Worksheet data could not be read'
        );
      }

      // Expected header row is A3, but keep A1 fallback for backward compatibility.
      const headerA3 = parseHeaderKeys(sheet, 2);
      const headerA1 = parseHeaderKeys(sheet, 0);
      const rowsFromA3 = parseRows(sheet, 2);
      const rowsFromA1 = parseRows(sheet, 0);
      const validationA3 = validateHeaders(headerA3);
      const validationA1 = validateHeaders(headerA1);

      if (validationA3.isValid) {
        return {
          rows: rowsFromA3,
          meta: {
            sourceName: sourceName || sourceUrl,
            headerRow: 'A3',
            matchedRequiredColumns: validationA3.matched,
            totalRequiredColumns: validationA3.total,
          },
        };
      }
      if (validationA1.isValid) {
        return {
          rows: rowsFromA1,
          meta: {
            sourceName: sourceName || sourceUrl,
            headerRow: 'A1',
            matchedRequiredColumns: validationA1.matched,
            totalRequiredColumns: validationA1.total,
          },
        };
      }

      const best = validationA3.matched >= validationA1.matched ? validationA3 : validationA1;
      throw this.createMewpExternalValidationError(
        tableType,
        sourceName || sourceUrl,
        '',
        best.matched,
        best.total,
        best.missingLabels,
        `Missing required columns: ${best.missingLabels.join(', ')}. Expected header row at A3 (fallback A1 was also checked).`
      );
    } catch (error: any) {
      if (error instanceof MewpExternalFileValidationError) {
        throw error;
      }
      const msg = String(error?.message || error || '').trim();
      logger.warn(`Could not load external MEWP source '${sourceUrl}': ${msg}`);
      throw this.createMewpExternalValidationError(
        tableType,
        sourceName || sourceUrl,
        '',
        0,
        requiredColumns.length,
        requiredColumns.map((item) => item.label),
        `Unable to load or parse the file: ${msg}`
      );
    }
  }

  private getRequiredColumns(tableType: MewpExternalTableType): MewpRequiredColumn[] {
    return tableType === 'bugs'
      ? MewpExternalTableUtils.EXTERNAL_BUGS_REQUIRED_COLUMNS
      : MewpExternalTableUtils.EXTERNAL_L3L4_REQUIRED_COLUMNS;
  }

  private createMewpExternalValidationError(
    tableType: MewpExternalTableType,
    sourceName: string,
    headerRow: 'A3' | 'A1' | '',
    matchedRequiredColumns: number,
    totalRequiredColumns: number,
    missingRequiredColumns: string[],
    message: string
  ): MewpExternalFileValidationError {
    const details: MewpExternalTableValidationResult = {
      tableType,
      sourceName,
      valid: false,
      headerRow,
      matchedRequiredColumns,
      totalRequiredColumns,
      missingRequiredColumns,
      rowCount: 0,
      message,
    };
    return new MewpExternalFileValidationError(message, details);
  }

  private resolveMewpExternalSourceUrl(source: MewpExternalFileRef | null | undefined): string {
    const directUrl = String(source?.url || '').trim();
    if (directUrl) return directUrl;

    const bucketName = String(source?.bucketName || '').trim();
    const objectName = String(source?.objectName || source?.text || '').trim();
    const minioBase = String(process.env.MINIOSERVER || '').trim().replace(/\/+$/g, '');
    if (!bucketName || !objectName || !minioBase) return '';
    const encodedObjectName = objectName
      .split('/')
      .map((segment) => encodeURIComponent(segment))
      .join('/');
    return `${minioBase}/${bucketName}/${encodedObjectName}`;
  }

  private validateMewpExternalSourceIsolation(
    source: MewpExternalFileRef | null | undefined,
    sourceUrl: string
  ): { valid: boolean; message: string } {
    const dedicatedBucket = String(
      process.env.MEWP_EXTERNAL_INGESTION_BUCKET || 'mewp-external-ingestion'
    ).trim();
    const declaredSourceType = String(source?.sourceType || '').trim().toLowerCase();
    if (declaredSourceType && declaredSourceType !== 'mewpexternalingestion') {
      return {
        valid: false,
        message: `Unsupported sourceType '${source?.sourceType}'. Expected 'mewpExternalIngestion'.`,
      };
    }

    const inferred = this.extractMewpBucketAndObjectFromUrl(sourceUrl);
    const bucketName = String(source?.bucketName || inferred.bucketName || '').trim();
    const objectName = String(source?.objectName || source?.text || inferred.objectName || '').trim();

    if (bucketName && bucketName !== dedicatedBucket) {
      return {
        valid: false,
        message: `Invalid storage bucket '${bucketName}'. Expected '${dedicatedBucket}'.`,
      };
    }

    if (objectName) {
      const normalizedObject = objectName.toLowerCase();
      if (!normalizedObject.includes('/mewp-external-ingestion/')) {
        return {
          valid: false,
          message: `Invalid object path '${objectName}'. Expected '/mewp-external-ingestion/' prefix segment.`,
        };
      }
    }

    return { valid: true, message: '' };
  }

  private extractMewpBucketAndObjectFromUrl(url: string): { bucketName: string; objectName: string } {
    try {
      const parsed = new URL(String(url || '').trim());
      const segments = decodeURIComponent(parsed.pathname || '')
        .replace(/^\/+/g, '')
        .split('/')
        .filter((item) => !!item);
      if (segments.length < 2) return { bucketName: '', objectName: '' };
      return {
        bucketName: String(segments[0] || '').trim(),
        objectName: segments.slice(1).join('/'),
      };
    } catch {
      return { bucketName: '', objectName: '' };
    }
  }

  private resolveMewpExternalSourceExtension(source: MewpExternalFileRef | null | undefined): string {
    const candidates = [source?.name, source?.objectName, source?.text, source?.url]
      .map((value) => String(value || '').trim())
      .filter((value) => !!value);
    for (const candidate of candidates) {
      const clean = candidate.split('?')[0].split('#')[0];
      const match = /\.([a-z0-9]+)$/i.exec(clean);
      if (!match) continue;
      return `.${String(match[1] || '').toLowerCase()}`;
    }
    return '';
  }

  private resolveMewpExternalMaxFileSize(): number {
    const configured = Number(process.env.MEWP_EXTERNAL_MAX_FILE_SIZE_BYTES || 0);
    if (Number.isFinite(configured) && configured > 0) return configured;
    return MewpExternalTableUtils.DEFAULT_EXTERNAL_FILE_MAX_BYTES;
  }

  private normalizeExternalColumnKey(value: any): string {
    return String(value || '')
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '');
  }
}
