#!/usr/bin/env node

import { createHash } from 'crypto';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
import mammoth from 'mammoth';
import { PDFParse } from 'pdf-parse';
import Database from 'better-sqlite3';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';
import { loadSecureAuthStore, resolveAuthStorePath } from './auth-store.js';
import { PaddleOcrService } from 'ppu-paddle-ocr';
import { createCanvas } from '@napi-rs/canvas';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { ImageProcessor } from 'ppu-ocv';

const GOOGLE_APPS_EXPORT_MIME: Record<string, string> = {
  'application/vnd.google-apps.document': 'text/plain',
  'application/vnd.google-apps.spreadsheet': 'text/csv',
  'application/vnd.google-apps.presentation': 'text/plain',
};

const EXTRA_TEXT_MIME_TYPES = new Set([
  'application/json',
  'application/ld+json',
  'application/xml',
  'application/javascript',
  'application/x-javascript',
  'application/x-sh',
  'application/x-httpd-php',
  'application/sql',
]);

const DOCX_MIME_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/msword',
]);

const PDF_MIME_TYPES = new Set([
  'application/pdf',
]);

const PPTX_MIME_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'application/vnd.ms-powerpoint',
]);

const XLSX_MIME_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel',
]);

const IMAGE_MIME_TYPES = new Set([
  'image/png',
  'image/jpeg',
  'image/gif',
  'image/bmp',
  'image/tiff',
  'image/webp',
]);

const IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif', '.webp'];

const OCR_TIMEOUT_MS = 120_000; // 2 minutes for large server models
const OCR_MAX_IMAGE_SIZE = 10 * 1024 * 1024;
const OCR_MAX_IMAGES_PER_DOC = 10;
const SCANNED_PDF_TEXT_THRESHOLD = 100;

const PADDLE_MODEL_BASE = 'https://media.githubusercontent.com/media/PT-Perkasa-Pilar-Utama/ppu-paddle-ocr-models/main';
const PADDLE_DICT_BASE = 'https://raw.githubusercontent.com/PT-Perkasa-Pilar-Utama/ppu-paddle-ocr-models/main';

const CYRILLIC_RE = /[\u0400-\u04FF]/;

type EntityType = 'assignments' | 'courses' | 'announcements';
type SearchSort = 'relevance' | 'dueDate' | 'updatedAt';

type SearchDoc = {
  key: string;
  entityType: EntityType;
  id: string;
  courseId?: string;
  courseName?: string;
  title: string;
  body?: string;
  titleNormalized: string;
  bodyNormalized: string;
  state?: string;
  dueDate?: string;
  dueTimestamp?: number;
  updatedAt?: string;
  updatedTimestamp?: number;
  alternateLink?: string;
  assignedGrade?: number;
  maxPoints?: number;
  isMissing?: boolean;
};

type SearchInput = {
  query?: string;
  entityTypes?: EntityType[];
  courseIds?: string[];
  states?: string[];
  dueFrom?: string;
  dueTo?: string;
  missingOnly?: boolean;
  gradedOnly?: boolean;
  limit?: number;
  cursor?: string;
  sort?: SearchSort;
  forceRefresh?: boolean;
};

type SearchHit = {
  entityType: EntityType;
  id: string;
  courseId?: string;
  courseName?: string;
  title: string;
  snippet?: string;
  dueDate?: string;
  state?: string;
  score: number;
  scoreReason: string[];
  alternateLink?: string;
  updatedAt?: string;
  assignedGrade?: number;
  maxPoints?: number;
};

type SearchOutput = {
  items: SearchHit[];
  facets: {
    entityTypes: Record<string, number>;
    states: Record<string, number>;
    courses: Array<{ courseId: string; courseName?: string; count: number }>;
  };
  totalApprox: number;
  nextCursor: string | null;
  meta: {
    indexVersion: number;
    lastRefreshAt: string | null;
    cacheStatus: 'hit' | 'refresh';
    warnings: string[];
  };
};

type SuggestInput = {
  prefix: string;
  entityTypes?: EntityType[];
  limit?: number;
  forceRefresh?: boolean;
};

type SuggestOutput = {
  suggestions: Array<{ text: string; kind: EntityType; weight: number }>;
  meta: {
    indexVersion: number;
    lastRefreshAt: string | null;
    cacheStatus: 'hit' | 'refresh';
    warnings: string[];
  };
};

type CursorPayload = {
  offset: number;
  queryHash: string;
  filtersHash: string;
  indexVersion: number;
};

type RefreshSnapshot = {
  docs: SearchDoc[];
  warnings: string[];
  sourceFailures: number;
};

type RefreshState = {
  refreshed: boolean;
  warnings: string[];
};

type CachedMaterialDoc = {
  docRef: string;
  courseId: string;
  courseWorkId: string;
  assignmentTitle?: string | null;
  materialIndex: number;
  type: string;
  title?: string | null;
  url?: string | null;
  driveFileId?: string;
  mimeType?: string | null;
  text: string;
  textNormalized: string;
  extractedAt: string;
  expiresAtMs: number;
};

class SearchValidationError extends Error {
  constructor(public readonly code: 'INVALID_CURSOR' | 'INVALID_FILTER', message: string) {
    super(message);
  }
}

class SearchService {
  private readonly store = new Map<EntityType, Map<string, SearchDoc>>([
    ['assignments', new Map<string, SearchDoc>()],
    ['courses', new Map<string, SearchDoc>()],
    ['announcements', new Map<string, SearchDoc>()],
  ]);
  private version = 0;
  private lastRefreshAt: number | null = null;
  private expiresAt = 0;
  private lastWarnings: string[] = [];

  constructor(
    private readonly ttlMs: number,
    private readonly refreshSnapshot: () => Promise<RefreshSnapshot>
  ) {}

  public async search(input: SearchInput): Promise<SearchOutput> {
    const refreshState = await this.ensureIndex(Boolean(input.forceRefresh));
    const now = Date.now();
    const limit = clampNumber(input.limit, 20, 1, 100);
    const sort: SearchSort = input.sort ?? 'relevance';
    if (!['relevance', 'dueDate', 'updatedAt'].includes(sort)) {
      throw new SearchValidationError('INVALID_FILTER', `Unsupported sort value: ${String(input.sort)}`);
    }

    const entityTypes = sanitizeEntityTypes(input.entityTypes);
    const queryNormalized = normalizeText(input.query || '');
    const dueFromTs = parseDateFilter(input.dueFrom, 'dueFrom');
    const dueToTs = parseDateFilter(input.dueTo, 'dueTo');
    if (dueFromTs != null && dueToTs != null && dueFromTs > dueToTs) {
      throw new SearchValidationError('INVALID_FILTER', 'dueFrom must be before or equal to dueTo');
    }

    const courseIds = new Set((input.courseIds || []).filter(Boolean));
    const states = new Set((input.states || []).map((value) => value.toUpperCase()));

    const filterSignature = hashString(
      JSON.stringify({
        entityTypes: [...entityTypes].sort(),
        courseIds: [...courseIds].sort(),
        states: [...states].sort(),
        dueFrom: input.dueFrom || null,
        dueTo: input.dueTo || null,
        missingOnly: Boolean(input.missingOnly),
        gradedOnly: Boolean(input.gradedOnly),
        sort,
        limit,
      })
    );
    const queryHash = hashString(queryNormalized);

    let offset = 0;
    if (input.cursor) {
      const decoded = decodeCursor(input.cursor);
      if (decoded.indexVersion !== this.version) {
        throw new SearchValidationError('INVALID_CURSOR', 'Cursor index version is stale');
      }
      if (decoded.queryHash !== queryHash || decoded.filtersHash !== filterSignature) {
        throw new SearchValidationError('INVALID_CURSOR', 'Cursor does not match current query or filters');
      }
      offset = decoded.offset;
    }

    const allDocs = this.collectDocs(entityTypes);
    const ranked = allDocs
      .map((doc) => {
        const rankedDoc = this.rankDoc(doc, queryNormalized, now);
        return rankedDoc;
      })
      .filter((rankedDoc) => {
        const doc = rankedDoc.doc;
        if (queryNormalized && rankedDoc.textScore <= 0) return false;
        if (courseIds.size > 0 && (!doc.courseId || !courseIds.has(doc.courseId))) return false;
        if (states.size > 0 && (!doc.state || !states.has(doc.state.toUpperCase()))) return false;
        if (Boolean(input.gradedOnly) && doc.entityType === 'assignments') {
          if (doc.assignedGrade == null || doc.maxPoints == null) return false;
        }
        if (Boolean(input.missingOnly) && doc.entityType === 'assignments' && !doc.isMissing) return false;
        if (dueFromTs != null || dueToTs != null) {
          if (doc.dueTimestamp == null) return false;
          if (dueFromTs != null && doc.dueTimestamp < dueFromTs) return false;
          if (dueToTs != null && doc.dueTimestamp > dueToTs) return false;
        }
        return true;
      });

    ranked.sort((a, b) => {
      if (sort === 'dueDate') {
        const leftDue = a.doc.dueTimestamp ?? Number.MAX_SAFE_INTEGER;
        const rightDue = b.doc.dueTimestamp ?? Number.MAX_SAFE_INTEGER;
        if (leftDue !== rightDue) return leftDue - rightDue;
        return b.score - a.score;
      }

      if (sort === 'updatedAt') {
        const leftUpdated = a.doc.updatedTimestamp ?? 0;
        const rightUpdated = b.doc.updatedTimestamp ?? 0;
        if (leftUpdated !== rightUpdated) return rightUpdated - leftUpdated;
        return b.score - a.score;
      }

      if (b.score !== a.score) return b.score - a.score;
      const leftDue = a.doc.dueTimestamp ?? Number.MAX_SAFE_INTEGER;
      const rightDue = b.doc.dueTimestamp ?? Number.MAX_SAFE_INTEGER;
      if (leftDue !== rightDue) return leftDue - rightDue;
      return (b.doc.updatedTimestamp ?? 0) - (a.doc.updatedTimestamp ?? 0);
    });

    if (offset < 0 || offset > ranked.length) {
      throw new SearchValidationError('INVALID_CURSOR', 'Cursor offset is out of range');
    }

    const paged = ranked.slice(offset, offset + limit);
    const nextOffset = offset + paged.length;
    const nextCursor = nextOffset < ranked.length
      ? encodeCursor({
          offset: nextOffset,
          queryHash,
          filtersHash: filterSignature,
          indexVersion: this.version,
        })
      : null;

    const facets = buildFacets(ranked.map((item) => item.doc));

    return {
      items: paged.map((item) => ({
        entityType: item.doc.entityType,
        id: item.doc.id,
        courseId: item.doc.courseId,
        courseName: item.doc.courseName,
        title: item.doc.title,
        snippet: item.snippet,
        dueDate: item.doc.dueDate,
        state: item.doc.state,
        score: item.score,
        scoreReason: item.reasons,
        alternateLink: item.doc.alternateLink,
        updatedAt: item.doc.updatedAt,
        assignedGrade: item.doc.assignedGrade,
        maxPoints: item.doc.maxPoints,
      })),
      facets,
      totalApprox: ranked.length,
      nextCursor,
      meta: {
        indexVersion: this.version,
        lastRefreshAt: this.lastRefreshAt ? new Date(this.lastRefreshAt).toISOString() : null,
        cacheStatus: refreshState.refreshed ? 'refresh' : 'hit',
        warnings: refreshState.warnings,
      },
    };
  }

  public async suggest(input: SuggestInput): Promise<SuggestOutput> {
    const refreshState = await this.ensureIndex(Boolean(input.forceRefresh));
    const prefix = normalizeText(input.prefix || '');
    if (!prefix) {
      throw new SearchValidationError('INVALID_FILTER', 'prefix is required');
    }

    const limit = clampNumber(input.limit, 10, 1, 25);
    const entityTypes = sanitizeEntityTypes(input.entityTypes);
    const docs = this.collectDocs(entityTypes);
    const buckets = new Map<string, { text: string; kind: EntityType; weight: number }>();

    for (const doc of docs) {
      const title = (doc.title || '').trim();
      if (normalizeText(title).startsWith(prefix)) {
        const key = `${doc.entityType}|${title}`;
        const existing = buckets.get(key);
        buckets.set(key, {
          text: title,
          kind: doc.entityType,
          weight: (existing?.weight || 0) + 6,
        });
      }

      const terms = tokenizeForSuggest(`${doc.title} ${doc.body || ''}`);
      for (const term of terms) {
        if (!term.startsWith(prefix)) continue;
        const key = `${doc.entityType}|${term}`;
        const existing = buckets.get(key);
        buckets.set(key, {
          text: term,
          kind: doc.entityType,
          weight: (existing?.weight || 0) + (doc.titleNormalized.includes(term) ? 3 : 1),
        });
      }
    }

    const suggestions = [...buckets.values()]
      .sort((a, b) => {
        if (b.weight !== a.weight) return b.weight - a.weight;
        return a.text.localeCompare(b.text);
      })
      .slice(0, limit);

    return {
      suggestions,
      meta: {
        indexVersion: this.version,
        lastRefreshAt: this.lastRefreshAt ? new Date(this.lastRefreshAt).toISOString() : null,
        cacheStatus: refreshState.refreshed ? 'refresh' : 'hit',
        warnings: refreshState.warnings,
      },
    };
  }

  private async ensureIndex(forceRefresh: boolean): Promise<RefreshState> {
    const now = Date.now();
    const needsRefresh =
      forceRefresh ||
      this.version === 0 ||
      this.lastRefreshAt == null ||
      now >= this.expiresAt;

    if (!needsRefresh) {
      return { refreshed: false, warnings: this.lastWarnings };
    }

    const snapshot = await this.refreshSnapshot();
    if (snapshot.docs.length === 0 || snapshot.sourceFailures >= 3) {
      throw new Error('Search index refresh failed: no usable data returned from Google Classroom');
    }

    this.store.get('assignments')!.clear();
    this.store.get('courses')!.clear();
    this.store.get('announcements')!.clear();

    for (const doc of snapshot.docs) {
      this.store.get(doc.entityType)!.set(doc.key, doc);
    }

    this.version += 1;
    this.lastRefreshAt = now;
    this.expiresAt = now + this.ttlMs;
    this.lastWarnings = snapshot.warnings;

    return { refreshed: true, warnings: snapshot.warnings };
  }

  private collectDocs(entityTypes: EntityType[]): SearchDoc[] {
    const docs: SearchDoc[] = [];
    for (const type of entityTypes) {
      docs.push(...this.store.get(type)!.values());
    }
    return docs;
  }

  private rankDoc(doc: SearchDoc, normalizedQuery: string, now: number): {
    doc: SearchDoc;
    score: number;
    reasons: string[];
    snippet?: string;
    textScore: number;
  } {
    const reasons: string[] = [];
    let score = 0;
    let textScore = 0;

    if (normalizedQuery) {
      if (doc.titleNormalized.includes(normalizedQuery)) {
        score += 0.55;
        textScore += 0.55;
        reasons.push('title_match');
      }
      if (doc.bodyNormalized.includes(normalizedQuery)) {
        score += 0.25;
        textScore += 0.25;
        reasons.push('content_match');
      }

      const queryTokens = tokenizeForSuggest(normalizedQuery);
      if (queryTokens.length > 0) {
        const tokenMatches = queryTokens.filter((token) =>
          doc.titleNormalized.includes(token) || doc.bodyNormalized.includes(token)
        ).length;
        if (tokenMatches > 0) {
          const tokenBoost = Math.min(0.15, tokenMatches * 0.05);
          score += tokenBoost;
          textScore += tokenBoost;
          reasons.push('token_overlap');
        }
      }
    } else {
      score += 0.05;
      reasons.push('no_query_penalty');
    }

    if (doc.entityType === 'assignments' && doc.dueTimestamp != null) {
      const msPerDay = 24 * 60 * 60 * 1000;
      const daysToDue = Math.floor((doc.dueTimestamp - now) / msPerDay);
      if (daysToDue >= 0 && daysToDue <= 14) {
        const dueBoost = Math.max(0.05, (14 - daysToDue) / 100);
        score += dueBoost;
        reasons.push('due_soon_boost');
      }
      if (doc.isMissing) {
        score += 0.2;
        reasons.push('missing_status_boost');
      }
    }

    if (doc.updatedTimestamp != null) {
      const ageDays = (now - doc.updatedTimestamp) / (24 * 60 * 60 * 1000);
      if (ageDays <= 30) {
        score += 0.1;
        reasons.push('freshness_boost');
      }
    }

    score = Math.max(0, Math.min(1, Number(score.toFixed(4))));
    return {
      doc,
      score,
      reasons,
      snippet: buildSnippet(doc.body || doc.title, normalizedQuery),
      textScore,
    };
  }
}

class OCRProcessor {
  private service: PaddleOcrService | null = null;
  private initPromise: Promise<void> | null = null;
  private currentLang: 'eng' | 'cyrillic' = 'eng';
  private destroyed = false;

  private static readonly MODELS_DIR = path.join(
    path.dirname(fs.realpathSync(process.argv[1] || __filename)),
    '..',
    'models'
  );

  private static readonly LOCAL_MODELS = {
    detection: path.join(OCRProcessor.MODELS_DIR, 'PP-OCRv5_server_det_infer.onnx'),
    enRecognition: path.join(OCRProcessor.MODELS_DIR, 'PP-OCRv5_server_rec_infer.onnx'),
    enDict: path.join(OCRProcessor.MODELS_DIR, 'ppocrv5_dict.txt'),
    cyrillicRecognition: path.join(OCRProcessor.MODELS_DIR, 'cyrillic_PP-OCRv5_mobile_rec_infer.onnx'),
    cyrillicDict: path.join(OCRProcessor.MODELS_DIR, 'ppocrv5_cyrillic_dict.txt'),
  };

  private resolveModel(localPath: string, remoteUrl: string): string {
    return fs.existsSync(localPath) ? localPath : remoteUrl;
  }

  async initialize(): Promise<void> {
    if (this.destroyed) return;
    if (this.service && this.initPromise) {
      await this.initPromise;
      return;
    }
    if (this.initPromise) {
      await this.initPromise;
      return;
    }
    const M = OCRProcessor.LOCAL_MODELS;
    this.service = new PaddleOcrService({
      model: {
        detection: this.resolveModel(M.detection, `${PADDLE_MODEL_BASE}/detection/PP-OCRv5_server_det_infer.onnx`),
        recognition: this.resolveModel(M.enRecognition, `${PADDLE_MODEL_BASE}/recognition/PP-OCRv5_server_rec_infer.onnx`),
        charactersDictionary: this.resolveModel(M.enDict, `${PADDLE_DICT_BASE}/recognition/ppocrv5_dict.txt`),
      },
      processing: { engine: 'opencv' },
      detection: {
        maxSideLength: 1440,
        paddingVertical: 0.5,
        paddingHorizontal: 0.8,
        minimumAreaThreshold: 10,
      },
      session: {
        executionProviders: ['dml', 'cpu'],
        graphOptimizationLevel: 'all',
      },
    });
    this.initPromise = this.service.initialize();
    await this.initPromise;
    console.error('OCR: PaddleOCR initialized with execution providers: dml -> cpu (DirectML GPU acceleration enabled)');
  }

  private async switchToCyrillic(): Promise<void> {
    if (!this.service || this.currentLang === 'cyrillic') return;
    try {
      const M = OCRProcessor.LOCAL_MODELS;
      await this.service.changeRecognitionModel(
        this.resolveModel(M.cyrillicRecognition, `${PADDLE_MODEL_BASE}/recognition/multi/cyrillic/v5/cyrillic_PP-OCRv5_mobile_rec_infer.onnx`)
      );
      await this.service.changeTextDictionary(
        this.resolveModel(M.cyrillicDict, `${PADDLE_DICT_BASE}/recognition/multi/cyrillic/v5/ppocrv5_cyrillic_dict.txt`)
      );
      this.currentLang = 'cyrillic';
    } catch {
      console.error('OCR: Failed to switch to Cyrillic model, keeping English');
    }
  }

  private ocvReady = false;

  private async preprocessImage(imageBuffer: Buffer): Promise<Buffer> {
    try {
      if (!this.ocvReady) {
        await ImageProcessor.initRuntime();
        this.ocvReady = true;
      }
      const { loadImage } = await import('@napi-rs/canvas');
      const img = await loadImage(imageBuffer);
      const w = img.width;
      const h = img.height;
      const srcCanvas = createCanvas(w, h);
      const srcCtx = srcCanvas.getContext('2d');
      srcCtx.drawImage(img, 0, 0);

      const processor = new ImageProcessor(srcCanvas as any);
      processor.grayscale().adaptiveThreshold();
      const resultCanvas = processor.toCanvas() as any;
      processor.destroy();

      return resultCanvas.toBuffer('image/png') as Buffer;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error(`OCR preprocessImage failed, using raw image: ${msg}`);
      return imageBuffer;
    }
  }

  async recognizeImage(imageBuffer: Buffer): Promise<string> {
    if (this.destroyed) return '';
    if (imageBuffer.length > OCR_MAX_IMAGE_SIZE) return '';
    try {
      await this.initialize();
      if (!this.service) return '';

      const arrayBuffer = imageBuffer.slice().buffer as ArrayBuffer;

      const result = await Promise.race([
        this.service.recognize(arrayBuffer),
        new Promise<never>((_, reject) =>
          setTimeout(() => reject(new Error('OCR timeout')), OCR_TIMEOUT_MS)
        ),
      ]);

      const text = (result.text || '').trim();
      if (!text) return '';

      if (this.currentLang === 'eng' && CYRILLIC_RE.test(text)) {
        const cyrillicChars = (text.match(/[\u0400-\u04FF]/g) || []).length;
        if (cyrillicChars / text.length > 0.15) {
          await this.switchToCyrillic();
          const retryResult = await this.service.recognize(arrayBuffer);
          const retryText = (retryResult.text || '').trim();
          return retryText.length >= text.length * 0.7 ? retryText : text;
        }
      }

      return text;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error(`OCR recognizeImage failed: ${msg}`);
      return '';
    }
  }

  async recognizeImages(buffers: Buffer[]): Promise<string[]> {
    const results: string[] = [];
    const limited = buffers.slice(0, OCR_MAX_IMAGES_PER_DOC);
    for (let i = 0; i < limited.length; i++) {
      const text = await this.recognizeImage(limited[i]);
      if (text) {
        results.push(`[Image ${results.length + 1}]: ${text}`);
      }
    }
    return results;
  }

  async destroy(): Promise<void> {
    if (this.destroyed) return;
    this.destroyed = true;
    if (this.service) {
      try {
        await this.service.destroy();
      } catch {}
    }
    this.service = null;
    this.initPromise = null;
  }
}

class MaterialCacheStore {
  private readonly db: Database.Database;

  constructor(dbPath: string) {
    this.db = new Database(dbPath);
    this.db.pragma('journal_mode = WAL');
    this.db.pragma('synchronous = NORMAL');
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS material_cache (
        doc_ref TEXT PRIMARY KEY,
        course_id TEXT NOT NULL,
        course_work_id TEXT NOT NULL,
        assignment_title TEXT,
        material_index INTEGER NOT NULL,
        type TEXT NOT NULL,
        title TEXT,
        url TEXT,
        drive_file_id TEXT,
        mime_type TEXT,
        text TEXT NOT NULL,
        text_normalized TEXT NOT NULL,
        extracted_at TEXT NOT NULL,
        expires_at_ms INTEGER NOT NULL
      );
      CREATE INDEX IF NOT EXISTS idx_material_cache_course ON material_cache(course_id);
      CREATE INDEX IF NOT EXISTS idx_material_cache_coursework ON material_cache(course_work_id);
      CREATE INDEX IF NOT EXISTS idx_material_cache_expires ON material_cache(expires_at_ms);
    `);
  }

public cleanupExpired(nowMs: number): number {
    const result = this.db.prepare('DELETE FROM material_cache WHERE expires_at_ms <= ?').run(nowMs);
    return Number(result.changes || 0);
  }

  public clearAll(): number {
    const result = this.db.prepare('DELETE FROM material_cache').run();
    return Number(result.changes || 0);
  }

  public upsert(doc: CachedMaterialDoc) {
    this.db.prepare(`
      INSERT INTO material_cache (
        doc_ref, course_id, course_work_id, assignment_title, material_index, type, title, url, drive_file_id,
        mime_type, text, text_normalized, extracted_at, expires_at_ms
      ) VALUES (
        @docRef, @courseId, @courseWorkId, @assignmentTitle, @materialIndex, @type, @title, @url, @driveFileId,
        @mimeType, @text, @textNormalized, @extractedAt, @expiresAtMs
      )
      ON CONFLICT(doc_ref) DO UPDATE SET
        course_id = excluded.course_id,
        course_work_id = excluded.course_work_id,
        assignment_title = excluded.assignment_title,
        material_index = excluded.material_index,
        type = excluded.type,
        title = excluded.title,
        url = excluded.url,
        drive_file_id = excluded.drive_file_id,
        mime_type = excluded.mime_type,
        text = excluded.text,
        text_normalized = excluded.text_normalized,
        extracted_at = excluded.extracted_at,
        expires_at_ms = excluded.expires_at_ms
    `).run(doc);
  }

  public count(nowMs: number): number {
    const row = this.db.prepare('SELECT COUNT(*) as count FROM material_cache WHERE expires_at_ms > ?').get(nowMs) as { count: number };
    return row?.count ?? 0;
  }

  public get(docRef: string, nowMs: number): CachedMaterialDoc | null {
    const row = this.db.prepare(
      'SELECT * FROM material_cache WHERE doc_ref = ? AND expires_at_ms > ?'
    ).get(docRef, nowMs) as any;
    return row ? this.rowToDoc(row) : null;
  }

  public list(filters: { courseId?: string; courseWorkId?: string; limit: number; nowMs: number }): CachedMaterialDoc[] {
    const where: string[] = ['expires_at_ms > ?'];
    const params: any[] = [filters.nowMs];
    if (filters.courseId) {
      where.push('course_id = ?');
      params.push(filters.courseId);
    }
    if (filters.courseWorkId) {
      where.push('course_work_id = ?');
      params.push(filters.courseWorkId);
    }
    params.push(filters.limit);
    const rows = this.db.prepare(
      `SELECT * FROM material_cache WHERE ${where.join(' AND ')} ORDER BY extracted_at DESC LIMIT ?`
    ).all(...params) as any[];
    return rows.map((row) => this.rowToDoc(row));
  }

  public search(filters: {
    queryNormalized: string;
    courseId?: string;
    courseWorkId?: string;
    limit: number;
    nowMs: number;
  }): CachedMaterialDoc[] {
    const where: string[] = ['expires_at_ms > ?', 'text_normalized LIKE ?'];
    const params: any[] = [filters.nowMs, `%${filters.queryNormalized}%`];
    if (filters.courseId) {
      where.push('course_id = ?');
      params.push(filters.courseId);
    }
    if (filters.courseWorkId) {
      where.push('course_work_id = ?');
      params.push(filters.courseWorkId);
    }
    params.push(Math.max(filters.limit * 5, 25));
    const rows = this.db.prepare(
      `SELECT * FROM material_cache WHERE ${where.join(' AND ')} ORDER BY extracted_at DESC LIMIT ?`
    ).all(...params) as any[];
    return rows.map((row) => this.rowToDoc(row));
  }

  private rowToDoc(row: any): CachedMaterialDoc {
    return {
      docRef: row.doc_ref,
      courseId: row.course_id,
      courseWorkId: row.course_work_id,
      assignmentTitle: row.assignment_title ?? null,
      materialIndex: Number(row.material_index),
      type: row.type,
      title: row.title ?? null,
      url: row.url ?? null,
      driveFileId: row.drive_file_id ?? undefined,
      mimeType: row.mime_type ?? null,
      text: row.text,
      textNormalized: row.text_normalized,
      extractedAt: row.extracted_at,
      expiresAtMs: Number(row.expires_at_ms),
    };
  }
}

function sanitizeEntityTypes(values?: EntityType[]): EntityType[] {
  if (!values || values.length === 0) {
    return ['assignments', 'courses', 'announcements'];
  }

  const allowed: EntityType[] = ['assignments', 'courses', 'announcements'];
  const filtered = values.filter((value): value is EntityType => allowed.includes(value));
  return filtered.length > 0 ? [...new Set(filtered)] : ['assignments', 'courses', 'announcements'];
}

function parseDateFilter(value: string | undefined, name: string): number | null {
  if (!value) return null;
  const timestamp = Date.parse(value);
  if (Number.isNaN(timestamp)) {
    throw new SearchValidationError('INVALID_FILTER', `${name} must be a valid date string`);
  }
  return timestamp;
}

function normalizeText(value: string): string {
  return value
    .toLowerCase()
    .normalize('NFKD')
    .replace(/\s+/g, ' ')
    .trim();
}

function tokenizeForSuggest(value: string): string[] {
  return normalizeText(value)
    .split(/[^\p{L}\p{N}_'-]+/u)
    .map((token) => token.trim())
    .filter((token) => token.length >= 2);
}

function buildSnippet(value: string, normalizedQuery: string): string | undefined {
  const text = value.trim();
  if (!text) return undefined;
  if (!normalizedQuery) {
    return text.length <= 180 ? text : `${text.slice(0, 177)}...`;
  }

  const normalizedText = normalizeText(text);
  const index = normalizedText.indexOf(normalizedQuery);
  if (index < 0) {
    return text.length <= 180 ? text : `${text.slice(0, 177)}...`;
  }

  const start = Math.max(0, index - 60);
  const end = Math.min(text.length, index + 120);
  const snippet = text.slice(start, end).trim();
  return `${start > 0 ? '...' : ''}${snippet}${end < text.length ? '...' : ''}`;
}

function hashString(value: string): string {
  return createHash('sha1').update(value).digest('hex');
}

function encodeCursor(cursor: CursorPayload): string {
  return Buffer.from(JSON.stringify(cursor), 'utf8').toString('base64url');
}

function decodeCursor(cursor: string): CursorPayload {
  try {
    const decoded = JSON.parse(Buffer.from(cursor, 'base64url').toString('utf8')) as CursorPayload;
    if (
      typeof decoded.offset !== 'number' ||
      typeof decoded.queryHash !== 'string' ||
      typeof decoded.filtersHash !== 'string' ||
      typeof decoded.indexVersion !== 'number'
    ) {
      throw new Error('Invalid cursor payload structure');
    }
    return decoded;
  } catch {
    throw new SearchValidationError('INVALID_CURSOR', 'Cursor value is malformed');
  }
}

function clampNumber(value: number | undefined, fallback: number, min: number, max: number): number {
  if (typeof value !== 'number' || Number.isNaN(value)) return fallback;
  return Math.max(min, Math.min(max, Math.floor(value)));
}

function buildFacets(docs: SearchDoc[]): {
  entityTypes: Record<string, number>;
  states: Record<string, number>;
  courses: Array<{ courseId: string; courseName?: string; count: number }>;
} {
  const entityTypes: Record<string, number> = {};
  const states: Record<string, number> = {};
  const coursesMap = new Map<string, { courseId: string; courseName?: string; count: number }>();

  for (const doc of docs) {
    entityTypes[doc.entityType] = (entityTypes[doc.entityType] || 0) + 1;
    if (doc.state) {
      states[doc.state] = (states[doc.state] || 0) + 1;
    }
    if (doc.courseId) {
      const current = coursesMap.get(doc.courseId);
      if (current) {
        current.count += 1;
      } else {
        coursesMap.set(doc.courseId, {
          courseId: doc.courseId,
          courseName: doc.courseName,
          count: 1,
        });
      }
    }
  }

  const courses = [...coursesMap.values()]
    .sort((a, b) => {
      if (b.count !== a.count) return b.count - a.count;
      return a.courseId.localeCompare(b.courseId);
    })
    .slice(0, 25);

  return { entityTypes, states, courses };
}

// Default OAuth credentials — injected at publish time via inject-credentials.js
const DEFAULT_CLIENT_ID = '__GOOGLE_CLIENT_ID__';
const DEFAULT_CLIENT_SECRET = '__GOOGLE_CLIENT_SECRET__';
const SEARCH_INDEX_TTL_MS = 5 * 60 * 1000;
const MATERIAL_CACHE_TTL_MS = 6 * 60 * 60 * 1000;
const MATERIAL_CACHE_MAX_TEXT_CHARS = 1_000_000;
const DEFAULT_MATERIAL_CACHE_DB = 'material-cache.sqlite';

class GoogleClassroomMCPServer {
  private server: Server;
  private auth: OAuth2Client | null = null;
  private classroom: any = null;
  private drive: any = null;
  private forms: any = null;
  private readonly materialCacheStore: MaterialCacheStore;
  private searchService: SearchService;
  private readonly ocrProcessor: OCRProcessor;

  constructor() {
    this.server = new Server(
      {
        name: 'google-classroom-mcp-server',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.searchService = new SearchService(SEARCH_INDEX_TTL_MS, async () => this.buildSearchSnapshot());
    const resolvedAuthStorePath = resolveAuthStorePath(process.env.GOOGLE_AUTH_STORE);
    const materialDbFallbackPath = path.join(path.dirname(resolvedAuthStorePath), DEFAULT_MATERIAL_CACHE_DB);
    const materialDbPath = process.env.GOOGLE_MATERIAL_CACHE_DB || materialDbFallbackPath;
    this.materialCacheStore = new MaterialCacheStore(materialDbPath);
    this.ocrProcessor = new OCRProcessor();

    this.setupToolHandlers();
    this.setupAuth();
  }

  private async setupAuth() {
    try {
      const clientId = process.env.GOOGLE_CLIENT_ID || DEFAULT_CLIENT_ID;
      const clientSecret = process.env.GOOGLE_CLIENT_SECRET || DEFAULT_CLIENT_SECRET;
      const redirectUri = process.env.GOOGLE_REDIRECT_URI || 'urn:ietf:wg:oauth:2.0:oob';

      // 1) Try plain refresh token from environment (legacy/simple mode)
      if (process.env.GOOGLE_REFRESH_TOKEN) {
        this.auth = new google.auth.OAuth2(clientId, clientSecret, redirectUri);
        this.auth.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
        this.classroom = google.classroom({ version: 'v1', auth: this.auth });
        this.drive = google.drive({ version: 'v3', auth: this.auth });
        this.forms = google.forms({ version: 'v1', auth: this.auth });
        console.error('Authenticated via environment variables');
        return;
      }

      // 2) Try secure token store (recommended)
      const authStorePath = resolveAuthStorePath(process.env.GOOGLE_AUTH_STORE);
      if (fs.existsSync(authStorePath)) {
        const secureAuth = loadSecureAuthStore(authStorePath);
        if (secureAuth) {
          this.auth = new google.auth.OAuth2(
            secureAuth.clientId,
            secureAuth.clientSecret,
            secureAuth.redirectUri
          );
          this.auth.setCredentials({ refresh_token: secureAuth.refreshToken });
          this.classroom = google.classroom({ version: 'v1', auth: this.auth });
          this.drive = google.drive({ version: 'v3', auth: this.auth });
          this.forms = google.forms({ version: 'v1', auth: this.auth });
          console.error(`Authenticated via secure auth store: ${authStorePath}`);
          return;
        }
      }

      // 3) Fallback to legacy tokens.json file (backward compatibility)
      const tokensPath = path.join(process.cwd(), 'tokens.json');
      const credentialsPath = path.join(process.cwd(), 'credentials.json');
      if (fs.existsSync(tokensPath) && fs.existsSync(credentialsPath)) {
        const credentials = JSON.parse(fs.readFileSync(credentialsPath, 'utf8'));
        const tokens = JSON.parse(fs.readFileSync(tokensPath, 'utf8'));
        const { client_id, client_secret, redirect_uris } = credentials.web || credentials.installed;
        this.auth = new google.auth.OAuth2(
          client_id,
          client_secret,
          redirect_uris[0] || 'urn:ietf:wg:oauth:2.0:oob'
        );
        this.auth.setCredentials(tokens);
        this.classroom = google.classroom({ version: 'v1', auth: this.auth });
        this.drive = google.drive({ version: 'v3', auth: this.auth });
        this.forms = google.forms({ version: 'v1', auth: this.auth });
        console.error('Authenticated via tokens.json (legacy mode)');
        return;
      }

      console.error('No authentication found. Please run: npm run setup-auth');
      console.error('Expected one of: GOOGLE_REFRESH_TOKEN, secure auth store, or tokens.json + credentials.json');
    } catch (error) {
      console.error('Authentication setup failed:', error);
    }
  }

  private setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      return {
        tools: [
          {
            name: 'search',
            description: 'Unified AI-first search across assignments, courses, and announcements with ranking, facets, and cursor pagination',
            inputSchema: {
              type: 'object',
              properties: {
                query: { type: 'string', description: 'Search query text' },
                entityTypes: {
                  type: 'array',
                  items: { type: 'string', enum: ['assignments', 'courses', 'announcements'] },
                  description: 'Entities to search. Defaults to all.',
                },
                courseIds: { type: 'array', items: { type: 'string' }, description: 'Limit to specific course IDs' },
                states: { type: 'array', items: { type: 'string' }, description: 'Filter by entity state' },
                dueFrom: { type: 'string', description: 'Due date lower bound (ISO date/time)' },
                dueTo: { type: 'string', description: 'Due date upper bound (ISO date/time)' },
                missingOnly: { type: 'boolean', description: 'Assignments only: include only missing assignments' },
                gradedOnly: { type: 'boolean', description: 'Assignments only: include only graded assignments' },
                limit: { type: 'number', description: 'Page size (1..100), default 20' },
                cursor: { type: 'string', description: 'Opaque cursor from previous response' },
                sort: { type: 'string', enum: ['relevance', 'dueDate', 'updatedAt'], description: 'Result sort order' },
                forceRefresh: { type: 'boolean', description: 'Force refresh of in-memory index' },
              },
            },
          },
          {
            name: 'suggest_search_terms',
            description: 'Autocomplete and query expansion suggestions for AI search flows',
            inputSchema: {
              type: 'object',
              properties: {
                prefix: { type: 'string', description: 'Prefix for suggestions' },
                entityTypes: {
                  type: 'array',
                  items: { type: 'string', enum: ['assignments', 'courses', 'announcements'] },
                  description: 'Entities to suggest from. Defaults to all.',
                },
                limit: { type: 'number', description: 'Maximum number of suggestions (1..25), default 10' },
                forceRefresh: { type: 'boolean', description: 'Force refresh of in-memory index' },
              },
              required: ['prefix'],
            },
          },
          // Legacy tool names for backward compatibility
          {
            name: 'courses',
            description: 'Deprecated: use search. Get a list of all your Google Classroom courses (legacy)',
            inputSchema: {
              type: 'object',
              properties: {
                courseStates: {
                  type: 'array',
                  items: { type: 'string' },
                  description: 'Filter by course states (ACTIVE, ARCHIVED, PROVISIONED, DECLINED, SUSPENDED)',
                },
              },
            },
          },
          {
            name: 'course-details',
            description: 'Deprecated: use search. Get detailed information about a specific course including announcements (legacy)',
            inputSchema: {
              type: 'object',
              properties: {
                courseId: { type: 'string', description: 'The ID of the course to retrieve' },
              },
              required: ['courseId'],
            },
          },
          {
            name: 'assignments',
            description: 'Deprecated: use search. Get assignments and coursework for a specific course (legacy)',
            inputSchema: {
              type: 'object',
              properties: {
                courseId: { type: 'string', description: 'The ID of the course' },
              },
              required: ['courseId'],
            },
          },
          {
            name: 'list_courses',
            description: 'Deprecated: use search. List all courses with advanced filtering options',
            inputSchema: {
              type: 'object',
              properties: {
                courseStates: {
                  type: 'array',
                  items: { type: 'string' },
                  description: 'Filter by course states (ACTIVE, ARCHIVED, PROVISIONED, DECLINED, SUSPENDED)',
                },
                teacherId: { type: 'string', description: 'Filter courses by teacher ID' },
                studentId: { type: 'string', description: 'Filter courses by student ID' },
              },
            },
          },
          {
            name: 'get_course',
            description: 'Deprecated: use search. Get detailed information about a specific course',
            inputSchema: {
              type: 'object',
              properties: {
                courseId: { type: 'string', description: 'The ID of the course to retrieve' },
              },
              required: ['courseId'],
            },
          },
          {
            name: 'list_coursework',
            description: 'Deprecated: use search. List assignments and coursework for a course',
            inputSchema: {
              type: 'object',
              properties: {
                courseId: { type: 'string', description: 'The ID of the course' },
                courseWorkStates: {
                  type: 'array',
                  items: { type: 'string' },
                  description: 'Filter by coursework states (PUBLISHED, DRAFT, DELETED)',
                },
              },
              required: ['courseId'],
            },
          },
          {
            name: 'get_coursework',
            description: 'Deprecated: use search. Get detailed information about a specific assignment',
            inputSchema: {
              type: 'object',
              properties: {
                courseId: { type: 'string', description: 'The ID of the course' },
                courseWorkId: { type: 'string', description: 'The ID of the coursework/assignment' },
              },
              required: ['courseId', 'courseWorkId'],
            },
          },
          {
            name: 'list_submissions',
            description: 'Deprecated: use search. View your own submissions for an assignment, including state, grade, and attachments',
            inputSchema: {
              type: 'object',
              properties: {
                courseId: { type: 'string', description: 'The ID of the course' },
                courseWorkId: { type: 'string', description: 'The ID of the coursework/assignment' },
              },
              required: ['courseId', 'courseWorkId'],
            },
          },
          {
            name: 'list_announcements',
            description: 'Deprecated: use search. List announcements for a course',
            inputSchema: {
              type: 'object',
              properties: {
                courseId: { type: 'string', description: 'The ID of the course' },
                announcementStates: {
                  type: 'array',
                  items: { type: 'string' },
                  description: 'Filter by announcement states (PUBLISHED, DRAFT, DELETED)',
                },
              },
              required: ['courseId'],
            },
          },
          {
            name: 'get_upcoming_assignments',
            description: 'Deprecated: use search. Get upcoming assignments due within the next N days across all active courses. Defaults to 7 days.',
            inputSchema: {
              type: 'object',
              properties: {
                days: { type: 'number', description: 'Number of days to look ahead. Defaults to 7.' },
              },
            },
          },
          {
            name: 'get_missing_assignments',
            description: 'Deprecated: use search. Get all past-due assignments that have not been submitted across all active courses',
            inputSchema: {
              type: 'object',
              properties: {},
            },
          },
          {
            name: 'get_assignments',
            description: 'Deprecated: use search. Get all published assignments for a specific course, formatted for easy reading',
            inputSchema: {
              type: 'object',
              properties: {
                courseId: { type: 'string', description: 'The ID of the course' },
              },
              required: ['courseId'],
            },
          },
          {
            name: 'calculate_grade',
            description: 'Deprecated: use search. Calculate your overall grade percentage for a course based on all graded assignments',
            inputSchema: {
              type: 'object',
              properties: {
                courseId: { type: 'string', description: 'The ID of the course' },
              },
              required: ['courseId'],
            },
          },
          {
            name: 'get_assignment_materials',
            description: 'Deprecated: use search. Get all materials and attachments for a specific assignment (Drive files, links, YouTube videos, forms)',
            inputSchema: {
              type: 'object',
              properties: {
                courseId: { type: 'string', description: 'The ID of the course' },
                courseWorkId: { type: 'string', description: 'The ID of the assignment' },
              },
              required: ['courseId', 'courseWorkId'],
            },
          },
          {
            name: 'get_assignment_material_text',
            description: 'Extract text from assignment materials and Drive attachments when possible (Google Docs/Sheets/Slides + text-like files)',
            inputSchema: {
              type: 'object',
              properties: {
                courseId: { type: 'string', description: 'The ID of the course' },
                courseWorkId: { type: 'string', description: 'The ID of the assignment' },
                materialIndex: { type: 'number', description: 'Optional material index to extract a single item only (0-based)' },
                maxCharsPerMaterial: { type: 'number', description: 'Max extracted characters per material (default 12000, max 50000)' },
                previewChars: { type: 'number', description: 'Preview size returned inline when includeText=false (default 700, max 5000)' },
                includeText: { type: 'boolean', description: 'When true, include extracted text inline (truncated by maxCharsPerMaterial). Defaults to false.' },
              },
              required: ['courseId', 'courseWorkId'],
            },
          },
          {
            name: 'search_material_cache',
            description: 'Search inside previously extracted material text cache',
            inputSchema: {
              type: 'object',
              properties: {
                query: { type: 'string', description: 'Text query to find in cached materials' },
                courseId: { type: 'string', description: 'Optional filter by course ID' },
                courseWorkId: { type: 'string', description: 'Optional filter by assignment ID' },
                limit: { type: 'number', description: 'Maximum results (1..50), default 10' },
                snippetChars: { type: 'number', description: 'Snippet size per hit (100..2000), default 500' },
              },
              required: ['query'],
            },
          },
          {
            name: 'read_material_cache',
            description: 'Read cached material text in chunks by docRef',
            inputSchema: {
              type: 'object',
              properties: {
                docRef: { type: 'string', description: 'Reference returned by get_assignment_material_text/search_material_cache' },
                offset: { type: 'number', description: 'Character offset (0-based), default 0' },
                maxChars: { type: 'number', description: 'Chunk size (200..20000), default 3000' },
              },
              required: ['docRef'],
            },
          },
{
            name: 'list_material_cache',
            description: 'List cached material documents with metadata',
            inputSchema: {
              type: 'object',
              properties: {
                courseId: { type: 'string', description: 'Optional filter by course ID' },
                courseWorkId: { type: 'string', description: 'Optional filter by assignment ID' },
                limit: { type: 'number', description: 'Maximum docs (1..200), default 50' },
              },
            },
          },
          {
            name: 'clear_material_cache',
            description: 'Delete all cached material documents. Useful for forcing re-extraction with updated OCR or forcing cache refresh.',
            inputSchema: { type: 'object', properties: {} },
          },
          {
            name: 'get_grades',
            description: 'Deprecated: use search. Get your grades across all active courses, showing assigned grade, max points, and submission state for each assignment',
            inputSchema: { type: 'object', properties: {} },
          },
          {
            name: 'get_dashboard',
            description: 'Get a compact dashboard: active courses, upcoming assignments, missing assignments, and grade summary',
            inputSchema: {
              type: 'object',
              properties: {
                days: { type: 'number', description: 'How many days ahead to use for upcoming assignments. Defaults to 7.' },
                upcomingLimit: { type: 'number', description: 'Maximum number of upcoming assignments to include. Defaults to 10.' },
                missingLimit: { type: 'number', description: 'Maximum number of missing assignments to include. Defaults to 10.' },
              },
            },
          },
          {
            name: 'search_assignments',
            description: "Deprecated: use search. Legacy wrapper over unified search (entityTypes=['assignments'])",
            inputSchema: {
              type: 'object',
              properties: {
                query: {
                  type: 'string',
                  description: 'Keyword to match in assignment title or description. If omitted, returns by due-date filters only.',
                },
                courseId: { type: 'string', description: 'Optional course ID to limit search to one course.' },
                daysAhead: {
                  type: 'number',
                  description: 'Include only assignments due within the next N days. If omitted, no upper due-date limit is applied.',
                },
                includeNoDueDate: {
                  type: 'boolean',
                  description: 'Include assignments that do not have a due date. Defaults to false.',
                },
                limit: { type: 'number', description: 'Maximum number of results. Defaults to 50.' },
              },
            },
          },
        ],
      };
    });

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      console.error(`[MCP] Tool called: ${request.params.name}`);
      if (!this.classroom) {
        throw new McpError(
          ErrorCode.InternalError,
          'Google Classroom API not initialized. Please run: npm run setup-auth'
        );
      }

      try {
        switch (request.params.name) {
          case 'search':
            return await this.searchTool(request.params.arguments as SearchInput || {});
          case 'suggest_search_terms':
            return await this.suggestSearchTermsTool(request.params.arguments as SuggestInput);

          // Legacy tool compatibility
          case 'courses':
          case 'list_courses':
            return await this.listCourses(request.params.arguments || {});
          case 'course-details':
          case 'get_course':
            return await this.getCourse(request.params.arguments as { courseId: string });
          case 'assignments':
          case 'list_coursework':
            return await this.listCoursework(request.params.arguments as { courseId: string; courseWorkStates?: string[] });
          case 'get_coursework':
            return await this.getCoursework(request.params.arguments as { courseId: string; courseWorkId: string });
          case 'list_submissions':
            return await this.listSubmissions(request.params.arguments as { courseId: string; courseWorkId: string });
          case 'list_announcements':
            return await this.listAnnouncements(request.params.arguments as { courseId: string; announcementStates?: string[] });
          case 'get_upcoming_assignments':
            return await this.getUpcomingAssignments(request.params.arguments as { days?: number } || {});
          case 'get_missing_assignments':
            return await this.getMissingAssignments();
          case 'get_assignments':
            return await this.getAssignments(request.params.arguments as { courseId: string });
          case 'calculate_grade':
            return await this.calculateGrade(request.params.arguments as { courseId: string });
          case 'get_assignment_materials':
            return await this.getAssignmentMaterials(request.params.arguments as { courseId: string; courseWorkId: string });
case 'get_assignment_material_text': {
            const args = request.params.arguments as {
              courseId: string;
              courseWorkId: string;
              materialIndex?: number;
              maxCharsPerMaterial?: number;
              previewChars?: number;
              includeText?: boolean;
            };
            console.error(`[MCP] get_assignment_material_text: courseId=${args.courseId}, courseWorkId=${args.courseWorkId}, includeText=${args.includeText}`);
            return await this.getAssignmentMaterialText(args);
          }
          case 'search_material_cache':
            return await this.searchMaterialCache(
              request.params.arguments as {
                query: string;
                courseId?: string;
                courseWorkId?: string;
                limit?: number;
                snippetChars?: number;
              }
            );
          case 'read_material_cache':
            return await this.readMaterialCache(
              request.params.arguments as {
                docRef: string;
                offset?: number;
                maxChars?: number;
              }
            );
case 'list_material_cache':
            return await this.listMaterialCache(
              request.params.arguments as {
                courseId?: string;
                courseWorkId?: string;
                limit?: number;
              } || {}
            );
          case 'clear_material_cache':
            return await this.clearMaterialCache();
          case 'get_grades':
            return await this.getGrades();
          case 'get_dashboard':
            return await this.getDashboard(
              request.params.arguments as { days?: number; upcomingLimit?: number; missingLimit?: number } || {}
            );
          case 'search_assignments':
            return await this.searchAssignments(
              request.params.arguments as {
                query?: string;
                courseId?: string;
                daysAhead?: number;
                includeNoDueDate?: boolean;
                limit?: number;
              } || {}
            );
          default:
            throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${request.params.name}`);
        }
      console.error(`[MCP] Tool ${request.params.name} completed successfully`);
      } catch (error) {
        if (error instanceof SearchValidationError) {
          throw new McpError(ErrorCode.InvalidParams, `${error.code}: ${error.message}`);
        }
        console.error(`[MCP] Error executing tool ${request.params.name}:`, error);
        throw new McpError(ErrorCode.InternalError, `Tool execution failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      }
    });
  }

  private async buildSearchSnapshot(): Promise<RefreshSnapshot> {
    if (!this.classroom) {
      throw new Error('Google Classroom API not initialized');
    }

    const docs: SearchDoc[] = [];
    const warnings: string[] = [];
    let sourceFailures = 0;

    const coursesResponse = await this.classroom.courses.list({ courseStates: ['ACTIVE'] });
    const courses: any[] = coursesResponse.data.courses || [];
    const courseNameById = new Map<string, string>();

    if (courses.length === 0) {
      warnings.push('No active courses found while refreshing search index.');
    }

    for (const course of courses) {
      courseNameById.set(course.id, course.name || '');
      docs.push({
        key: `course:${course.id}`,
        entityType: 'courses',
        id: course.id,
        courseId: course.id,
        courseName: course.name || '',
        title: course.name || '',
        body: `${course.section || ''} ${course.descriptionHeading || ''} ${course.description || ''}`.trim(),
        titleNormalized: normalizeText(course.name || ''),
        bodyNormalized: normalizeText(`${course.section || ''} ${course.descriptionHeading || ''} ${course.description || ''}`),
        state: course.courseState || undefined,
        updatedAt: course.updateTime || course.creationTime || undefined,
        updatedTimestamp: course.updateTime ? Date.parse(course.updateTime) : course.creationTime ? Date.parse(course.creationTime) : undefined,
        alternateLink: course.alternateLink || undefined,
      });
    }

    let assignmentSuccessCount = 0;
    let announcementSuccessCount = 0;

    await Promise.all(
      courses.map(async (course) => {
        const [courseWorkResult, submissionsResult, announcementsResult] = await Promise.allSettled([
          this.classroom.courses.courseWork.list({
            courseId: course.id,
            courseWorkStates: ['PUBLISHED'],
          }),
          this.classroom.courses.courseWork.studentSubmissions.list({
            courseId: course.id,
            courseWorkId: '-',
            userId: 'me',
          }),
          this.classroom.courses.announcements.list({
            courseId: course.id,
            announcementStates: ['PUBLISHED'],
          }),
        ]);

        const submissionsMap: Record<string, any> = {};
        if (submissionsResult.status === 'fulfilled') {
          for (const sub of submissionsResult.value.data.studentSubmissions || []) {
            submissionsMap[sub.courseWorkId] = sub;
          }
        } else {
          warnings.push(`Submissions unavailable for course ${course.id}: ${String(submissionsResult.reason)}`);
        }

        if (courseWorkResult.status === 'fulfilled') {
          assignmentSuccessCount += 1;
          const courseWorks: any[] = courseWorkResult.value.data.courseWork || [];
          for (const cw of courseWorks) {
            const dueTimestamp = this.getDueTimestamp(cw.dueDate);
            const state = submissionsMap[cw.id]?.state ?? 'NOT_STARTED';
            const isMissing =
              dueTimestamp != null &&
              dueTimestamp < Date.now() &&
              state !== 'TURNED_IN' &&
              state !== 'RETURNED';
            docs.push({
              key: `assignment:${course.id}:${cw.id}`,
              entityType: 'assignments',
              id: cw.id,
              courseId: course.id,
              courseName: course.name || '',
              title: cw.title || '',
              body: cw.description || '',
              titleNormalized: normalizeText(cw.title || ''),
              bodyNormalized: normalizeText(cw.description || ''),
              state,
              dueDate: this.formatDueDate(cw.dueDate) || undefined,
              dueTimestamp: dueTimestamp ?? undefined,
              updatedAt: cw.updateTime || cw.creationTime || undefined,
              updatedTimestamp: cw.updateTime ? Date.parse(cw.updateTime) : cw.creationTime ? Date.parse(cw.creationTime) : undefined,
              alternateLink: cw.alternateLink || undefined,
              assignedGrade: submissionsMap[cw.id]?.assignedGrade ?? undefined,
              maxPoints: cw.maxPoints ?? undefined,
              isMissing,
            });
          }
        } else {
          warnings.push(`Coursework unavailable for course ${course.id}: ${String(courseWorkResult.reason)}`);
        }

        if (announcementsResult.status === 'fulfilled') {
          announcementSuccessCount += 1;
          const announcements: any[] = announcementsResult.value.data.announcements || [];
          for (const announcement of announcements) {
            const text = announcement.text || '';
            docs.push({
              key: `announcement:${course.id}:${announcement.id}`,
              entityType: 'announcements',
              id: announcement.id,
              courseId: course.id,
              courseName: courseNameById.get(course.id) || course.name || '',
              title: text.length > 80 ? `${text.slice(0, 77)}...` : text || 'Announcement',
              body: text,
              titleNormalized: normalizeText(text.length > 80 ? `${text.slice(0, 77)}...` : text || 'Announcement'),
              bodyNormalized: normalizeText(text),
              state: announcement.announcementState || 'PUBLISHED',
              updatedAt: announcement.updateTime || announcement.creationTime || undefined,
              updatedTimestamp: announcement.updateTime
                ? Date.parse(announcement.updateTime)
                : announcement.creationTime
                ? Date.parse(announcement.creationTime)
                : undefined,
              alternateLink: announcement.alternateLink || undefined,
            });
          }
        } else {
          warnings.push(`Announcements unavailable for course ${course.id}: ${String(announcementsResult.reason)}`);
        }
      })
    );

    if (assignmentSuccessCount === 0) {
      sourceFailures += 1;
      warnings.push('Assignments source failed for all courses.');
    }
    if (announcementSuccessCount === 0) {
      sourceFailures += 1;
      warnings.push('Announcements source failed for all courses.');
    }
    if (coursesResponse == null) {
      sourceFailures += 1;
      warnings.push('Courses source failed.');
    }

    return { docs, warnings, sourceFailures };
  }

  private formatDueDate(dueDate: any): string | null {
    if (!dueDate?.year || !dueDate?.month || !dueDate?.day) return null;
    return `${dueDate.year}-${String(dueDate.month).padStart(2, '0')}-${String(dueDate.day).padStart(2, '0')}`;
  }

  private getDueTimestamp(dueDate: any): number | null {
    if (!dueDate?.year || !dueDate?.month || !dueDate?.day) return null;
    return new Date(dueDate.year, dueDate.month - 1, dueDate.day).getTime();
  }

  private async searchTool(args: SearchInput): Promise<{ content: Array<{ type: 'text'; text: string }> }> {
    const result = await this.searchService.search(args);
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }

  private async suggestSearchTermsTool(args: SuggestInput): Promise<{ content: Array<{ type: 'text'; text: string }> }> {
    const result = await this.searchService.suggest(args);
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }

  private async listCourses(args: any) {
    const response = await this.classroom.courses.list({
      courseStates: args.courseStates,
      teacherId: args.teacherId,
      studentId: args.studentId,
    });
    return { content: [{ type: 'text', text: JSON.stringify(response.data, null, 2) }] };
  }

  private async getCourse(args: { courseId: string }) {
    const response = await this.classroom.courses.get({ id: args.courseId });
    return { content: [{ type: 'text', text: JSON.stringify(response.data, null, 2) }] };
  }

  private async listCoursework(args: { courseId: string; courseWorkStates?: string[] }) {
    const response = await this.classroom.courses.courseWork.list({
      courseId: args.courseId,
      courseWorkStates: args.courseWorkStates,
    });
    return { content: [{ type: 'text', text: JSON.stringify(response.data, null, 2) }] };
  }

  private async getCoursework(args: { courseId: string; courseWorkId: string }) {
    const response = await this.classroom.courses.courseWork.get({
      courseId: args.courseId,
      id: args.courseWorkId,
    });
    return { content: [{ type: 'text', text: JSON.stringify(response.data, null, 2) }] };
  }

  private async listSubmissions(args: { courseId: string; courseWorkId: string }) {
    const response = await this.classroom.courses.courseWork.studentSubmissions.list({
      courseId: args.courseId,
      courseWorkId: args.courseWorkId,
      userId: 'me',
    });
    return { content: [{ type: 'text', text: JSON.stringify(response.data, null, 2) }] };
  }

  private async listAnnouncements(args: { courseId: string; announcementStates?: string[] }) {
    const response = await this.classroom.courses.announcements.list({
      courseId: args.courseId,
      announcementStates: args.announcementStates,
    });
    return { content: [{ type: 'text', text: JSON.stringify(response.data, null, 2) }] };
  }

  private async getUpcomingAssignments(args: { days?: number } = {}) {
    const days = args.days ?? 7;
    const now = new Date();
    now.setHours(0, 0, 0, 0);
    const dueFrom = now.toISOString();
    const dueTo = new Date(now.getTime() + days * 24 * 60 * 60 * 1000).toISOString();
    const result = await this.searchService.search({
      entityTypes: ['assignments'],
      dueFrom,
      dueTo,
      sort: 'dueDate',
      limit: 100,
    });

    const assignments = result.items.map((item) => ({
      courseId: item.courseId,
      courseName: item.courseName,
      assignmentId: item.id,
      title: item.title,
      dueDate: item.dueDate ?? null,
      maxPoints: item.maxPoints ?? null,
      alternateLink: item.alternateLink ?? null,
    }));

    return { content: [{ type: 'text', text: JSON.stringify(assignments, null, 2) }] };
  }

  private async getMissingAssignments() {
    const result = await this.searchService.search({
      entityTypes: ['assignments'],
      missingOnly: true,
      sort: 'dueDate',
      limit: 200,
    });

    const missing = result.items.map((item) => ({
      courseId: item.courseId,
      courseName: item.courseName,
      assignmentId: item.id,
      title: item.title,
      dueDate: item.dueDate ?? null,
      maxPoints: item.maxPoints ?? null,
      submissionState: item.state ?? 'NOT_STARTED',
      alternateLink: item.alternateLink ?? null,
    }));

    return { content: [{ type: 'text', text: JSON.stringify(missing, null, 2) }] };
  }

  private async getAssignments(args: { courseId: string }) {
    const result = await this.searchService.search({
      entityTypes: ['assignments'],
      courseIds: [args.courseId],
      sort: 'dueDate',
      limit: 200,
    });

    const assignments = result.items.map((item) => ({
      id: item.id,
      title: item.title,
      description: item.snippet ?? null,
      maxPoints: item.maxPoints ?? null,
      dueDate: item.dueDate ?? null,
      alternateLink: item.alternateLink ?? null,
    }));

    return { content: [{ type: 'text', text: JSON.stringify(assignments, null, 2) }] };
  }

  private async getGradesData() {
    const coursesResponse = await this.classroom.courses.list({ courseStates: ['ACTIVE'] });
    const courses: any[] = coursesResponse.data.courses || [];
    const perCourse = await Promise.all(
      courses.map(async (course: any) => {
        try {
          const [cwResponse, subResponse] = await Promise.all([
            this.classroom.courses.courseWork.list({
              courseId: course.id,
              courseWorkStates: ['PUBLISHED'],
            }),
            this.classroom.courses.courseWork.studentSubmissions.list({
              courseId: course.id,
              courseWorkId: '-',
              userId: 'me',
            }),
          ]);

          const courseworkMap: Record<string, any> = {};
          for (const cw of cwResponse.data.courseWork || []) {
            courseworkMap[cw.id] = cw;
          }

          return (subResponse.data.studentSubmissions || []).map((sub: any) => {
            const cw = courseworkMap[sub.courseWorkId] || {};
            return {
              courseId: course.id,
              courseName: course.name,
              assignmentId: sub.courseWorkId,
              title: cw.title ?? null,
              state: sub.state,
              assignedGrade: sub.assignedGrade ?? null,
              maxPoints: cw.maxPoints ?? null,
              dueDate: this.formatDueDate(cw.dueDate),
              alternateLink: sub.alternateLink ?? null,
            };
          });
        } catch {
          return [];
        }
      })
    );
    return perCourse.flat();
  }

  private async getGrades() {
    const grades = await this.getGradesData();
    return { content: [{ type: 'text', text: JSON.stringify(grades, null, 2) }] };
  }

  private async calculateGrade(args: { courseId: string }) {
    const [cwResponse, subResponse] = await Promise.all([
      this.classroom.courses.courseWork.list({
        courseId: args.courseId,
        courseWorkStates: ['PUBLISHED'],
      }),
      this.classroom.courses.courseWork.studentSubmissions.list({
        courseId: args.courseId,
        courseWorkId: '-',
        userId: 'me',
      }),
    ]);

    const courseworkMap: Record<string, any> = {};
    for (const cw of cwResponse.data.courseWork || []) {
      courseworkMap[cw.id] = cw;
    }

    let totalEarned = 0;
    let totalPossible = 0;
    const breakdown: any[] = [];

    for (const sub of subResponse.data.studentSubmissions || []) {
      const cw = courseworkMap[sub.courseWorkId];
      if (!cw || cw.maxPoints == null || sub.assignedGrade == null) continue;
      totalEarned += sub.assignedGrade;
      totalPossible += cw.maxPoints;
      breakdown.push({
        assignmentId: sub.courseWorkId,
        title: cw.title ?? null,
        earned: sub.assignedGrade,
        possible: cw.maxPoints,
        percentage: cw.maxPoints > 0 ? Math.round((sub.assignedGrade / cw.maxPoints) * 1000) / 10 : null,
      });
    }

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          courseId: args.courseId,
          totalEarned,
          totalPossible,
          overallPercentage: totalPossible > 0 ? Math.round((totalEarned / totalPossible) * 1000) / 10 : null,
          gradedAssignments: breakdown.length,
          breakdown,
        }, null, 2),
      }],
    };
  }

  private async getAssignmentMaterials(args: { courseId: string; courseWorkId: string }) {
    const response = await this.classroom.courses.courseWork.get({
      courseId: args.courseId,
      id: args.courseWorkId,
    });

    const cw = response.data;
    const materials = (cw.materials || []).map((m: any) => {
      if (m.driveFile) {
        return {
          type: 'driveFile',
          title: m.driveFile.driveFile?.title ?? null,
          url: m.driveFile.driveFile?.alternateLink ?? null,
          shareMode: m.driveFile.shareMode ?? null,
        };
      } else if (m.youTubeVideo) {
        return {
          type: 'youTubeVideo',
          title: m.youTubeVideo.title ?? null,
          url: m.youTubeVideo.alternateLink ?? null,
          thumbnailUrl: m.youTubeVideo.thumbnailUrl ?? null,
        };
      } else if (m.link) {
        return {
          type: 'link',
          title: m.link.title ?? null,
          url: m.link.url ?? null,
        };
      } else if (m.form) {
        return {
          type: 'form',
          title: m.form.title ?? null,
          url: m.form.formUrl ?? null,
          responseUrl: m.form.responseUrl ?? null,
        };
      }
      return { type: 'unknown', raw: m };
    });

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          assignmentId: cw.id,
          title: cw.title,
          description: cw.description ?? null,
          materials,
          alternateLink: cw.alternateLink ?? null,
        }, null, 2),
      }],
    };
  }

private async getAssignmentMaterialText(args: {
    courseId: string;
    courseWorkId: string;
    materialIndex?: number;
    maxCharsPerMaterial?: number;
    previewChars?: number;
    includeText?: boolean;
  }) {
    console.error(`[MCP] getAssignmentMaterialText: Fetching assignment ${args.courseWorkId} from course ${args.courseId}`);
    this.cleanupMaterialCache();
    const response = await this.classroom.courses.courseWork.get({
      courseId: args.courseId,
      id: args.courseWorkId,
    });

    const cw = response.data;
    const sourceMaterials: any[] = cw.materials || [];
    console.error(`[MCP] getAssignmentMaterialText: Found ${sourceMaterials.length} material(s) in assignment "${cw.title}"`);
    const maxCharsPerMaterial = clampNumber(args.maxCharsPerMaterial, 12000, 1000, 50000);
    const previewChars = clampNumber(args.previewChars, 700, 150, 5000);
    const includeText = Boolean(args.includeText);
    const hasMaterialIndex = typeof args.materialIndex === 'number' && Number.isFinite(args.materialIndex);
    const materialIndex = hasMaterialIndex ? Math.floor(args.materialIndex as number) : null;
    if (materialIndex != null && (materialIndex < 0 || materialIndex >= sourceMaterials.length)) {
      throw new SearchValidationError(
        'INVALID_FILTER',
        `materialIndex must be between 0 and ${Math.max(0, sourceMaterials.length - 1)}`
      );
    }

    const materials = materialIndex != null
      ? [{ index: materialIndex, material: sourceMaterials[materialIndex] }]
      : sourceMaterials.map((material, index) => ({ index, material }));

    const extracted: any[] = [];
    const warnings: string[] = [];

    for (const item of materials) {
      const extractedItem = await this.extractMaterialText({
        courseId: args.courseId,
        courseWorkId: args.courseWorkId,
        assignmentTitle: cw.title ?? null,
        materialIndex: item.index,
        material: item.material,
        maxCharsPerMaterial,
        previewChars,
        includeText,
      });
      extracted.push(extractedItem);
      if (extractedItem.warning) warnings.push(extractedItem.warning);
    }

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          assignmentId: cw.id,
          title: cw.title,
          description: cw.description ?? null,
          totalMaterials: sourceMaterials.length,
          extractedCount: extracted.length,
          maxCharsPerMaterial,
          includeText,
          previewChars,
          materials: extracted,
          cacheStats: {
            cachedDocs: this.materialCacheStore.count(Date.now()),
          },
          warnings: warnings.length > 0 ? warnings : undefined,
          alternateLink: cw.alternateLink ?? null,
        }, null, 2),
      }],
    };
  }

  private async extractMaterialText(args: {
    courseId: string;
    courseWorkId: string;
    assignmentTitle?: string | null;
    materialIndex: number;
    material: any;
    maxCharsPerMaterial: number;
    previewChars: number;
    includeText: boolean;
  }) {
    const { courseId, courseWorkId, assignmentTitle, materialIndex, material, maxCharsPerMaterial, previewChars, includeText } = args;
    if (material?.driveFile?.driveFile?.id) {
      const driveFile = material.driveFile.driveFile;
      const extraction = await this.extractTextFromDriveFile(driveFile.id, driveFile.title || null);
      const docRef = this.buildMaterialDocRef(courseId, courseWorkId, materialIndex, driveFile.id);
      const { textForCache, textForInline, previewText, extractedChars, warning } = this.prepareMaterialTextForResponse(
        extraction.text ?? null,
        includeText,
        maxCharsPerMaterial,
        previewChars,
        extraction.warning ?? null
      );
      if (textForCache) {
        this.putMaterialCache({
          docRef,
          courseId,
          courseWorkId,
          assignmentTitle: assignmentTitle ?? null,
          materialIndex,
          type: 'driveFile',
          title: driveFile.title ?? null,
          url: driveFile.alternateLink ?? null,
          driveFileId: driveFile.id,
          mimeType: extraction.mimeType ?? null,
          text: textForCache,
          textNormalized: normalizeText(textForCache),
          extractedAt: new Date().toISOString(),
          expiresAtMs: Date.now() + MATERIAL_CACHE_TTL_MS,
        });
      }
      return {
        docRef,
        index: materialIndex,
        type: 'driveFile',
        title: driveFile.title ?? null,
        url: driveFile.alternateLink ?? null,
        driveFileId: driveFile.id,
        mimeType: extraction.mimeType ?? null,
        extractionStatus: extraction.status,
        extractedText: textForInline,
        previewText,
        extractedChars,
        warning,
      };
    }

    if (material?.link) {
      return {
        index: materialIndex,
        type: 'link',
        title: material.link.title ?? null,
        url: material.link.url ?? null,
        extractionStatus: 'skipped',
        extractedText: null,
        previewText: null,
        extractedChars: 0,
        warning: 'Link material does not include retrievable file content via Classroom API.',
      };
    }

    if (material?.youTubeVideo) {
      return {
        index: materialIndex,
        type: 'youTubeVideo',
        title: material.youTubeVideo.title ?? null,
        url: material.youTubeVideo.alternateLink ?? null,
        extractionStatus: 'skipped',
        extractedText: null,
        previewText: null,
        extractedChars: 0,
        warning: 'YouTube material text extraction is not supported in this tool.',
      };
    }

    if (material?.form) {
      const formTitle = material.form.title ?? null;
      const formUrl = material.form.formUrl ?? null;
      const formExtraction = await this.extractTextFromGoogleForm(formUrl, formTitle);
      const formDocRef = this.buildMaterialDocRef(
        courseId,
        courseWorkId,
        materialIndex,
        formExtraction.formId || `form-${materialIndex}`
      );
      const { textForCache, textForInline, previewText, extractedChars, warning } = this.prepareMaterialTextForResponse(
        formExtraction.text ?? null,
        includeText,
        maxCharsPerMaterial,
        previewChars,
        formExtraction.warning ?? null
      );
      if (textForCache) {
        this.putMaterialCache({
          docRef: formDocRef,
          courseId,
          courseWorkId,
          assignmentTitle: assignmentTitle ?? null,
          materialIndex,
          type: 'form',
          title: formTitle,
          url: formUrl,
          driveFileId: formExtraction.formId ?? undefined,
          mimeType: 'application/vnd.google-apps.form',
          text: textForCache,
          textNormalized: normalizeText(textForCache),
          extractedAt: new Date().toISOString(),
          expiresAtMs: Date.now() + MATERIAL_CACHE_TTL_MS,
        });
      }
      return {
        docRef: formDocRef,
        index: materialIndex,
        type: 'form',
        title: formTitle,
        url: formUrl,
        extractionStatus: formExtraction.status,
        extractedText: textForInline,
        previewText,
        extractedChars,
        warning,
      };
    }

    return {
      index: materialIndex,
      type: 'unknown',
      extractionStatus: 'skipped',
      extractedText: null,
      previewText: null,
      extractedChars: 0,
      warning: 'Unsupported material type.',
    };
  }

  private async extractTextFromDriveFile(fileId: string, fallbackTitle: string | null) {
    if (!this.drive) {
      return {
        status: 'error',
        warning: 'Google Drive API client is not initialized.',
        mimeType: null,
      };
    }

    try {
      const fileMeta = await this.drive.files.get({
        fileId,
        fields: 'id,name,mimeType,size,webViewLink',
        supportsAllDrives: true,
      });
      const mimeType = fileMeta.data.mimeType || null;
      const fileName = fileMeta.data.name || fallbackTitle || fileId;

      if (!mimeType) {
        return {
          status: 'skipped',
          warning: `Skipping ${fileName}: mimeType is unavailable.`,
          mimeType: null,
        };
      }

console.error(`[OCR] Starting extraction for: ${fileName} (${mimeType})`);
      let rawText: string | null = null;
      let ocrTexts: string[] = [];

      if (GOOGLE_APPS_EXPORT_MIME[mimeType]) {
        console.error(`[OCR] Exporting Google Docs file as text...`);
        if (mimeType === 'application/vnd.google-apps.document') {
          const exported = await this.drive.files.export(
            { fileId, mimeType: 'text/plain' },
            { responseType: 'arraybuffer' }
          );
          rawText = this.decodeBinaryText(exported.data);
          console.error(`[OCR] Text extracted: ${rawText?.length || 0} chars`);
          try {
            console.error(`[OCR] Looking for embedded images in HTML...`);
            const htmlExported = await this.drive.files.export(
              { fileId, mimeType: 'text/html' },
              { responseType: 'arraybuffer' }
            );
            const html = this.decodeBinaryText(htmlExported.data);
            const imageBuffers = await this.extractImagesFromGoogleDocHtml(html, fileId);
            if (imageBuffers.length > 0) {
              console.error(`[OCR] Found ${imageBuffers.length} embedded image(s), running OCR...`);
              ocrTexts = await this.ocrProcessor.recognizeImages(imageBuffers);
              const totalOcrChars = ocrTexts.join('').length;
              console.error(`[OCR] OCR completed: ${ocrTexts.length} image(s), ~${totalOcrChars} chars recognized`);
            } else {
              console.error(`[OCR] No embedded images found`);
            }
          } catch (e) {
            console.error(`[OCR] Failed to extract images: ${e}`);
          }
        } else {
          const exportMimeType = GOOGLE_APPS_EXPORT_MIME[mimeType];
          const exported = await this.drive.files.export(
            { fileId, mimeType: exportMimeType },
            { responseType: 'arraybuffer' }
          );
          rawText = this.decodeBinaryText(exported.data);
          console.error(`[OCR] Text extracted: ${rawText?.length || 0} chars`);
        }
      } else if (this.isImageMimeType(mimeType, fileName)) {
        console.error(`[OCR] Image file detected, running OCR...`);
        const downloaded = await this.drive.files.get(
          { fileId, alt: 'media', supportsAllDrives: true },
          { responseType: 'arraybuffer' }
        );
        const imgBuf = this.toBuffer(downloaded.data);
        console.error(`[OCR] Image loaded: ${imgBuf.length} bytes`);
        const ocrResult = await this.ocrProcessor.recognizeImage(imgBuf);
        rawText = ocrResult || null;
        if (rawText) {
          console.error(`[OCR] Image OCR completed: ~${rawText.length} chars recognized`);
        } else {
          console.error(`[OCR] No text recognized in image`);
        }
      } else if (this.isDocxMimeType(mimeType, fileName)) {
        console.error(`[OCR] DOCX file detected, extracting text and embedded images...`);
        const downloaded = await this.drive.files.get(
          { fileId, alt: 'media', supportsAllDrives: true },
          { responseType: 'arraybuffer' }
        );
        rawText = await this.extractDocxText(downloaded.data);
      } else if (this.isPptxMimeType(mimeType, fileName)) {
        console.error(`[OCR] PPTX file detected, extracting text and embedded images...`);
        const downloaded = await this.drive.files.get(
          { fileId, alt: 'media', supportsAllDrives: true },
          { responseType: 'arraybuffer' }
        );
        rawText = await this.extractPptxText(downloaded.data);
      } else if (this.isXlsxMimeType(mimeType, fileName)) {
        console.error(`[OCR] XLSX file detected, extracting text...`);
        const downloaded = await this.drive.files.get(
          { fileId, alt: 'media', supportsAllDrives: true },
          { responseType: 'arraybuffer' }
        );
        rawText = await this.extractXlsxText(downloaded.data);
      } else if (this.isPdfMimeType(mimeType, fileName)) {
        console.error(`[OCR] PDF file detected, extracting text...`);
        const downloaded = await this.drive.files.get(
          { fileId, alt: 'media', supportsAllDrives: true },
          { responseType: 'arraybuffer' }
        );
        rawText = await this.extractPdfText(downloaded.data);
      } else if (this.isTextLikeMimeType(mimeType, fileName)) {
        console.error(`[OCR] Text file detected, decoding...`);
        const downloaded = await this.drive.files.get(
          { fileId, alt: 'media', supportsAllDrives: true },
          { responseType: 'arraybuffer' }
        );
        rawText = this.decodeBinaryText(downloaded.data);
        console.error(`[OCR] Text decoded: ${rawText?.length || 0} chars`);
      } else {
        return {
          status: 'skipped',
          warning: `Skipping ${fileName}: unsupported mimeType ${mimeType}.`,
          mimeType,
        };
      }

if ((!rawText || !rawText.trim()) && ocrTexts.length === 0) {
        return {
          status: 'empty',
          warning: `No text could be extracted from ${fileName}.`,
          mimeType,
        };
      }

      let finalText = rawText || '';
      if (ocrTexts.length > 0) {
        const separator = finalText.trim() ? '\n\n--- Document Images OCR ---\n' : '--- Image OCR ---\n';
        finalText = finalText.trim() + separator + ocrTexts.join('\n');
      }

      return {
        status: 'ok',
        text: finalText,
        mimeType,
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (message.includes('insufficient') || message.includes('scope') || message.includes('permission')) {
        return {
          status: 'error',
          warning: 'Drive access denied. Re-run `npm run setup-auth` to grant drive.readonly scope.',
          mimeType: null,
        };
      }
      return {
        status: 'error',
        warning: `Drive extraction failed: ${message}`,
        mimeType: null,
      };
    }
  }

  private toBuffer(data: unknown): Buffer {
    if (Buffer.isBuffer(data)) return data;
    if (data instanceof ArrayBuffer) return Buffer.from(data);
    if (ArrayBuffer.isView(data)) return Buffer.from((data as ArrayBufferView).buffer, (data as ArrayBufferView).byteOffset, (data as ArrayBufferView).byteLength);
    if (typeof data === 'string') return Buffer.from(data, 'utf8');
    return Buffer.alloc(0);
  }

  private async extractImagesFromZip(buffer: Buffer, pathPrefix: string): Promise<Buffer[]> {
    const images: Buffer[] = [];
    try {
      const zip = await JSZip.loadAsync(buffer);
      const imageExts = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif', '.webp'];
      const imageFiles = Object.keys(zip.files)
        .filter((name) => {
          if (!name.startsWith(pathPrefix)) return false;
          const ext = path.extname(name).toLowerCase();
          return imageExts.includes(ext);
        })
        .sort();
      for (const name of imageFiles) {
        const data = await zip.file(name)?.async('nodebuffer');
        if (data && data.length <= OCR_MAX_IMAGE_SIZE) {
          images.push(data);
        }
      }
    } catch {}
    return images;
  }

  private isGoogleDriveImageUrl(url: string): boolean {
    try {
      const u = new URL(url);
      return u.hostname === 'lh3.googleusercontent.com'
        || u.hostname === 'lh4.googleusercontent.com'
        || u.hostname === 'lh5.googleusercontent.com'
        || u.hostname === 'lh6.googleusercontent.com'
        || u.hostname === 'drive.google.com'
        || u.hostname === 'docs.google.com';
    } catch {
      return false;
    }
  }

  private async extractImagesFromGoogleDocHtml(html: string, fileId: string): Promise<Buffer[]> {
    const images: Buffer[] = [];
    const imgRegex = /<img[^>]+src=["']([^"']+)["']/gi;
    let match: RegExpExecArray | null;

    while ((match = imgRegex.exec(html)) !== null) {
      const src = match[1];
      try {
        if (src.startsWith('data:')) {
          const base64Match = src.match(/^data:image\/[^;]+;base64,(.+)$/);
          if (base64Match) {
            const buf = Buffer.from(base64Match[1], 'base64');
            if (buf.length <= OCR_MAX_IMAGE_SIZE) images.push(buf);
          }
        } else if (this.isGoogleDriveImageUrl(src)) {
          const downloaded = await this.drive.files.get(
            { fileId, alt: 'media', supportsAllDrives: true },
            { responseType: 'arraybuffer' }
          );
          const buf = this.toBuffer(downloaded.data);
          if (buf.length <= OCR_MAX_IMAGE_SIZE) images.push(buf);
        } else if (src.startsWith('http')) {
          const resp = await fetch(src, { redirect: 'follow' });
          if (resp.ok) {
            const ab = await resp.arrayBuffer();
            const buf = Buffer.from(ab);
            if (buf.length <= OCR_MAX_IMAGE_SIZE) images.push(buf);
          }
        }
      } catch {}
      if (images.length >= OCR_MAX_IMAGES_PER_DOC) break;
    }
    return images;
  }

  private decodeBinaryText(data: unknown): string {
    let buffer: Buffer;
    if (Buffer.isBuffer(data)) {
      buffer = data;
    } else if (data instanceof ArrayBuffer) {
      buffer = Buffer.from(data);
    } else if (ArrayBuffer.isView(data)) {
      buffer = Buffer.from(data.buffer, data.byteOffset, data.byteLength);
    } else if (typeof data === 'string') {
      return data;
    } else {
      return '';
    }

    if (buffer.length >= 2 && buffer[0] === 0xff && buffer[1] === 0xfe) {
      return buffer.toString('utf16le').replace(/^\uFEFF/, '');
    }
    if (buffer.length >= 2 && buffer[0] === 0xfe && buffer[1] === 0xff) {
      const swapped = Buffer.allocUnsafe(buffer.length);
      for (let i = 0; i + 1 < buffer.length; i += 2) {
        swapped[i] = buffer[i + 1];
        swapped[i + 1] = buffer[i];
      }
      return swapped.toString('utf16le').replace(/^\uFEFF/, '');
    }
    return buffer.toString('utf8').replace(/^\uFEFF/, '');
  }

  private isImageMimeType(mimeType: string, fileName: string): boolean {
    if (IMAGE_MIME_TYPES.has(mimeType)) return true;
    const ext = path.extname(fileName).toLowerCase();
    return IMAGE_EXTENSIONS.includes(ext);
  }

  private isTextLikeMimeType(mimeType: string, fileName: string): boolean {
    if (mimeType.startsWith('text/')) return true;
    if (EXTRA_TEXT_MIME_TYPES.has(mimeType)) return true;
    const ext = path.extname(fileName).toLowerCase();
    return ['.txt', '.md', '.csv', '.json', '.xml', '.yaml', '.yml', '.html', '.htm', '.log', '.tsv'].includes(ext);
  }

  private isDocxMimeType(mimeType: string, fileName: string): boolean {
    if (DOCX_MIME_TYPES.has(mimeType)) return true;
    const ext = path.extname(fileName).toLowerCase();
    return ext === '.docx' || ext === '.doc';
  }

private async extractDocxText(data: unknown): Promise<string> {
    const buffer = this.toBuffer(data);
    if (!buffer.length) return '';

    console.error(`[OCR] DOCX: Extracting text with mammoth...`);
    const result = await mammoth.extractRawText({ buffer });
    let text = result.value.replace(/\r\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
    console.error(`[OCR] DOCX: Text extracted: ${text.length} chars`);

    try {
      console.error(`[OCR] DOCX: Looking for embedded images...`);
      const imageBuffers = await this.extractImagesFromZip(buffer, 'word/media/');
      if (imageBuffers.length > 0) {
        console.error(`[OCR] DOCX: Found ${imageBuffers.length} embedded image(s), running OCR...`);
        const ocrTexts = await this.ocrProcessor.recognizeImages(imageBuffers);
        if (ocrTexts.length > 0) {
          const totalOcrChars = ocrTexts.join('').length;
          text += '\n\n--- Embedded Images OCR ---\n' + ocrTexts.join('\n');
          console.error(`[OCR] DOCX: OCR completed: ${ocrTexts.length} image(s), ~${totalOcrChars} chars recognized`);
        } else {
          console.error(`[OCR] DOCX: No text recognized in embedded images`);
        }
      } else {
        console.error(`[OCR] DOCX: No embedded images found`);
      }
    } catch (e) {
      console.error(`[OCR] DOCX: Failed to process images: ${e}`);
    }

    return text;
  }

  private isPdfMimeType(mimeType: string, fileName: string): boolean {
    if (PDF_MIME_TYPES.has(mimeType)) return true;
    const ext = path.extname(fileName).toLowerCase();
    return ext === '.pdf';
  }

  private isPptxMimeType(mimeType: string, fileName: string): boolean {
    if (PPTX_MIME_TYPES.has(mimeType)) return true;
    const ext = path.extname(fileName).toLowerCase();
    return ext === '.pptx' || ext === '.ppt';
  }

  private isXlsxMimeType(mimeType: string, fileName: string): boolean {
    if (XLSX_MIME_TYPES.has(mimeType)) return true;
    const ext = path.extname(fileName).toLowerCase();
    return ext === '.xlsx' || ext === '.xls' || ext === '.xlsm';
  }

private async extractPdfText(data: unknown): Promise<string> {
    const buffer = this.toBuffer(data);
    if (!buffer.length) return '';

    console.error(`[OCR] PDF: Parsing text with pdf-parse...`);
    const parser = new PDFParse({ data: buffer });
    try {
      const result = await parser.getText();
      const text = (result.text || '').replace(/\r\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
      console.error(`[OCR] PDF: Text extracted: ${text.length} chars`);

      if (text.length > SCANNED_PDF_TEXT_THRESHOLD) {
        console.error(`[OCR] PDF: Sufficient text found (${text.length} chars), no OCR needed`);
        return text;
      }

      console.error(`[OCR] PDF: Low text amount (${text.length} chars), possible scanned PDF. Running OCR...`);
      try {
        const ocrText = await this.extractPdfTextViaOCR(buffer);
        if (ocrText) {
          const totalChars = text.length + ocrText.length;
          console.error(`[OCR] PDF: OCR completed: ~${ocrText.length} chars from scanned pages`);
          return text ? text + '\n\n--- Scanned Pages OCR ---\n' + ocrText : ocrText;
        }
      } catch (e) {
        console.error(`[OCR] PDF: OCR failed: ${e}`);
      }

      return text;
    } finally {
      await parser.destroy();
    }
  }

  private async extractPdfTextViaOCR(pdfBuffer: Buffer): Promise<string> {
    const doc = await getDocument({ data: new Uint8Array(pdfBuffer) }).promise;
    const maxPages = Math.min(doc.numPages, OCR_MAX_IMAGES_PER_DOC);
    console.error(`[OCR] PDF: Rendering ${maxPages} page(s) for OCR...`);
    const pageTexts: string[] = [];

    for (let i = 1; i <= maxPages; i++) {
      try {
        console.error(`[OCR] PDF: Rendering page ${i}/${maxPages}...`);
        const page = await doc.getPage(i);
        const viewport = page.getViewport({ scale: 3.0 });
        const canvas = createCanvas(viewport.width, viewport.height);
        const ctx = canvas.getContext('2d');

        await page.render({ canvas: canvas as any, canvasContext: ctx as any, viewport }).promise;

        const imageBuffer = canvas.toBuffer('image/png');
        console.error(`[OCR] PDF: Running OCR on page ${i}...`);
        const text = await this.ocrProcessor.recognizeImage(imageBuffer);
        if (text?.trim()) {
          pageTexts.push(`[Page ${i}]: ${text.trim()}`);
          console.error(`[OCR] PDF: Page ${i} recognized: ~${text.length} chars`);
        } else {
          console.error(`[OCR] PDF: Page ${i}: no text recognized`);
        }
      } catch (e) {
        console.error(`[OCR] PDF: Failed to process page ${i}: ${e}`);
      }
    }

    console.error(`[OCR] PDF: Total pages with text: ${pageTexts.length}/${maxPages}`);
    return pageTexts.join('\n\n');
  }

private async extractPptxText(data: unknown): Promise<string> {
    const buffer = this.toBuffer(data);
    if (!buffer.length) return '';

    console.error(`[OCR] PPTX: Parsing slides...`);
    const zip = await JSZip.loadAsync(buffer);
    const slideFiles = Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
      .sort((a, b) => {
        const ai = Number((a.match(/slide(\d+)\.xml/i) || [])[1] || 0);
        const bi = Number((b.match(/slide(\d+)\.xml/i) || [])[1] || 0);
        return ai - bi;
      });

    console.error(`[OCR] PPTX: Found ${slideFiles.length} slide(s)`);
    const slides: string[] = [];
    for (const slideFile of slideFiles) {
      const xml = await zip.file(slideFile)?.async('text');
      if (!xml) continue;
      const slideText = this.extractXmlText(xml);
      if (slideText) {
        const num = Number((slideFile.match(/slide(\d+)\.xml/i) || [])[1] || slides.length + 1);
        slides.push(`Slide ${num}\n${slideText}`);
      }
    }

    let text = slides.join('\n\n').trim();
    console.error(`[OCR] PPTX: Text extracted from ${slides.length} slide(s): ${text.length} chars`);

    try {
      console.error(`[OCR] PPTX: Looking for slide images...`);
      const imageBuffers = await this.extractImagesFromZip(buffer, 'ppt/media/');
      if (imageBuffers.length > 0) {
        console.error(`[OCR] PPTX: Found ${imageBuffers.length} image(s), running OCR...`);
        const ocrTexts = await this.ocrProcessor.recognizeImages(imageBuffers);
        if (ocrTexts.length > 0) {
          const totalOcrChars = ocrTexts.join('').length;
          text += '\n\n--- Slide Images OCR ---\n' + ocrTexts.join('\n');
          console.error(`[OCR] PPTX: OCR completed: ${ocrTexts.length} image(s), ~${totalOcrChars} chars recognized`);
        } else {
          console.error(`[OCR] PPTX: No text recognized in images`);
        }
      } else {
        console.error(`[OCR] PPTX: No slide images found`);
      }
    } catch (e) {
      console.error(`[OCR] PPTX: Failed to process images: ${e}`);
    }

    return text;
  }

  private async extractXlsxText(data: unknown): Promise<string> {
    let buffer: Buffer;
    if (Buffer.isBuffer(data)) {
      buffer = data;
    } else if (data instanceof ArrayBuffer) {
      buffer = Buffer.from(data);
    } else if (ArrayBuffer.isView(data)) {
      buffer = Buffer.from(data.buffer, data.byteOffset, data.byteLength);
    } else {
      return '';
    }

    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const parts: string[] = [];
    for (const name of workbook.SheetNames) {
      const sheet = workbook.Sheets[name];
      const csv = XLSX.utils.sheet_to_csv(sheet, { blankrows: false }).trim();
      if (!csv) continue;
      parts.push(`Sheet: ${name}\n${csv}`);
    }
    return parts.join('\n\n').trim();
  }

  private extractXmlText(xml: string): string {
    const fromTextNodes = [...xml.matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/gi)]
      .map((match) => this.decodeXmlEntities(match[1] || '').trim())
      .filter(Boolean);
    if (fromTextNodes.length > 0) {
      return fromTextNodes.join('\n');
    }

    return this.decodeXmlEntities(
      xml
        .replace(/<[^>]+>/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
    );
  }

  private decodeXmlEntities(value: string): string {
    return value
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#10;/g, '\n')
      .replace(/&#13;/g, '\n')
      .replace(/&#9;/g, '\t');
  }

private async extractTextFromGoogleForm(formUrl: string | null, fallbackTitle: string | null): Promise<{
    status: 'ok' | 'empty' | 'error' | 'skipped';
    text?: string;
    warning?: string | null;
    formId?: string | null;
  }> {
    const formId = this.extractGoogleFormId(formUrl);
    console.error(`[OCR] Google Form: Extracting from form ID: ${formId || 'N/A'}`);
    let apiErrorMessage: string | null = null;

    if (this.forms && formId) {
      try {
        console.error(`[OCR] Google Form: Fetching form data via API...`);
        const response = await this.forms.forms.get({ formId });
        const form = response.data || {};
        console.error(`[OCR] Google Form: Form title: ${form.info?.title || fallbackTitle || 'Untitled'}`);
        const lines: string[] = [];
        lines.push(`Form: ${form.info?.title || fallbackTitle || formId}`);
        if (form.info?.description) lines.push(`Description: ${form.info.description}`);

        const items = (form.items || []) as Array<Record<string, unknown>>;
        console.error(`[OCR] Google Form: Found ${items.length} item(s)`);
        let questionIndex = 1;
        const imageUrls: string[] = [];
        for (const item of items) {
          const itemTitle = item.title ? String(item.title) : '';
          const questionItem = item.questionItem as Record<string, unknown> | undefined;
          const question = (questionItem?.question ?? null) as Record<string, unknown> | null;
          if (question) {
            const title = itemTitle || String(question.questionId || `Question ${questionIndex}`);
            lines.push(`${questionIndex}. ${title}`);
            if (question.required) lines.push('Required: yes');
            const choiceQuestion = (question.choiceQuestion ?? null) as Record<string, unknown> | null;
            const choiceOptions = Array.isArray(choiceQuestion?.options)
              ? (choiceQuestion!.options as Array<Record<string, unknown>>)
              : [];
            if (choiceOptions.length) {
              const opts = choiceOptions
                .map((opt) => String(opt.value || '').trim())
                .filter(Boolean);
              if (opts.length) lines.push(`Options: ${opts.join(' | ')}`);
            }
            questionIndex += 1;
          } else if (item.textItem) {
            if (itemTitle) lines.push(`Text item: ${itemTitle}`);
          } else if (itemTitle) {
            lines.push(`Item: ${itemTitle}`);
          }
          const imageItem = item.imageItem as Record<string, unknown> | undefined;
          if (imageItem?.imageUri) {
            imageUrls.push(String(imageItem.imageUri));
          } else if (imageItem?.sourceUri) {
            imageUrls.push(String(imageItem.sourceUri));
          }
          if ((item.image as Record<string, unknown> | undefined)?.contentUri) {
            imageUrls.push(String((item.image as Record<string, unknown>).contentUri));
          }
        }

        if (imageUrls.length > 0) {
          console.error(`[OCR] Google Form: Found ${imageUrls.length} image(s), running OCR...`);
          const ocrTexts: string[] = [];
          for (const url of imageUrls.slice(0, OCR_MAX_IMAGES_PER_DOC)) {
            try {
              const resp = await fetch(url, { redirect: 'follow' });
              if (resp.ok) {
                const ab = await resp.arrayBuffer();
                const buf = Buffer.from(ab);
                if (buf.length <= OCR_MAX_IMAGE_SIZE) {
                  const text = await this.ocrProcessor.recognizeImage(buf);
                  if (text?.trim()) ocrTexts.push(`[Form Image ${ocrTexts.length + 1}]: ${text.trim()}`);
                }
              }
            } catch {}
          }
          if (ocrTexts.length > 0) {
            lines.push('\n--- Form Images OCR ---');
            lines.push(...ocrTexts);
            console.error(`[OCR] Google Form: OCR completed: ${ocrTexts.length} image(s) processed`);
          }
        }

        const text = lines.join('\n').trim();
        if (text) {
          console.error(`[OCR] Google Form: Extracted ${text.length} chars total`);
          return {
            status: 'ok',
            text,
            formId,
          };
        }
        apiErrorMessage = 'Google Form returned no extractable text.';
      } catch (error) {
        apiErrorMessage = error instanceof Error ? error.message : String(error);
        console.error(`[OCR] Google Form: API error: ${apiErrorMessage}`);
      }
    } else if (!this.forms) {
      apiErrorMessage = 'Google Forms API client is not initialized.';
      console.error(`[OCR] Google Form: API not initialized`);
    } else if (!formId) {
      apiErrorMessage = 'Could not parse Google Form ID from URL.';
      console.error(`[OCR] Google Form: Could not parse form ID from URL`);
    }

    console.error(`[OCR] Google Form: Falling back to public form scraping...`);
    const publicFallback = await this.extractTextFromPublicGoogleForm(formUrl, fallbackTitle);
    if (publicFallback.status === 'ok' || publicFallback.status === 'empty') {
      return {
        ...publicFallback,
        formId,
        warning: apiErrorMessage
          ? `Forms API unavailable for this link; extracted via public viewform. API reason: ${apiErrorMessage}`
          : (publicFallback.warning ?? null),
      };
    }

    const message = apiErrorMessage || publicFallback.warning || 'Unknown forms extraction error';
    if (message.includes('insufficient') || message.includes('scope') || message.includes('permission')) {
      return {
        status: 'error',
        warning: 'Google Forms access denied. Re-run `npm run setup-auth` with forms.body.readonly scope.',
        formId,
      };
    }
    if (message.includes('not been used') || message.includes('disabled')) {
      return {
        status: 'error',
        warning: 'Google Forms API is not enabled in your Google Cloud project.',
        formId,
      };
    }
    return {
      status: 'error',
      warning: `Forms extraction failed: ${message}`,
      formId,
    };
  }

  private async extractTextFromPublicGoogleForm(formUrl: string | null, fallbackTitle: string | null): Promise<{
    status: 'ok' | 'empty' | 'error' | 'skipped';
    text?: string;
    warning?: string | null;
  }> {
    if (!formUrl) {
      return {
        status: 'skipped',
        warning: 'Form URL is empty.',
      };
    }

    let parsed: URL;
    try {
      parsed = new URL(formUrl);
    } catch {
      return {
        status: 'skipped',
        warning: 'Form URL is invalid.',
      };
    }
    if (parsed.hostname !== 'docs.google.com') {
      return {
        status: 'skipped',
        warning: 'Only docs.google.com forms are supported.',
      };
    }

const viewFormUrl = `${parsed.origin}${parsed.pathname}${parsed.search || ''}`;
    console.error(`[OCR] Google Form Public: Fetching form from ${viewFormUrl}`);
    try {
      const response = await fetch(viewFormUrl, {
        method: 'GET',
        redirect: 'follow',
      });
      if (!response.ok) {
        console.error(`[OCR] Google Form Public: Fetch failed with HTTP ${response.status}`);
        return {
          status: 'error',
          warning: `Public form fetch failed with HTTP ${response.status}.`,
        };
      }

      console.error(`[OCR] Google Form Public: Parsing HTML...`);
      const html = await response.text();
      const parsedText = this.extractStructuredTextFromGoogleFormHtml(html, fallbackTitle);
      console.error(`[OCR] Google Form Public: Parsed text: ${parsedText?.length || 0} chars`);

      console.error(`[OCR] Google Form Public: Looking for images in HTML...`);
      const imgRegex = /<img[^>]+src=["']([^"']+)["']/gi;
      const imgSrcs: string[] = [];
      let imgMatch: RegExpExecArray | null;
      while ((imgMatch = imgRegex.exec(html)) !== null) {
        const src = imgMatch[1];
        if (!src.startsWith('data:image/svg') && !src.includes('fonts.googleapis') && !src.includes('favicon')) {
          imgSrcs.push(src);
        }
      }
      console.error(`[OCR] Google Form Public: Found ${imgSrcs.length} image(s)`);

      const ocrTexts: string[] = [];
      for (const src of imgSrcs.slice(0, OCR_MAX_IMAGES_PER_DOC)) {
        try {
          let buf: Buffer;
          if (src.startsWith('data:')) {
            const b64Match = src.match(/^data:image\/[^;]+;base64,(.+)$/);
            if (b64Match) {
              buf = Buffer.from(b64Match[1], 'base64');
            } else {
              continue;
            }
          } else {
            const imgResp = await fetch(src, { redirect: 'follow' });
            if (!imgResp.ok) continue;
            buf = Buffer.from(await imgResp.arrayBuffer());
          }
          if (buf.length <= OCR_MAX_IMAGE_SIZE && buf.length > 100) {
            const text = await this.ocrProcessor.recognizeImage(buf);
            if (text?.trim()) ocrTexts.push(`[Form Image ${ocrTexts.length + 1}]: ${text.trim()}`);
          }
        } catch {}
      }
      if (ocrTexts.length > 0) {
        console.error(`[OCR] Google Form Public: OCR completed: ${ocrTexts.length} image(s) processed`);
      }

      let finalText = parsedText || '';
      if (ocrTexts.length > 0) {
        const separator = finalText.trim() ? '\n\n--- Form Images OCR ---\n' : '--- Form Images OCR ---\n';
        finalText = finalText.trim() + separator + ocrTexts.join('\n');
      }

      if (finalText) {
        console.error(`[OCR] Google Form Public: Total extracted: ${finalText.length} chars`);
        return {
          status: 'ok',
          text: finalText,
          warning: null,
        };
      }

      return {
        status: 'empty',
        warning: 'Public form page was loaded, but no extractable content was found.',
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return {
        status: 'error',
        warning: `Public form fallback failed: ${message}`,
      };
    }
  }

  private extractStructuredTextFromGoogleFormHtml(html: string, fallbackTitle: string | null): string | null {
    const titleMatch = html.match(/<title>([\s\S]*?)<\/title>/i);
    const htmlTitle = titleMatch ? this.normalizeFormText(titleMatch[1]) : '';
    const payloadMatch = html.match(/FB_PUBLIC_LOAD_DATA_\s*=\s*(\[[\s\S]*?\]);/i);
    if (!payloadMatch) {
      if (htmlTitle || fallbackTitle) {
        return `Form: ${htmlTitle || fallbackTitle}`;
      }
      return null;
    }

    const payload = payloadMatch[1];
    const parsedData = this.parseGoogleFormsPublicPayload(payload);
    if (!parsedData) {
      if (htmlTitle || fallbackTitle) {
        return `Form: ${htmlTitle || fallbackTitle}`;
      }
      return null;
    }

    const root = parsedData as unknown[];
    const formBlock = Array.isArray(root[1]) ? (root[1] as unknown[]) : [];
    const formTitle = this.normalizeFormText(formBlock[8]) || htmlTitle || fallbackTitle || '';
    const items = Array.isArray(formBlock[1]) ? (formBlock[1] as unknown[]) : [];

    const lines: string[] = [];
    if (formTitle) {
      lines.push(`Form: ${formTitle}`);
    }

    let index = 1;
    for (const rawItem of items) {
      if (!Array.isArray(rawItem)) continue;
      const item = rawItem as unknown[];
      const itemTitle = this.normalizeFormText(item[1]) || this.normalizeFormText(Array.isArray(item[11]) ? (item[11] as unknown[])[1] : null);
      if (!itemTitle) continue;
      lines.push(`${index}. ${itemTitle}`);

      const typeCode = typeof item[3] === 'number' ? (item[3] as number) : null;
      if (typeCode != null) {
        lines.push(`Type: ${this.describeGoogleFormQuestionType(typeCode)}`);
      }

      const options = this.extractGoogleFormOptions(item[4]);
      if (options.length > 0) {
        lines.push(`Options: ${options.join(' | ')}`);
      }
      index += 1;
    }

    const text = lines.join('\n').trim();
    return text || null;
  }

  private parseGoogleFormsPublicPayload(payload: string): unknown[] | null {
    try {
      const parsed = Function(`"use strict"; return (${payload});`)() as unknown;
      return Array.isArray(parsed) ? (parsed as unknown[]) : null;
    } catch {
      return null;
    }
  }

  private extractGoogleFormOptions(questionBlock: unknown): string[] {
    if (!Array.isArray(questionBlock)) {
      return [];
    }

    const options: string[] = [];
    for (const subQuestion of questionBlock) {
      if (!Array.isArray(subQuestion)) continue;
      const optionEntries = subQuestion[1];
      if (!Array.isArray(optionEntries)) continue;
      for (const rawOption of optionEntries) {
        const optionText = this.normalizeFormText(Array.isArray(rawOption) ? rawOption[0] : rawOption);
        if (!optionText) continue;
        options.push(optionText);
      }
    }

    const uniq: string[] = [];
    const seen = new Set<string>();
    for (const option of options) {
      if (seen.has(option)) continue;
      seen.add(option);
      uniq.push(option);
      if (uniq.length >= 25) break;
    }
    return uniq;
  }

  private describeGoogleFormQuestionType(typeCode: number): string {
    const map: Record<number, string> = {
      0: 'Short answer',
      1: 'Paragraph',
      2: 'Multiple choice',
      3: 'Dropdown',
      4: 'Checkboxes',
      5: 'Linear scale',
      6: 'Grid',
      7: 'Date',
      8: 'Time',
      9: 'File upload',
    };
    return map[typeCode] || `Type ${typeCode}`;
  }

  private normalizeFormText(value: unknown): string {
    if (typeof value !== 'string') return '';
    return this.decodeXmlEntities(value)
      .replace(/\u00A0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  private extractGoogleFormId(formUrl: string | null): string | null {
    if (!formUrl) return null;
    try {
      const parsed = new URL(formUrl);
      const match = parsed.pathname.match(/\/forms\/d\/e\/([^/]+)/i) || parsed.pathname.match(/\/forms\/d\/([^/]+)/i);
      return match?.[1] || null;
    } catch {
      return null;
    }
  }

  private prepareMaterialTextForResponse(
    rawText: string | null,
    includeText: boolean,
    maxCharsPerMaterial: number,
    previewChars: number,
    initialWarning: string | null
  ): {
    textForCache: string | null;
    textForInline: string | null;
    previewText: string | null;
    extractedChars: number;
    warning: string | null;
  } {
    if (!rawText) {
      return {
        textForCache: null,
        textForInline: null,
        previewText: null,
        extractedChars: 0,
        warning: initialWarning,
      };
    }

    const extractedChars = rawText.length;
    let warning = initialWarning;
    let cachedText = rawText;
    if (cachedText.length > MATERIAL_CACHE_MAX_TEXT_CHARS) {
      cachedText = cachedText.slice(0, MATERIAL_CACHE_MAX_TEXT_CHARS);
      warning = warning
        ? `${warning} Text was trimmed to ${MATERIAL_CACHE_MAX_TEXT_CHARS} chars for cache safety.`
        : `Text was trimmed to ${MATERIAL_CACHE_MAX_TEXT_CHARS} chars for cache safety.`;
    }

    let textForInline: string | null = null;
    if (includeText) {
      textForInline = cachedText.length > maxCharsPerMaterial
        ? `${cachedText.slice(0, maxCharsPerMaterial)}\n\n...[truncated ${cachedText.length - maxCharsPerMaterial} chars]`
        : cachedText;
      if (cachedText.length > maxCharsPerMaterial) {
        warning = warning
          ? `${warning} Inline text was truncated to ${maxCharsPerMaterial} chars.`
          : `Inline text was truncated to ${maxCharsPerMaterial} chars.`;
      }
    }

    const previewText = cachedText.length > previewChars
      ? `${cachedText.slice(0, previewChars)}...`
      : cachedText;

    return {
      textForCache: cachedText,
      textForInline,
      previewText,
      extractedChars: cachedText.length,
      warning,
    };
  }

  private buildMaterialDocRef(courseId: string, courseWorkId: string, materialIndex: number, driveFileId?: string) {
    const suffix = driveFileId ? hashString(driveFileId).slice(0, 10) : 'na';
    return `${courseId}:${courseWorkId}:${materialIndex}:${suffix}`;
  }

  private putMaterialCache(doc: CachedMaterialDoc) {
    this.materialCacheStore.upsert(doc);
  }

private cleanupMaterialCache() {
    this.materialCacheStore.cleanupExpired(Date.now());
  }

  private async clearMaterialCache() {
    const deleted = this.materialCacheStore.clearAll();
    return {
      content: [{
        type: 'text',
        text: JSON.stringify({ cleared: true, deletedEntries: deleted }, null, 2),
      }],
    };
  }

  private async listMaterialCache(args: { courseId?: string; courseWorkId?: string; limit?: number } = {}) {
    this.cleanupMaterialCache();
    const limit = clampNumber(args.limit, 50, 1, 200);
    const now = Date.now();
    const docs = this.materialCacheStore.list({
      courseId: args.courseId,
      courseWorkId: args.courseWorkId,
      limit,
      nowMs: now,
    });
    const items = docs
      .map((doc) => ({
        docRef: doc.docRef,
        courseId: doc.courseId,
        courseWorkId: doc.courseWorkId,
        assignmentTitle: doc.assignmentTitle ?? null,
        materialIndex: doc.materialIndex,
        type: doc.type,
        title: doc.title ?? null,
        mimeType: doc.mimeType ?? null,
        extractedChars: doc.text.length,
        extractedAt: doc.extractedAt,
        expiresAt: new Date(doc.expiresAtMs).toISOString(),
      }));

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          totalCached: this.materialCacheStore.count(now),
          returned: items.length,
          items,
        }, null, 2),
      }],
    };
  }

  private async searchMaterialCache(args: {
    query: string;
    courseId?: string;
    courseWorkId?: string;
    limit?: number;
    snippetChars?: number;
  }) {
    this.cleanupMaterialCache();
    const queryNormalized = normalizeText(args.query || '');
    if (!queryNormalized) {
      throw new SearchValidationError('INVALID_FILTER', 'query is required');
    }
    const limit = clampNumber(args.limit, 10, 1, 50);
    const snippetChars = clampNumber(args.snippetChars, 500, 100, 2000);
    const now = Date.now();

    const hits = this.materialCacheStore.search({
      queryNormalized,
      courseId: args.courseId,
      courseWorkId: args.courseWorkId,
      limit,
      nowMs: now,
    })
      .map((doc) => {
        const titleScore = normalizeText(doc.title || '').includes(queryNormalized) ? 2 : 0;
        const bodyIncludes = doc.textNormalized.includes(queryNormalized);
        const bodyScore = bodyIncludes ? 1 : 0;
        const score = titleScore + bodyScore;
        const snippet = bodyIncludes
          ? this.buildMaterialSnippet(doc.text, queryNormalized, snippetChars)
          : (doc.text.length > snippetChars ? `${doc.text.slice(0, snippetChars)}...` : doc.text);
        return {
          doc,
          score,
          snippet,
        };
      })
      .sort((a, b) => {
        if (b.score !== a.score) return b.score - a.score;
        return b.doc.extractedAt.localeCompare(a.doc.extractedAt);
      })
      .slice(0, limit)
      .map((item) => ({
        docRef: item.doc.docRef,
        courseId: item.doc.courseId,
        courseWorkId: item.doc.courseWorkId,
        assignmentTitle: item.doc.assignmentTitle ?? null,
        title: item.doc.title ?? null,
        materialIndex: item.doc.materialIndex,
        mimeType: item.doc.mimeType ?? null,
        score: item.score,
        snippet: item.snippet,
        extractedAt: item.doc.extractedAt,
      }));

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          query: args.query,
          totalHits: hits.length,
          hits,
        }, null, 2),
      }],
    };
  }

  private buildMaterialSnippet(text: string, queryNormalized: string, snippetChars: number) {
    const normalized = normalizeText(text);
    const idx = normalized.indexOf(queryNormalized);
    if (idx < 0) {
      return text.length > snippetChars ? `${text.slice(0, snippetChars)}...` : text;
    }
    const start = Math.max(0, idx - Math.floor(snippetChars / 3));
    const end = Math.min(text.length, start + snippetChars);
    return `${start > 0 ? '...' : ''}${text.slice(start, end)}${end < text.length ? '...' : ''}`;
  }

  private async readMaterialCache(args: { docRef: string; offset?: number; maxChars?: number }) {
    this.cleanupMaterialCache();
    const doc = this.materialCacheStore.get(args.docRef, Date.now());
    if (!doc) {
      throw new SearchValidationError('INVALID_FILTER', `docRef not found or expired: ${args.docRef}`);
    }
    const offset = clampNumber(args.offset, 0, 0, Math.max(0, doc.text.length));
    const maxChars = clampNumber(args.maxChars, 3000, 200, 20000);
    const end = Math.min(doc.text.length, offset + maxChars);
    const chunk = doc.text.slice(offset, end);

    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          docRef: doc.docRef,
          courseId: doc.courseId,
          courseWorkId: doc.courseWorkId,
          assignmentTitle: doc.assignmentTitle ?? null,
          title: doc.title ?? null,
          mimeType: doc.mimeType ?? null,
          offset,
          end,
          totalChars: doc.text.length,
          hasMore: end < doc.text.length,
          nextOffset: end < doc.text.length ? end : null,
          text: chunk,
        }, null, 2),
      }],
    };
  }

  private async getDashboard(args: { days?: number; upcomingLimit?: number; missingLimit?: number } = {}) {
    const days = args.days ?? 7;
    const upcomingLimit = clampNumber(args.upcomingLimit, 10, 1, 100);
    const missingLimit = clampNumber(args.missingLimit, 10, 1, 100);
    const now = new Date();
    now.setHours(0, 0, 0, 0);
    const dueFrom = now.toISOString();
    const dueTo = new Date(now.getTime() + days * 24 * 60 * 60 * 1000).toISOString();

    const [coursesResult, upcomingResult, missingResult, grades] = await Promise.all([
      this.searchService.search({ entityTypes: ['courses'], limit: 1000 }),
      this.searchService.search({
        entityTypes: ['assignments'],
        dueFrom,
        dueTo,
        sort: 'dueDate',
        limit: upcomingLimit,
      }),
      this.searchService.search({
        entityTypes: ['assignments'],
        missingOnly: true,
        sort: 'dueDate',
        limit: missingLimit,
      }),
      this.getGradesData(),
    ]);

    let totalEarned = 0;
    let totalPossible = 0;
    for (const grade of grades) {
      if (typeof grade.assignedGrade === 'number' && typeof grade.maxPoints === 'number' && grade.maxPoints > 0) {
        totalEarned += grade.assignedGrade;
        totalPossible += grade.maxPoints;
      }
    }

    const dashboard = {
      generatedAt: new Date().toISOString(),
      activeCourses: coursesResult.totalApprox,
      upcomingWindowDays: days,
      upcomingCount: upcomingResult.totalApprox,
      missingCount: missingResult.totalApprox,
      gradeSummary: {
        gradedAssignments: grades.filter((g) => g.assignedGrade != null && g.maxPoints != null).length,
        totalEarned,
        totalPossible,
        overallPercentage: totalPossible > 0 ? Math.round((totalEarned / totalPossible) * 1000) / 10 : null,
      },
      upcoming: upcomingResult.items,
      missing: missingResult.items,
      meta: {
        warnings: [...new Set([...(upcomingResult.meta.warnings || []), ...(missingResult.meta.warnings || [])])],
      },
    };

    return { content: [{ type: 'text', text: JSON.stringify(dashboard, null, 2) }] };
  }

  private async searchAssignments(args: {
    query?: string;
    courseId?: string;
    daysAhead?: number;
    includeNoDueDate?: boolean;
    limit?: number;
  }) {
    const now = new Date();
    now.setHours(0, 0, 0, 0);
    const dueFrom = args.daysAhead != null ? now.toISOString() : undefined;
    const dueTo = args.daysAhead != null
      ? new Date(now.getTime() + args.daysAhead * 24 * 60 * 60 * 1000).toISOString()
      : undefined;

    const unified = await this.searchService.search({
      query: args.query,
      entityTypes: ['assignments'],
      courseIds: args.courseId ? [args.courseId] : undefined,
      dueFrom,
      dueTo,
      sort: 'dueDate',
      limit: args.limit,
    });

    let items = unified.items;
    if (args.includeNoDueDate) {
      const extra = await this.searchService.search({
        query: args.query,
        entityTypes: ['assignments'],
        courseIds: args.courseId ? [args.courseId] : undefined,
        sort: 'relevance',
        limit: args.limit,
      });
      const withNoDue = extra.items.filter((item) => !item.dueDate);
      const seen = new Set(items.map((item) => `${item.courseId || ''}:${item.id}`));
      for (const candidate of withNoDue) {
        const key = `${candidate.courseId || ''}:${candidate.id}`;
        if (!seen.has(key)) {
          items.push(candidate);
          seen.add(key);
        }
      }
    }

    items = items.slice(0, clampNumber(args.limit, 50, 1, 200));
    const legacy = items.map((item) => ({
      courseId: item.courseId,
      courseName: item.courseName,
      assignmentId: item.id,
      title: item.title,
      description: item.snippet ?? null,
      dueDate: item.dueDate ?? null,
      maxPoints: item.maxPoints ?? null,
      workType: null,
      alternateLink: item.alternateLink ?? null,
    }));

    return { content: [{ type: 'text', text: JSON.stringify(legacy, null, 2) }] };
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);

    const cleanup = async () => {
      try {
        await this.ocrProcessor.destroy();
      } catch {}
      process.exit(0);
    };
    process.on('SIGINT', cleanup);
    process.on('SIGTERM', cleanup);

    console.error('Google Classroom MCP server running on stdio');
  }
}

const server = new GoogleClassroomMCPServer();
server.run().catch(console.error);
