import { db } from '@/lib/db';
import { sql } from 'kysely';
import { NotFoundError } from '@/lib/errors';
import * as XLSX from 'xlsx';

/**
 * Column mapping per assessment spec (ndis_excel_import_logic.sql)
 */
const COLUMNS = {
   ITEM_NUMBER: 'A',
   ITEM_NAME: 'B',
   CATEGORY_NUMBER: 'F',
   CATEGORY_NAME: 'H',
   UNIT: 'I',
   IS_QUOTE_REQUIRED: 'J',
   START_DATE: 'K',
   END_DATE: 'L',
   ACT: 'M',
   NSW: 'N',
   NT: 'O',
   QLD: 'P',
   SA: 'Q',
   TAS: 'R',
   VIC: 'S',
   WA: 'T',
   REMOTE: 'U',
   VERY_REMOTE: 'V',
   IS_NF2F_SUPPORT_PROVISION: 'W',
   IS_PROVIDER_TRAVEL: 'X',
   IS_SHORT_NOTICE_CANCEL: 'Y',
   IS_NDIA_REQUESTED_REPORTS: 'Z',
   IS_IRREGULAR_SIL_SUPPORTS: 'AA',
   SUPPORT_TYPE: 'AB',
} as const;

// Column -> fixed attribute code (label comes from the header cell at runtime)
const ATTRIBUTE_COLUMNS: Array<{ column: string; code: string }> = [
   { column: COLUMNS.IS_QUOTE_REQUIRED, code: 'IS_QUOTE_REQUIRED' },
   {
      column: COLUMNS.IS_NF2F_SUPPORT_PROVISION,
      code: 'IS_NF2F_SUPPORT_PROVISION',
   },
   { column: COLUMNS.IS_PROVIDER_TRAVEL, code: 'IS_PROVIDER_TRAVEL' },
   { column: COLUMNS.IS_SHORT_NOTICE_CANCEL, code: 'IS_SHORT_NOTICE_CANCEL' },
   {
      column: COLUMNS.IS_NDIA_REQUESTED_REPORTS,
      code: 'IS_NDIA_REQUESTED_REPORTS',
   },
   {
      column: COLUMNS.IS_IRREGULAR_SIL_SUPPORTS,
      code: 'IS_IRREGULAR_SIL_SUPPORTS',
   },
];

// Column -> fixed full_label (code/label come from the header cell at runtime)
const PRICING_REGION_COLUMNS: Array<{ column: string; fullLabel: string }> = [
   { column: COLUMNS.ACT, fullLabel: 'Australian Capital Territory' },
   { column: COLUMNS.NSW, fullLabel: 'New South Wales' },
   { column: COLUMNS.NT, fullLabel: 'Northern Territory' },
   { column: COLUMNS.QLD, fullLabel: 'Queensland' },
   { column: COLUMNS.SA, fullLabel: 'South Australia' },
   { column: COLUMNS.TAS, fullLabel: 'Tasmania' },
   { column: COLUMNS.VIC, fullLabel: 'Victoria' },
   { column: COLUMNS.WA, fullLabel: 'Western Australia' },
   { column: COLUMNS.REMOTE, fullLabel: 'Remote' },
   { column: COLUMNS.VERY_REMOTE, fullLabel: 'Very Remote' },
];

type ColumnKeys = (typeof COLUMNS)[keyof typeof COLUMNS];

interface ParsedRow {
   itemNumber: string;
   itemName: string;
   categoryNumber: string;
   categoryName: string;
   unit: string;
   startDate: string;
   endDate: string | null;
   supportTypeLabel: string | null; // raw Column AB cell value for this row
   attributes: Record<string, boolean>;
   prices: Record<string, number | null>;
}

function createRowReader(row: Record<string, unknown>) {
   return {
      string: (column: string) => String(row[column] ?? '').trim(),
      value: (column: string) => row[column],
   };
}

function toIsoString(cellValue: unknown): string | null {
   const dateStr = String(cellValue || '');
   if (!/^\d{8}$/.test(dateStr)) return null;
   if (dateStr === '99991231') return null; // NDIS's "infinite" end-date sentinel

   const year = parseInt(dateStr.substring(0, 4), 10);
   const month = parseInt(dateStr.substring(4, 6), 10) - 1;
   const day = parseInt(dateStr.substring(6, 8), 10);

   const date = new Date(year, month, day);
   const isValidDate =
      date.getFullYear() === year &&
      date.getMonth() === month &&
      date.getDate() === day;
   if (!isValidDate) return null;

   return `${dateStr.substring(0, 4)}-${dateStr.substring(4, 6)}-${dateStr.substring(6, 8)}`;
}

function toYesNoBoolean(value: unknown): boolean {
   if (typeof value !== 'string') return false;
   return /^y(es)?$/i.test(value.trim());
}

/**
 * Formats a DB timestamptz (returned as a JS Date) to the same YYYY-MM-DD
 */
function toIsoDate(date: Date): string {
   return date.toISOString().slice(0, 10);
}

/**
 * Shared "SCREAMING_SNAKE_CASE" derivation rule used by the spec
 */
function toCode(label: string): string {
   return label.trim().toUpperCase().replace(/\s+/g, '_');
}

interface WorkbookParseResult {
   rows: ParsedRow[];
   headerRow: Record<string, unknown>;
}

/**
 * Parses every worksheet in the workbook into normalised rows
 */
function parseWorkbook(buffer: Buffer): WorkbookParseResult {
   const workbook = XLSX.read(buffer);
   const rows: ParsedRow[] = [];
   const seen = new Set<string>(); // `${itemNumber}::${categoryNumber}::${startDate}::${endDate}`
   let headerRow: Record<string, unknown> | null = null;

   for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const allRows: Record<ColumnKeys, unknown>[] = XLSX.utils.sheet_to_json(
         sheet,
         {
            header: 'A',
            defval: null,
         },
      );
      if (allRows.length === 0) continue;

      if (!headerRow) headerRow = allRows[0]; // capture once, from the first non-empty sheet
      const dataRows = allRows.slice(1);

      for (const raw of dataRows) {
         const reader = createRowReader(raw);
         const itemNumber = reader.string(COLUMNS.ITEM_NUMBER);
         if (!itemNumber) continue; // skip blank/footer rows

         const categoryNumber = reader.string(COLUMNS.CATEGORY_NUMBER);
         const startDate = toIsoString(reader.value(COLUMNS.START_DATE));
         const endDate = toIsoString(reader.value(COLUMNS.END_DATE));

         if (!startDate) continue; // start_date is NOT NULL downstream; unparsable rows can't be imported

         const dedupeKey = `${itemNumber}::${categoryNumber}::${startDate}::${endDate ?? ''}`;
         if (seen.has(dedupeKey)) continue;
         seen.add(dedupeKey);

         const attributes: Record<string, boolean> = {};
         for (const { column, code } of ATTRIBUTE_COLUMNS) {
            attributes[code] = toYesNoBoolean(reader.value(column));
         }

         const prices: Record<string, number | null> = {};
         for (const { column } of PRICING_REGION_COLUMNS) {
            const value = reader.value(column);
            prices[column] =
               value === null || value === '' ? null : Number(value);
         }

         const supportTypeLabel = reader.string(COLUMNS.SUPPORT_TYPE) || null;

         rows.push({
            itemNumber,
            itemName: reader.string(COLUMNS.ITEM_NAME),
            categoryNumber,
            categoryName: reader.string(COLUMNS.CATEGORY_NAME),
            unit: reader.string(COLUMNS.UNIT),
            startDate,
            endDate,
            supportTypeLabel,
            attributes,
            prices, // keyed by column letter here; mapped to region code during import
         });
      }
   }

   return { rows, headerRow: headerRow ?? {} };
}

export interface ImportSummary {
   categoriesCreated: number;
   categoriesUpdated: number;
   categoriesDeactivated: number;
   itemsCreated: number;
   itemsUpdated: number;
   itemsDeactivated: number;
   pricesCreated: number;
   pricesUpdated: number;
   pricesRemoved: number;
}

/**
 * Idempotently imports a parsed NDIS pricing workbook into an existing rate_set
 */
const WRITE_BATCH_SIZE = 250;

function chunks<T>(items: T[], size = WRITE_BATCH_SIZE): T[][] {
   const batches: T[][] = [];
   for (let i = 0; i < items.length; i += size) batches.push(items.slice(i, i + size));
   return batches;
}

async function writeBatches<T>(items: T[], write: (batch: T[]) => Promise<unknown>) {
   for (const batch of chunks(items)) await write(batch);
}

function isoDateKey(value: Date | string | null): string {
   if (value === null) return '';
   return (value instanceof Date ? value : new Date(value)).toISOString().slice(0, 10);
}

function makePriceKey(
   itemId: number,
   typeId: number | null,
   regionCode: string | null,
   startDate: string,
   endDate: string | null,
): string {
   return [itemId, typeId ?? 'null', regionCode ?? 'null', startDate, endDate ?? ''].join('::');
}

/**
 * Imports the workbook atomically.
 *
 * Existing rows are fetched in bulk and matched in memory. Writes are batched
 * in groups of 250 to avoid thousands of sequential database round-trips.
 */
export async function importRateSetExcel(rateSetId: number, file: File) {
   const rateSet = await db
      .selectFrom('rate_set')
      .select('id')
      .where('id', '=', rateSetId)
      .executeTakeFirst();

   if (!rateSet) throw new NotFoundError('Rate set not found.');

   const buffer = Buffer.from(await file.arrayBuffer());
   const { rows, headerRow } = parseWorkbook(buffer);
   const headerReader = createRowReader(headerRow);

   const summary: ImportSummary = {
      categoriesCreated: 0,
      categoriesUpdated: 0,
      categoriesDeactivated: 0,
      itemsCreated: 0,
      itemsUpdated: 0,
      itemsDeactivated: 0,
      pricesCreated: 0,
      pricesUpdated: 0,
      pricesRemoved: 0,
   };

   if (rows.length === 0) return summary;

   await db.transaction().execute(async (trx) => {
      const now = new Date();

      /* Global pricing-region reference data */
      const regionRows = PRICING_REGION_COLUMNS.flatMap(({ column, fullLabel }) => {
         const label = headerReader.string(column);
         return label ? [{ code: toCode(label), label, full_label: fullLabel, column }] : [];
      });
      const regionCodeByColumn = new Map(regionRows.map((region) => [region.column, region.code]));
      const existingRegions = regionRows.length
         ? await trx.selectFrom('rate_set_support_item_pricing_region')
              .selectAll()
              .where('code', 'in', regionRows.map((region) => region.code))
              .execute()
         : [];
      const regionByCode = new Map(existingRegions.map((region) => [region.code, region]));

      for (const region of regionRows) {
         const existing = regionByCode.get(region.code);
         if (existing && (existing.label !== region.label || existing.full_label !== region.full_label)) {
            await trx.updateTable('rate_set_support_item_pricing_region')
               .set({ label: region.label, full_label: region.full_label })
               .where('code', '=', region.code).execute();
         } else if (!existing) {
            await trx.insertInto('rate_set_support_item_pricing_region')
               .values({ code: region.code, label: region.label, full_label: region.full_label })
               .onConflict((oc) => oc.column('code').doNothing()).execute();
         }
      }

      /* Global attribute types */
      const attributeTypeRows = ATTRIBUTE_COLUMNS.flatMap(({ column, code }) => {
         const label = headerReader.string(column);
         return label ? [{ code, label }] : [];
      });
      const existingAttributeTypes = attributeTypeRows.length
         ? await trx.selectFrom('rate_set_support_item_attribute_type')
              .selectAll()
              .where('code', 'in', attributeTypeRows.map((item) => item.code))
              .execute()
         : [];
      const attributeTypeByCode = new Map(existingAttributeTypes.map((item) => [item.code, item]));

      for (const attribute of attributeTypeRows) {
         const existing = attributeTypeByCode.get(attribute.code);
         if (existing && existing.label !== attribute.label) {
            await trx.updateTable('rate_set_support_item_attribute_type')
               .set({ label: attribute.label }).where('code', '=', attribute.code).execute();
         } else if (!existing) {
            await trx.insertInto('rate_set_support_item_attribute_type')
               .values(attribute).onConflict((oc) => oc.column('code').doNothing()).execute();
         }
      }

      /* Global support-item types */
      const distinctTypes = new Map<string, string>();
      for (const row of rows) {
         if (row.supportTypeLabel) distinctTypes.set(toCode(row.supportTypeLabel), row.supportTypeLabel);
      }
      const typeCodes = [...distinctTypes.keys()];
      const existingTypes = typeCodes.length
         ? await trx.selectFrom('rate_set_support_item_type')
              .selectAll().where('code', 'in', typeCodes).execute()
         : [];
      const typeByCode = new Map(existingTypes.map((type) => [type.code, type]));
      const typeIdByCode = new Map(existingTypes.map((type) => [type.code, type.id]));

      for (const [code, label] of distinctTypes) {
         const existing = typeByCode.get(code);
         if (existing && existing.label !== label) {
            await trx.updateTable('rate_set_support_item_type')
               .set({ label }).where('id', '=', existing.id).execute();
         } else if (!existing) {
            const inserted = await trx.insertInto('rate_set_support_item_type')
               .values({ code, label })
               .onConflict((oc) => oc.column('code').doNothing())
               .returning('id').executeTakeFirst();
            if (inserted) typeIdByCode.set(code, inserted.id);
            else {
               const concurrent = await trx.selectFrom('rate_set_support_item_type')
                  .select('id').where('code', '=', code).executeTakeFirstOrThrow();
               typeIdByCode.set(code, concurrent.id);
            }
         }
      }

      /* Categories: one read, batched upsert, one bulk deactivation */
      const distinctCategories = new Map<string, string>();
      for (const row of rows) distinctCategories.set(row.categoryNumber, row.categoryName);
      const sortedCategories = [...distinctCategories.entries()].sort((a, b) =>
         a[0].localeCompare(b[0], undefined, { numeric: true }),
      );
      const categoryNumbers = sortedCategories.map(([number]) => number);
      const existingCategories = await trx.selectFrom('rate_set_category')
         .selectAll().where('rate_set_id', '=', rateSetId)
         .where('category_number', 'in', categoryNumbers).execute();
      const categoryByNumber = new Map(existingCategories.map((item) => [item.category_number, item]));

      const categoryWrites: Array<{
         id?: number; rate_set_id: number; category_number: string; category_name: string;
         sorting: number; deleted_at?: Date | null; deactivated_at?: Date | null; updated_at?: Date;
      }> = [];

      sortedCategories.forEach(([categoryNumber, categoryName], index) => {
         const existing = categoryByNumber.get(categoryNumber);
         if (existing) {
            if (existing.category_name !== categoryName || existing.deleted_at || existing.deactivated_at) {
               summary.categoriesUpdated++;
            }
            categoryWrites.push({
               id: existing.id, rate_set_id: rateSetId, category_number: categoryNumber,
               category_name: categoryName, sorting: index + 1,
               deleted_at: null, deactivated_at: null, updated_at: now,
            });
         } else {
            categoryWrites.push({
               rate_set_id: rateSetId, category_number: categoryNumber,
               category_name: categoryName, sorting: index + 1,
            });
            summary.categoriesCreated++;
         }
      });

      await writeBatches(categoryWrites, (batch) =>
         trx.insertInto('rate_set_category').values(batch)
            .onConflict((oc) => oc.column('id').doUpdateSet({
               category_name: sql.ref('excluded.category_name'),
               sorting: sql.ref('excluded.sorting'),
               deleted_at: null, deactivated_at: null, updated_at: now,
            })).execute(),
      );

      const allCategories = await trx.selectFrom('rate_set_category')
         .selectAll().where('rate_set_id', '=', rateSetId)
         .where('deleted_at', 'is', null).execute();
      const categoryIdByNumber = new Map(allCategories.map((item) => [item.category_number, item.id]));
      const missingCategoryIds = allCategories
         .filter((item) => !distinctCategories.has(item.category_number))
         .map((item) => item.id);

      if (missingCategoryIds.length) {
         await trx.updateTable('rate_set_category').set({ deactivated_at: now })
            .where('id', 'in', missingCategoryIds).execute();
         summary.categoriesDeactivated = missingCategoryIds.length;
      }

      /* Support items: resolve all item IDs with one bulk upsert */
      const distinctItems = new Map<string, { row: ParsedRow; categoryId: number; sorting: number }>();
      let itemSort = 1;
      for (const row of rows) {
         const workbookKey = row.categoryNumber + '::' + row.itemNumber;
         if (distinctItems.has(workbookKey)) continue;
         const categoryId = categoryIdByNumber.get(row.categoryNumber);
         if (categoryId === undefined) throw new Error('Unable to resolve category ' + row.categoryNumber);
         distinctItems.set(workbookKey, { row, categoryId, sorting: itemSort++ });
      }

      const categoryIds = [...new Set([...distinctItems.values()].map((item) => item.categoryId))];
      const existingItems = await trx.selectFrom('rate_set_support_item')
         .selectAll().where('rate_set_id', '=', rateSetId)
         .where('category_id', 'in', categoryIds).execute();
      const existingItemByKey = new Map(existingItems.map((item) => [
         item.category_id + '::' + item.item_number, item,
      ]));

      const itemWrites: Array<{
         id?: number; rate_set_id: number; category_id: number; item_number: string;
         item_name: string; unit: string | null; sorting: number;
         deleted_at?: Date | null; deactivated_at?: Date | null; updated_at?: Date;
      }> = [];

      for (const { row, categoryId, sorting } of distinctItems.values()) {
         const existing = existingItemByKey.get(categoryId + '::' + row.itemNumber);
         if (existing) {
            if (existing.item_name !== row.itemName || existing.unit !== (row.unit || null) ||
                existing.deleted_at || existing.deactivated_at) summary.itemsUpdated++;
            itemWrites.push({
               id: existing.id, rate_set_id: rateSetId, category_id: categoryId,
               item_number: row.itemNumber, item_name: row.itemName, unit: row.unit || null,
               sorting, deleted_at: null, deactivated_at: null, updated_at: now,
            });
         } else {
            itemWrites.push({
               rate_set_id: rateSetId, category_id: categoryId,
               item_number: row.itemNumber, item_name: row.itemName,
               unit: row.unit || null, sorting,
            });
            summary.itemsCreated++;
         }
      }

      const itemIdByWorkbookKey = new Map<string, number>();
      const writtenItems: Array<{ id: number; category_id: number; item_number: string }> = [];
      for (const batch of chunks(itemWrites)) {
         const written = await trx.insertInto('rate_set_support_item').values(batch)
            .onConflict((oc) => oc.column('id').doUpdateSet({
               item_name: sql.ref('excluded.item_name'),
               unit: sql.ref('excluded.unit'),
               sorting: sql.ref('excluded.sorting'),
               deleted_at: null, deactivated_at: null, updated_at: now,
            }))
            .returning(['id', 'category_id', 'item_number']).execute();
         writtenItems.push(...written);
      }
      for (const item of writtenItems) {
         itemIdByWorkbookKey.set(item.category_id + '::' + item.item_number, item.id);
      }

      /* Attributes: read once, then batch only changed/new values */
      const itemIds = writtenItems.map((item) => item.id);
      const existingAttributes = itemIds.length
         ? await trx.selectFrom('rate_set_support_item_attribute')
              .selectAll().where('support_item_id', 'in', itemIds).execute()
         : [];
      const attributeByKey = new Map(existingAttributes.map((item) => [
         item.support_item_id + '::' + item.attribute_code, item,
      ]));
      const attributeWrites: Array<{
         id?: number; support_item_id: number; attribute_code: string; value: boolean;
      }> = [];

      for (const { row, categoryId } of distinctItems.values()) {
         const itemId = itemIdByWorkbookKey.get(categoryId + '::' + row.itemNumber);
         if (itemId === undefined) continue;
         for (const [code, value] of Object.entries(row.attributes)) {
            const existing = attributeByKey.get(itemId + '::' + code);
            if (existing && existing.value === value) continue;
            attributeWrites.push({
               ...(existing ? { id: existing.id } : {}),
               support_item_id: itemId, attribute_code: code, value,
            });
         }
      }

      await writeBatches(attributeWrites, (batch) =>
         trx.insertInto('rate_set_support_item_attribute').values(batch)
            .onConflict((oc) => oc.column('id').doUpdateSet({
               value: sql.ref('excluded.value'),
            })).execute(),
      );

      /* Prices: read once for this rate set, then batch changed/new prices */
      const existingPrices = await trx.selectFrom('rate_set_support_item_price')
         .selectAll().where('rate_set_id', '=', rateSetId).execute();
      const existingPriceByKey = new Map<string, (typeof existingPrices)[number]>();
      for (const price of existingPrices) {
         const key = makePriceKey(
            price.support_item_id, price.type_id, price.pricing_region_code,
            isoDateKey(price.start_date), price.end_date ? isoDateKey(price.end_date) : null,
         );
         if (!existingPriceByKey.has(key)) existingPriceByKey.set(key, price);
      }

      const seenPriceKeys = new Set<string>();
      const priceWrites: Array<{
         id?: number; rate_set_id: number; support_item_id: number; type_id: number | null;
         pricing_region_code: string; unit_price: string; start_date: string; end_date: string | null;
      }> = [];

      // Iterate all rows here: the same support item can have multiple effective-date ranges.
      for (const row of rows) {
         const workbookItemKey = row.categoryNumber + '::' + row.itemNumber;
         const itemId = itemIdByWorkbookKey.get(workbookItemKey);
         if (itemId === undefined) continue;
         const typeId = row.supportTypeLabel ? (typeIdByCode.get(toCode(row.supportTypeLabel)) ?? null) : null;

         for (const [column, unitPrice] of Object.entries(row.prices)) {
            if (unitPrice === null) continue;
            const regionCode = regionCodeByColumn.get(column);
            if (!regionCode) continue;

            const key = makePriceKey(itemId, typeId, regionCode, row.startDate, row.endDate);
            seenPriceKeys.add(key);
            const existing = existingPriceByKey.get(key);
            const roundedPrice = unitPrice.toFixed(4);

            if (existing && Number(existing.unit_price) === Number(roundedPrice)) continue;

            priceWrites.push({
               ...(existing ? { id: existing.id } : {}),
               rate_set_id: rateSetId, support_item_id: itemId, type_id: typeId,
               pricing_region_code: regionCode, unit_price: roundedPrice,
               start_date: row.startDate, end_date: row.endDate,
            });
            if (existing) summary.pricesUpdated++;
            else summary.pricesCreated++;
         }
      }

      await writeBatches(priceWrites, (batch) =>
         trx.insertInto('rate_set_support_item_price').values(batch)
            .onConflict((oc) => oc.column('id').doUpdateSet({
               unit_price: sql.ref('excluded.unit_price'), updated_at: now,
            })).execute(),
      );

      /* Deactivate missing items in one statement */
      const seenItemKeys = new Set([...distinctItems.keys()]);
      const allItems = await trx.selectFrom('rate_set_support_item')
         .select(['id', 'category_id', 'item_number'])
         .where('rate_set_id', '=', rateSetId)
         .where('deleted_at', 'is', null).execute();
      const missingItemIds = allItems.filter((item) =>
         !seenItemKeys.has(item.category_id + '::' + item.item_number),
      ).map((item) => item.id);

      if (missingItemIds.length) {
         await trx.updateTable('rate_set_support_item').set({ deactivated_at: now })
            .where('id', 'in', missingItemIds).execute();
         summary.itemsDeactivated = missingItemIds.length;
      }

      /* Delete missing prices in one statement */
      const missingPriceIds = existingPrices.filter((price) => {
         const key = makePriceKey(
            price.support_item_id, price.type_id, price.pricing_region_code,
            isoDateKey(price.start_date), price.end_date ? isoDateKey(price.end_date) : null,
         );
         return !seenPriceKeys.has(key);
      }).map((price) => price.id);

      if (missingPriceIds.length) {
         await trx.deleteFrom('rate_set_support_item_price')
            .where('id', 'in', missingPriceIds).execute();
         summary.pricesRemoved = missingPriceIds.length;
      }
   });

   return summary;
}
