import { router, json, error, db, ai, requireAuth } from '@appdeploy/sdk';

const safeKey = (value: string) =>
  value.replace(/[^A-Za-z0-9_.-]/g, '_').slice(0, 90) || 'user';
const table = (userId: string, kind: string) =>
  'hs_' + kind + '_' + safeKey(userId);

const backupKinds: Record<string, string> = {
  assets: 'assets',
  moves: 'moves',
  asset_updates: 'asset_updates',
  people: 'people',
  boarding_attendance: 'boarding_attendance',
  classes: 'structure_classes',
  blocks: 'structure_blocks',
  suppliers: 'structure_suppliers',
};

export const handler = router({
  'GET /api/backup/:kind': [requireAuth(), async ctx => {
    const tableKind = backupKinds[ctx.params.kind];
    if (!tableKind) return error('Unsupported backup data type', 400);
    const options: { limit: number; nextToken?: string } = { limit: 100 };
    if (ctx.query.nextToken) options.nextToken = ctx.query.nextToken;
    const page = await db.list(table(ctx.user!.userId, tableKind), options);
    return json({ kind: ctx.params.kind, records: page.items, nextToken: page.nextToken || null });
  }],
  'POST /api/backup/restore': [requireAuth(), async ctx => {
    const body = ctx.body as { kind?: string; records?: Array<Record<string, unknown>> };
    const tableKind = backupKinds[body.kind || ''];
    if (!tableKind || !Array.isArray(body.records) || body.records.length < 1 || body.records.length > 3) return error('Choose a supported data type and restore 1 to 3 records per request', 400);
    if (body.records.some(record => !record || typeof record.id !== 'string' || !record.id.trim())) return error('Every backup record must contain its original record ID', 400);
    const uid = ctx.user!.userId;
    const ids = body.records.map(record => String(record.id));
    if (new Set(ids).size !== ids.length) return error('Duplicate record IDs found in this restore batch', 400);
    const existing = await db.get(table(uid, tableKind), ids, { consistent: true });
    const updates: Array<{ id: string; record: Record<string, unknown> }> = [];
    const additions: Array<Record<string, unknown>> = [];
    body.records.forEach((record, index) => {
      const safeRecord = { ...record, ownerId: uid };
      if (existing[index]) updates.push({ id: ids[index], record: safeRecord });
      else additions.push({ ...safeRecord, id: ids[index] });
    });
    let updatedCount = 0;
    let addedCount = 0;
    if (updates.length) {
      const outcomes = await db.update(table(uid, tableKind), updates);
      updatedCount = outcomes.filter(Boolean).length;
      if (updatedCount !== updates.length) return error('Some existing records could not be restored. Retry the same backup file to reconcile the remaining records.', 409);
    }
    if (additions.length) {
      const outcomes = await db.tx.add(table(uid, tableKind), additions);
      addedCount = outcomes.filter(Boolean).length;
      if (addedCount !== additions.length) return error('Some missing records could not be restored. Retry the same backup file to reconcile the remaining records.', 409);
    }
    return json({ restored: updatedCount + addedCount, updated: updatedCount, added: addedCount });
  }],
  'GET /api/_healthcheck': [async () => json({ message: 'Success' })],
  'GET /api/inventory': [
    requireAuth(),
    async ctx => {
      try {
        const uid = ctx.user!.userId;
        const [assets, movements, legacyUpdates] = await Promise.all([
          db.list(table(uid, 'assets'), { limit: 100 }),
          db.list(table(uid, 'moves'), { limit: 100 }),
          db.list(table(uid, 'asset_updates'), { limit: 500 }),
        ]);
        // Backward compatibility: older versions stored updated quantities in
        // asset_updates rather than updating the main asset record.
        const recovered = new Map<string, Record<string, unknown>>();
        for (const entry of legacyUpdates.items as Array<Record<string, unknown>>) {
          if (typeof entry.assetId === 'string' && entry.updated && typeof entry.updated === 'object') {
            recovered.set(entry.assetId, entry.updated as Record<string, unknown>);
          }
        }
        const assetRows = (assets.items as Array<Record<string, unknown>>).map(asset => ({
          ...asset,
          ...(recovered.get(String(asset.id)) || {}),
          id: asset.id,
        }));
        return json({ assets: assetRows, movements: movements.items });
      } catch (err) {
        console.error('Inventory read failed', err);
        return json({
          assets: [],
          movements: [],
          message: 'No saved records yet.',
        });
      }
    },
  ],
  'GET /api/reports/summary': [requireAuth(), async ctx => { const uid = ctx.user!.userId; const [assets, people, moves, attendance, classes, blocks, suppliers] = await Promise.all([db.list(table(uid, 'assets'), { limit: 500 }), db.list(table(uid, 'people'), { limit: 500 }), db.list(table(uid, 'moves'), { limit: 500 }), db.list(table(uid, 'boarding_attendance'), { limit: 500 }), db.list(table(uid, 'structure_classes'), { limit: 200 }), db.list(table(uid, 'structure_blocks'), { limit: 200 }), db.list(table(uid, 'structure_suppliers'), { limit: 200 })]); const assetRows = assets.items as Array<Record<string, unknown>>; const peopleRows = people.items as Array<Record<string, unknown>>; return json({ summary: { totalItems: assetRows.length, totalQuantity: assetRows.reduce((n, r) => n + (Number(r.quantity) || 0), 0), students: peopleRows.filter(r => r.recordType === 'Student').length, staff: peopleRows.filter(r => r.recordType === 'Staff').length, movements: moves.items.length, boardingAttendanceEvents: attendance.items.length, classes: classes.items.length, blocks: blocks.items.length, suppliers: suppliers.items.length, damagedItems: assetRows.filter(r => String(r.condition || '').toLowerCase().includes('damag')).length }, generatedAt: new Date().toISOString() }); }],
  'POST /api/inventory/assets': [
    requireAuth(),
    async ctx => {
      const body = ctx.body as Record<string, unknown>;
      if (typeof body.name !== 'string' || !body.name.trim())
        return error('Item name is required', 400);
      const id =
        typeof body.id === 'string'
          ? body.id.replace(/[^A-Za-z0-9_.-]/g, '_').slice(0, 100)
          : '';
      const record = {
        ...body,
        ownerId: ctx.user!.userId,
        createdAt: new Date().toISOString(),
      };
      const [savedId] = await db.add(table(ctx.user!.userId, 'assets'), [
        { ...record, id: id || undefined },
      ]);
      if (!savedId)
        return error(
          'Could not save item. Check whether this item ID already exists.',
          409
        );
      return json({ id: savedId, saved: true }, 201);
    },
  ],
  'PUT /api/inventory/assets/:id': [requireAuth(), async ctx => { const body = ctx.body as Record<string, unknown>; if (typeof body.name !== 'string' || !body.name.trim()) return error('Item name is required', 400); const [ok] = await db.update(table(ctx.user!.userId, 'assets'), [{ id: ctx.params.id, record: { ...body, ownerId: ctx.user!.userId, updatedAt: new Date().toISOString() } }]); return ok ? json({ updated: true, id: ctx.params.id }) : error('Item not found or could not be updated', 404); }],
  'DELETE /api/inventory/assets/:id': [requireAuth(), async ctx => { const [ok] = await db.delete(table(ctx.user!.userId, 'assets'), [ctx.params.id]); return ok ? json({ deleted: true }) : error('Item not found', 404); }],
  'POST /api/inventory/movements': [
    requireAuth(),
    async ctx => {
      const body = ctx.body as Record<string, unknown>;
      if (typeof body.assetName !== 'string' || !body.assetName.trim())
        return error('Item is required', 400);
      const uid = ctx.user!.userId;
      const assetId = typeof body.assetId === 'string' ? body.assetId : '';
      const updated = body.updatedAsset && typeof body.updatedAsset === 'object'
        ? body.updatedAsset as Record<string, unknown>
        : null;
      // Persist the new stock balance in the canonical asset record. The
      // movement log alone is not sufficient to maintain current stock.
      if (assetId && updated) {
        const [assetSaved] = await db.update(table(uid, 'assets'), [{
          id: assetId,
          record: { ...updated, ownerId: uid, updatedAt: new Date().toISOString() },
        }]);
        if (!assetSaved) return error('Stock movement was not recorded because the asset balance could not be updated. Refresh and try again.', 409);
      }
      const [savedId] = await db.add(table(uid, 'moves'), [{
        ...body,
        ownerId: uid,
        createdAt: new Date().toISOString(),
      }]);
      if (!savedId) return error('Stock balance may have changed, but the movement log could not be saved. Contact an administrator before retrying.', 500);
      return json({ id: savedId, saved: true, stockUpdated: Boolean(assetId && updated) }, 201);
    },
  ],
  'GET /api/inventory/movements': [requireAuth(), async ctx => { const rows = await db.list(table(ctx.user!.userId, 'moves'), { limit: 500 }); return json({ movements: rows.items }); }],
  'GET /api/people': [requireAuth(), async ctx => { const people = await db.list(table(ctx.user!.userId, 'people'), { limit: 200 }); return json({ people: people.items }); }],
  'POST /api/people': [requireAuth(), async ctx => { const body = ctx.body as Record<string, unknown>; const uid = ctx.user!.userId; if (typeof body.fullName !== 'string' || !body.fullName.trim()) return error('Full name is required', 400); if (body.recordType !== 'Student' && body.recordType !== 'Staff') return error('Record type must be Student or Staff', 400); if (body.recordType === 'Student' && (typeof body.admissionNo !== 'string' || !body.admissionNo.trim())) return error('Admission number is required', 400); if (body.recordType === 'Student' && body.residence === 'Boarder' && (typeof body.dormitory !== 'string' || !body.dormitory.trim())) return error('Dormitory is required for a boarder', 400); if (body.recordType === 'Staff' && (typeof body.staffNo !== 'string' || !body.staffNo.trim())) return error('TSC or national ID number is required', 400); if (body.recordType === 'Student') { const existing = await db.list(table(uid, 'people'), { limit: 500 }); if ((existing.items as Array<Record<string, unknown>>).some(row => row.recordType === 'Student' && String(row.admissionNo || '').trim().toLowerCase() === String(body.admissionNo).trim().toLowerCase())) return error('Admission number already exists. Use a unique admission number.', 409); } const [id] = await db.add(table(uid, 'people'), [{ ...body, dormitory: body.residence === 'Day scholar' ? '' : body.dormitory || '', parentName: typeof body.parentName === 'string' ? body.parentName.trim() : '', ownerId: uid, createdAt: new Date().toISOString() }]); if (!id) return error('Could not save person. Please retry.', 409); return json({ id, saved: true }, 201); }],
  'PUT /api/people/:id': [requireAuth(), async ctx => { const body = ctx.body as Record<string, unknown>; const uid = ctx.user!.userId; if (typeof body.fullName !== 'string' || !body.fullName.trim()) return error('Full name is required', 400); if (body.recordType === 'Student' && (typeof body.admissionNo !== 'string' || !body.admissionNo.trim())) return error('Admission number is required', 400); if (body.recordType === 'Student' && body.residence === 'Boarder' && (typeof body.dormitory !== 'string' || !body.dormitory.trim())) return error('Dormitory is required for a boarder', 400); if (body.recordType === 'Student') { const existing = await db.list(table(uid, 'people'), { limit: 500 }); if ((existing.items as Array<Record<string, unknown>>).some(row => row.id !== ctx.params.id && row.recordType === 'Student' && String(row.admissionNo || '').trim().toLowerCase() === String(body.admissionNo).trim().toLowerCase())) return error('Admission number already exists. Use a unique admission number.', 409); } const record = { ...body, dormitory: body.residence === 'Day scholar' ? '' : body.dormitory || '', parentName: typeof body.parentName === 'string' ? body.parentName.trim() : '', ownerId: uid, updatedAt: new Date().toISOString() }; const [ok] = await db.update(table(uid, 'people'), [{ id: ctx.params.id, record }]); return ok ? json({ updated: true, id: ctx.params.id }) : error('Person not found or could not be updated', 404); }],
  'DELETE /api/people/:id': [requireAuth(), async ctx => { const [ok] = await db.delete(table(ctx.user!.userId, 'people'), [ctx.params.id]); return ok ? json({ deleted: true }) : error('Person not found', 404); }],
  'GET /api/boarding-attendance': [requireAuth(), async ctx => { const rows = await db.list(table(ctx.user!.userId, 'boarding_attendance'), { limit: 500 }); return json({ attendance: rows.items }); }],
  'PUT /api/boarding-attendance/:id': [requireAuth(), async ctx => { const body = ctx.body as Record<string, unknown>; if (body.event !== 'Check-in' && body.event !== 'Check-out') return error('Invalid attendance event', 400); const [ok] = await db.update(table(ctx.user!.userId, 'boarding_attendance'), [{ id: ctx.params.id, record: { ...body, ownerId: ctx.user!.userId, updatedAt: new Date().toISOString() } }]); return ok ? json({ updated: true, id: ctx.params.id }) : error('Attendance record not found', 404); }],
  'DELETE /api/boarding-attendance/:id': [requireAuth(), async ctx => { const [ok] = await db.delete(table(ctx.user!.userId, 'boarding_attendance'), [ctx.params.id]); return ok ? json({ deleted: true }) : error('Attendance record not found', 404); }],
  'GET /api/school-structure': [requireAuth(), async ctx => { const kinds = ['classes', 'blocks', 'suppliers']; const results = await Promise.all(kinds.map(kind => db.list(table(ctx.user!.userId, 'structure_' + kind), { limit: 200 }))); return json({ structure: { classes: results[0].items, blocks: results[1].items, suppliers: results[2].items } }); }],
  'PUT /api/school-structure/:kind/:id': [requireAuth(), async ctx => { if (!['classes', 'blocks', 'suppliers'].includes(ctx.params.kind)) return error('Invalid structure type', 400); const body = ctx.body as Record<string, unknown>; if (typeof body.name !== 'string' || !body.name.trim()) return error('Name is required', 400); const [ok] = await db.update(table(ctx.user!.userId, 'structure_' + ctx.params.kind), [{ id: ctx.params.id, record: { ...body, ownerId: ctx.user!.userId, updatedAt: new Date().toISOString() } }]); return ok ? json({ updated: true, id: ctx.params.id }) : error('Record not found', 404); }],
  'DELETE /api/school-structure/:kind/:id': [requireAuth(), async ctx => { if (!['classes', 'blocks', 'suppliers'].includes(ctx.params.kind)) return error('Invalid structure type', 400); const [ok] = await db.delete(table(ctx.user!.userId, 'structure_' + ctx.params.kind), [ctx.params.id]); return ok ? json({ deleted: true }) : error('Record not found', 404); }],
  'POST /api/school-structure': [requireAuth(), async ctx => { const body = ctx.body as { kind?: string; record?: Record<string, unknown> }; if (!['classes', 'blocks', 'suppliers'].includes(body.kind || '') || !body.record || typeof body.record.name !== 'string' || !body.record.name.trim()) return error('A valid structure type and name are required', 400); const [id] = await db.add(table(ctx.user!.userId, 'structure_' + body.kind), [{ ...body.record, id: typeof body.record.id === 'string' ? body.record.id : undefined, ownerId: ctx.user!.userId, createdAt: new Date().toISOString() }]); if (!id) return error('Could not save school structure record', 409); return json({ id, saved: true }, 201); }],
  'GET /api/suppliers': [requireAuth(), async ctx => { const rows = await db.list(table(ctx.user!.userId, 'structure_suppliers'), { limit: 500 }); return json({ suppliers: rows.items }); }],
  'POST /api/suppliers': [requireAuth(), async ctx => { const body = ctx.body as Record<string, unknown>; if (typeof body.name !== 'string' || !body.name.trim()) return error('Supplier name is required', 400); const [id] = await db.add(table(ctx.user!.userId, 'structure_suppliers'), [{ ...body, ownerId: ctx.user!.userId, createdAt: new Date().toISOString() }]); return id ? json({ id, saved: true }, 201) : error('Could not save supplier', 409); }],
  'GET /api/classes': [requireAuth(), async ctx => { const rows = await db.list(table(ctx.user!.userId, 'structure_classes'), { limit: 500 }); return json({ classes: rows.items }); }],
  'GET /api/blocks': [requireAuth(), async ctx => { const rows = await db.list(table(ctx.user!.userId, 'structure_blocks'), { limit: 500 }); return json({ blocks: rows.items }); }],
  'POST /api/bulk-import': [requireAuth(), async ctx => { const body = ctx.body as { kind?: string; records?: Array<Record<string, unknown>> }; if (!['students', 'items'].includes(body.kind || '') || !Array.isArray(body.records) || body.records.length < 1 || body.records.length > 200) return error('Choose students or items and import between 1 and 200 records per batch', 400); const records = body.records; if (body.kind === 'students' && records.some(r => typeof r.fullName !== 'string' || !r.fullName.trim() || typeof r.admissionNo !== 'string' || !r.admissionNo.trim())) return error('Every student requires a full name and admission number', 400); if (body.kind === 'items' && records.some(r => typeof r.name !== 'string' || !r.name.trim() || !Number.isFinite(Number(r.quantity)) || Number(r.quantity) < 0)) return error('Every item requires a name and a valid non-negative quantity', 400); const target = body.kind === 'students' ? 'people' : 'assets'; const saved = await db.add(table(ctx.user!.userId, target), records.map(record => ({ ...record, ownerId: ctx.user!.userId, createdAt: new Date().toISOString() }))); const imported = saved.filter(Boolean).length; if (imported !== records.length) return json({ imported, total: records.length, warning: 'Some rows could not be saved; review duplicates and retry only failed rows.' }, 207); return json({ imported, total: records.length, saved: true }, 201); }],
  'POST /api/boarding-attendance': [requireAuth(), async ctx => { const body = ctx.body as Record<string, unknown>; if (typeof body.admissionNo !== 'string' || !body.admissionNo.trim()) return error('Admission number is required', 400); if (body.event !== 'Check-in' && body.event !== 'Check-out') return error('Invalid attendance event', 400); if (body.method !== 'Biometric device' || typeof body.deviceRef !== 'string' || !body.deviceRef.trim()) return error('A registered biometric device reference is required', 400); const [id] = await db.add(table(ctx.user!.userId, 'boarding_attendance'), [{ ...body, ownerId: ctx.user!.userId, recordedAt: new Date().toISOString() }]); if (!id) return error('Could not save attendance event', 409); return json({ id, saved: true }, 201); }],
  'POST /api/ai/stock-advice': [
    requireAuth(),
    async ctx => {
      const body = ctx.body as {
        assets?: Array<{
          name?: string;
          category?: string;
          quantity?: number;
          condition?: string;
          location?: string;
          cost?: number;
        }>;
        movements?: Array<{
          assetName?: string;
          action?: string;
          quantity?: number;
        }>;
      };
      const assets = Array.isArray(body.assets) ? body.assets.slice(0, 80) : [];
      const movements = Array.isArray(body.movements)
        ? body.movements.slice(0, 30)
        : [];
      try {
        const result = await ai.generate({
          system:
            'You are a careful school inventory analyst for a Kenyan school. Give concise practical recommendations based only on supplied records. Never invent facts. Prioritize shortages, damaged assets, textbook availability, and audit controls. Use Kenyan shillings where relevant.',
          prompt:
            'Review this school inventory and give 3-5 prioritized actions in plain text. Inventory: ' +
            JSON.stringify(assets) +
            ' Recent movements: ' +
            JSON.stringify(movements),
          maxTokens: 500,
          thinkingMode: 'FAST',
        });
        return json({ advice: result.text });
      } catch (err) {
        console.error('AI stock advice failed', err);
        return error('AI stock advice is temporarily unavailable.', 503);
      }
    },
  ],
});
