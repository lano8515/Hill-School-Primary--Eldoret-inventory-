import { useEffect, useMemo, useState } from 'react';
import { api, auth } from '@appdeploy/client';
import * as XLSX from 'xlsx';
import JsBarcode from 'jsbarcode';
import {
  Activity,
  Archive,
  ArrowDownToLine,
  ArrowLeftRight,
  BarChart3,
  BookOpen,
  Boxes,
  CheckCircle2,
  UserPlus,
  UserRound,
  Fingerprint,
  ChevronDown,
  CircleHelp,
  Download,
  Database,
  FileText,
  GraduationCap,
  LayoutDashboard,
  Menu,
  PackagePlus,
  Plus,
  QrCode,
  Search,
  Settings2,
  ShieldCheck,
  Sparkles,
  Users,
  Warehouse,
  Wrench,
  X,
  AlertTriangle,
  Building2,
  School,
  Truck,
  Upload,
} from 'lucide-react';

type Asset = {
  id: string;
  name: string;
  category: string;
  quantity: number;
  condition: string;
  location: string;
  cost: number;
  barcode: string;
  kind: string;
  updatedAt: string;
};
type Movement = {
  id: string;
  assetName: string;
  action: string;
  person: string;
  quantity: number;
  date: string;
  barcode: string;
};
type Person = { id: string; recordType: 'Student' | 'Staff'; fullName: string; admissionNo: string; staffNo: string; staffType: 'Teaching' | 'Non-teaching'; gender: string; grade: string; stream: string; residence: 'Boarder' | 'Day scholar' | 'Staff'; dormitory?: string; parentName?: string; phone: string; role: string; status: 'Active' | 'Inactive' | 'Transferred'; createdAt: string; biometricRef?: string };
type AttendanceRecord = { id: string; admissionNo: string; fullName: string; grade: string; stream: string; event: 'Check-in' | 'Check-out'; timestamp: string; method: 'Biometric device'; deviceRef: string };
type RecordData = Asset | Movement | Person | AttendanceRecord;

const starterAssets: Asset[] = [
  {
    id: 'seed-1',
    name: 'Grade 7 Mathematics Textbook',
    category: 'Textbooks',
    quantity: 184,
    condition: 'Good',
    location: 'Store A',
    cost: 850,
    barcode: 'BK-G7-MATH-001',
    kind: 'Returnable',
    updatedAt: 'Today',
  },
  {
    id: 'seed-2',
    name: 'Student wooden desk',
    category: 'Furniture',
    quantity: 96,
    condition: 'Good',
    location: 'Grade 6 East',
    cost: 4200,
    barcode: 'FUR-DESK-096',
    kind: 'Returnable',
    updatedAt: 'Today',
  },
  {
    id: 'seed-3',
    name: 'HP ProDesk computer',
    category: 'ICT',
    quantity: 28,
    condition: 'Good',
    location: 'Computer Lab',
    cost: 65000,
    barcode: 'ICT-HP-028',
    kind: 'Returnable',
    updatedAt: 'Yesterday',
  },
  {
    id: 'seed-4',
    name: 'Metallic bed',
    category: 'Boarding',
    quantity: 12,
    condition: 'Needs repair',
    location: 'Block B · Dorm 2',
    cost: 11500,
    barcode: 'BED-B-012',
    kind: 'Returnable',
    updatedAt: 'Yesterday',
  },
  {
    id: 'seed-5',
    name: 'A4 photocopy paper',
    category: 'Stationery',
    quantity: 7,
    condition: 'Good',
    location: 'Admin Store',
    cost: 650,
    barcode: 'ST-A4-007',
    kind: 'Consumable',
    updatedAt: 'Today',
  },
  {
    id: 'seed-6',
    name: 'Grade 8 Integrated Science',
    category: 'Textbooks',
    quantity: 9,
    condition: 'Good',
    location: 'Library',
    cost: 920,
    barcode: 'BK-G8-SCI-009',
    kind: 'Returnable',
    updatedAt: 'Today',
  },
];
const starterMovements: Movement[] = [
  {
    id: 'mv-1',
    assetName: 'Grade 7 Mathematics Textbook',
    action: 'Issued',
    person: 'Grade 7 · Stream A',
    quantity: 12,
    date: 'Today, 09:42',
    barcode: 'BK-G7-MATH-001',
  },
  {
    id: 'mv-2',
    assetName: 'Student wooden desk',
    action: 'Allocated',
    person: 'Grade 6 East',
    quantity: 4,
    date: 'Today, 08:16',
    barcode: 'FUR-DESK-096',
  },
  {
    id: 'mv-3',
    assetName: 'HP ProDesk computer',
    action: 'Maintenance',
    person: 'ICT Department',
    quantity: 1,
    date: 'Yesterday',
    barcode: 'ICT-HP-028',
  },
];

const money = (n: number) =>
  new Intl.NumberFormat('en-KE', {
    style: 'currency',
    currency: 'KES',
    maximumFractionDigits: 0,
  }).format(n);
const nav = [
  { label: 'Overview', icon: LayoutDashboard },
  { label: 'Classes & streams', icon: School },
  { label: 'Blocks & dormitories', icon: Building2 },
  { label: 'Suppliers', icon: Truck },
  { label: 'Bulk import', icon: Upload },
  { label: 'Inventory', icon: Boxes },
  { label: 'Textbooks', icon: BookOpen },
  { label: 'Furniture coding', icon: QrCode },
  { label: 'Issue & return', icon: ArrowLeftRight },
  { label: 'Locations', icon: Warehouse },
  { label: 'People', icon: Users },
  { label: 'Boarding attendance', icon: Fingerprint },
  { label: 'Reports', icon: BarChart3 },
  { label: 'Data backup', icon: Database },
];

function App() {
  const [signedInUser, setSignedInUser] = useState<Awaited<ReturnType<typeof auth.getUser>>>(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [authError, setAuthError] = useState('');
  const [signingIn, setSigningIn] = useState(false);
  const [section, setSection] = useState('Overview');
  const [assets, setAssets] = useState<Asset[]>(starterAssets);
  const [movements, setMovements] = useState<Movement[]>(starterMovements);
  const [people, setPeople] = useState<Person[]>([]);
  const [peopleType, setPeopleType] = useState<'All' | 'Student' | 'Staff'>('Student');
  const [personSearch, setPersonSearch] = useState('');
  const [attendance, setAttendance] = useState<AttendanceRecord[]>([]);
  const [attendanceSearch, setAttendanceSearch] = useState('');
  const [attendanceEvent, setAttendanceEvent] = useState<'Check-in' | 'Check-out'>('Check-in');
  const [attendanceAdmission, setAttendanceAdmission] = useState('');
  const [attendanceDeviceRef, setAttendanceDeviceRef] = useState('');
  const [attendanceBusy, setAttendanceBusy] = useState(false);
  type StructureRow = { id: string; name: string; detail?: string; parent?: string; streams?: string[]; dormitories?: string[]; contact?: string; phone?: string; email?: string; address?: string; supplied?: string };
  const [schoolStructure, setSchoolStructure] = useState<Record<string, StructureRow[]>>({ classes: [], blocks: [], suppliers: [] });
  const [structureForm, setStructureForm] = useState({ name: '', detail: '', parent: '', streams: 'A,B,C,D', dormitories: 'Dormitory 1,Dormitory 2', contact: '', phone: '', email: '', address: '', supplied: '' });
  const [importType, setImportType] = useState<'students' | 'items'>('students');
  const [importText, setImportText] = useState('');
  const [excelRows, setExcelRows] = useState<Record<string, string>[]>([]);
  const [excelFileName, setExcelFileName] = useState('');
  const [importBusy, setImportBusy] = useState(false);
  const [furnitureGrade, setFurnitureGrade] = useState('G8');
  const [furnitureStream, setFurnitureStream] = useState('N');
  const [furnitureItem, setFurnitureItem] = useState('CH');
  const [furnitureCount, setFurnitureCount] = useState('30');
  const [furnitureCondition, setFurnitureCondition] = useState('Good');
  const [furnitureBusy, setFurnitureBusy] = useState(false);
  const [furnitureFilterGrade, setFurnitureFilterGrade] = useState('All grades');
  const [furnitureFilterStream, setFurnitureFilterStream] = useState('All streams');
  const [furnitureFilterItem, setFurnitureFilterItem] = useState('All items');
  const [childDrafts, setChildDrafts] = useState<Record<string, string>>({});
  const [showPersonForm, setShowPersonForm] = useState(false);
  const [editingPersonId, setEditingPersonId] = useState<string | null>(null);
  const [personForm, setPersonForm] = useState({ recordType: 'Student' as 'Student' | 'Staff', fullName: '', admissionNo: '', staffNo: '', staffType: 'Teaching' as 'Teaching' | 'Non-teaching', gender: '', grade: 'Grade 1', stream: 'A', residence: 'Day scholar' as 'Boarder' | 'Day scholar' | 'Staff', dormitory: '', parentName: '', phone: '', role: '', status: 'Active' as 'Active' | 'Inactive' | 'Transferred', biometricRef: '' });
  const [search, setSearch] = useState('');
  const [category, setCategory] = useState('All categories');
  const [showForm, setShowForm] = useState(false);
  const [showMove, setShowMove] = useState(false);
  const [aiText, setAiText] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
  const [notice, setNotice] = useState('');
  const [loaded, setLoaded] = useState(false);
  const [backupBusy, setBackupBusy] = useState(false);
  const [restoreBusy, setRestoreBusy] = useState(false);
  const [backupStatus, setBackupStatus] = useState('');
  const [form, setForm] = useState({
    name: '',
    category: 'Textbooks',
    quantity: '1',
    condition: 'Good',
    location: 'Main Store',
    cost: '',
    barcode: '',
    kind: 'Returnable',
  });
  const [move, setMove] = useState({
    barcode: '',
    action: 'Issued',
    person: '',
    quantity: '1',
  });

  useEffect(() => {
    let active = true;
    auth.getUser()
      .then(user => { if (active) setSignedInUser(user); })
      .catch(() => { if (active) setSignedInUser(null); })
      .finally(() => { if (active) setAuthLoading(false); });
    return () => { active = false; };
  }, []);

  async function handleSignIn() {
    setSigningIn(true);
    setAuthError('');
    try {
      const result = await auth.signIn({ scope: 'openid email profile offline_access' });
      setSignedInUser(result.user);
    } catch (err) {
      const code = (err as { code?: string })?.code;
      setAuthError(code === 'popup_blocked' ? 'Your browser blocked the sign-in popup. Allow popups for this site and try again.' : code === 'popup_closed' ? 'Sign-in was cancelled. Select Sign in to try again.' : 'Sign-in could not be completed. Please try again.');
    } finally {
      setSigningIn(false);
    }
  }

  useEffect(() => {
    if (section !== 'Furniture coding') return;
    document.querySelectorAll<SVGSVGElement>('svg[data-code128]').forEach(svg => {
      const code = svg.dataset.code128;
      if (!code) return;
      try { JsBarcode(svg, code, { format: 'CODE128', displayValue: false, margin: 0, height: 32, width: 1.25 }); } catch { /* Invalid labels remain readable as text. */ }
    });
  }, [section, assets, furnitureFilterGrade, furnitureFilterStream, furnitureFilterItem]);

  useEffect(() => {
    if (!signedInUser) { setLoaded(true); return; }
    api
      .get('/api/inventory')
      .then(({ data }) => {
        if (Array.isArray(data.assets) && data.assets.length)
          setAssets(data.assets);
        if (Array.isArray(data.movements) && data.movements.length)
          setMovements(data.movements);
        if (Array.isArray(data.people)) setPeople(data.people);
        setLoaded(true);
      })
      .catch(() => setLoaded(true));
    api.get('/api/people').then(({ data }) => { if (Array.isArray(data.people)) setPeople(data.people); }).catch(() => {});
    api.get('/api/boarding-attendance').then(({ data }) => { if (Array.isArray(data.attendance)) setAttendance(data.attendance); }).catch(() => {});
    api.get('/api/school-structure').then(({ data }) => { if (data.structure) setSchoolStructure({ classes: data.structure.classes || [], blocks: data.structure.blocks || [], suppliers: data.structure.suppliers || [] }); }).catch(() => {});
  }, [signedInUser?.userId]);

  const filtered = useMemo(
    () =>
      assets.filter(
        a =>
          (category === 'All categories' || a.category === category) &&
          [a.name, a.category, a.location, a.barcode].some(v =>
            v.toLowerCase().includes(search.toLowerCase())
          )
      ),
    [assets, category, search]
  );
  const lowStock = assets.filter(a => a.quantity <= 10);
  const value = assets.reduce((sum, a) => sum + a.quantity * a.cost, 0);

  async function persistAsset(asset: Asset) {
    try {
      await api.post('/api/inventory/assets', asset);
    } catch {
      setNotice(
        'Saved in this session. Sign in/backend storage may need setup before shared persistence is available.'
      );
    }
  }
  function addAsset(e: React.FormEvent) {
    e.preventDefault();
    const asset: Asset = {
      ...form,
      id: 'asset-' + Date.now(),
      quantity: Math.max(0, Number(form.quantity)),
      cost: Math.max(0, Number(form.cost)),
      updatedAt: 'Just now',
      barcode:
        form.barcode.trim() ||
        'HS-' + Math.random().toString(36).slice(2, 8).toUpperCase(),
    };
    setAssets(prev => [asset, ...prev]);
    void persistAsset(asset);
    setShowForm(false);
    setForm({
      name: '',
      category: 'Textbooks',
      quantity: '1',
      condition: 'Good',
      location: 'Main Store',
      cost: '',
      barcode: '',
      kind: 'Returnable',
    });
    setNotice('Inventory item added.');
  }
  async function saveMovement(e: React.FormEvent) {
    e.preventDefault();
    const found = assets.find(
      a => a.barcode.toLowerCase() === move.barcode.trim().toLowerCase()
    );
    if (!found) {
      setNotice('Barcode not found. Check the code or add the item first.');
      return;
    }
    const qty = Math.max(1, Number(move.quantity));
    if (found.kind === 'Consumable' && move.action === 'Returned') {
      setNotice('Consumable items cannot be returned to stock. Record a stock receipt or adjustment instead.');
      return;
    }
    const recipient = people.find(p => p.recordType === 'Student' && p.admissionNo.toLowerCase() === move.person.trim().toLowerCase());
    if (recipient && recipient.residence === 'Day scholar' && (found.category !== 'Furniture' || !/^Grade \\d+/i.test(found.location))) {
      setNotice('Day scholars may only be issued furniture assigned to their classroom. This item/location is not eligible.');
      return;
    }
    if (recipient && recipient.residence === 'Boarder' && !/(Grade \\d+|class|block|dorm|store|library)/i.test(found.location)) {
      setNotice('For boarders, issue items only to a registered classroom or boarding location.');
      return;
    }
    if (
      (move.action === 'Issued' || move.action === 'Allocated') &&
      qty > found.quantity
    ) {
      setNotice('Not enough stock available for that transaction.');
      return;
    }
    const updated = {
      ...found,
      quantity:
        found.quantity +
        (move.action === 'Returned'
          ? qty
          : move.action === 'Issued' || move.action === 'Allocated'
            ? -qty
            : 0),
      updatedAt: 'Just now',
    };
    setAssets(prev => prev.map(a => (a.id === found.id ? updated : a)));
    const record: Movement = {
      id: 'mv-' + Date.now(),
      assetName: found.name,
      action: move.action,
      person: move.person || 'Not specified',
      quantity: qty,
      date: new Date().toLocaleString(),
      barcode: found.barcode,
    };
    setMovements(prev => [record, ...prev]);
    try {
      await api.post('/api/inventory/movements', {
        ...record,
        assetId: found.id,
        updatedAsset: updated,
      });
    } catch {}
    setShowMove(false);
    setMove({ barcode: '', action: 'Issued', person: '', quantity: '1' });
    setNotice(move.action + ' transaction recorded.');
  }
  function resetPersonForm() {
    setEditingPersonId(null);
    setPersonForm({ recordType: 'Student', fullName: '', admissionNo: '', staffNo: '', staffType: 'Teaching', gender: '', grade: 'Grade 1', stream: 'A', residence: 'Day scholar', dormitory: '', parentName: '', phone: '', role: '', status: 'Active', biometricRef: '' });
  }
  function openPersonEditor(person: Person) {
    setEditingPersonId(person.id);
    setPersonForm({ recordType: person.recordType, fullName: person.fullName || '', admissionNo: person.admissionNo || '', staffNo: person.staffNo || '', staffType: person.staffType || 'Teaching', gender: person.gender || '', grade: person.grade || 'Grade 1', stream: person.stream || 'A', residence: person.residence || 'Day scholar', dormitory: person.dormitory || '', parentName: person.parentName || '', phone: person.phone || '', role: person.role || '', status: person.status || 'Active', biometricRef: person.biometricRef || '' });
    setShowPersonForm(true);
  }
  async function savePerson(e: React.FormEvent) {
    e.preventDefault();
    const existing = editingPersonId ? people.find(p => p.id === editingPersonId) : undefined;
    const person: Person = { ...personForm, id: editingPersonId || 'person-' + Date.now(), createdAt: existing?.createdAt || new Date().toISOString() };
    if (!person.fullName.trim()) { setNotice('Learner or staff name is required.'); return; }
    if (person.recordType === 'Student' && !person.admissionNo.trim()) { setNotice('Admission number is required for learners.'); return; }
    if (person.recordType === 'Student' && person.residence === 'Boarder' && !person.dormitory.trim()) { setNotice('Select a dormitory for a boarder before saving.'); return; }
    if (person.recordType === 'Student' && person.residence === 'Day scholar') person.dormitory = '';
    if (person.recordType === 'Staff' && !person.staffNo.trim()) { setNotice(person.staffType === 'Teaching' ? 'TSC number is required for teaching staff.' : 'National ID number is required for non-teaching staff.'); return; }
    try {
      if (editingPersonId) { await api.put('/api/people/' + encodeURIComponent(editingPersonId), person); setPeople(prev => prev.map(p => p.id === editingPersonId ? person : p)); setNotice('Record updated successfully.'); }
      else { await api.post('/api/people', person); setPeople(prev => [person, ...prev]); setNotice(person.recordType === 'Student' ? 'Learner registered successfully.' : 'Staff registered successfully.'); }
      setShowPersonForm(false); resetPersonForm();
    } catch { setNotice('The server did not confirm saving. Check for a duplicate admission number and try again.'); }
  }
  async function recordAttendance(e: React.FormEvent) {
    e.preventDefault();
    const learner = people.find(p => p.recordType === 'Student' && p.admissionNo.trim().toLowerCase() === attendanceAdmission.trim().toLowerCase() && p.residence === 'Boarder' && p.status === 'Active');
    if (!learner) { setNotice('Only active registered boarders can be recorded in boarding attendance. Check admission number and boarder status.'); return; }
    if (!learner.biometricRef) { setNotice('This boarder has no biometric device reference. Register the device-assigned user ID in the student record first.'); return; }
    if (!attendanceDeviceRef.trim() || attendanceDeviceRef.trim() !== learner.biometricRef) { setNotice('Biometric device reference does not match the registered learner. No attendance recorded.'); return; }
    setAttendanceBusy(true);
    const record: AttendanceRecord = { id: 'att-' + Date.now(), admissionNo: learner.admissionNo, fullName: learner.fullName, grade: learner.grade, stream: learner.stream, event: attendanceEvent, timestamp: new Date().toISOString(), method: 'Biometric device', deviceRef: learner.biometricRef };
    try { await api.post('/api/boarding-attendance', record); setAttendance(prev => [record, ...prev]); setNotice('Boarding attendance recorded for ' + learner.fullName + '.'); setAttendanceAdmission(''); setAttendanceDeviceRef(''); }
    catch { setNotice('Attendance was not confirmed by the server. Check authentication and biometric-device integration.'); }
    finally { setAttendanceBusy(false); }
  }
  async function addStructureChild(parent: StructureRow, key: 'classes' | 'blocks', child: string) {
    const value = child.trim();
    if (!value) { setNotice(key === 'classes' ? 'Enter a stream name first.' : 'Enter a dormitory name first.'); return; }
    const field = key === 'classes' ? 'streams' : 'dormitories';
    const current = (parent[field] || []) as string[];
    if (current.some(v => v.toLowerCase() === value.toLowerCase())) { setNotice('That ' + (key === 'classes' ? 'stream' : 'dormitory') + ' already exists under this ' + (key === 'classes' ? 'class.' : 'block.')); return; }
    const updated: StructureRow = { ...parent, [field]: [...current, value] };
    setSchoolStructure(prev => ({ ...prev, [key]: (prev[key] || []).map(row => row.id === parent.id ? updated : row) }));
    setChildDrafts(prev => ({ ...prev, [parent.id]: '' }));
    try {
      await api.put('/api/school-structure/' + key + '/' + encodeURIComponent(parent.id), updated);
      setNotice('Added ' + value + ' under ' + parent.name + '.');
    } catch {
      setNotice('Added ' + value + ' in this screen, but the server did not confirm saving. Please retry after checking the connection.');
    }
  }

  async function removeStructureChild(parent: StructureRow, key: 'classes' | 'blocks', child: string) {
    const field = key === 'classes' ? 'streams' : 'dormitories';
    const label = key === 'classes' ? 'stream' : 'dormitory';
    if (!window.confirm('Delete ' + label + ' "' + child + '" from ' + parent.name + '? This cannot be undone.')) return;
    const updated: StructureRow = { ...parent, [field]: ((parent[field] || []) as string[]).filter(value => value !== child) };
    try {
      await api.put('/api/school-structure/' + key + '/' + encodeURIComponent(parent.id), updated);
      setSchoolStructure(prev => ({ ...prev, [key]: (prev[key] || []).map(row => row.id === parent.id ? updated : row) }));
      setNotice('Deleted ' + label + ' ' + child + '.');
    } catch {
      setNotice('Could not delete ' + label + '. The server did not confirm the change; please retry.');
    }
  }

  async function removeStructureRow(row: StructureRow, key: 'classes' | 'blocks' | 'suppliers') {
    const label = key === 'classes' ? 'class' : key === 'blocks' ? 'block' : 'supplier';
    const children = key === 'classes'
      ? ' and its streams'
      : key === 'blocks'
        ? ' and its dormitories'
        : '';
    const warning = key === 'classes' || key === 'blocks'
      ? ' Existing learner or inventory records that refer to it will not be deleted automatically and may need reassignment.'
      : '';
    if (!window.confirm('Delete ' + label + ' "' + row.name + '"' + children + '? This cannot be undone.' + warning)) return;
    try {
      await api.delete('/api/school-structure/' + key + '/' + encodeURIComponent(row.id));
      setSchoolStructure(prev => ({
        ...prev,
        [key]: (prev[key] || []).filter(item => item.id !== row.id),
      }));
      setNotice('Deleted ' + label + ' ' + row.name + '.');
    } catch {
      setNotice('Could not delete ' + label + ' ' + row.name + '. The server did not confirm the deletion; please retry.');
    }
  }

  async function saveStructure(e: React.FormEvent) {
    e.preventDefault();
    const key = section === 'Classes & streams' ? 'classes' : section === 'Blocks & dormitories' ? 'blocks' : 'suppliers';
    const row: StructureRow = { id: key.slice(0, -1) + '-' + Date.now(), name: structureForm.name.trim(), detail: structureForm.detail.trim(), parent: structureForm.parent.trim(), streams: structureForm.streams.split(',').map(v => v.trim()).filter(Boolean), dormitories: structureForm.dormitories.split(',').map(v => v.trim()).filter(Boolean), contact: structureForm.contact.trim(), phone: structureForm.phone.trim(), email: structureForm.email.trim(), address: structureForm.address.trim(), supplied: structureForm.supplied.trim() };
    if (!row.name) { setNotice('Enter a name before saving.'); return; }
    const next = { ...schoolStructure, [key]: [row, ...(schoolStructure[key] || [])] };
    setSchoolStructure(next);
    try { await api.post('/api/school-structure', { kind: key, record: row }); setNotice('Saved ' + row.name + '.'); }
    catch { setNotice('Added to this session, but server authentication prevented saving.'); }
    setStructureForm({ name: '', detail: '', parent: '', streams: 'A,B,C,D', dormitories: 'Dormitory 1,Dormitory 2', contact: '', phone: '', email: '', address: '', supplied: '' });
  }
  function parseCsv(text: string) {
    const lines = text.trim().split(/\\r?\\n/).filter(line => line.trim());
    if (lines.length < 2) return [] as Record<string, string>[];
    const split = (line: string) => line.match(/(?:^|,)\\s*(?:"((?:[^"]|"")*)"|([^,]*))/g)?.map(cell => cell.replace(/^,/, '').trim().replace(/^"|"$/g, '').replace(/""/g, '"')) || [];
    const headers = split(lines[0]).map(h => h.toLowerCase().replace(/[ _-]+/g, ''));
    return lines.slice(1).map(line => { const cells = split(line); return Object.fromEntries(headers.map((h, i) => [h, cells[i] || ''])); }).filter(row => Object.values(row).some(Boolean));
  }
  async function handleExcelFile(file?: File) {
    if (!file) return;
    const allowed = /\\.(xlsx|xls|csv)$/i.test(file.name);
    if (!allowed) { setNotice('Choose an Excel workbook (.xlsx or .xls) or a CSV exported from Excel.'); return; }
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });
      const rows = raw.map(row => Object.fromEntries(Object.entries(row).map(([key, value]) => [key.toLowerCase().replace(/[ _-]+/g, ''), String(value ?? '').trim()])));
      if (!rows.length) { setNotice('The first worksheet is empty or has no data rows.'); return; }
      setExcelRows(rows);
      setExcelFileName(file.name);
      setNotice('Loaded ' + rows.length + ' rows from ' + file.name + '. Review the import type and then import.');
    } catch {
      setNotice('Could not read this file. Save it as .xlsx or .csv in Excel and try again.');
    }
  }

  function downloadExcelTemplate() {
    const headers = importType === 'students'
      ? ['fullName','admissionNo','grade','stream','gender','residence','phone','status']
      : ['name','category','quantity','condition','location','cost','barcode','kind'];
    const sample = importType === 'students'
      ? ['Example Learner','HS/2026/001','Grade 7','A','Female','Boarder','0712345678','Active']
      : ['Mathematics Textbook','Textbooks',30,'Good','Main Store',850,'BK-MATH-001','Returnable'];
    const sheet = XLSX.utils.aoa_to_sheet([headers, sample]);
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, importType === 'students' ? 'Students' : 'Inventory');
    XLSX.writeFile(book, 'hill-school-' + importType + '-template.xlsx');
  }

  async function runBulkImport() {
    const rows = excelRows;
    if (!rows.length) { setNotice('Choose an Excel file first, or download the template and fill it in.'); return; }
    setImportBusy(true);
    try {
      const records = rows.map((r, i) => importType === 'students' ? ({ id: 'bulk-student-' + Date.now() + '-' + i, recordType: 'Student', fullName: r.fullname || r.name || '', admissionNo: r.admissionno || r.admissionnumber || '', gender: r.gender || '', grade: r.grade || 'Grade 1', stream: r.stream || 'A', residence: (r.residence || r.boarder || '').toLowerCase().includes('board') ? 'Boarder' : 'Day scholar', phone: r.phone || r.guardianphone || '', status: r.status || 'Active', staffNo: '', staffType: 'Teaching', role: '', biometricRef: '', createdAt: new Date().toISOString() }) : ({ id: 'bulk-item-' + Date.now() + '-' + i, name: r.name || r.itemname || '', category: r.category || 'General', quantity: Number(r.quantity || 0), condition: r.condition || 'Good', location: r.location || 'Main Store', cost: Number(r.cost || r.unitcost || 0), barcode: r.barcode || 'HS-' + (Date.now() + i), kind: (r.kind || r.type || 'Returnable').toLowerCase().includes('consum') ? 'Consumable' : 'Returnable', updatedAt: 'Just now' }));
      const valid = records.filter((r: any) => importType === 'students' ? r.fullName.trim() && r.admissionNo.trim() : r.name.trim());
      if (!valid.length) { setNotice(importType === 'students' ? 'Each student row needs fullName and admissionNo columns.' : 'Each item row needs a name column.'); return; }
      if (importType === 'students') { setPeople(prev => [...valid as Person[], ...prev]); }
      else { setAssets(prev => [...valid as Asset[], ...prev]); }
      const { data } = await api.post('/api/bulk-import', { kind: importType, records: valid });
      setNotice('Imported ' + (data.imported ?? valid.length) + ' ' + importType + '.');
      setExcelRows([]);
      setExcelFileName('');
      setImportText('');
    } catch { setNotice('The server could not confirm saving the imported records. Your file was read, but verify records before importing again to avoid duplicates.'); }
    finally { setImportBusy(false); }
  }
  async function askAI() {
    setAiBusy(true);
    setAiText('');
    try {
      const { data } = await api.post('/api/ai/stock-advice', {
        assets,
        movements,
      });
      setAiText(data.advice || 'No recommendations returned.');
    } catch {
      setAiText(
        'AI advice is temporarily unavailable. Review low-stock items, damaged assets and recent issue/return movements, then try again.'
      );
    } finally {
      setAiBusy(false);
    }
  }
  async function generateFurnitureCodes(e: React.FormEvent) {
    e.preventDefault();
    const count = Math.floor(Number(furnitureCount));
    if (!Number.isFinite(count) || count < 1 || count > 200) { setNotice('Choose between 1 and 200 seats per generation batch.'); return; }
    const existingCodes = new Set(assets.map(asset => asset.barcode));
    const records: Asset[] = [];
    for (let seat = 1; seat <= count; seat += 1) {
      const code = `HPS-${furnitureGrade}-${furnitureStream}-${furnitureItem}-${String(seat).padStart(2, '0')}`;
      if (existingCodes.has(code)) continue;
      records.push({ id: `furniture-${Date.now()}-${seat}`, name: `${furnitureItem === 'CH' ? 'Chair' : furnitureItem === 'DK' ? 'Desk' : 'Locker'} seat code`, category: 'Furniture', quantity: 1, condition: furnitureCondition, location: `${furnitureGrade} / Stream ${furnitureStream}`, cost: 0, barcode: code, kind: 'Returnable', updatedAt: new Date().toISOString() });
    }
    if (!records.length) { setNotice('These seat codes already exist. Choose a different class, stream or item type, or remove the existing codes first.'); return; }
    setFurnitureBusy(true);
    try {
      const { data } = await api.post('/api/bulk-import', { kind: 'items', records });
      const imported = Number(data.imported ?? 0);
      if (imported !== records.length) throw new Error('The server did not confirm every seat code.');
      setAssets(prev => [...records, ...prev]);
      setNotice(`Generated and saved ${records.length} furniture seat codes in the existing inventory register. No learner details are stored in these records.`);
    } catch (err) {
      setNotice(err instanceof Error ? `Furniture codes were not confirmed by the server: ${err.message}` : 'Furniture codes were not confirmed by the server. Please retry.');
    } finally { setFurnitureBusy(false); }
  }

  function exportFurnitureCodes() {
    const rows = furnitureAssets.map(asset => {
      const [, grade = '', stream = '', item = '', seat = ''] = asset.barcode.split('-');
      return { Code: asset.barcode, Grade: grade, Stream: stream, Item: item, 'Seat No': seat, Condition: asset.condition, Location: asset.location };
    });
    const sheet = XLSX.utils.json_to_sheet(rows);
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, 'Furniture codes');
    XLSX.writeFile(book, 'hill-school-furniture-codes.xlsx');
  }

  function printFurnitureLabels() {
    window.print();
  }

  const furnitureAssets = assets.filter(asset => asset.category === 'Furniture' && /^HPS-(PP1|PP2|G[1-9])-(A|B|C|E|M|N|S|W)-(CH|DK|LK)-\\d+$/i.test(asset.barcode))
    .filter(asset => furnitureFilterGrade === 'All grades' || asset.barcode.split('-')[1] === furnitureFilterGrade)
    .filter(asset => furnitureFilterStream === 'All streams' || asset.barcode.split('-')[2] === furnitureFilterStream)
    .filter(asset => furnitureFilterItem === 'All items' || asset.barcode.split('-')[3] === furnitureFilterItem);

  function exportCsv() {
    const rows = [
      [
        'Name',
        'Category',
        'Quantity',
        'Condition',
        'Location',
        'Unit cost (KES)',
        'Barcode',
        'Type',
      ],
      ...filtered.map(a => [
        a.name,
        a.category,
        String(a.quantity),
        a.condition,
        a.location,
        String(a.cost),
        a.barcode,
        a.kind,
      ]),
    ];
    const csv = rows
      .map(row => row.map(v => '"' + v.replaceAll('"', '""') + '"').join(','))
      .join('\n');
    const url = URL.createObjectURL(
      new Blob([csv], { type: 'text/csv;charset=utf-8;' })
    );
    const link = document.createElement('a');
    link.href = url;
    link.download = 'hill-school-inventory.csv';
    link.click();
    URL.revokeObjectURL(url);
  }

  const backupKinds = ['assets', 'moves', 'asset_updates', 'people', 'boarding_attendance', 'classes', 'blocks', 'suppliers'];

  async function downloadFullBackup() {
    setBackupBusy(true);
    setBackupStatus('Collecting saved records from the database…');
    try {
      const records: Record<string, unknown[]> = {};
      let total = 0;
      for (const kind of backupKinds) {
        const rows: unknown[] = [];
        let nextToken = '';
        let pageCount = 0;
        do {
          const query = nextToken ? '?nextToken=' + encodeURIComponent(nextToken) : '';
          const { data } = await api.get('/api/backup/' + kind + query);
          if (!Array.isArray(data.records)) throw new Error('The server returned an invalid backup page for ' + kind + '.');
          rows.push(...data.records);
          nextToken = typeof data.nextToken === 'string' ? data.nextToken : '';
          pageCount += 1;
          if (pageCount > 1000) throw new Error('Backup stopped because the page safety limit was reached. Contact support before relying on this file.');
          setBackupStatus('Collected ' + rows.length + ' ' + kind.replaceAll('_', ' ') + ' records…');
        } while (nextToken);
        records[kind] = rows;
        total += rows.length;
      }
      const backup = { format: 'hill-school-backup-v1', exportedAt: new Date().toISOString(), app: 'Hill School Inventory', records };
      const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'hill-school-backup-' + new Date().toISOString().slice(0, 10) + '.json';
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
      setBackupStatus('Backup downloaded successfully. ' + total + ' records were included. Store the file somewhere secure.');
      setNotice('Database backup downloaded. Keep the file in a secure location.');
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Backup failed. Please try again.';
      setBackupStatus('Backup incomplete: ' + message);
      setNotice('Backup could not be confirmed complete. Do not rely on a partial file.');
    } finally {
      setBackupBusy(false);
    }
  }

  async function restoreBackupFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setRestoreBusy(true);
    setBackupStatus('Checking backup file…');
    try {
      if (file.size > 25 * 1024 * 1024) throw new Error('The backup file is larger than 25 MB. Contact support for a large-file restore.');
      const parsed = JSON.parse(await file.text()) as { format?: string; records?: Record<string, unknown> };
      if (parsed.format !== 'hill-school-backup-v1' || !parsed.records || typeof parsed.records !== 'object') throw new Error('This is not a supported Hill School backup file.');
      const kinds = backupKinds.filter(kind => Array.isArray(parsed.records?.[kind]));
      if (!kinds.length) throw new Error('No supported data tables were found in this backup.');
      const total = kinds.reduce((sum, kind) => sum + (parsed.records?.[kind] as unknown[]).length, 0);
      if (!window.confirm('Restore ' + total + ' records from this backup? Existing records with matching IDs will be updated; missing records will be added. This merge does not delete current records. Make a fresh backup first.')) return;
      let restored = 0;
      for (const kind of kinds) {
        const rows = parsed.records[kind] as unknown[];
        for (let start = 0; start < rows.length; start += 3) {
          const batch = rows.slice(start, start + 3);
          const { data } = await api.post('/api/backup/restore', { kind, records: batch });
          restored += Number(data.restored || 0);
          setBackupStatus('Restored ' + restored + ' of ' + total + ' records…');
        }
      }
      setBackupStatus('Restore completed. ' + restored + ' records were processed. Refresh the page to reload all restored records.');
      setNotice('Backup restore finished. Refresh the app to reload restored data.');
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Restore failed. Some earlier batches may already have been restored; you can safely retry the same file.';
      setBackupStatus('Restore incomplete: ' + message);
      setNotice('Restore did not finish completely. Review the backup status before retrying.');
    } finally {
      setRestoreBusy(false);
      e.target.value = '';
    }
  }

  if (authLoading) {
    return <main className="auth-screen"><section className="auth-card"><div className="auth-logo"><GraduationCap size={28} /></div><p className="auth-eyebrow">HILL SCHOOL</p><h1>Preparing your workspace</h1><p className="auth-description">Checking your secure school account…</p><div className="auth-loader" /></section></main>;
  }
  if (!signedInUser) {
    return <main className="auth-screen"><section className="auth-card"><div className="auth-logo"><GraduationCap size={30} /></div><p className="auth-eyebrow">HILL SCHOOL · INVENTORY INTELLIGENCE</p><h1>Welcome back</h1><p className="auth-description">Sign in securely to access school inventory, learner and staff records, textbook circulation, and boarding attendance.</p><button className="auth-signin" onClick={handleSignIn} disabled={signingIn}>{signingIn ? 'Connecting securely…' : 'Sign in to Hill School'}</button>{authError && <p className="auth-error" role="alert">{authError}</p>}<p className="auth-footnote"><ShieldCheck size={15} /> School records are available only after sign-in.</p></section></main>;
  }

  return (
    <div className="shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark">
            <GraduationCap size={23} />
          </div>
          <div>
            <strong>HILL SCHOOL</strong>
            <span>Inventory intelligence</span>
          </div>
        </div>
        <div className="school-switch">
          <div className="school-avatar">HS</div>
          <div>
            <b>Hill School</b>
            <small>Primary & Junior School</small>
          </div>
          <ChevronDown size={15} />
        </div>
        <div className="nav-caption">SCHOOL WORKSPACE</div>
        <nav>
          {nav.map(item => (
            <button
              key={item.label}
              className={'nav-item ' + (section === item.label ? 'active' : '')}
              onClick={() => setSection(item.label)}
            >
              <item.icon size={18} />
              <span>{item.label}</span>
              {item.label === 'Textbooks' && (
                <span className="nav-count">2</span>
              )}
            </button>
          ))}
        </nav>
        <div className="nav-caption management-caption">MANAGEMENT</div>
        <button className="nav-item" onClick={() => setSection('Maintenance')}>
          <Wrench size={18} />
          <span>Maintenance</span>
        </button>
        <button className="nav-item" onClick={() => setSection('Settings')}>
          <Settings2 size={18} />
          <span>Settings</span>
        </button>
        <button className="nav-item auth-signout" onClick={async () => { await auth.signOut(); setSignedInUser(null); setAuthError(''); }}>
          <ShieldCheck size={18} />
          <span>Sign out</span>
        </button>
        <div className="sidebar-bottom">
          <div className="help-card">
            <div className="help-icon">
              <CircleHelp size={18} />
            </div>
            <b>Need a hand?</b>
            <p>Learn how to manage school assets efficiently.</p>
            <button
              onClick={() =>
                setNotice(
                  'Help centre: use Inventory to add assets and Issue & return to track movement by barcode.'
                )
              }
            >
              Open help centre <span>↗</span>
            </button>
          </div>
          <div className="profile">
            <div className="profile-avatar">AD</div>
            <div>
              <b>School Administrator</b>
              <small>Inventory manager</small>
            </div>
            <button
              aria-label="Profile menu"
              onClick={() =>
                setNotice(
                  'Role-based access can be configured as the next security setup step.'
                )
              }
            >
              <ChevronDown size={16} />
            </button>
          </div>
        </div>
      </aside>
      <main className="main">
        <header className="topbar">
          <div className="breadcrumb">
            <span>Workspace</span>
            <span className="crumb-slash">/</span>
            <b>{section}</b>
          </div>
          <div className="top-actions">
            <div className="status">
              <span className="status-dot"></span> School year 2026
            </div>
            <button
              className="icon-button"
              aria-label="Activity"
              onClick={() => setSection('Issue & return')}
            >
              <Activity size={18} />
            </button>
            <button
              className="avatar-button"
              onClick={() =>
                setNotice(
                  'Signed in as demo administrator. Configure authentication before production use.'
                )
              }
            >
              AD
            </button>
          </div>
        </header>
        <div className="content">
          {notice && (
            <div className="notice">
              <CheckCircle2 size={17} />
              <span>{notice}</span>
              <button onClick={() => setNotice('')} aria-label="Dismiss">
                <X size={16} />
              </button>
            </div>
          )}
          <div className="page-heading">
            <div>
              <div className="eyebrow">
                SCHOOL OPERATIONS <span>•</span> 2026
              </div>
              <h1>{section === 'Overview' ? 'Inventory overview' : section}</h1>
              <p>
                Know what your school owns, where it is, and what needs
                attention.
              </p>
            </div>
            <div className="heading-actions">
              <button className="btn btn-secondary" onClick={exportCsv}>
                <Download size={16} /> Export report
              </button>
              <button
                className="btn btn-primary"
                onClick={() =>
                  section === 'Issue & return'
                    ? setShowMove(true)
                    : setShowForm(true)
                }
              >
                <Plus size={17} />
                {section === 'Issue & return'
                  ? 'New transaction'
                  : 'Add inventory'}
              </button>
            </div>
          </div>
          {section === 'Overview' && (
            <>
              <div className="welcome-strip">
                <div className="welcome-icon">
                  <Sparkles size={20} />
                </div>
                <div>
                  <b>Your school at a glance</b>
                  <p>
                    Track assets, textbooks, allocations and stock movements
                    from one workspace.
                  </p>
                </div>
                <span className="live-pill">
                  <span /> LIVE WORKSPACE
                </span>
              </div>
              <div className="metrics">
                <Metric
                  icon={<Boxes size={19} />}
                  label="Inventory items"
                  value={assets.length.toLocaleString()}
                  note="Tracked item types"
                  tone="blue"
                />
                <Metric
                  icon={<Warehouse size={19} />}
                  label="Total units"
                  value={assets
                    .reduce((s, a) => s + a.quantity, 0)
                    .toLocaleString()}
                  note="Across all locations"
                  tone="violet"
                />
                <Metric
                  icon={<ShieldCheck size={19} />}
                  label="Estimated value"
                  value={money(value)}
                  note="Based on recorded unit costs"
                  tone="green"
                />
                <Metric
                  icon={<AlertTriangle size={19} />}
                  label="Needs attention"
                  value={String(
                    lowStock.length +
                      assets.filter(a => a.condition !== 'Good').length
                  )}
                  note="Low stock or condition issues"
                  tone="orange"
                />
              </div>
              <div className="dashboard-grid">
                <section className="panel inventory-panel">
                  <div className="panel-head">
                    <div>
                      <h2>Inventory snapshot</h2>
                      <p>Latest items tracked by your school</p>
                    </div>
                    <button
                      className="text-button"
                      onClick={() => setSection('Inventory')}
                    >
                      View inventory <span>→</span>
                    </button>
                  </div>
                  <InventoryTable
                    assets={assets.slice(0, 5)}
                    onSelect={a => {
                      setSearch(a.name);
                      setSection('Inventory');
                    }}
                  />
                </section>
                <section className="panel attention-panel">
                  <div className="panel-head">
                    <div>
                      <h2>Attention needed</h2>
                      <p>Items that may need action</p>
                    </div>
                    <span className="alert-count">{lowStock.length}</span>
                  </div>
                  {lowStock.length ? (
                    lowStock.slice(0, 4).map(a => (
                      <div className="attention-row" key={a.id}>
                        <div className="attention-symbol">
                          <AlertTriangle size={17} />
                        </div>
                        <div className="attention-copy">
                          <b>{a.name}</b>
                          <small>{a.location}</small>
                        </div>
                        <span className="stock-badge">{a.quantity} left</span>
                      </div>
                    ))
                  ) : (
                    <div className="empty-inline">
                      No low-stock items. Looking good!
                    </div>
                  )}
                  <button
                    className="wide-outline"
                    onClick={askAI}
                    disabled={aiBusy}
                  >
                    <Sparkles size={16} />
                    {aiBusy ? 'Analysing stock…' : 'Ask AI for stock advice'}
                  </button>
                  {aiText && (
                    <div className="ai-advice">
                      <b>
                        <Sparkles size={14} /> AI stock advice
                      </b>
                      <p>{aiText}</p>
                    </div>
                  )}
                </section>
              </div>
              <div className="bottom-grid">
                <section className="panel">
                  <div className="panel-head">
                    <div>
                      <h2>Recent activity</h2>
                      <p>Latest inventory movements</p>
                    </div>
                    <button
                      className="text-button"
                      onClick={() => setSection('Issue & return')}
                    >
                      All activity →
                    </button>
                  </div>
                  <div className="activity-list">
                    {movements.slice(0, 4).map(m => (
                      <div className="activity-row" key={m.id}>
                        <div
                          className={
                            'activity-icon ' +
                            m.action.toLowerCase().replaceAll(' ', '-')
                          }
                        >
                          <ArrowLeftRight size={16} />
                        </div>
                        <div className="activity-desc">
                          <b>{m.assetName}</b>
                          <small>
                            {m.action} · {m.person}
                          </small>
                        </div>
                        <div className="activity-meta">
                          <b>
                            {m.quantity} unit{m.quantity === 1 ? '' : 's'}
                          </b>
                          <small>{m.date}</small>
                        </div>
                      </div>
                    ))}
                  </div>
                </section>
                <section className="panel quick-panel">
                  <div className="panel-head">
                    <div>
                      <h2>Quick actions</h2>
                      <p>Common tasks, one click away</p>
                    </div>
                  </div>
                  <div className="quick-actions">
                    <button onClick={() => setShowForm(true)}>
                      <span className="quick-icon blue">
                        <PackagePlus size={18} />
                      </span>
                      <span>
                        <b>Add new stock</b>
                        <small>Register a new asset or supply</small>
                      </span>
                      <span className="quick-arrow">↗</span>
                    </button>
                    <button
                      onClick={() => {
                        setSection('Textbooks');
                        setShowMove(true);
                      }}
                    >
                      <span className="quick-icon purple">
                        <QrCode size={18} />
                      </span>
                      <span>
                        <b>Issue a textbook</b>
                        <small>Scan or enter a book barcode</small>
                      </span>
                      <span className="quick-arrow">↗</span>
                    </button>
                    <button onClick={() => setSection('Reports')}>
                      <span className="quick-icon green">
                        <FileText size={18} />
                      </span>
                      <span>
                        <b>Generate a report</b>
                        <small>Export stock and valuation data</small>
                      </span>
                      <span className="quick-arrow">↗</span>
                    </button>
                  </div>
                </section>
              </div>
            </>
          )}
          {(section === 'Inventory' || section === 'Textbooks') && (
            <section className="panel full-panel">
              <div className="table-toolbar">
                <div className="search-box">
                  <Search size={17} />
                  <input
                    value={search}
                    onChange={e => setSearch(e.target.value)}
                    placeholder="Search by item, category, location or barcode…"
                  />
                </div>
                <select
                  value={category}
                  onChange={e => setCategory(e.target.value)}
                >
                  <option>All categories</option>
                  {[
                    'Textbooks',
                    'Furniture',
                    'ICT',
                    'Boarding',
                    'Stationery',
                    'Sports',
                    'Laboratory',
                    'Cleaning',
                  ].map(c => (
                    <option key={c}>{c}</option>
                  ))}
                </select>
                <button className="btn btn-secondary" onClick={exportCsv}>
                  <ArrowDownToLine size={16} /> Export CSV
                </button>
              </div>
              <InventoryTable
                assets={
                  section === 'Textbooks'
                    ? filtered.filter(a => a.category === 'Textbooks')
                    : filtered
                }
                onSelect={a =>
                  setNotice(
                    a.name + ' · Barcode ' + a.barcode + ' · ' + a.condition
                  )
                }
              />
              {filtered.length === 0 && (
                <div className="empty-state">
                  No items match your search. Try another term or add an
                  inventory item.
                </div>
              )}
            </section>
          )}
          {section === 'Furniture coding' && (
            <div className="furniture-workspace">
              <section className="panel furniture-builder">
                <div className="panel-head"><div><h2>Furniture seat coding</h2><p>Create unique school-owned codes without learner names or admission numbers.</p></div><span className="category-chip"><QrCode size={14}/> HPS label standard</span></div>
                <form className="furniture-form" onSubmit={generateFurnitureCodes}>
                  <label>Grade<select value={furnitureGrade} onChange={e => setFurnitureGrade(e.target.value)}>{['PP1','PP2',...Array.from({length:9},(_,i)=>`G${i+1}`)].map(grade => <option key={grade} value={grade}>{grade}</option>)}</select></label>
                  <label>Stream<select value={furnitureStream} onChange={e => setFurnitureStream(e.target.value)}>{[['A','Alpha'],['B','Beta'],['C','Central'],['E','East'],['M','Mid'],['N','North'],['S','South'],['W','West']].map(([code,name]) => <option key={code} value={code}>{code} · {name}</option>)}</select></label>
                  <label>Item type<select value={furnitureItem} onChange={e => setFurnitureItem(e.target.value)}><option value="CH">CH · Chair</option><option value="DK">DK · Desk</option><option value="LK">LK · Locker</option></select></label>
                  <label>Number of seats<input type="number" min="1" max="200" value={furnitureCount} onChange={e => setFurnitureCount(e.target.value)} required/><small>Generate 1–200 per batch</small></label>
                  <label>Condition<select value={furnitureCondition} onChange={e => setFurnitureCondition(e.target.value)}><option>Good</option><option>Fair</option><option>Damaged</option></select></label>
                  <div className="furniture-code-preview"><span>Example code</span><strong>{`HPS-${furnitureGrade}-${furnitureStream}-${furnitureItem}-${String(Math.max(1, Number(furnitureCount) || 1)).padStart(2,'0')}`}</strong><small>Each code identifies one seat only.</small></div>
                  <button className="btn btn-primary" type="submit" disabled={furnitureBusy}><Plus size={16}/>{furnitureBusy ? 'Generating codes…' : 'Generate and save codes'}</button>
                </form>
              </section>
              <section className="panel furniture-register">
                <div className="panel-head"><div><h2>Furniture code register</h2><p>{furnitureAssets.length} coded furniture seats · saved to the existing inventory register</p></div><div className="furniture-tools"><button className="btn btn-secondary" type="button" onClick={exportFurnitureCodes} disabled={!furnitureAssets.length}><Download size={15}/> Excel</button><button className="btn btn-secondary" type="button" onClick={printFurnitureLabels} disabled={!furnitureAssets.length}><FileText size={15}/> Print labels / PDF</button></div></div>
                <div className="furniture-filters"><label>Grade<select value={furnitureFilterGrade} onChange={e => setFurnitureFilterGrade(e.target.value)}><option>All grades</option>{['PP1','PP2',...Array.from({length:9},(_,i)=>`G${i+1}`)].map(grade => <option key={grade}>{grade}</option>)}</select></label><label>Stream<select value={furnitureFilterStream} onChange={e => setFurnitureFilterStream(e.target.value)}><option>All streams</option>{['A','B','C','E','M','N','S','W'].map(stream => <option key={stream}>{stream}</option>)}</select></label><label>Item type<select value={furnitureFilterItem} onChange={e => setFurnitureFilterItem(e.target.value)}><option>All items</option><option value="CH">Chair</option><option value="DK">Desk</option><option value="LK">Locker</option></select></label></div>
                <div className="table-wrap furniture-table-wrap"><table><thead><tr><th>CODE / SEAT</th><th>GRADE</th><th>STREAM</th><th>ITEM</th><th>CONDITION</th><th>LOCATION</th></tr></thead><tbody>{furnitureAssets.map(asset => { const parts = asset.barcode.split('-'); return <tr key={asset.id}><td><strong>{asset.barcode}</strong><small>Seat identifier only</small></td><td>{parts[1]}</td><td>{parts[2]}</td><td>{parts[3] === 'CH' ? 'Chair' : parts[3] === 'DK' ? 'Desk' : 'Locker'}</td><td><span className={'condition ' + (asset.condition === 'Good' ? 'good' : asset.condition === 'Damaged' ? 'damaged' : 'repair')}>{asset.condition}</span></td><td>{asset.location}</td></tr>; })}</tbody></table>{!furnitureAssets.length && <div className="furniture-empty"><QrCode size={28}/><strong>No furniture codes yet</strong><p>Choose a grade, stream and item type above, then generate a seat-code batch.</p></div>}</div>
                <div className="furniture-label-grid">{furnitureAssets.map(asset => <div className="furniture-label" key={'label-'+asset.id}><div className="furniture-label-school">HILL SCHOOL · FURNITURE</div><svg data-code128={asset.barcode} aria-label={'Code 128 barcode for '+asset.barcode}/><strong>{asset.barcode}</strong><small>{asset.condition} · Seat identifier</small></div>)}</div>
              </section>
            </div>
          )}
          {section === 'Issue & return' && (
            <section className="panel full-panel">
              <div className="panel-head">
                <div>
                  <h2>Issue, return & allocation log</h2>
                  <p>
                    Every transaction is linked to an item barcode for
                    traceability.
                  </p>
                </div>
                <button
                  className="btn btn-primary"
                  onClick={() => setShowMove(true)}
                >
                  <Plus size={16} /> New transaction
                </button>
              </div>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>Item</th>
                      <th>Transaction</th>
                      <th>Recipient / location</th>
                      <th>Quantity</th>
                      <th>Barcode</th>
                      <th>Date</th>
                    </tr>
                  </thead>
                  <tbody>
                    {movements.map(m => (
                      <tr key={m.id}>
                        <td>
                          <b>{m.assetName}</b>
                        </td>
                        <td>
                          <span className="condition good">{m.action}</span>
                        </td>
                        <td>{m.person}</td>
                        <td>{m.quantity}</td>
                        <td>
                          <code>{m.barcode}</code>
                        </td>
                        <td>{m.date}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          )}
          {(section === 'Classes & streams' || section === 'Blocks & dormitories' || section === 'Suppliers') && (() => {
            const key = section === 'Classes & streams' ? 'classes' : section === 'Blocks & dormitories' ? 'blocks' : 'suppliers';
            const rows = schoolStructure[key] || [];
            return <section className="panel full-panel structure-panel"><div className="panel-head"><div><h2>{section}</h2><p>{key === 'classes' ? 'Create each class first, then add its streams directly under that class.' : key === 'blocks' ? 'Create each boarding block first, then add its dormitories directly under that block.' : 'Keep supplier contacts and purchasing details in one register.'}</p></div><span className="category-chip">{rows.length} records</span></div>
              <form className="structure-form" onSubmit={saveStructure}><div className="form-grid">
                <label>{key === 'classes' ? 'Class / grade name' : key === 'blocks' ? 'Block name' : 'Supplier / company name'}<input required value={structureForm.name} onChange={e => setStructureForm({...structureForm, name:e.target.value})} placeholder={key === 'classes' ? 'e.g. Grade 7' : key === 'blocks' ? 'e.g. Block A' : 'Supplier registered name'}/></label>
                {key === 'blocks' && <label>Capacity / notes<input value={structureForm.detail} onChange={e => setStructureForm({...structureForm, detail:e.target.value})} placeholder="e.g. 80 beds"/></label>}
                {key === 'suppliers' && <><label>Contact person<input value={structureForm.contact} onChange={e => setStructureForm({...structureForm, contact:e.target.value})} placeholder="Contact name"/></label><label>Phone number<input value={structureForm.phone} onChange={e => setStructureForm({...structureForm, phone:e.target.value})} placeholder="Supplier phone"/></label><label>Email<input type="email" value={structureForm.email} onChange={e => setStructureForm({...structureForm, email:e.target.value})} placeholder="accounts@supplier.co.ke"/></label><label>Address / location<input value={structureForm.address} onChange={e => setStructureForm({...structureForm, address:e.target.value})} placeholder="Town / address"/></label><label>Items supplied<input value={structureForm.supplied} onChange={e => setStructureForm({...structureForm, supplied:e.target.value})} placeholder="Furniture, stationery, textbooks…"/></label></>}
              </div><button className="btn btn-primary" type="submit"><Plus size={16}/> Add {key === 'classes' ? 'class' : key === 'blocks' ? 'block' : 'supplier'}</button></form>
              <div className="structure-list">{rows.map(row => <article className="structure-row" key={row.id}><div className="structure-icon">{key === 'classes' ? <School size={19}/> : key === 'blocks' ? <Building2 size={19}/> : <Truck size={19}/>}</div><div className="structure-copy"><div className="structure-title-row"><b>{row.name}</b>{key !== 'suppliers' && <button type="button" className="structure-delete" aria-label={'Delete ' + (key === 'classes' ? 'class ' : 'block ') + row.name} onClick={() => void removeStructureRow(row, key as 'classes' | 'blocks')}>Delete {key === 'classes' ? 'class' : 'block'}</button>}</div>{key === 'classes' ? <><p>Streams under this class:</p><div className="child-items">{(row.streams || []).length ? (row.streams || []).map(stream => <span className="child-item" key={stream}>{stream}<button type="button" className="child-delete" aria-label={'Delete stream ' + stream} onClick={() => void removeStructureChild(row, 'classes', stream)}>Delete</button></span>) : <span className="child-empty">No streams added yet</span>}</div><div className="child-add"><input aria-label={'New stream under ' + row.name} value={childDrafts[row.id] || ''} onChange={e => setChildDrafts(prev => ({...prev, [row.id]:e.target.value}))} placeholder="Enter stream name, e.g. East"/><button type="button" className="btn btn-secondary" onClick={() => void addStructureChild(row, 'classes', childDrafts[row.id] || '')}><Plus size={14}/> Add stream</button></div></> : key === 'blocks' ? <><p>Dormitories under this block:{row.detail ? ' ' + row.detail : ''}</p><div className="child-items">{(row.dormitories || []).length ? (row.dormitories || []).map(dormitory => <span className="child-item" key={dormitory}>{dormitory}<button type="button" className="child-delete" aria-label={'Delete dormitory ' + dormitory} onClick={() => void removeStructureChild(row, 'blocks', dormitory)}>Delete</button></span>) : <span className="child-empty">No dormitories added yet</span>}</div><div className="child-add"><input aria-label={'New dormitory under ' + row.name} value={childDrafts[row.id] || ''} onChange={e => setChildDrafts(prev => ({...prev, [row.id]:e.target.value}))} placeholder="Enter dormitory name, e.g. Dorm 1"/><button type="button" className="btn btn-secondary" onClick={() => void addStructureChild(row, 'blocks', childDrafts[row.id] || '')}><Plus size={14}/> Add dormitory</button></div></> : <p>{[row.contact, row.phone, row.email, row.address, row.supplied].filter(Boolean).join(' · ') || 'No contact details added'}</p>}</div></article>)}{rows.length === 0 && <div className="empty-state">No {key} records yet. Add the first one using the form above.</div>}</div>
            </section>;
          })()}
          {section === 'Bulk import' && <section className="panel full-panel bulk-panel"><div className="panel-head"><div><h2>Excel bulk import</h2><p>Upload a completed Excel workbook (.xlsx/.xls) to import student or inventory records in batches.</p></div><Upload size={22}/></div><div className="bulk-type"><button type="button" className={'btn ' + (importType === 'students' ? 'btn-primary' : 'btn-secondary')} onClick={() => {setImportType('students');setExcelRows([]);setExcelFileName('');}}>Import students</button><button type="button" className={'btn ' + (importType === 'items' ? 'btn-primary' : 'btn-secondary')} onClick={() => {setImportType('items');setExcelRows([]);setExcelFileName('');}}>Import items</button></div><div className="import-help"><b>Required Excel column headings for {importType}:</b><p>{importType === 'students' ? 'fullName, admissionNo, grade, stream, gender, residence, phone, status' : 'name, category, quantity, condition, location, cost, barcode, kind'}</p><p>Use the first worksheet. Keep the heading row at the top. Student rows require fullName and admissionNo; item rows require name and quantity.</p><button type="button" className="btn btn-secondary" onClick={downloadExcelTemplate}><Download size={15}/> Download Excel template (.xlsx)</button></div><label className="csv-label">Choose Excel file<input type="file" accept=".xlsx,.xls,.csv" onChange={e => void handleExcelFile(e.target.files?.[0])}/></label>{excelFileName && <div className="excel-file"><FileText size={18}/><div><b>{excelFileName}</b><span>{excelRows.length} data rows loaded</span></div><button type="button" className="text-button" onClick={() => {setExcelRows([]);setExcelFileName('');}}>Remove</button></div>}<div className="modal-actions"><button className="btn btn-secondary" type="button" onClick={() => {setExcelRows([]);setExcelFileName('');}}>Clear file</button><button className="btn btn-primary" type="button" disabled={importBusy || !excelRows.length} onClick={runBulkImport}><Upload size={16}/>{importBusy ? 'Importing…' : 'Validate and import ' + excelRows.length + ' rows'}</button></div><p className="import-warning">Review the selected import type before importing. Check duplicate admission numbers and barcodes first. Server confirmation is required for durable shared records. Import up to 200 records per batch.</p></section>}
          {section === 'Locations' && (
            <div className="location-grid">
              {[
                'Main Store',
                'Library',
                'Computer Lab',
                'Admin Store',
                'Grade 6 East',
                'Block B · Dorm 2',
              ].map((loc, i) => (
                <div className="panel location-card" key={loc}>
                  <div className="location-symbol">
                    <Warehouse size={20} />
                  </div>
                  <span className="location-label">SCHOOL LOCATION</span>
                  <h3>{loc}</h3>
                  <p>
                    {assets
                      .filter(a => a.location === loc)
                      .reduce((s, a) => s + a.quantity, 0)}{' '}
                    tracked units
                  </p>
                  <button
                    className="text-button"
                    onClick={() => {
                      setSearch(loc);
                      setSection('Inventory');
                    }}
                  >
                    View items →
                  </button>
                </div>
              ))}
            </div>
          )}
          {section === 'People' && (
            <section className="panel full-panel people-panel">
              <div className="panel-head people-heading"><div><h2>{peopleType === 'Student' ? 'Learners register' : peopleType === 'Staff' ? 'Staff register' : 'Learners & staff register'}</h2><p>Manage learner admission details, class and stream placement, boarding assignments, parent contacts and status.</p></div><button className="btn btn-primary" onClick={() => { resetPersonForm(); setPersonForm(prev => ({ ...prev, recordType: peopleType === 'Staff' ? 'Staff' : 'Student', residence: peopleType === 'Staff' ? 'Staff' : 'Day scholar' })); setShowPersonForm(true); }}><UserPlus size={16}/> {peopleType === 'Staff' ? 'Register staff' : 'Add learner'}</button></div>
              <div className="people-stats"><div><span>Students</span><b>{people.filter(p => p.recordType === 'Student').length}</b></div><div><span>Staff</span><b>{people.filter(p => p.recordType === 'Staff').length}</b></div><div><span>Boarders</span><b>{people.filter(p => p.recordType === 'Student' && p.residence === 'Boarder').length}</b></div><div><span>Day scholars</span><b>{people.filter(p => p.recordType === 'Student' && p.residence === 'Day scholar').length}</b></div></div>
              <div className="table-toolbar"><div className="search-box"><Search size={17}/><input value={personSearch} onChange={e => setPersonSearch(e.target.value)} placeholder="Search name, admission number, TSC or ID number…"/></div><select value={peopleType} onChange={e => setPeopleType(e.target.value as 'All' | 'Student' | 'Staff')}><option value="All">All people</option><option value="Student">Students</option><option value="Staff">Staff</option></select></div>
              <div className="table-wrap"><table><thead>{peopleType === 'Student' ? <tr><th>ADMISSION NO</th><th>NAME</th><th>TYPE</th><th>CLASS</th><th>STREAM</th><th>DORMITORY</th><th>PARENT / GUARDIAN</th><th>STATUS</th><th>ACTIONS</th></tr> : <tr><th>NAME / IDENTIFIER</th><th>TYPE</th><th>CLASS / STAFF CATEGORY</th><th>GENDER</th><th>RESIDENCE</th><th>PHONE</th><th>STATUS</th><th>ACTIONS</th></tr>}</thead><tbody>{people.filter(p => (peopleType === 'All' || p.recordType === peopleType) && [p.fullName,p.admissionNo,p.staffNo,p.grade,p.stream,p.dormitory || '',p.parentName || '',p.phone,p.role].some(v => v.toLowerCase().includes(personSearch.toLowerCase()))).map(p => peopleType === 'Student' ? <tr key={p.id}><td><b>{p.admissionNo}</b></td><td><div className="item-cell"><div className="item-symbol book-symbol"><GraduationCap size={17}/></div><div><b>{p.fullName}</b><small>{p.gender || 'Gender not set'}</small></div></div></td><td><span className="category-chip">{p.residence === 'Boarder' ? 'Boarder' : 'Day Scholar'}</span></td><td>{p.grade || '—'}</td><td>{p.stream || '—'}</td><td>{p.residence === 'Boarder' ? (p.dormitory || 'Not assigned') : '—'}</td><td><div><b>{p.parentName || '—'}</b><small>{p.phone || 'No contact'}</small></div></td><td><span className={'condition ' + (p.status === 'Active' ? 'good' : 'repair')}>{p.status}</span></td><td><button className="btn btn-secondary learner-edit" onClick={() => openPersonEditor(p)}><Settings2 size={14}/> Edit</button></td></tr> : <tr key={p.id}><td><div className="item-cell"><div className="item-symbol book-symbol"><UserRound size={17}/></div><div><b>{p.fullName}</b><small>{p.recordType === 'Student' ? 'Admission: ' + p.admissionNo : (p.staffType === 'Teaching' ? 'TSC: ' : 'ID: ') + p.staffNo}</small></div></div></td><td><span className="category-chip">{p.recordType}</span></td><td>{p.recordType === 'Student' ? p.grade + ' · Stream ' + p.stream : p.staffType + (p.role ? ' · ' + p.role : '')}</td><td>{p.gender || '—'}</td><td>{p.residence}</td><td>{p.phone || '—'}</td><td><span className={'condition ' + (p.status === 'Active' ? 'good' : 'repair')}>{p.status}</span></td><td><button className="btn btn-secondary learner-edit" onClick={() => openPersonEditor(p)}><Settings2 size={14}/> Edit</button></td></tr>)}</tbody></table></div>
              {people.length === 0 && <div className="empty-state">No student or staff records registered yet. Select Register person to create the first record.</div>}
            </section>
          )}
          {section === 'Boarding attendance' && (
            <section className="panel full-panel attendance-panel">
              <div className="panel-head"><div><h2>Boarding biometric attendance</h2><p>Attendance register for all active boarders. Store only device-assigned references, never raw fingerprints or face templates.</p></div><span className="category-chip"><Fingerprint size={14}/> Device workflow</span></div>
              <div className="attendance-stats"><div><span>Registered boarders</span><b>{people.filter(p => p.recordType === 'Student' && p.residence === 'Boarder' && p.status === 'Active').length}</b></div><div><span>Attendance events</span><b>{attendance.length}</b></div><div><span>Check-ins today</span><b>{attendance.filter(a => a.event === 'Check-in' && new Date(a.timestamp).toDateString() === new Date().toDateString()).length}</b></div><div><span>Check-outs today</span><b>{attendance.filter(a => a.event === 'Check-out' && new Date(a.timestamp).toDateString() === new Date().toDateString()).length}</b></div></div>
              <div className="device-notice"><Fingerprint size={22}/><div><b>Biometric hardware connection required</b><p>This screen validates the learner's admission number and device-assigned user reference. Automatic verified scans require a compatible fingerprint/face terminal and its supported API or SDK. A typed reference alone is not proof of a biometric scan.</p></div></div>
              <form className="attendance-form" onSubmit={recordAttendance}><div className="form-grid"><label>Attendance event<select value={attendanceEvent} onChange={e => setAttendanceEvent(e.target.value as 'Check-in' | 'Check-out')}><option>Check-in</option><option>Check-out</option></select></label><label>Boarder admission number<input required value={attendanceAdmission} onChange={e => setAttendanceAdmission(e.target.value)} placeholder="Enter admission number"/></label><label>Device user reference<input required value={attendanceDeviceRef} onChange={e => setAttendanceDeviceRef(e.target.value)} placeholder="Reference returned by device"/></label></div><button className="btn btn-primary" type="submit" disabled={attendanceBusy}><Fingerprint size={16}/>{attendanceBusy ? 'Verifying…' : 'Validate and record event'}</button></form>
              <div className="table-toolbar"><div className="search-box"><Search size={17}/><input value={attendanceSearch} onChange={e => setAttendanceSearch(e.target.value)} placeholder="Search attendance by name or admission number…"/></div></div>
              <div className="table-wrap"><table><thead><tr><th>BOARDER</th><th>GRADE</th><th>EVENT</th><th>TIME</th><th>METHOD</th></tr></thead><tbody>{attendance.filter(a => [a.fullName,a.admissionNo,a.grade].some(v => v.toLowerCase().includes(attendanceSearch.toLowerCase()))).map(a => <tr key={a.id}><td><b>{a.fullName}</b><small>{a.admissionNo}</small></td><td>{a.grade} · {a.stream}</td><td><span className={'condition ' + (a.event === 'Check-in' ? 'good' : 'repair')}>{a.event}</span></td><td>{new Date(a.timestamp).toLocaleString()}</td><td>{a.method}</td></tr>)}</tbody></table></div>
              {attendance.length === 0 && <div className="empty-state">No attendance events yet. Register boarder device references and connect the supported biometric-device integration.</div>}
            </section>
          )}
          {section === 'Reports' && (
            <div className="report-grid">
              <ReportCard
                icon={<Boxes />}
                title="Stock register"
                desc="Current quantities, categories, condition and location."
                action="Export stock CSV"
                onClick={exportCsv}
              />
              <ReportCard
                icon={<BarChart3 />}
                title="Asset valuation"
                desc="Estimated value calculated from quantity and unit cost."
                action="Export valuation data"
                onClick={exportCsv}
              />
              <ReportCard
                icon={<ArrowLeftRight />}
                title="Movement log"
                desc="Issues, returns, allocations and maintenance history."
                action="View movement log"
                onClick={() => setSection('Issue & return')}
              />
            </div>
          )}
          {section === 'Data backup' && (
            <section className="panel full-panel backup-panel">
              <div className="panel-head"><div><h2>Database backup & recovery</h2><p>Download a copy of your saved school records and restore records from a previous Hill School backup.</p></div><span className="category-chip"><ShieldCheck size={14}/> Signed-in account</span></div>
              <div className="backup-warning"><AlertTriangle size={19}/><div><b>Protect school records</b><p>Download backups regularly and store copies in a secure location outside this computer. Backup files may contain learner, staff and inventory information. Restoring merges matching record IDs and adds missing records; it does not delete existing records.</p></div></div>
              <div className="backup-actions">
                <div className="backup-action-card"><div className="backup-action-icon"><Download size={20}/></div><h3>Download database backup</h3><p>Export inventory items, stock movements, legacy stock updates, people, boarding attendance, classes, blocks and suppliers to a JSON file.</p><button className="btn btn-primary" onClick={downloadFullBackup} disabled={backupBusy || restoreBusy}>{backupBusy ? 'Preparing backup…' : 'Download full backup'}</button></div>
                <div className="backup-action-card"><div className="backup-action-icon"><Upload size={20}/></div><h3>Restore from backup</h3><p>Select a Hill School JSON backup. Records are merged by their saved IDs, so current records not in the file are not deleted.</p><label className="btn btn-secondary backup-file-label">{restoreBusy ? 'Restoring…' : 'Choose backup file'}<input type="file" accept="application/json,.json" onChange={restoreBackupFile} disabled={restoreBusy || backupBusy}/></label></div>
              </div>
              <div className="backup-details"><b>Included data</b><div className="feature-pills"><span>Inventory</span><span>Stock movements</span><span>Students & staff</span><span>Classes & streams</span><span>Blocks & dormitories</span><span>Suppliers</span><span>Boarding attendance</span></div><p>For a complete backup, wait for the success message and keep the downloaded file. Large datasets may take time to export.</p></div>
              {backupStatus && <div className={'backup-status ' + (backupStatus.toLowerCase().includes('incomplete') ? 'backup-status-error' : '')} role="status">{backupBusy || restoreBusy ? <Activity size={16}/> : <CheckCircle2 size={16}/>}<span>{backupStatus}</span></div>}
            </section>
          )}
          {section === 'Maintenance' && (
            <section className="panel placeholder-panel">
              <div className="placeholder-icon">
                <Wrench size={25} />
              </div>
              <h2>Maintenance & repairs</h2>
              <p>
                Items marked “Needs repair” appear in the attention list. A
                dedicated work-order workflow can record faults, assigned
                technicians, costs and completion.
              </p>
              <button
                className="btn btn-primary"
                onClick={() =>
                  setNotice(
                    'Maintenance work-order workflow identified for the next release.'
                  )
                }
              >
                Review maintenance queue
              </button>
            </section>
          )}
          {section === 'Settings' && (
            <section className="panel placeholder-panel">
              <div className="placeholder-icon">
                <Settings2 size={25} />
              </div>
              <h2>School settings</h2>
              <p>
                Configure school year, categories, storage locations, user
                roles, backup policy and audit retention before rolling out to
                all departments.
              </p>
              <div className="feature-pills">
                <span>Kenya · KES</span>
                <span>School year 2026</span>
                <span>Barcode-ready</span>
                <span>
                  {loaded
                    ? 'Connected to app backend'
                    : 'Connecting to backend…'}
                </span>
              </div>
            </section>
          )}
          <footer className="footer">
            <span>© 2026 Hill School · Inventory intelligence</span>
            <span>
              <span className="footer-dot"></span>{' '}
              {loaded ? 'Workspace ready' : 'Connecting…'}{' '}
              <span className="footer-sep">·</span> Built for accountable school
              operations
            </span>
          </footer>
        </div>
      </main>
      {showForm && (
        <div
          className="modal-backdrop"
          onMouseDown={e => {
            if (e.target === e.currentTarget) setShowForm(false);
          }}
        >
          <form className="modal" onSubmit={addAsset}>
            <div className="modal-head">
              <div>
                <span className="eyebrow">INVENTORY REGISTER</span>
                <h2>Add inventory item</h2>
              </div>
              <button
                type="button"
                className="icon-button"
                onClick={() => setShowForm(false)}
              >
                <X size={18} />
              </button>
            </div>
            <label>
              Item name
              <input
                required
                value={form.name}
                onChange={e => setForm({ ...form, name: e.target.value })}
                placeholder="e.g. Grade 8 Science textbook"
              />
            </label>
            <div className="form-grid">
              <label>
                Category
                <select
                  value={form.category}
                  onChange={e => setForm({ ...form, category: e.target.value })}
                >
                  {[
                    'Textbooks',
                    'Furniture',
                    'ICT',
                    'Boarding',
                    'Stationery',
                    'Sports',
                    'Laboratory',
                    'Cleaning',
                    'Other',
                  ].map(c => (
                    <option key={c}>{c}</option>
                  ))}
                </select>
              </label>
              <label>
                Item type
                <select
                  value={form.kind}
                  onChange={e => setForm({ ...form, kind: e.target.value })}
                >
                  <option>Returnable</option>
                  <option>Consumable</option>
                </select>
              </label>
              <label>
                Quantity
                <input
                  required
                  min="0"
                  type="number"
                  value={form.quantity}
                  onChange={e => setForm({ ...form, quantity: e.target.value })}
                />
              </label>
              <label>
                Unit cost (KES)
                <input
                  required
                  min="0"
                  type="number"
                  value={form.cost}
                  onChange={e => setForm({ ...form, cost: e.target.value })}
                  placeholder="0"
                />
              </label>
              <label>
                Condition
                <select
                  value={form.condition}
                  onChange={e =>
                    setForm({ ...form, condition: e.target.value })
                  }
                >
                  <option>Good</option>
                  <option>Needs repair</option>
                  <option>Damaged</option>
                </select>
              </label>
              <label>
                Location
                <input
                  required
                  value={form.location}
                  onChange={e => setForm({ ...form, location: e.target.value })}
                  placeholder="Class, store, block or dorm"
                />
              </label>
            </div>
            <label>
              Barcode / asset code
              <input
                value={form.barcode}
                onChange={e => setForm({ ...form, barcode: e.target.value })}
                placeholder="Leave blank to generate a code"
              />
            </label>
            <div className="modal-actions">
              <button
                type="button"
                className="btn btn-secondary"
                onClick={() => setShowForm(false)}
              >
                Cancel
              </button>
              <button type="submit" className="btn btn-primary">
                <Plus size={16} /> Save item
              </button>
            </div>
          </form>
        </div>
      )}
      {showPersonForm && (
        <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) setShowPersonForm(false); }}>
          <form className="modal person-modal" onSubmit={savePerson}>
            <div className="modal-head"><div><span className="eyebrow">SCHOOL DIRECTORY</span><h2>{editingPersonId ? 'Edit ' + (personForm.recordType === 'Student' ? 'learner' : 'staff') + ' record' : 'Register learner or staff'}</h2></div><button type="button" className="icon-button" onClick={() => { setShowPersonForm(false); resetPersonForm(); }}><X size={18}/></button></div>
            {!editingPersonId && <label>Record type<select value={personForm.recordType} onChange={e => setPersonForm({...personForm, recordType: e.target.value as 'Student' | 'Staff', residence: e.target.value === 'Staff' ? 'Staff' : 'Day scholar'})}><option value="Student">Student</option><option value="Staff">Staff</option></select></label>}
            <label>Full name<input required value={personForm.fullName} onChange={e => setPersonForm({...personForm, fullName:e.target.value})} placeholder="Enter full name as recorded by the school"/></label>
            {personForm.recordType === 'Student' ? <div className="form-grid"><label>Admission number<input required value={personForm.admissionNo} onChange={e => setPersonForm({...personForm,admissionNo:e.target.value})} placeholder="e.g. HS/2026/001"/></label><label>Class<select required value={personForm.grade} onChange={e => { const nextClass = e.target.value; const classRow = schoolStructure.classes.find(row => row.name === nextClass); const streams = classRow?.streams || ['A','B','C','D']; setPersonForm({...personForm,grade:nextClass,stream:streams.includes(personForm.stream) ? personForm.stream : streams[0] || ''}); }}>{[...new Set([...schoolStructure.classes.map(row => row.name), personForm.grade || 'Grade 1'])].filter(Boolean).map(name => <option key={name}>{name}</option>)}</select></label><label>Stream<select required value={personForm.stream} onChange={e => setPersonForm({...personForm,stream:e.target.value})}>{(schoolStructure.classes.find(row => row.name === personForm.grade)?.streams || ['A','B','C','D']).map(s => <option key={s}>{s}</option>)}</select></label><label>Type<select value={personForm.residence as 'Boarder' | 'Day scholar' | 'Staff'} onChange={e => setPersonForm({...personForm,residence:e.target.value as 'Boarder' | 'Day scholar' | 'Staff',dormitory:e.target.value === 'Boarder' ? personForm.dormitory : ''})}><option value="Day scholar">Day Scholar</option><option value="Boarder">Boarder</option></select></label><label>Gender<select value={personForm.gender} onChange={e => setPersonForm({...personForm,gender:e.target.value})}><option value="">Select gender</option><option>Female</option><option>Male</option><option>Prefer not to say</option></select></label><label>Parent / guardian name<input value={personForm.parentName} onChange={e => setPersonForm({...personForm,parentName:e.target.value})} placeholder="Enter parent or guardian full name"/></label><label>Parent / guardian phone<input value={personForm.phone} onChange={e => setPersonForm({...personForm,phone:e.target.value})} placeholder="Optional contact number"/></label>{personForm.residence === 'Boarder' && <label>Dormitory<select required value={personForm.dormitory} onChange={e => setPersonForm({...personForm,dormitory:e.target.value})}><option value="">Select dormitory</option>{schoolStructure.blocks.flatMap(block => (block.dormitories || []).map(dorm => ({value:block.name + ' / ' + dorm, label:block.name + ' / ' + dorm}))).map(dorm => <option key={dorm.value} value={dorm.value}>{dorm.label}</option>)}{personForm.dormitory && !schoolStructure.blocks.some(block => (block.dormitories || []).some(dorm => block.name + ' / ' + dorm === personForm.dormitory)) && <option value={personForm.dormitory}>{personForm.dormitory}</option>}</select><small>Choose a dormitory registered under Blocks & dormitories.</small></label>}</div> : <div className="form-grid"><label>Staff category<select value={personForm.staffType} onChange={e => setPersonForm({...personForm,staffType:e.target.value as 'Teaching' | 'Non-teaching'})}><option>Teaching</option><option>Non-teaching</option></select></label><label>{personForm.staffType === 'Teaching' ? 'TSC number' : 'National ID number'}<input required value={personForm.staffNo} onChange={e => setPersonForm({...personForm,staffNo:e.target.value})} placeholder={personForm.staffType === 'Teaching' ? 'Enter TSC number' : 'Enter national ID number'}/></label><label>Role / designation<input value={personForm.role} onChange={e => setPersonForm({...personForm,role:e.target.value})} placeholder="e.g. Teacher, bursar, cook"/></label><label>Gender<select value={personForm.gender} onChange={e => setPersonForm({...personForm,gender:e.target.value})}><option value="">Select gender</option><option>Female</option><option>Male</option><option>Prefer not to say</option></select></label><label>Phone<input value={personForm.phone} onChange={e => setPersonForm({...personForm,phone:e.target.value})} placeholder="Optional contact number"/></label></div>}
            {personForm.recordType === 'Student' && personForm.residence === 'Boarder' && <label>Biometric device user reference<input value={personForm.biometricRef} onChange={e => setPersonForm({...personForm, biometricRef:e.target.value})} placeholder="Device-assigned user ID; do not enter biometric data"/><small>Use the identifier supplied by a compatible attendance terminal. This field does not enrol a fingerprint or face.</small></label>}
            <label>Record status<select value={personForm.status} onChange={e => setPersonForm({...personForm,status:e.target.value as 'Active' | 'Inactive' | 'Transferred'})}><option>Active</option><option>Transferred</option><option>Inactive</option></select></label>
            <div className="modal-actions"><button type="button" className="btn btn-secondary" onClick={() => { setShowPersonForm(false); resetPersonForm(); }}>Cancel</button><button type="submit" className="btn btn-primary"><UserPlus size={16}/> {editingPersonId ? 'Save changes' : 'Save registration'}</button></div>
          </form>
        </div>
      )}
      {showMove && (
        <div
          className="modal-backdrop"
          onMouseDown={e => {
            if (e.target === e.currentTarget) setShowMove(false);
          }}
        >
          <form className="modal" onSubmit={saveMovement}>
            <div className="modal-head">
              <div>
                <span className="eyebrow">BARCODE WORKFLOW</span>
                <h2>Issue, return or allocate</h2>
              </div>
              <button
                type="button"
                className="icon-button"
                onClick={() => setShowMove(false)}
              >
                <X size={18} />
              </button>
            </div>
            <div className="scan-note">
              <QrCode size={22} />
              <span>
                <b>Barcode scanner ready</b>
                <small>
                  Connect a USB scanner or type the barcode below. Most scanners
                  act like a keyboard.
                </small>
              </span>
            </div>
            <label>
              Barcode / asset code
              <input
                required
                autoFocus
                value={move.barcode}
                onChange={e => setMove({ ...move, barcode: e.target.value })}
                placeholder="Scan or enter barcode"
              />
            </label>
            <div className="form-grid">
              <label>
                Transaction
                <select
                  value={move.action}
                  onChange={e => setMove({ ...move, action: e.target.value })}
                >
                  <option>Issued</option>
                  <option>Returned</option>
                  <option>Allocated</option>
                  <option>Maintenance</option>
                </select>
              </label>
              <label>
                Quantity
                <input
                  type="number"
                  required
                  min="1"
                  value={move.quantity}
                  onChange={e => setMove({ ...move, quantity: e.target.value })}
                />
              </label>
            </div>
            <label>
              Recipient / class / location
              <input
                value={move.person}
                onChange={e => setMove({ ...move, person: e.target.value })}
                placeholder="e.g. Grade 7 Stream A or learner admission no."
              />
            </label>
            <div className="modal-actions">
              <button
                type="button"
                className="btn btn-secondary"
                onClick={() => setShowMove(false)}
              >
                Cancel
              </button>
              <button type="submit" className="btn btn-primary">
                <ArrowLeftRight size={16} /> Record transaction
              </button>
            </div>
          </form>
        </div>
      )}
    </div>
  );
}

function Metric({
  icon,
  label,
  value,
  note,
  tone,
}: {
  icon: React.ReactNode;
  label: string;
  value: string;
  note: string;
  tone: string;
}) {
  return (
    <div className="metric-card">
      <div className={'metric-icon ' + tone}>{icon}</div>
      <div className="metric-label">{label}</div>
      <div className="metric-value">{value}</div>
      <div className="metric-note">{note}</div>
    </div>
  );
}
function InventoryTable({
  assets,
  onSelect,
}: {
  assets: Asset[];
  onSelect: (a: Asset) => void;
}) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>ITEM / ASSET</th>
            <th>CATEGORY</th>
            <th>QUANTITY</th>
            <th>CONDITION</th>
            <th>LOCATION</th>
            <th>UNIT COST</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          {assets.map(a => (
            <tr key={a.id} onClick={() => onSelect(a)} className="click-row">
              <td>
                <div className="item-cell">
                  <div
                    className={
                      'item-symbol ' +
                      (a.category === 'Textbooks'
                        ? 'book-symbol'
                        : a.category === 'ICT'
                          ? 'ict-symbol'
                          : 'furniture-symbol')
                    }
                  >
                    {a.category === 'Textbooks' ? (
                      <BookOpen size={17} />
                    ) : a.category === 'ICT' ? (
                      <Activity size={17} />
                    ) : (
                      <Archive size={17} />
                    )}
                  </div>
                  <div>
                    <b>{a.name}</b>
                    <small>{a.barcode}</small>
                  </div>
                </div>
              </td>
              <td>
                <span className="category-chip">{a.category}</span>
              </td>
              <td>
                <b>{a.quantity}</b> <small>units</small>
              </td>
              <td>
                <span
                  className={
                    'condition ' +
                    (a.condition === 'Good'
                      ? 'good'
                      : a.condition === 'Damaged'
                        ? 'damaged'
                        : 'repair')
                  }
                >
                  {a.condition}
                </span>
              </td>
              <td>{a.location}</td>
              <td>{money(a.cost)}</td>
              <td>
                <button
                  className="row-more"
                  onClick={e => {
                    e.stopPropagation();
                    onSelect(a);
                  }}
                  aria-label={'Details for ' + a.name}
                >
                  ···
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
function ReportCard({
  icon,
  title,
  desc,
  action,
  onClick,
}: {
  icon: React.ReactNode;
  title: string;
  desc: string;
  action: string;
  onClick: () => void;
}) {
  return (
    <section className="panel report-card">
      <div className="report-icon">{icon}</div>
      <h3>{title}</h3>
      <p>{desc}</p>
      <button className="text-button" onClick={onClick}>
        {action} →
      </button>
    </section>
  );
}

export default App;
