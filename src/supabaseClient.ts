import { createClient } from '@supabase/supabase-js';

const env = (import.meta as any).env || {};
const SUPABASE_URL: string = env.VITE_SUPABASE_URL || '';
const SUPABASE_KEY: string = env.VITE_SUPABASE_ANON_KEY || '';
const ADMIN_EMAIL = 'ahhmadaliyou2000@gmail.com';

export const supabase = createClient(
  SUPABASE_URL || 'https://missing-url.supabase.co',
  SUPABASE_KEY || 'missing-key'
);

/** Make every error look like the shape App.tsx already reads. */
function fail(message: string): never {
  const err: any = new Error(message);
  err.response = { data: { error: message } };
  throw err;
}
function check<T>(result: { data: T; error: any }): T {
  if (result.error) {
    const code = result.error.code;
    if (code === '23505')
      fail('A staff member with that Staff ID or email already exists.');
    fail(result.error.message || 'Something went wrong. Please try again.');
  }
  return result.data;
}
function lagosToday(): string {
  return new Date().toLocaleDateString('en-CA', { timeZone: 'Africa/Lagos' });
}
function hhmm(value: string | null | undefined, fallback: string): string {
  return value ? String(value).slice(0, 5) : fallback;
}

function mapStaff(row: any) {
  return {
    id: row.id,
    fullName: row.full_name,
    staffId: row.staff_id,
    email: row.email,
    phone: row.phone || '',
    department: row.department || '',
    designation: row.designation || '',
    status: row.status,
    createdAt: row.created_at,
  };
}
function mapAttendance(row: any, staffById: Record<string, any>) {
  const person = staffById[row.staff_record_id] || {};
  return {
    id: row.id,
    staffRecordId: row.staff_record_id,
    staffName: person.fullName || 'Unknown',
    staffId: person.staffId || '',
    email: person.email || '',
    date: row.date,
    checkIn: row.check_in,
    checkOut: row.check_out || undefined,
    totalHours: row.total_hours != null ? Number(row.total_hours) : undefined,
    status: row.status,
  };
}

async function currentEmail(): Promise<string> {
  const { data } = await supabase.auth.getSession();
  return (data.session?.user?.email || '').toLowerCase();
}

async function bootstrap() {
  const email = await currentEmail();
  if (!email) fail('You are not signed in.');
  const isAdmin = email === ADMIN_EMAIL;

  const settingsRow = check(
    await supabase.from('settings').select('*').eq('id', 1).maybeSingle()
  ) as any;
  const settings = {
    checkInStart: hhmm(settingsRow?.check_in_start, '07:30'),
    resumption: hhmm(settingsRow?.resumption, '08:00'),
    closing: hhmm(settingsRow?.closing, '16:00'),
    graceMinutes: settingsRow?.grace_minutes ?? 0,
  };

  const staffRows = check(
    await supabase.from('staff').select('*').order('full_name')
  ) as any[];
  const staffList = staffRows.map(mapStaff);
  const profile =
    staffList.find(person => person.email.toLowerCase() === email) || null;

  if (!isAdmin) {
    if (!profile)
      fail(
        'Your email is not registered. Ask the IT Unit administrator to register your email.'
      );
    if (profile.status !== 'active')
      fail('Your staff account is inactive. Contact the IT Unit administrator.');
  }

  const attendanceRows = check(
    await supabase
      .from('attendance')
      .select('*')
      .order('date', { ascending: false })
      .limit(5000)
  ) as any[];
  const staffById: Record<string, any> = {};
  staffList.forEach(person => (staffById[person.id] = person));

  return {
    data: {
      isAdmin,
      profile,
      staff: isAdmin ? staffList : [],
      attendance: attendanceRows.map(row => mapAttendance(row, staffById)),
      settings,
      today: lagosToday(),
    },
  };
}

function staffPayload(body: any) {
  return {
    full_name: String(body.fullName || '').trim(),
    staff_id: String(body.staffId || '').trim(),
    email: String(body.email || '').trim().toLowerCase(),
    phone: String(body.phone || '').trim() || null,
    department: String(body.department || '').trim() || null,
    designation: String(body.designation || '').trim() || null,
  };
}

export const api = {
  async get(path: string) {
    if (path === '/api/bootstrap') return bootstrap();
    if (path === '/api/admin/qr') {
      const data = check(await supabase.rpc('admin_new_qr'));
      return { data };
    }
    fail('Unknown request: ' + path);
  },
  async post(path: string, body: any) {
    if (path === '/api/attendance/checkin' || path === '/api/attendance/checkout') {
      const action = path.endsWith('checkout') ? 'checkout' : 'checkin';
      const data = check(
        await supabase.rpc('record_attendance', {
          p_action: action,
          p_token: body?.token,
        })
      );
      return { data };
    }
    if (path === '/api/admin/staff') {
      const data = check(
        await supabase.from('staff').insert(staffPayload(body)).select().single()
      );
      return { data };
    }
    fail('Unknown request: ' + path);
  },
  async put(path: string, body: any) {
    if (path.startsWith('/api/admin/staff/')) {
      const id = decodeURIComponent(path.replace('/api/admin/staff/', ''));
      const data = check(
        await supabase
          .from('staff')
          .update(staffPayload(body))
          .eq('id', id)
          .select()
          .single()
      );
      return { data };
    }
    if (path === '/api/admin/settings') {
      const data = check(
        await supabase
          .from('settings')
          .update({
            check_in_start: body.checkInStart,
            resumption: body.resumption,
            closing: body.closing,
            grace_minutes: Number(body.graceMinutes) || 0,
          })
          .eq('id', 1)
          .select()
          .single()
      );
      return { data };
    }
    fail('Unknown request: ' + path);
  },
  async delete(path: string) {
    if (path.startsWith('/api/admin/staff/')) {
      const id = decodeURIComponent(path.replace('/api/admin/staff/', ''));
      check(
        await supabase.from('staff').update({ status: 'inactive' }).eq('id', id)
      );
      return { data: { ok: true } };
    }
    fail('Unknown request: ' + path);
  },
};

export const auth = {
  async getUser() {
    const { data } = await supabase.auth.getSession();
    const user = data.session?.user;
    if (!user) return null;
    return {
      email: user.email || '',
      name: (user.user_metadata as any)?.full_name || user.email || '',
    };
  },
  async signInWithPassword(email: string, password: string) {
    const { data, error } = await supabase.auth.signInWithPassword({
      email: email.trim().toLowerCase(),
      password,
    });
    if (error) fail(error.message);
    const user = data.user;
    return { email: user?.email || email, name: user?.email || email };
  },
  async signUp(email: string, password: string) {
    const { data, error } = await supabase.auth.signUp({
      email: email.trim().toLowerCase(),
      password,
    });
    if (error) fail(error.message);
    if (!data.session)
      fail(
        'Account created. Please check your email to confirm it, then sign in.'
      );
    const user = data.user;
    return { email: user?.email || email, name: user?.email || email };
  },
  async signOut() {
    await supabase.auth.signOut();
  },
};
