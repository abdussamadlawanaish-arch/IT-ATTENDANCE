import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { auth, api } from '@appdeploy/client';
import { Html5Qrcode } from 'html5-qrcode';
import { QRCodeSVG } from 'qrcode.react';
import * as XLSX from 'xlsx';
import {
  Activity,
  AlertCircle,
  ArrowDownToLine,
  BarChart3,
  CalendarDays,
  Check,
  CheckCircle2,
  ChevronDown,
  Clock3,
  Download,
  FileSpreadsheet,
  Filter,
  LayoutDashboard,
  LogOut,
  Menu,
  Monitor,
  QrCode,
  RefreshCw,
  Search,
  Settings,
  ShieldCheck,
  Smartphone,
  Users,
  X,
  Camera,
  UserPlus,
  Building2,
  CircleHelp,
  ClipboardList,
  CircleAlert,
  ScanLine,
} from 'lucide-react';

type Role = 'admin' | 'staff';
type Staff = {
  id: string;
  fullName: string;
  staffId: string;
  email: string;
  phone?: string;
  department?: string;
  designation?: string;
  status: 'active' | 'inactive';
  createdAt?: string;
};
type Attendance = {
  id: string;
  staffRecordId: string;
  staffName: string;
  staffId: string;
  email: string;
  date: string;
  checkIn: string;
  checkOut?: string;
  localCheckOut?: string;
  totalHours?: number;
  status: 'PRESENT' | 'LATE';
};
type SettingsData = {
  checkInStart: string;
  resumption: string;
  closing: string;
  graceMinutes: number;
};
type Bootstrap = {
  isAdmin: boolean;
  profile: Staff | null;
  staff: Staff[];
  attendance: Attendance[];
  settings: SettingsData;
  today: string;
};

const ADMIN_EMAIL = 'ahhmadaliyou2000@gmail.com';
const DEFAULT_SETTINGS: SettingsData = {
  checkInStart: '07:30',
  resumption: '08:00',
  closing: '16:00',
  graceMinutes: 0,
};
const navAdmin = [
  { id: 'overview', label: 'Overview', icon: LayoutDashboard },
  { id: 'staff', label: 'Staff directory', icon: Users },
  { id: 'attendance', label: 'Attendance records', icon: ClipboardList },
  { id: 'qr', label: 'QR display', icon: QrCode },
  { id: 'reports', label: 'Reports', icon: BarChart3 },
  { id: 'settings', label: 'Settings', icon: Settings },
];
const navStaff = [
  { id: 'overview', label: 'My dashboard', icon: LayoutDashboard },
  { id: 'scan', label: 'Scan attendance QR', icon: ScanLine },
];

function dateLabel(value?: string) {
  if (!value) return '—';
  return new Date(value + 'T12:00:00').toLocaleDateString('en-NG', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
  });
}
function timeLabel(value?: string) {
  if (!value) return '—';
  return new Date(value).toLocaleTimeString('en-NG', {
    hour: '2-digit',
    minute: '2-digit',
    hour12: true,
    timeZone: 'Africa/Lagos',
  });
}
function initials(name = '') {
  return (
    name
      .split(' ')
      .filter(Boolean)
      .slice(0, 2)
      .map(part => part[0])
      .join('')
      .toUpperCase() || 'IT'
  );
}
function csvDownload(rows: Record<string, unknown>[], filename: string) {
  const sheet = XLSX.utils.json_to_sheet(rows);
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, 'Attendance');
  XLSX.writeFile(book, filename);
}
function StatusPill({ status }: { status: string }) {
  const cls =
    status === 'PRESENT' || status === 'active'
      ? 'success'
      : status === 'LATE'
        ? 'warning'
        : 'muted';
  return (
    <span className={'status-pill ' + cls}>
      <span />
      {status}
    </span>
  );
}

export default function App() {
  const [user, setUser] = useState<any>(null);
  const [boot, setBoot] = useState<Bootstrap | null>(null);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [active, setActive] = useState('overview');
  const [mobileNav, setMobileNav] = useState(false);
  const [search, setSearch] = useState('');
  const [filterDate, setFilterDate] = useState('');
  const [showStaffForm, setShowStaffForm] = useState(false);
  const [editing, setEditing] = useState<Staff | null>(null);
  const [form, setForm] = useState({
    fullName: '',
    staffId: '',
    email: '',
    phone: '',
    department: 'Information Technology',
    designation: '',
  });
  const [qr, setQr] = useState<{
    token: string;
    expiresAt: number;
    date: string;
  } | null>(null);
  const [settingsForm, setSettingsForm] =
    useState<SettingsData>(DEFAULT_SETTINGS);
  const [scannerMessage, setScannerMessage] = useState('');
  const scannerRef = useRef<Html5Qrcode | null>(null);

  const isAdmin = Boolean(boot?.isAdmin);
  const nav = isAdmin ? navAdmin : navStaff;

  const loadData = useCallback(async () => {
    const response = await api.get('/api/bootstrap');
    const data = response.data as Bootstrap;
    setBoot(data);
    setSettingsForm(data.settings || DEFAULT_SETTINGS);
    return data;
  }, []);

  useEffect(() => {
    let mounted = true;
    auth
      .getUser()
      .then(async current => {
        if (!mounted) return;
        setUser(current);
        if (current) {
          try {
            const response = await api.get('/api/bootstrap');
            if (mounted) {
              setBoot(response.data as Bootstrap);
              setSettingsForm(
                (response.data as Bootstrap).settings || DEFAULT_SETTINGS
              );
            }
          } catch (err: any) {
            if (mounted)
              setError(
                err?.response?.data?.error ||
                  err?.message ||
                  'Unable to load your account. Ask the IT Unit administrator to register your email.'
              );
          }
        }
        if (mounted) setLoading(false);
      })
      .catch(() => {
        if (mounted) {
          setError('Could not restore your sign-in session. Please try again.');
          setLoading(false);
        }
      });
    return () => {
      mounted = false;
    };
  }, []);

  useEffect(() => {
    if (!user) return;
    const timer = window.setInterval(() => {
      loadData().catch(() => undefined);
    }, 30000);
    return () => window.clearInterval(timer);
  }, [user, loadData]);

  useEffect(() => {
    if (active !== 'qr' || !isAdmin) return;
    let cancelled = false;
    const refreshQr = async () => {
      try {
        const response = await api.get('/api/admin/qr');
        if (!cancelled) setQr(response.data);
      } catch (err: any) {
        if (!cancelled)
          setError(
            err?.response?.data?.error ||
              'Could not generate the attendance QR code.'
          );
      }
    };
    refreshQr();
    const timer = window.setInterval(refreshQr, 45000);    return () => {
      cancelled = true;
      window.clearInterval(timer);
    };
  }, [active, isAdmin]);

  const stopScanner = useCallback(async () => {
    const scanner = scannerRef.current;
    if (scanner) {
      try {
        if (scanner.isScanning) await scanner.stop();
      } catch {
        /* camera may already be stopped */
      }
      try {
        scanner.clear();
      } catch {
        /* ignore cleanup */
      }
      scannerRef.current = null;
    }
  }, []);

  const submitScannedToken = useCallback(
    async (token: string) => {
      if (!token || busy) return;
      setBusy(true);
      setScannerMessage('Verifying QR code…');
      try {
        const todayRecord = boot?.attendance?.find(record => record.date === boot.today);
        const action = todayRecord?.checkIn && !todayRecord.checkOut ? 'checkout' : 'checkin';
        const response = await api.post('/api/attendance/' + action, { token });
        const recordedTime = action === 'checkout' ? response.data.checkOut : response.data.checkIn;
        setScannerMessage(
          (action === 'checkout' ? 'Check-out recorded successfully at ' : 'Check-in recorded successfully at ') +
            timeLabel(recordedTime) +
            (response.data.totalHours != null ? '. Total time: ' + response.data.totalHours.toFixed(2) + ' hours.' : '.')
        );
        setNotice(action === 'checkout' ? 'Your check-out has been recorded.' : 'Your check-in has been recorded.');
        await loadData();
        await stopScanner();
      } catch (err: any) {
        const message =
          err?.response?.data?.error ||
          err?.message ||
          'Unable to record attendance.';
        setScannerMessage(message);
      } finally {
        setBusy(false);
      }
    },
    [busy, boot, loadData, stopScanner]
  );

  useEffect(() => {
    if (active !== 'scan' || !user || isAdmin) {
      stopScanner();
      return;
    }
    let cancelled = false;
    const start = async () => {
      try {
        await new Promise(resolve => window.setTimeout(resolve, 250));
        if (cancelled) return;
        const scanner = new Html5Qrcode('fmc-qr-reader');
        scannerRef.current = scanner;
        await scanner.start(
          { facingMode: 'environment' },
          { fps: 10, qrbox: { width: 230, height: 230 }, aspectRatio: 1 },
          decoded => {
            if (!cancelled) submitScannedToken(decoded);
          },
          () => undefined
        );
      } catch {
        if (!cancelled)
          setScannerMessage(
            'Camera access is unavailable. Allow camera permission and reload, or ask your supervisor for assistance.'
          );
      }
    };
    start();
    return () => {
      cancelled = true;
      stopScanner();
    };
  }, [active, user, isAdmin, submitScannedToken, stopScanner]);

  const showMessage = (message: string) => {
    setNotice(message);
    window.setTimeout(() => setNotice(''), 3500);
  };
  const signIn = async () => {
    setError('');
    try {
      await auth.signInWithRedirect();
    } catch (err: any) {
      setError(err?.message || 'Sign-in could not start. Please retry.');
    }
  };
  const signOut = async () => {
    await stopScanner();
    await auth.signOut();
    setUser(null);
    setBoot(null);
    setActive('overview');
  };
  const openNewStaff = (staff?: Staff) => {
    setEditing(staff || null);
    setForm(
      staff
        ? {
            fullName: staff.fullName,
            staffId: staff.staffId,
            email: staff.email,
            phone: staff.phone || '',
            department: staff.department || 'Information Technology',
            designation: staff.designation || '',
          }
        : {
            fullName: '',
            staffId: '',
            email: '',
            phone: '',
            department: 'Information Technology',
            designation: '',
          }
    );
    setShowStaffForm(true);
  };
  const saveStaff = async (event: React.FormEvent) => {
    event.preventDefault();
    setBusy(true);
    setError('');
    try {
      if (editing)
        await api.put(
          '/api/admin/staff/' + encodeURIComponent(editing.id),
          form
        );
      else await api.post('/api/admin/staff', form);
      await loadData();
      setShowStaffForm(false);
      showMessage(
        editing ? 'Staff record updated.' : 'Staff member registered.'
      );
    } catch (err: any) {
      setError(
        err?.response?.data?.error ||
          err?.message ||
          'Could not save staff record.'
      );
    } finally {
      setBusy(false);
    }
  };
  const deactivateStaff = async (staff: Staff) => {
    if (
      !window.confirm(
        'Deactivate ' +
          staff.fullName +
          '? They will no longer be able to record attendance.'
      )
    )
      return;
    try {
      await api.delete('/api/admin/staff/' + encodeURIComponent(staff.id));
      await loadData();
      showMessage('Staff account deactivated.');
    } catch (err: any) {
      setError(
        err?.response?.data?.error || 'Could not deactivate staff account.'
      );
    }
  };
  const saveSettings = async (event: React.FormEvent) => {
    event.preventDefault();
    setBusy(true);
    try {
      await api.put('/api/admin/settings', settingsForm);
      await loadData();
      showMessage('Attendance settings saved.');
    } catch (err: any) {
      setError(err?.response?.data?.error || 'Could not save settings.');
    } finally {
      setBusy(false);
    }
  };

  const staff = boot?.staff || [];
  const attendance = boot?.attendance || [];
  const filteredStaff = useMemo(
    () =>
      staff.filter(person =>
        [person.fullName, person.staffId, person.email, person.department]
          .join(' ')
          .toLowerCase()
          .includes(search.toLowerCase())
      ),
    [staff, search]
  );
  const filteredAttendance = useMemo(
    () =>
      attendance.filter(
        record =>
          (!filterDate || record.date === filterDate) &&
          [record.staffName, record.staffId, record.email, record.status]
            .join(' ')
            .toLowerCase()
            .includes(search.toLowerCase())
      ),
    [attendance, filterDate, search]
  );
  const todayAttendance = attendance.filter(
    record => record.date === boot?.today
  );
  const presentCount = todayAttendance.filter(
    record => record.status === 'PRESENT'
  ).length;
  const lateCount = todayAttendance.filter(
    record => record.status === 'LATE'
  ).length;
  const absentCount = Math.max(
    0,
    staff.filter(person => person.status === 'active').length -
      todayAttendance.length
  );
  const attendanceRate = staff.filter(person => person.status === 'active')
    .length
    ? Math.round(
        (todayAttendance.length /
          staff.filter(person => person.status === 'active').length) *
          100
      )
    : 0;
  const ownRecords = attendance
    .slice()
    .sort((a, b) => b.checkIn.localeCompare(a.checkIn));
  const ownToday = ownRecords.find(record => record.date === boot?.today);
  const monthRecords = ownRecords.filter(
    record => record.date?.slice(0, 7) === boot?.today?.slice(0, 7)
  );
  const qrSeconds = qr
    ? Math.max(0, Math.ceil((qr.expiresAt - Date.now()) / 1000))
    : 0;

  if (loading)
    return (
      <div className="loading-screen">
        <div className="brand-mark">
          <ShieldCheck size={26} />
        </div>
        <div className="spinner" />        <p>Loading FMC IT Unit Attendance…</p>
      </div>
    );

  if (!user)
    return (
      <div className="login-shell">
        <div className="login-visual">
          <div className="login-brand">
            <div className="brand-mark">
              <ShieldCheck size={27} />
            </div>
            <div>
              <strong>FMC BIRNIN KUDU</strong>
              <span>INFORMATION TECHNOLOGY UNIT</span>
            </div>
          </div>
          <div className="login-copy">
            <span className="eyebrow">STAFF ATTENDANCE MANAGEMENT</span>
            <h1>
              Every day.
              <br />
              <em>Accounted for.</em>
            </h1>
            <p>
              A secure digital attendance platform for the Federal Medical
              Centre, Birnin Kudu IT Unit.
            </p>
            <div className="login-points">
              <span>
                <CheckCircle2 size={17} /> Secure QR check-in
              </span>
              <span>
                <CheckCircle2 size={17} /> Accurate attendance records
              </span>
              <span>
                <CheckCircle2 size={17} /> Clear reporting and accountability
              </span>
            </div>
          </div>
          <div className="login-footer">
            FEDERAL MEDICAL CENTRE, BIRNIN KUDU · JIGAWA STATE
          </div>
        </div>
        <div className="login-panel">
          <div className="login-card">
            <div className="mobile-brand">
              <div className="brand-mark">
                <ShieldCheck size={24} />
              </div>
              <span>FMC IT UNIT</span>
            </div>
            <span className="eyebrow">WELCOME</span>
            <h2>Sign in to continue</h2>
            <p>
              Use your registered Google or email sign-in account to access the
              attendance system.
            </p>
            {error && (
              <div className="alert error">
                <AlertCircle size={17} />
                {error}
              </div>
            )}
            <button className="primary-btn full" onClick={signIn}>
              Continue to secure sign in <ArrowDownToLine size={17} />
            </button>
            <div className="login-note">
              <ShieldCheck size={16} />
              <span>
                Access is restricted to registered IT Unit staff. The initial
                administrator is <b>{ADMIN_EMAIL}</b>.
              </span>
            </div>
          </div>
          <div className="panel-foot">
            Need access? Contact the IT Unit system administrator.
          </div>
        </div>
      </div>
    );

  if (user && !boot && error)
    return (
      <div className="login-shell">
        <div className="login-panel recovery">
          <div className="login-card">
            <div className="brand-mark">
              <ShieldCheck size={26} />
            </div>
            <h2>Account access</h2>
            <div className="alert error">{error}</div>
            <p>
              Your email must be registered by the IT Unit administrator before
              staff attendance features become available.
            </p>
            <button
              className="primary-btn full"
              onClick={() =>
                loadData()
                  .then(setBoot)
                  .catch(e =>
                    setError(
                      e?.response?.data?.error ||
                        'Still unable to load account.'
                    )
                  )
              }
            >
              Retry
            </button>
            <button className="text-btn" onClick={signOut}>
              Sign out
            </button>
          </div>
        </div>
      </div>
    );

  return (
    <div className="app-shell">
      <aside className={'sidebar ' + (mobileNav ? 'mobile-open' : '')}>
        <div className="sidebar-brand">
          <div className="brand-mark">
            <ShieldCheck size={23} />
          </div>
          <div>
            <b>FMC IT UNIT</b>
            <small>ATTENDANCE SYSTEM</small>
          </div>
          <button className="mobile-close" onClick={() => setMobileNav(false)}>
            <X size={18} />
          </button>
        </div>
        <div className="facility-label">
          <Building2 size={14} /> FEDERAL MEDICAL CENTRE
        </div>
        <div className="nav-label">WORKSPACE</div>
        <nav>
          {nav.map(item => {
            const Icon = item.icon;
            return (
              <button
                key={item.id}
                className={'nav-item ' + (active === item.id ? 'active' : '')}
                onClick={() => {
                  setActive(item.id);
                  setMobileNav(false);
                  setError('');
                }}
              >
                <Icon size={18} />
                <span>{item.label}</span>
                {item.id === 'scan' && <span className="nav-live" />}
              </button>
            );
          })}
        </nav>
        <div className="sidebar-bottom">
          <div className="security-card">
            <ShieldCheck size={18} />
            <div>
              <b>Secure workspace</b>
              <span>Protected staff access</span>
            </div>
          </div>
          <div className="user-mini">
            <div className="avatar">
              {initials(boot?.profile?.fullName || user.name || user.email)}
            </div>
            <div className="user-mini-copy">
              <b>{boot?.profile?.fullName || user.name || 'Signed-in user'}</b>
              <span>{isAdmin ? 'Super Administrator' : 'IT Unit Staff'}</span>
            </div>
            <button title="Sign out" onClick={signOut}>
              <LogOut size={17} />
            </button>
          </div>
        </div>
      </aside>
      {mobileNav && (
        <button
          className="mobile-overlay"
          onClick={() => setMobileNav(false)}
          aria-label="Close menu"
        />
      )}
      <main className="main-area">
        <header className="topbar">
          <div className="topbar-left">
            <button className="menu-btn" onClick={() => setMobileNav(true)}>
              <Menu size={20} />
            </button>
            <div>
              <div className="breadcrumb">
                FMC BIRNIN KUDU <span>/</span> IT UNIT
              </div>
              <h1>
                {nav.find(item => item.id === active)?.label || 'Dashboard'}
              </h1>
            </div>
          </div>
          <div className="topbar-right">
            <div className="date-chip">
              <CalendarDays size={16} />
              <span>{dateLabel(boot?.today)}</span>
            </div>
            <button
              className="icon-btn"
              onClick={() =>
                loadData()
                  .then(() => showMessage('Attendance data refreshed.'))
                  .catch(() => setError('Refresh failed.'))
              }
              title="Refresh data"
            >
              <RefreshCw size={17} />
            </button>
            <div className="top-avatar">
              {initials(boot?.profile?.fullName || user.name || user.email)}
            </div>
          </div>
        </header>
        <div className="page-content">
          {error && (
            <div className="alert error page-alert">
              <AlertCircle size={17} />
              {error}
              <button onClick={() => setError('')}>
                <X size={16} />
              </button>
            </div>
          )}
          {notice && (
            <div className="toast">
              <CheckCircle2 size={17} />
              {notice}
            </div>
          )}

          {active === 'overview' && isAdmin && (
            <section>
              <div className="welcome-row">
                <div>
                  <span className="eyebrow">ADMINISTRATION OVERVIEW</span>
                  <h2>Good day, Administrator</h2>
                  <p>Here is the IT Unit attendance summary for today.</p>
                </div>
                <button className="outline-btn" onClick={() => setActive('qr')}>
                  <QrCode size={17} /> Open QR display                </button>
              </div>
              <div className="stat-grid">
                <div className="stat-card">
                  <div className="stat-top">
                    <span>Total registered staff</span>
                    <span className="stat-icon blue">
                      <Users size={18} />
                    </span>
                  </div>
                  <strong>
                    {staff.filter(s => s.status === 'active').length}
                  </strong>
                  <small>Active IT Unit accounts</small>
                </div>
                <div className="stat-card">
                  <div className="stat-top">
                    <span>Present today</span>
                    <span className="stat-icon green">
                      <CheckCircle2 size={18} />
                    </span>
                  </div>
                  <strong>{presentCount}</strong>
                  <small>Checked in on time</small>
                </div>
                <div className="stat-card">
                  <div className="stat-top">
                    <span>Late arrivals</span>
                    <span className="stat-icon amber">
                      <Clock3 size={18} />
                    </span>
                  </div>
                  <strong>{lateCount}</strong>
                  <small>Checked in after resumption</small>
                </div>
                <div className="stat-card">
                  <div className="stat-top">
                    <span>Not yet recorded</span>
                    <span className="stat-icon red">
                      <CircleAlert size={18} />
                    </span>
                  </div>
                  <strong>{absentCount}</strong>
                  <small>Absence is provisional until closing</small>
                </div>
              </div>
              <div className="overview-grid">
                <div className="panel attendance-panel">
                  <div className="panel-heading">
                    <div>
                      <h3>Today's attendance</h3>
                      <p>Live view · updates automatically</p>
                    </div>
                    <button
                      className="text-btn"
                      onClick={() => setActive('attendance')}
                    >
                      View all <span>→</span>
                    </button>
                  </div>
                  <div className="rate-block">
                    <div>
                      <strong>{attendanceRate}%</strong>
                      <span>Attendance rate</span>
                    </div>
                    <div className="rate-track">
                      <span style={{ width: attendanceRate + '%' }} />
                    </div>
                  </div>
                  <div className="mini-legend">
                    <span>
                      <i className="dot green-dot" /> {todayAttendance.length}{' '}
                      checked in
                    </span>
                    <span>
                      <i className="dot amber-dot" /> {lateCount} late
                    </span>
                    <span>
                      <i className="dot grey-dot" /> {absentCount} pending
                    </span>
                  </div>
                  <div className="table-wrap">
                    <table>
                      <thead>
                        <tr>
                          <th>STAFF MEMBER</th>
                          <th>CHECK-IN</th>
                          <th>STATUS</th>
                        </tr>
                      </thead>
                      <tbody>
                        {todayAttendance
                          .slice()
                          .sort((a, b) => b.checkIn.localeCompare(a.checkIn))
                          .slice(0, 6)
                          .map(record => (
                            <tr key={record.id}>
                              <td>
                                <div className="person-cell">
                                  <div className="avatar small">
                                    {initials(record.staffName)}
                                  </div>
                                  <div>
                                    <b>{record.staffName}</b>
                                    <small>{record.staffId}</small>
                                  </div>
                                </div>
                              </td>
                              <td>{timeLabel(record.checkIn)}</td>
                              <td>
                                <StatusPill status={record.status} />
                              </td>
                            </tr>
                          ))}
                        {!todayAttendance.length && (
                          <tr>
                            <td colSpan={3}>
                              <div className="empty-state">
                                <QrCode size={25} />
                                <b>No check-ins recorded today</b>
                                <span>
                                  Attendance entries will appear here after
                                  staff scan the QR code.
                                </span>
                              </div>
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
                <div className="panel quick-panel">
                  <div className="panel-heading">
                    <div>
                      <h3>Quick actions</h3>
                      <p>Common administration tasks</p>
                    </div>
                  </div>
                  <button
                    className="quick-action"
                    onClick={() => openNewStaff()}
                  >
                    <span className="quick-icon blue">
                      <UserPlus size={19} />
                    </span>
                    <span>
                      <b>Register staff</b>
                      <small>Add a staff member to the system</small>
                    </span>
                    <span className="arrow">→</span>
                  </button>
                  <button
                    className="quick-action"
                    onClick={() => setActive('qr')}
                  >
                    <span className="quick-icon violet">
                      <QrCode size={19} />
                    </span>
                    <span>
                      <b>Display attendance QR</b>
                      <small>Open entrance QR display screen</small>
                    </span>
                    <span className="arrow">→</span>
                  </button>
                  <button
                    className="quick-action"
                    onClick={() => setActive('reports')}
                  >
                    <span className="quick-icon green">
                      <FileSpreadsheet size={19} />
                    </span>
                    <span>
                      <b>Export attendance</b>
                      <small>Download attendance data</small>
                    </span>
                    <span className="arrow">→</span>
                  </button>
                  <div className="tip-box">
                    <CircleHelp size={18} />
                    <div>
                      <b>Attendance reminder</b>
                      <p>
                        Staff should scan the active QR code at the entrance
                        using their own signed-in account.
                      </p>
                    </div>
                  </div>
                </div>
              </div>
            </section>
          )}

          {active === 'overview' && !isAdmin && (
            <section>
              <div className="welcome-row">
                <div>
                  <span className="eyebrow">STAFF PORTAL</span>
                  <h2>
                    Welcome,{' '}
                    {boot?.profile?.fullName?.split(' ')[0] ||
                      user.name ||
                      'Staff member'}
                  </h2>
                  <p>Your personal attendance overview for the IT Unit.</p>
                </div>
                <button
                  className="primary-btn"
                  onClick={() => setActive('scan')}
                >
                  <ScanLine size={17} /> {ownToday?.checkIn && !ownToday?.checkOut ? 'Check out' : 'Scan attendance QR'}
                </button>
              </div>
              <div className="staff-status-card">
                <div className="staff-status-copy">
                  <span className="eyebrow">TODAY'S ATTENDANCE</span>
                  <h2>
                    {ownToday?.checkIn && ownToday?.checkOut
                      ? 'Attendance completed'
                      : ownToday?.checkIn
                        ? 'You are checked in'
                        : 'You have not checked in yet'}
                  </h2>
                  <p>
                    {ownToday?.checkIn
                      ? 'Check-in: ' + timeLabel(ownToday.checkIn) +
                        (ownToday.checkOut ? ' · Check-out: ' + timeLabel(ownToday.checkOut) + (ownToday.totalHours != null ? ' · Total: ' + ownToday.totalHours.toFixed(2) + ' hours.' : '.') : ' · Remember to scan again when leaving to check out.')
                      : 'Scan the active QR code displayed at the FMC complex entrance when you arrive.'}
                  </p>
                </div>
                <div className="staff-status-mark">
                  {ownToday ? <CheckCircle2 size={34} /> : <QrCode size={34} />}
                </div>
              </div>
              <div className="stat-grid three">
                <div className="stat-card">
                  <div className="stat-top">
                    <span>Days recorded this month</span>
                    <span className="stat-icon blue">
                      <CalendarDays size={18} />
                    </span>
                  </div>
                  <strong>{monthRecords.length}</strong>
                  <small>Attendance entries</small>
                </div>
                <div className="stat-card">
                  <div className="stat-top">
                    <span>On-time arrivals</span>
                    <span className="stat-icon green">
                      <Check size={18} />
                    </span>                  </div>
                  <strong>
                    {monthRecords.filter(r => r.status === 'PRESENT').length}
                  </strong>
                  <small>This month</small>
                </div>
                <div className="stat-card">
                  <div className="stat-top">
                    <span>Late arrivals</span>
                    <span className="stat-icon amber">
                      <Clock3 size={18} />
                    </span>
                  </div>
                  <strong>
                    {monthRecords.filter(r => r.status === 'LATE').length}
                  </strong>
                  <small>This month</small>
                </div>
              </div>
              <div className="panel">
                <div className="panel-heading">
                  <div>
                    <h3>My attendance history</h3>
                    <p>Your recent check-in records</p>
                  </div>
                </div>
                <div className="table-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>DATE</th>
                        <th>CHECK-IN TIME</th>
                        <th>CHECK-OUT TIME</th>
                        <th>HOURS WORKED</th>
                        <th>STATUS</th>
                      </tr>
                    </thead>
                    <tbody>
                      {ownRecords.slice(0, 10).map(record => (
                        <tr key={record.id}>
                          <td>{dateLabel(record.date)}</td>
                          <td>{timeLabel(record.checkIn)}</td>
                          <td>{record.checkOut ? timeLabel(record.checkOut) : 'Pending'}</td>
                          <td>{record.totalHours != null ? record.totalHours.toFixed(2) + ' hrs' : '—'}</td>
                          <td>
                            <StatusPill status={record.status} />
                          </td>
                        </tr>
                      ))}
                      {!ownRecords.length && (
                        <tr>
                          <td colSpan={5}>
                            <div className="empty-state">
                              <CalendarDays size={25} />
                              <b>No attendance history yet</b>
                              <span>
                                Your recorded attendance will appear here.
                              </span>
                            </div>
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
          )}

          {active === 'staff' && isAdmin && (
            <section>
              <div className="section-toolbar">
                <div>
                  <h2>Staff directory</h2>
                  <p>Manage registered IT Unit staff accounts.</p>
                </div>
                <button className="primary-btn" onClick={() => openNewStaff()}>
                  <UserPlus size={17} /> Register staff
                </button>
              </div>
              <div className="panel">
                <div className="table-toolbar">
                  <div className="search-box">
                    <Search size={17} />
                    <input
                      value={search}
                      onChange={e => setSearch(e.target.value)}
                      placeholder="Search name, staff ID or email…"
                    />
                  </div>
                  <span className="record-count">
                    {filteredStaff.length} staff records
                  </span>
                </div>
                <div className="table-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>STAFF MEMBER</th>
                        <th>STAFF ID</th>
                        <th>DEPARTMENT</th>
                        <th>EMAIL</th>
                        <th>STATUS</th>
                        <th>ACTIONS</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredStaff.map(person => (
                        <tr key={person.id}>
                          <td>
                            <div className="person-cell">
                              <div className="avatar small">
                                {initials(person.fullName)}
                              </div>
                              <div>
                                <b>{person.fullName}</b>
                                <small>
                                  {person.designation || 'IT Unit Staff'}
                                </small>
                              </div>
                            </div>
                          </td>
                          <td>{person.staffId}</td>
                          <td>{person.department || '—'}</td>
                          <td>{person.email}</td>
                          <td>
                            <StatusPill status={person.status} />
                          </td>
                          <td>
                            <div className="row-actions">
                              <button onClick={() => openNewStaff(person)}>
                                Edit
                              </button>
                              {person.status === 'active' && (
                                <button
                                  className="danger-link"
                                  onClick={() => deactivateStaff(person)}
                                >
                                  Deactivate
                                </button>
                              )}
                            </div>
                          </td>
                        </tr>
                      ))}
                      {!filteredStaff.length && (
                        <tr>
                          <td colSpan={6}>
                            <div className="empty-state">
                              <Users size={25} />
                              <b>No staff records found</b>
                              <span>
                                Register staff individually. Bulk import can be
                                added in a later iteration.
                              </span>
                            </div>
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
          )}

          {active === 'attendance' && isAdmin && (
            <section>
              <div className="section-toolbar">
                <div>
                  <h2>Attendance records</h2>
                  <p>Review daily check-ins and staff attendance status.</p>
                </div>
                <button
                  className="outline-btn"
                  onClick={() =>
                    csvDownload(
                      filteredAttendance.map(r => ({
                        Date: r.date,
                        Name: r.staffName,
                        'Staff ID': r.staffId,
                        Email: r.email,
                        'Check-in': timeLabel(r.checkIn),
                        'Check-out': r.checkOut ? timeLabel(r.checkOut) : 'Pending',
                        'Hours worked': r.totalHours != null ? r.totalHours.toFixed(2) : '',
                        Status: r.status,
                      })),
                      'FMC-IT-Attendance.xlsx'
                    )
                  }
                >
                  <Download size={17} /> Export Excel
                </button>
              </div>
              <div className="panel">
                <div className="table-toolbar">
                  <div className="search-box">
                    <Search size={17} />
                    <input
                      value={search}
                      onChange={e => setSearch(e.target.value)}
                      placeholder="Search staff or status…"
                    />
                  </div>
                  <div className="toolbar-filters">
                    <input
                      type="date"
                      value={filterDate}
                      onChange={e => setFilterDate(e.target.value)}
                    />
                    <button
                      className="text-btn"
                      onClick={() => {
                        setFilterDate('');
                        setSearch('');
                      }}
                    >
                      Clear filters
                    </button>
                  </div>
                </div>
                <div className="table-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>DATE</th>
                        <th>STAFF MEMBER</th>
                        <th>STAFF ID</th>
                        <th>CHECK-IN</th>
                        <th>CHECK-OUT</th>
                        <th>HOURS</th>
                        <th>STATUS</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredAttendance
                        .slice()
                        .sort((a, b) => b.checkIn.localeCompare(a.checkIn))
                        .map(record => (
                          <tr key={record.id}>
                            <td>{dateLabel(record.date)}</td>
                            <td>
                              <div className="person-cell">
                                <div className="avatar small">
                                  {initials(record.staffName)}
                                </div>
                                <b>{record.staffName}</b>
                              </div>
                            </td>
                            <td>{record.staffId}</td>
                            <td>{timeLabel(record.checkIn)}</td>
                            <td>{record.checkOut ? timeLabel(record.checkOut) : 'Pending'}</td>
                            <td>{record.totalHours != null ? record.totalHours.toFixed(2) + ' hrs' : '—'}</td>
                            <td>
                              <StatusPill status={record.status} />
                            </td>
                          </tr>
                        ))}
                      {!filteredAttendance.length && (
                        <tr>                          <td colSpan={7}>
                            <div className="empty-state">
                              <ClipboardList size={25} />
                              <b>No matching attendance records</b>
                              <span>
                                Try changing the date or search terms.
                              </span>
                            </div>
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
          )}

          {active === 'qr' && isAdmin && (
            <section>
              <div className="welcome-row">
                <div>
                  <span className="eyebrow">ENTRANCE DISPLAY</span>
                  <h2>Daily attendance QR code</h2>
                  <p>
                    Display this screen on a tablet or computer at the FMC
                    complex entrance.
                  </p>
                </div>
                <button className="outline-btn" onClick={() => window.print()}>
                  <Monitor size={17} /> Print display
                </button>
              </div>
              <div className="qr-layout">
                <div className="qr-display-card">
                  <div className="qr-display-brand">
                    <div className="brand-mark">
                      <ShieldCheck size={23} />
                    </div>
                    <div>
                      <b>FEDERAL MEDICAL CENTRE</b>
                      <span>BIRNIN KUDU · INFORMATION TECHNOLOGY UNIT</span>
                    </div>
                  </div>
                  <div className="qr-main">
                    <span className="eyebrow">DAILY STAFF ATTENDANCE</span>
                    <h2>
                      Scan to record
                      <br />
                      your attendance
                    </h2>
                    <div className="qr-frame">
                      {qr?.token ? (
                        <QRCodeSVG
                          value={qr.token}
                          size={244}
                          level="H"
                          includeMargin
                        />
                      ) : (
                        <div className="qr-placeholder">
                          <RefreshCw size={26} />
                          <span>Generating secure QR…</span>
                        </div>
                      )}
                    </div>
                    <div className="qr-expiry">
                      <span className="pulse-dot" /> QR refreshes automatically{' '}
                      <b>{qrSeconds}s</b>
                    </div>
                    <p className="qr-instruction">
                      Open the FMC IT Unit Attendance app on your phone and scan
                      this code before entering the complex.
                    </p>
                  </div>
                  <div className="qr-display-foot">
                    <span>{dateLabel(boot?.today)}</span>
                    <span>SECURE ATTENDANCE SESSION</span>
                  </div>
                </div>
                <div className="qr-side">
                  <div className="panel">
                    <div className="panel-heading">
                      <div>
                        <h3>Display instructions</h3>
                        <p>For the entrance display device</p>
                      </div>
                    </div>
                    <ol className="instruction-list">
                      <li>
                        <span>1</span>
                        <div>
                          <b>Keep this page open</b>
                          <p>
                            Use a dedicated tablet, laptop or monitor at the
                            approved entrance.
                          </p>
                        </div>
                      </li>
                      <li>
                        <span>2</span>
                        <div>
                          <b>QR rotates automatically</b>
                          <p>
                            A short-lived QR token refreshes regularly to reduce
                            screenshot sharing.
                          </p>
                        </div>
                      </li>
                      <li>
                        <span>3</span>
                        <div>
                          <b>Staff scan individually</b>
                          <p>
                            Each person must sign in with their own registered
                            account before scanning.
                          </p>
                        </div>
                      </li>
                    </ol>
                  </div>
                  <div className="alert info">
                    <ShieldCheck size={18} />
                    <span>
                      Do not share this display link publicly. The QR token is
                      short-lived and validated by the server.
                    </span>
                  </div>
                </div>
              </div>
            </section>
          )}

          {active === 'scan' && !isAdmin && (
            <section>
              <div className="welcome-row">
                <div>
                  <span className="eyebrow">STAFF ATTENDANCE</span>
                  <h2>{boot?.attendance?.some(record => record.date === boot.today && record.checkIn && !record.checkOut) ? 'Scan to check out' : 'Scan to check in'}</h2>
                  <p>
                    Scan the current QR code displayed at the entrance using your phone camera. The system records check-in first and check-out on your next scan.
                  </p>
                </div>
              </div>
              <div className="scanner-layout">
                <div className="scanner-card">
                  <div className="scanner-head">
                    <span className="scanner-icon">
                      <Camera size={22} />
                    </span>
                    <div>
                      <b>{boot?.attendance?.some(record => record.date === boot.today && record.checkIn && !record.checkOut) ? 'QR check-out scanner' : 'QR check-in scanner'}</b>
                      <small>Camera access is required</small>
                    </div>
                    <span className="scanner-secure">
                      <ShieldCheck size={14} /> Secure
                    </span>
                  </div>
                  <div className="scanner-window">
                    <div id="fmc-qr-reader" />
                    <div className="scanner-overlay">
                      <ScanLine size={34} />
                    </div>
                  </div>
                  <div className="scanner-status">
                    {scannerMessage ? (
                      <>
                        <AlertCircle size={17} />
                        {scannerMessage}
                      </>
                    ) : (
                      <>
                        <span className="pulse-dot" /> Camera ready — point at
                        the entrance QR code.
                      </>
                    )}
                  </div>
                  <button
                    className="outline-btn full"
                    onClick={() => {
                      setScannerMessage('');
                      loadData().catch(() => undefined);
                    }}
                  >
                    {' '}
                    <RefreshCw size={16} /> Refresh scanner status
                  </button>
                </div>
                <div className="scanner-help panel">
                  <h3>Before you scan</h3>
                  <div className="help-item">
                    <CheckCircle2 size={18} />
                    <span>
                      Sign in using your own registered staff account.
                    </span>
                  </div>
                  <div className="help-item">
                    <CheckCircle2 size={18} />
                    <span>Allow camera access when your browser asks.</span>
                  </div>
                  <div className="help-item">
                    <CheckCircle2 size={18} />
                    <span>
                      Scan the current QR displayed at the FMC entrance.
                    </span>
                  </div>
                  <div className="help-item">
                    <CheckCircle2 size={18} />
                    <span>Wait for the attendance confirmation message.</span>
                  </div>
                  <div className="tip-box">
                    <CircleHelp size={18} />
                    <div>
                      <b>How check-out works</b>
                      <p>
                        Your first scan records check-in. Scan the same live entrance QR again when leaving to record check-out and calculate your hours worked. Each action can only be recorded once per day.
                      </p>
                    </div>
                  </div>
                </div>
              </div>
            </section>
          )}

          {active === 'reports' && isAdmin && (
            <section>
              <div className="section-toolbar">
                <div>
                  <h2>Reports & analytics</h2>
                  <p>Download attendance data for management review.</p>
                </div>
                <button
                  className="primary-btn"
                  onClick={() =>
                    csvDownload(
                      filteredAttendance.map(r => ({
                        Date: r.date,
                        Name: r.staffName,
                        'Staff ID': r.staffId,
                        Email: r.email,
                        'Check-in': timeLabel(r.checkIn),
                        'Check-out': r.checkOut ? timeLabel(r.checkOut) : 'Pending',
                        'Hours worked': r.totalHours != null ? r.totalHours.toFixed(2) : '',
                        Status: r.status,
                      })),
                      'FMC-IT-Unit-Attendance.xlsx'
                    )
                  }
                >
                  <FileSpreadsheet size={17} /> Download Excel                </button>
              </div>
              <div className="report-grid">
                <div className="report-card">
                  <div className="report-icon blue">
                    <CalendarDays size={21} />
                  </div>
                  <h3>Daily attendance</h3>
                  <p>Attendance records for the current day.</p>
                  <button
                    onClick={() =>
                      csvDownload(
                        todayAttendance.map(r => ({
                          Date: r.date,
                          Name: r.staffName,
                          'Staff ID': r.staffId,
                          'Check-in': timeLabel(r.checkIn),
                          'Check-out': r.checkOut ? timeLabel(r.checkOut) : 'Pending',
                          'Hours worked': r.totalHours != null ? r.totalHours.toFixed(2) : '',
                        'Check-out': r.checkOut ? timeLabel(r.checkOut) : 'Pending',
                        'Hours worked': r.totalHours != null ? r.totalHours.toFixed(2) : '',
                          Status: r.status,
                        })),
                        'FMC-IT-Daily-Attendance.xlsx'
                      )
                    }
                  >
                    Export daily report <ArrowDownToLine size={15} />
                  </button>
                </div>
                <div className="report-card">
                  <div className="report-icon green">
                    <BarChart3 size={21} />
                  </div>
                  <h3>Monthly attendance</h3>
                  <p>All available attendance records for the current month.</p>
                  <button
                    onClick={() =>
                      csvDownload(
                        attendance
                          .filter(
                            r =>
                              r.date?.slice(0, 7) === boot?.today?.slice(0, 7)
                          )
                          .map(r => ({
                            Date: r.date,
                            Name: r.staffName,
                            'Staff ID': r.staffId,
                            'Check-in': timeLabel(r.checkIn),
                          'Check-out': r.checkOut ? timeLabel(r.checkOut) : 'Pending',
                          'Hours worked': r.totalHours != null ? r.totalHours.toFixed(2) : '',
                        'Check-out': r.checkOut ? timeLabel(r.checkOut) : 'Pending',
                        'Hours worked': r.totalHours != null ? r.totalHours.toFixed(2) : '',
                            Status: r.status,
                          })),
                        'FMC-IT-Monthly-Attendance.xlsx'
                      )
                    }
                  >
                    Export monthly report <ArrowDownToLine size={15} />
                  </button>
                </div>
                <div className="report-card">
                  <div className="report-icon amber">
                    <Clock3 size={21} />
                  </div>
                  <h3>Late arrivals</h3>
                  <p>
                    Staff who checked in after the official resumption time.
                  </p>
                  <button
                    onClick={() =>
                      csvDownload(
                        attendance
                          .filter(r => r.status === 'LATE')
                          .map(r => ({
                            Date: r.date,
                            Name: r.staffName,
                            'Staff ID': r.staffId,
                            'Check-in': timeLabel(r.checkIn),
                          'Check-out': r.checkOut ? timeLabel(r.checkOut) : 'Pending',
                          'Hours worked': r.totalHours != null ? r.totalHours.toFixed(2) : '',
                        'Check-out': r.checkOut ? timeLabel(r.checkOut) : 'Pending',
                        'Hours worked': r.totalHours != null ? r.totalHours.toFixed(2) : '',
                            Status: r.status,
                          })),
                        'FMC-IT-Late-Arrivals.xlsx'
                      )
                    }
                  >
                    Export late report <ArrowDownToLine size={15} />
                  </button>
                </div>
              </div>
              <div className="panel report-summary">
                <div className="panel-heading">
                  <div>
                    <h3>Attendance snapshot</h3>
                    <p>Current records in the system</p>
                  </div>
                </div>
                <div className="summary-bars">
                  <div>
                    <span>Present</span>
                    <div className="bar">
                      <i
                        className="green-bar"
                        style={{
                          width:
                            (attendance.length
                              ? (attendance.filter(r => r.status === 'PRESENT')
                                  .length /
                                  attendance.length) *
                                100
                              : 0) + '%',
                        }}
                      />
                    </div>
                    <b>
                      {attendance.filter(r => r.status === 'PRESENT').length}
                    </b>
                  </div>
                  <div>
                    <span>Late</span>
                    <div className="bar">
                      <i
                        className="amber-bar"
                        style={{
                          width:
                            (attendance.length
                              ? (attendance.filter(r => r.status === 'LATE')
                                  .length /
                                  attendance.length) *
                                100
                              : 0) + '%',
                        }}
                      />
                    </div>
                    <b>{attendance.filter(r => r.status === 'LATE').length}</b>
                  </div>
                </div>
                <button className="outline-btn" onClick={() => window.print()}>
                  <Download size={16} /> Print / Save as PDF
                </button>
              </div>
            </section>
          )}

          {active === 'settings' && isAdmin && (
            <section>
              <div className="section-toolbar">
                <div>
                  <h2>Attendance settings</h2>
                  <p>Configure the official IT Unit attendance schedule.</p>
                </div>
              </div>
              <form className="panel settings-form" onSubmit={saveSettings}>
                <div className="panel-heading">
                  <div>
                    <h3>Working hours</h3>
                    <p>All times are interpreted in West Africa Time (WAT).</p>
                  </div>
                  <span className="settings-badge">
                    <ShieldCheck size={15} /> Admin only
                  </span>
                </div>
                <div className="settings-fields">
                  <label>
                    Check-in opens
                    <input
                      type="time"
                      value={settingsForm.checkInStart}
                      onChange={e =>
                        setSettingsForm({
                          ...settingsForm,
                          checkInStart: e.target.value,
                        })
                      }
                      required
                    />
                  </label>
                  <label>
                    Official resumption
                    <input
                      type="time"
                      value={settingsForm.resumption}
                      onChange={e =>
                        setSettingsForm({
                          ...settingsForm,
                          resumption: e.target.value,
                        })
                      }
                      required
                    />
                  </label>
                  <label>
                    Closing time
                    <input
                      type="time"
                      value={settingsForm.closing}
                      onChange={e =>
                        setSettingsForm({
                          ...settingsForm,
                          closing: e.target.value,
                        })
                      }
                      required
                    />
                  </label>
                  <label>
                    Grace period (minutes)
                    <input
                      type="number"
                      min="0"
                      max="120"
                      value={settingsForm.graceMinutes}
                      onChange={e =>
                        setSettingsForm({
                          ...settingsForm,
                          graceMinutes: Number(e.target.value),
                        })
                      }
                    />
                  </label>
                </div>
                <div className="settings-note">
                  <AlertCircle size={17} />
                  <span>
                    Staff checking in after the resumption time plus grace
                    period are marked LATE. Staff without a check-in are not
                    automatically marked absent until the reporting period
                    closes.
                  </span>
                </div>
                <button className="primary-btn" type="submit" disabled={busy}>
                  {busy ? 'Saving…' : 'Save attendance settings'}
                </button>
              </form>
              <div className="panel admin-info">
                <h3>System administrator</h3>
                <p>
                  The Super Admin account is restricted to this email address:
                </p>
                <div className="admin-email">
                  <ShieldCheck size={18} />
                  <b>{ADMIN_EMAIL}</b>
                </div>
              </div>
            </section>
          )}
        </div>
        <footer className="app-footer">
          <span>FMC BIRNIN KUDU · IT UNIT</span>
          <span>STAFF ATTENDANCE MANAGEMENT SYSTEM</span>
        </footer>
      </main>

      {showStaffForm && (
        <div
          className="modal-backdrop"
          onMouseDown={e => {
            if (e.target === e.currentTarget) setShowStaffForm(false);          }}
        >
          <form className="modal-card" onSubmit={saveStaff}>
            <div className="modal-head">
              <div>
                <span className="eyebrow">STAFF MANAGEMENT</span>
                <h2>{editing ? 'Edit staff record' : 'Register new staff'}</h2>
              </div>
              <button
                type="button"
                className="icon-btn"
                onClick={() => setShowStaffForm(false)}
              >
                <X size={19} />
              </button>
            </div>
            <div className="form-grid">
              <label>
                Full name
                <input
                  value={form.fullName}
                  onChange={e => setForm({ ...form, fullName: e.target.value })}
                  placeholder="Enter full name"
                  required
                />
              </label>
              <label>
                Staff ID
                <input
                  value={form.staffId}
                  onChange={e => setForm({ ...form, staffId: e.target.value })}
                  placeholder="e.g. FMC/IT/001"
                  required
                />
              </label>
              <label>
                Email address
                <input
                  type="email"
                  value={form.email}
                  onChange={e =>
                    setForm({ ...form, email: e.target.value.toLowerCase() })
                  }
                  placeholder="staff@fmc.gov.ng"
                  required
                />
              </label>
              <label>
                Phone number
                <input
                  value={form.phone}
                  onChange={e => setForm({ ...form, phone: e.target.value })}
                  placeholder="080…"
                />
              </label>
              <label>
                Department / Unit
                <input
                  value={form.department}
                  onChange={e =>
                    setForm({ ...form, department: e.target.value })
                  }
                  required
                />
              </label>
              <label>
                Designation
                <input
                  value={form.designation}
                  onChange={e =>
                    setForm({ ...form, designation: e.target.value })
                  }
                  placeholder="e.g. IT Officer"
                />
              </label>
            </div>
            <div className="modal-actions">
              <button
                type="button"
                className="outline-btn"
                onClick={() => setShowStaffForm(false)}
              >
                Cancel
              </button>
              <button className="primary-btn" type="submit" disabled={busy}>
                {busy ? 'Saving…' : editing ? 'Save changes' : 'Register staff'}
              </button>
            </div>
          </form>
        </div>
      )}
    </div>
  );
}
