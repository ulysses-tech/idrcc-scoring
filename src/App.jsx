/**
 * IDCC 2026 · DCC Rescue Flight — Judge & HQ Scoring System (v3)
 *
 * v3 추가 사항
 *  - [기능] 내장 초시계: 하단 고정 바에서 Start/Stop → 랩타임 자동 입력
 *           · 시각(타임스탬프) 기준 측정 → 화면이 꺼지거나 다른 앱으로 전환해도 정확
 *           · 새로고침·화면 이동 후에도 측정 계속 유지
 *           · 실수로 멈춘 경우 Undo로 측정 재개(멈춘 동안의 시간 포함)
 *           · 측정 중 화면 꺼짐 방지(지원 기기), 시작/정지 시 진동
 *           · 측정 중에는 제출 불가, PC에서는 스페이스바로 Start/Stop
 *  - [변경] 화면 표기 IDRCC → IDCC
 *
 * 주요 수정 사항
 *  - [버그] JSX 오류 속성(md:size) 제거 → 빌드 오류 원인
 *  - [버그] 순위를 전체 통합이 아닌 "카테고리별"로 산정, 동점(점수+시간) 공동 순위 처리
 *  - [버그] 동점 시간 비교에서 1/100초(ms) 누락 → 반영
 *  - [버그] HQ 수정 화면에 점수 항목이 없던 문제 → 전체 항목 수정 가능
 *  - [버그] 수정 저장 시 id 필드가 DB에 다시 기록되던 문제 제거
 *  - [버그] 정렬 시 필드 누락(undefined) 레코드에서 화면이 멈추던 문제 방지
 *  - [버그] 엑셀 파일명 날짜가 UTC 기준(한국 오전 9시 이전 전날 날짜)이던 문제 → 현지 날짜
 *  - [안정성] 서명 이미지를 scores와 분리(signatures/{id}) → 모든 기기가 서명 이미지까지
 *             매번 내려받던 부하 제거 (기존 레코드의 내장 서명도 그대로 읽힘)
 *  - [안정성] 화면 회전·크기 변경 시 서명이 지워지던 문제 → 서명 보존
 *  - [안정성] 온라인/오프라인 표시, 오프라인 제출 시 경고 후 자동 동기화
 *  - [검증] 동일 참가자·동일 차수 중복 제출 경고, 랩타임 0:00.00 제출 경고
 *  - [기능] 차수 전체/참가자별 최고 기록 보기, 카테고리 탭, 검색, 카테고리별 엑셀 시트
 *  - [UI] 밝은 톤, 단계별 섹션, 터치용 큰 버튼(+/-), 하단 고정 합계·제출 바
 */
import React, {
  useState,
  useEffect,
  useRef,
  useLayoutEffect,
  useMemo,
  useCallback,
  forwardRef,
  useImperativeHandle,
} from "react";
import { initializeApp, getApps } from "firebase/app";
import { getDatabase, ref, push, onValue, update, get } from "firebase/database";
import {
  LogOut,
  UserCheck,
  ShieldCheck,
  Download,
  Clock,
  Trash2,
  Pencil,
  X,
  Save,
  CheckCircle2,
  Plus,
  Minus,
  Wifi,
  WifiOff,
  AlertTriangle,
  Plane,
  Ban,
  Search,
  ChevronLeft,
  Medal,
  Code,
  PenLine,
  RotateCcw,
  Check,
  Lock,
  Users,
  Loader2,
  ArrowRight,
  Play,
  Square,
  Undo2,
  Timer,
} from "lucide-react";
import * as XLSX from "xlsx";
import SignatureCanvas from "react-signature-canvas";

// ─────────────────────────────────────────────────────────────
// Firebase
// ─────────────────────────────────────────────────────────────
const firebaseConfig = {
  apiKey: "AIzaSyCRYeoh01XG4-c4wopyH-5IUUoiVWvMd28",
  authDomain: "idrcc-d9cb5.firebaseapp.com",
  databaseURL:
    "https://idrcc-d9cb5-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "idrcc-d9cb5",
  storageBucket: "idrcc-d9cb5.firebasestorage.app",
  messagingSenderId: "320566419768",
  appId: "1:320566419768:web:887aa5fc5a634aaede21cd",
  measurementId: "G-19V571QJX6",
};

let db = null;
try {
  // [수정] 핫리로드 시 duplicate-app 오류 방지
  const app = getApps().length ? getApps()[0] : initializeApp(firebaseConfig);
  db = getDatabase(app);
} catch (e) {
  console.error("Firebase initialization failed:", e);
}

// ─────────────────────────────────────────────────────────────
// Competition constants
// ─────────────────────────────────────────────────────────────
const JUDGE_PASSWORD = "idrcc2026";
const HQ_PASSWORD = "admin2026";

const CATEGORIES = ["Elementary", "Middle School", "High School", "General"];
const CATEGORY_SHORT = {
  Elementary: "Elem.",
  "Middle School": "Middle",
  "High School": "High",
  General: "General",
};
const CATEGORY_STYLE = {
  Elementary: "bg-emerald-100 text-emerald-700 ring-emerald-200",
  "Middle School": "bg-sky-100 text-sky-700 ring-sky-200",
  "High School": "bg-violet-100 text-violet-700 ring-violet-200",
  General: "bg-amber-100 text-amber-800 ring-amber-200",
};
const TRIALS = ["1st", "2nd"];

const CODING_MAX = 20;
const MISSION_PTS = 10;
const MISSIONS = [
  { key: "takeoff", label: "Takeoff", short: "TKO" },
  { key: "obsA", label: "Obstacle A", short: "A" },
  { key: "obsB", label: "Obstacle B", short: "B" },
  { key: "obsC", label: "Obstacle C", short: "C" },
  { key: "resLnd", label: "Rescue Landing", short: "R-L" },
  { key: "resTko", label: "Rescue Takeoff", short: "R-T" },
];

const MAX_LANDINGS = 4;
const LANDING_ZONES = [
  {
    key: "lndBlueCount",
    scoreKey: "lndBlueScore",
    name: "Blue",
    pts: 10,
    pad: "bg-blue-500",
    soft: "bg-blue-50",
    ring: "ring-blue-200",
    text: "text-blue-600",
  },
  {
    key: "lndRedCount",
    scoreKey: "lndRedScore",
    name: "Red",
    pts: 20,
    pad: "bg-rose-500",
    soft: "bg-rose-50",
    ring: "ring-rose-200",
    text: "text-rose-600",
  },
  {
    key: "lndYellowCount",
    scoreKey: "lndYellowScore",
    name: "Yellow",
    pts: 30,
    pad: "bg-amber-400",
    soft: "bg-amber-50",
    ring: "ring-amber-200",
    text: "text-amber-600",
  },
];

const EMPTY_SCORES = {
  coding: 0,
  takeoff: 0,
  obsA: 0,
  obsB: 0,
  obsC: 0,
  resLnd: 0,
  resTko: 0,
  lndBlueCount: 0,
  lndRedCount: 0,
  lndYellowCount: 0,
  min: 0,
  sec: 0,
  ms: 0,
  isDQ: false,
};

// ─────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────
const n = (v) => Number(v) || 0;
const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));
const pad2 = (v) => String(n(v)).padStart(2, "0");
const selectAll = (e) => e.target.select();
const toInt = (raw) => parseInt(String(raw ?? "").replace(/\D/g, "") || "0", 10);

/** Obstacle A: Middle School 이상 / Obstacle C: High School·General */
function missionAllowed(key, category) {
  if (key === "obsA") return category !== "Elementary";
  if (key === "obsC") return category === "High School" || category === "General";
  return true;
}

function calcScores(d) {
  const obsA = missionAllowed("obsA", d.category) ? n(d.obsA) : 0;
  const obsC = missionAllowed("obsC", d.category) ? n(d.obsC) : 0;
  const lndBlueScore = n(d.lndBlueCount) * 10;
  const lndRedScore = n(d.lndRedCount) * 20;
  const lndYellowScore = n(d.lndYellowCount) * 30;
  const totalScore =
    n(d.coding) +
    n(d.takeoff) +
    obsA +
    n(d.obsB) +
    obsC +
    n(d.resLnd) +
    n(d.resTko) +
    lndBlueScore +
    lndRedScore +
    lndYellowScore;
  return { obsA, obsC, lndBlueScore, lndRedScore, lndYellowScore, totalScore };
}

/** 저장용 레코드 생성 (숫자형 정규화 + 점수 재계산) */
function buildRecord(form) {
  const out = {
    judgeName: String(form.judgeName || "").trim(),
    category: form.category,
    participantNumber: String(form.participantNumber || "").trim(),
    participantName: String(form.participantName || "").trim(),
    trial: form.trial,
  };
  Object.keys(EMPTY_SCORES).forEach((k) => {
    out[k] = k === "isDQ" ? !!form.isDQ : n(form[k]);
  });
  return { ...out, ...calcScores(out) };
}

/** [수정] 1/100초까지 포함한 시간(센티초) */
const timeToCs = (r) => n(r.min) * 6000 + n(r.sec) * 100 + n(r.ms);
const formatTime = (r) => `${n(r.min)}:${pad2(r.sec)}.${pad2(r.ms)}`;
const landingCount = (r) => n(r.lndBlueCount) + n(r.lndRedCount) + n(r.lndYellowCount);

function formatClock(ts) {
  if (!ts) return "";
  const d = new Date(ts);
  return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
}
function localStamp() {
  const d = new Date();
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}_${pad2(
    d.getHours()
  )}${pad2(d.getMinutes())}`;
}

/** 정렬 기준: DQ 최하위 → 점수 내림차순 → 시간 오름차순 → 먼저 제출 */
function compareRecords(a, b) {
  if (!!a.isDQ !== !!b.isDQ) return a.isDQ ? 1 : -1;
  const s = n(b.totalScore) - n(a.totalScore);
  if (s) return s;
  const t = timeToCs(a) - timeToCs(b);
  if (t) return t;
  return n(a.timestamp) - n(b.timestamp);
}

const participantKey = (r) =>
  `${r.category}__${String(r.participantNumber ?? "").trim().toLowerCase()}`;

/** 참가자(카테고리+번호)별 최고 기록 1건 */
function bestPerParticipant(records) {
  const m = new Map();
  records.forEach((r) => {
    const k = participantKey(r);
    const cur = m.get(k);
    if (!cur || compareRecords(r, cur) < 0) m.set(k, r);
  });
  return [...m.values()];
}

/** [수정] 카테고리별 순위 + 동점(점수·시간 동일) 공동 순위 */
function rankByCategory(records) {
  const out = [];
  const groups = [...CATEGORIES, "__other"];
  groups.forEach((cat) => {
    const list = records
      .filter((r) =>
        cat === "__other" ? !CATEGORIES.includes(r.category) : r.category === cat
      )
      .sort(compareRecords);
    let rank = 0;
    let prev = null;
    list.forEach((r, i) => {
      if (r.isDQ) {
        out.push({ ...r, rank: null });
        return;
      }
      if (
        !prev ||
        n(prev.totalScore) !== n(r.totalScore) ||
        timeToCs(prev) !== timeToCs(r)
      ) {
        rank = i + 1;
      }
      prev = r;
      out.push({ ...r, rank });
    });
  });
  return out;
}

const storage = {
  get(k, session = false) {
    try {
      return (session ? window.sessionStorage : window.localStorage).getItem(k);
    } catch {
      return null;
    }
  },
  set(k, v, session = false) {
    try {
      (session ? window.sessionStorage : window.localStorage).setItem(k, v);
    } catch {
      /* storage unavailable */
    }
  },
  remove(k, session = false) {
    try {
      (session ? window.sessionStorage : window.localStorage).removeItem(k);
    } catch {
      /* storage unavailable */
    }
  },
};

// ─────────────────────────────────────────────────────────────
// Hooks
// ─────────────────────────────────────────────────────────────
function useConnection() {
  const [connected, setConnected] = useState(null); // null = connecting
  useEffect(() => {
    if (!db) return undefined;
    return onValue(ref(db, ".info/connected"), (s) => setConnected(s.val() === true));
  }, []);
  return connected;
}

function useToast() {
  const [toast, setToast] = useState(null);
  const timer = useRef(null);
  const show = useCallback((msg, type = "success") => {
    clearTimeout(timer.current);
    setToast({ msg, type });
    timer.current = setTimeout(() => setToast(null), 2800);
  }, []);
  useEffect(() => () => clearTimeout(timer.current), []);
  return [toast, show];
}

/** 점수 폼 상태 — 카테고리 변경 시 해당 없는 장애물 자동 0점 */
function useScoreForm(initial) {
  const [form, setForm] = useState(initial);
  const patch = useCallback((p) => {
    setForm((prev) => {
      const next = { ...prev, ...p };
      MISSIONS.forEach((m) => {
        if (!missionAllowed(m.key, next.category)) next[m.key] = 0;
      });
      return next;
    });
  }, []);
  return [form, patch, setForm];
}

// ─────────────────────────────────────────────────────────────
// Stopwatch
// ─────────────────────────────────────────────────────────────
const SW_KEY = "idrcc_stopwatch";
const MAX_CS = 59 * 6000 + 59 * 100 + 99; // 59:59.99

/** 밀리초 → {min, sec, ms(1/100초)} (1/100초 미만은 버림) */
function msToTimeFields(ms) {
  const cs = Math.min(Math.floor(Math.max(0, ms) / 10), MAX_CS);
  return { min: Math.floor(cs / 6000), sec: Math.floor((cs % 6000) / 100), ms: cs % 100 };
}
const formatMs = (ms) => formatTime(msToTimeFields(ms));

function buzz(pattern) {
  try {
    if (navigator.vibrate) navigator.vibrate(pattern);
  } catch {
    /* vibration unsupported */
  }
}

/**
 * 초시계 — 시작 시각을 기준으로 계산하므로 화면 꺼짐·탭 전환·새로고침에도 정확.
 * 상태(idle | running | stopped)는 기기에 저장되어 새로고침 후에도 이어짐.
 */
function useStopwatch(onStop) {
  const [sw, setSw] = useState(() => {
    try {
      const s = JSON.parse(storage.get(SW_KEY) || "null");
      if (s && (s.status === "running" || s.status === "stopped") && n(s.startAt) > 0) return s;
    } catch {
      /* ignore corrupt state */
    }
    return { status: "idle", startAt: 0, elapsed: 0 };
  });
  const [now, setNow] = useState(() => Date.now());
  const onStopRef = useRef(onStop);
  onStopRef.current = onStop;

  // 상태 저장
  useEffect(() => {
    if (sw.status === "idle") storage.remove(SW_KEY);
    else storage.set(SW_KEY, JSON.stringify(sw));
  }, [sw]);

  // 새로고침 직후 '정지' 상태였다면 기록 시간을 폼에 다시 채움
  useEffect(() => {
    if (sw.status === "stopped") onStopRef.current?.(sw.elapsed);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // 측정 중: 화면 갱신 + 화면 꺼짐 방지
  useEffect(() => {
    if (sw.status !== "running") return undefined;
    const id = setInterval(() => setNow(Date.now()), 31);
    let lock = null;
    let done = false;
    const acquire = async () => {
      try {
        if (!("wakeLock" in navigator) || document.visibilityState !== "visible") return;
        const l = await navigator.wakeLock.request("screen");
        if (done) l.release();
        else lock = l;
      } catch {
        /* wake lock unsupported or denied */
      }
    };
    const onVis = () => {
      setNow(Date.now());
      if (document.visibilityState === "visible") acquire();
    };
    acquire();
    document.addEventListener("visibilitychange", onVis);
    return () => {
      done = true;
      clearInterval(id);
      document.removeEventListener("visibilitychange", onVis);
      try {
        if (lock) lock.release();
      } catch {
        /* already released */
      }
    };
  }, [sw.status]);

  const elapsed = sw.status === "running" ? Math.max(0, now - sw.startAt) : sw.elapsed;

  const start = () => {
    const t = Date.now();
    setNow(t);
    setSw({ status: "running", startAt: t, elapsed: 0 });
    buzz(40);
  };
  const stop = () => {
    if (sw.status !== "running") return;
    const e = Math.max(0, Date.now() - sw.startAt);
    setSw({ status: "stopped", startAt: sw.startAt, elapsed: e });
    onStopRef.current?.(e);
    buzz([40, 60, 40]);
  };
  /** 실수로 멈춘 경우: 처음 시작 시각 기준으로 측정을 이어감 */
  const undo = () => {
    if (sw.status !== "stopped") return;
    setNow(Date.now());
    setSw({ ...sw, status: "running" });
    buzz(40);
  };
  const reset = () => { setSw({ status: "idle", startAt: 0, elapsed: 0 }); onStopRef.current?.(0); };
  const toggle = () => {
    if (sw.status === "idle") start();
    else if (sw.status === "running") stop();
  };

  return { status: sw.status, elapsed, start, stop, undo, reset, toggle };
}

// ─────────────────────────────────────────────────────────────
// App
// ─────────────────────────────────────────────────────────────
export default function App() {
  const [isLoggedIn, setIsLoggedIn] = useState(
    () => storage.get("idrcc_auth", true) === "1"
  );
  const [view, setView] = useState("landing");
  const [records, setRecords] = useState([]);
  const [hqAuthenticated, setHqAuthenticated] = useState(false);
  const connected = useConnection();

  useEffect(() => {
    let meta = document.querySelector("meta[name=viewport]");
    if (!meta) {
      meta = document.createElement("meta");
      meta.name = "viewport";
      document.head.appendChild(meta);
    }
    meta.content =
      "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover";
    document.title = "IDCC 2026 · DCC Scoring";
  }, []);

  // [수정] 로그인 후에만 구독, 언마운트 시 구독 해제, 서명 이미지는 목록에서 제외
  useEffect(() => {
    if (!db || !isLoggedIn) return undefined;
    return onValue(
      ref(db, "scores"),
      (snap) => {
        const data = snap.val() || {};
        setRecords(
          Object.entries(data).map(([id, v]) => {
            const { signature, ...rest } = v || {};
            return { ...rest, id, hasLegacySignature: !!signature };
          })
        );
      },
      (err) => {
        console.error(err);
        alert("Could not load records from the database.\n" + err.message);
      }
    );
  }, [isLoggedIn]);

  const handleLogin = (pw) => {
    if (pw === JUDGE_PASSWORD) {
      storage.set("idrcc_auth", "1", true);
      setIsLoggedIn(true);
      return true;
    }
    return false;
  };

  const handleLogout = () => {
    storage.remove("idrcc_auth", true);
    setIsLoggedIn(false);
    setView("landing");
    setHqAuthenticated(false);
  };

  const goHQ = () => setView(hqAuthenticated ? "hq" : "hq_auth");

  if (!isLoggedIn) return <LoginScreen onLogin={handleLogin} />;

  return (
    <div className="min-h-screen bg-gradient-to-b from-sky-50 via-white to-amber-50/40 text-slate-800 antialiased selection:bg-sky-200">
      <AppHeader
        connected={connected}
        onLogout={handleLogout}
        onHome={() => setView("landing")}
      />

      {!db && (
        <div className="max-w-4xl mx-auto mt-4 px-4">
          <div className="flex items-center gap-2 bg-rose-50 text-rose-700 border border-rose-200 rounded-2xl p-4 font-bold">
            <AlertTriangle size={18} /> Database is not initialized. Check the Firebase configuration.
          </div>
        </div>
      )}

      {view === "landing" && (
        <Landing onJudge={() => setView("judge")} onHQ={goHQ} />
      )}
      {view === "judge" && (
        <JudgeView records={records} connected={connected} onGoToHQ={goHQ} />
      )}
      {view === "hq_auth" && (
        <HQAuth
          onAuth={() => {
            setHqAuthenticated(true);
            setView("hq");
          }}
          onBack={() => setView("landing")}
        />
      )}
      {view === "hq" && hqAuthenticated && (
        <HQView
          records={records}
          onGoToJudge={() => setView("judge")}
          onLock={() => {
            setHqAuthenticated(false);
            setView("landing");
          }}
        />
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// Layout
// ─────────────────────────────────────────────────────────────
function AppHeader({ connected, onLogout, onHome }) {
  return (
    <header className="sticky top-0 z-50 bg-white/90 backdrop-blur border-b border-sky-100 shadow-sm">
      <div className="h-1 bg-gradient-to-r from-sky-400 via-amber-300 to-rose-400" />
      <div className="max-w-7xl mx-auto px-3 md:px-6 py-2.5 flex items-center gap-3">
        <button onClick={onHome} className="flex items-center gap-3 min-w-0 text-left">
          <div className="w-10 h-10 rounded-2xl bg-gradient-to-br from-sky-400 to-blue-600 text-white grid place-items-center shadow-md shadow-sky-200 shrink-0">
            <Plane size={20} />
          </div>
          <div className="min-w-0">
            <div className="font-black text-lg md:text-xl leading-tight tracking-tight">
              IDCC 2026
            </div>
            <div className="text-xs md:text-sm font-semibold text-sky-600 truncate">
              DCC · Rescue Flight Scoring
            </div>
          </div>
        </button>
        <div className="ml-auto flex items-center gap-2">
          <ConnectionPill connected={connected} />
          <button
            onClick={onLogout}
            className="flex items-center gap-1.5 h-9 px-3 rounded-full border border-slate-200 bg-white text-slate-600 hover:text-rose-600 hover:border-rose-200 hover:bg-rose-50 transition text-sm font-bold"
          >
            <LogOut size={16} />
            <span className="hidden sm:inline">Logout</span>
          </button>
        </div>
      </div>
    </header>
  );
}

function ConnectionPill({ connected }) {
  if (connected === null)
    return (
      <span className="flex items-center gap-1.5 h-9 px-3 rounded-full bg-slate-100 text-slate-500 text-xs md:text-sm font-bold">
        <Loader2 size={14} className="animate-spin" /> Connecting
      </span>
    );
  return connected ? (
    <span className="flex items-center gap-1.5 h-9 px-3 rounded-full bg-emerald-50 text-emerald-700 ring-1 ring-emerald-200 text-xs md:text-sm font-bold">
      <Wifi size={14} /> Online
    </span>
  ) : (
    <span className="flex items-center gap-1.5 h-9 px-3 rounded-full bg-rose-50 text-rose-700 ring-1 ring-rose-200 text-xs md:text-sm font-bold animate-pulse">
      <WifiOff size={14} /> Offline
    </span>
  );
}

function Toast({ toast }) {
  if (!toast) return null;
  const tone =
    toast.type === "error"
      ? "bg-rose-600"
      : toast.type === "warn"
      ? "bg-amber-500"
      : "bg-emerald-600";
  return (
    <div className="fixed top-20 left-1/2 -translate-x-1/2 z-[200] px-4 w-full max-w-md pointer-events-none">
      <div
        className={`${tone} text-white rounded-2xl shadow-xl px-5 py-4 font-bold flex items-center gap-2`}
      >
        {toast.type === "success" ? <CheckCircle2 size={20} /> : <AlertTriangle size={20} />}
        {toast.msg}
      </div>
    </div>
  );
}

function Landing({ onJudge, onHQ }) {
  return (
    <main className="max-w-5xl mx-auto px-4 py-10 md:py-16">
      <div className="text-center mb-10">
        <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-sky-100 text-sky-700 text-xs font-black uppercase tracking-widest mb-4">
          <Plane size={14} /> Drone Coding Competition
        </div>
        <h2 className="text-3xl md:text-4xl font-black tracking-tight">Select your role</h2>
        <p className="mt-2 text-slate-500 font-medium">
          Choose how you will use the system today.
        </p>
      </div>
      <div className="grid md:grid-cols-2 gap-6">
        <RoleCard
          tone="sky"
          icon={<UserCheck size={40} />}
          title="Judge"
          desc="Score a flight, record the lap time and collect the participant's signature."
          cta="Start scoring"
          onClick={onJudge}
        />
        <RoleCard
          tone="violet"
          icon={<ShieldCheck size={40} />}
          title="HQ Admin"
          desc="Live rankings by category, record correction and Excel export."
          cta="Open dashboard"
          locked
          onClick={onHQ}
        />
      </div>
    </main>
  );
}

const ROLE_TONES = {
  sky: {
    card: "from-sky-50 to-white border-sky-100 hover:border-sky-300 hover:shadow-sky-100",
    icon: "bg-gradient-to-br from-sky-400 to-blue-500 shadow-sky-200",
    cta: "text-sky-600",
  },
  violet: {
    card: "from-violet-50 to-white border-violet-100 hover:border-violet-300 hover:shadow-violet-100",
    icon: "bg-gradient-to-br from-violet-400 to-indigo-500 shadow-violet-200",
    cta: "text-violet-600",
  },
};

function RoleCard({ tone, icon, title, desc, cta, onClick, locked }) {
  const t = ROLE_TONES[tone];
  return (
    <button
      onClick={onClick}
      className={`group text-left p-8 md:p-10 rounded-[2rem] border-2 bg-gradient-to-br ${t.card} shadow-lg hover:shadow-2xl hover:-translate-y-1 transition-all duration-300`}
    >
      <div className="flex items-start justify-between">
        <div className={`w-20 h-20 rounded-3xl ${t.icon} text-white grid place-items-center shadow-lg`}>
          {icon}
        </div>
        {locked && (
          <span className="flex items-center gap-1 text-xs font-bold text-slate-400 bg-white rounded-full px-3 py-1 border border-slate-200">
            <Lock size={12} /> Password
          </span>
        )}
      </div>
      <h3 className="mt-6 text-3xl font-black">{title}</h3>
      <p className="mt-2 text-slate-500 font-medium leading-relaxed">{desc}</p>
      <div className={`mt-6 inline-flex items-center gap-2 font-black ${t.cta}`}>
        {cta}
        <ArrowRight size={18} className="group-hover:translate-x-1 transition-transform" />
      </div>
    </button>
  );
}

// ─────────────────────────────────────────────────────────────
// Judge
// ─────────────────────────────────────────────────────────────
function JudgeView({ records, connected, onGoToHQ }) {
  const sigRef = useRef(null);
  const [toast, showToast] = useToast();
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [form, patch, setForm] = useScoreForm(() => ({
    judgeName: storage.get("idrcc_judge_name") || "",
    category: storage.get("idrcc_category") || "Elementary",
    participantNumber: "",
    participantName: "",
    trial: "1st",
    ...EMPTY_SCORES,
  }));

  // 심판 이름·카테고리는 기기에 기억 (새로고침해도 유지)
  useEffect(() => storage.set("idrcc_judge_name", form.judgeName), [form.judgeName]);
  useEffect(() => storage.set("idrcc_category", form.category), [form.category]);

  // 경기 시작 준비 알림 (드론 추격전과 동일):
  // 2번 Coding 점수를 누르면 바로 나타나 깜빡이고, 심판이 '확인'을 눌러야 멈춤.
  // 확인 후에는 하단 START 버튼이 깜빡여 초시계 시작을 유도. 제출하면 처음 상태로 돌아감.
  const [codingDone, setCodingDone] = useState(false);
  const [readyAck, setReadyAck] = useState(false);

  // 초시계: 정지하면 랩타임 칸에 자동 입력
  const sw = useStopwatch((ms) => patch(msToTimeFields(ms)));
  const swRef = useRef(sw);
  swRef.current = sw;

  // PC: 스페이스바로 Start/Stop (입력 칸·버튼·서명 패드 사용 중에는 동작 안 함)
  useEffect(() => {
    const onKey = (e) => {
      if (e.code !== "Space" || e.repeat) return;
      const t = e.target;
      const tag = String(t?.tagName || "").toLowerCase();
      if (["input", "textarea", "select", "button", "canvas"].includes(tag) || t?.isContentEditable)
        return;
      e.preventDefault();
      swRef.current.toggle();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  const calc = calcScores(form);

  const handleSubmit = async () => {
    if (!db) return alert("Database is not connected. Check the Firebase configuration.");
    if (sw.status === "running") {
      alert("The timer is still running.\nTap STOP before submitting.");
      return;
    }
    const rec = buildRecord(form);

    if (!rec.judgeName || !rec.participantNumber || !rec.participantName) {
      alert("Please fill in Judge Name, Participant No. and Participant Name.");
      window.scrollTo({ top: 0, behavior: "smooth" });
      return;
    }
    if (!sigRef.current || sigRef.current.isEmpty()) {
      alert("Participant signature is required.");
      return;
    }
    if (
      !rec.isDQ &&
      timeToCs(rec) === 0 &&
      !window.confirm("Lap time is 0:00.00.\nSubmit without a lap time?")
    )
      return;

    // [추가] 동일 카테고리·번호·차수 중복 경고
    const dup = records.find(
      (r) => participantKey(r) === participantKey(rec) && r.trial === rec.trial
    );
    if (
      dup &&
      !window.confirm(
        `A ${rec.trial} trial record already exists for No. ${rec.participantNumber} (${dup.participantName}, ${rec.category}).\n\nSubmit another record anyway?`
      )
    )
      return;

    const offline = connected === false;
    if (
      offline &&
      !window.confirm(
        "You are OFFLINE.\n\nThe record will be kept on this device and uploaded automatically when the connection returns.\nDo NOT close or refresh this page until the status shows Online.\n\nContinue?"
      )
    )
      return;

    setIsSubmitting(true);
    try {
      const key = push(ref(db, "scores")).key;
      const signature = sigRef.current.toDataURL();
      // [수정] 점수와 서명을 분리 저장 (원자적 다중 경로 쓰기)
      const writing = update(ref(db), {
        [`scores/${key}`]: { ...rec, hasSignature: true, timestamp: Date.now() },
        [`signatures/${key}`]: signature,
      });
      if (offline) {
        writing.catch((e) => console.error("Queued write failed:", e));
        showToast("Saved offline — will sync when online", "warn");
      } else {
        await writing;
        showToast(`Submitted · No. ${rec.participantNumber} · ${rec.totalScore} pts`);
      }
      sigRef.current.clear();
      sw.reset();
      setCodingDone(false);
      setReadyAck(false);
      setForm((prev) => ({
        ...prev,
        ...EMPTY_SCORES,
        participantNumber: "",
        participantName: "",
      }));
      window.scrollTo({ top: 0, behavior: "smooth" });
    } catch (e) {
      console.error(e);
      alert("Submission failed. Check the internet connection and try again.\n\n" + (e?.message || e));
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div className="max-w-4xl mx-auto px-3 md:px-6 pt-5 md:pt-8 pb-56 md:pb-40">
      <Toast toast={toast} />

      <div className="flex items-end justify-between gap-3 mb-5">
        <div>
          <div className="text-xs font-black uppercase tracking-widest text-sky-600">
            Judge Console
          </div>
          <h2 className="text-2xl md:text-3xl font-black tracking-tight">DCC Official Scoring</h2>
        </div>
        <button
          onClick={onGoToHQ}
          className="flex items-center gap-2 h-11 px-4 rounded-2xl bg-white border-2 border-violet-200 text-violet-600 font-bold hover:bg-violet-50 transition shrink-0"
        >
          <ShieldCheck size={18} /> <span className="hidden sm:inline">Go to</span> HQ
        </button>
      </div>

      <div className="space-y-5">
        <ParticipantFields form={form} patch={patch} />
        <ScoreFields
          form={form}
          patch={patch}
          stopwatch={sw}
          onCodingChange={() => setCodingDone(true)}
          afterCoding={
            codingDone && (
              <ReadyAlert acknowledged={readyAck} onAcknowledge={() => setReadyAck(true)} />
            )
          }
        />
        <Section step={6} title="Participant Signature" icon={<PenLine size={18} />} tone="slate"
          subtitle="Ask the participant to confirm the score and sign below.">
          <SignatureBox ref={sigRef} height={200} />
        </Section>
        <RecentSubmissions records={records} judgeName={form.judgeName} />
      </div>

      <SubmitBar
        total={calc.totalScore}
        form={form}
        isSubmitting={isSubmitting}
        onSubmit={handleSubmit}
        label="Finalize & Submit"
        stopwatch={sw}
        highlightStart={readyAck && sw.status === "idle"}
      />
    </div>
  );
}

function ParticipantFields({ form, patch, showJudge = true }) {
  return (
    <Section step={1} title="Participant" icon={<Users size={18} />} tone="sky">
      <div className="grid md:grid-cols-2 gap-4">
        {showJudge && (
          <TextField
            label="Judge Name"
            value={form.judgeName}
            onChange={(v) => patch({ judgeName: v })}
            placeholder="Your name"
          />
        )}
        <TextField
          label="Participant No."
          value={form.participantNumber}
          onChange={(v) => patch({ participantNumber: v })}
          placeholder="e.g. 101"
        />
        <TextField
          label="Participant Name"
          value={form.participantName}
          onChange={(v) => patch({ participantName: v })}
          placeholder="Full name"
          className={showJudge ? "md:col-span-2" : ""}
        />
      </div>
      <div className="grid md:grid-cols-[1fr_auto] gap-4 mt-4">
        <div>
          <FieldLabel>Category</FieldLabel>
          <Segmented
            options={CATEGORIES}
            value={form.category}
            onChange={(v) => patch({ category: v })}
            cols="grid-cols-2 sm:grid-cols-4"
          />
        </div>
        <div>
          <FieldLabel>Trial</FieldLabel>
          <Segmented
            options={TRIALS}
            value={form.trial}
            onChange={(v) => patch({ trial: v })}
            cols="grid-cols-2 md:w-48"
          />
        </div>
      </div>
    </Section>
  );
}

/**
 * 경기 시작 준비 알림 — 2번 Coding 점수를 누르면 나타나 깜빡이고,
 * 심판이 '확인'을 눌러야 깜빡임이 멈춤 (초시계 없이 경기를 시작하는 것을 방지)
 */
function ReadyAlert({ acknowledged, onAcknowledge }) {
  const boxRef = useRef(null);
  // 나타나는 즉시 화면 안으로 스크롤 (하단 고정 바에 가리지 않도록 여백 확보)
  useEffect(() => {
    try {
      boxRef.current?.scrollIntoView({ behavior: "smooth", block: "nearest" });
    } catch {
      /* scroll unsupported */
    }
  }, []);
  return (
    <div
      ref={boxRef}
      role="alert"
      style={{ scrollMarginTop: "6rem", scrollMarginBottom: "13rem" }}
      className={`flex flex-wrap items-center justify-between gap-3 rounded-2xl border-2 px-5 sm:px-6 py-4 transition-colors ${
        acknowledged ? "bg-emerald-50 border-emerald-300" : "bg-red-50 border-red-400"
      }`}
    >
      <p className={`text-lg font-black leading-snug ${acknowledged ? "text-emerald-800" : "text-red-600 animate-blink"}`}>
        ※ 학생의 경기 시작 준비가 완료되었습니다. 심판은 초시계를 준비하세요!
      </p>
      <button
        type="button"
        onClick={onAcknowledge}
        disabled={acknowledged}
        aria-pressed={acknowledged}
        className={`shrink-0 h-12 px-4 rounded-xl border-2 font-black flex items-center gap-2 transition active:scale-95 ${
          acknowledged
            ? "bg-emerald-600 border-emerald-600 text-white cursor-default"
            : "bg-white border-red-500 text-red-600 hover:bg-red-100"
        }`}
      >
        {acknowledged ? <CheckCircle2 size={20} /> : <span className="w-5 h-5 rounded-md border-2 border-red-500" />}
        {acknowledged ? "확인 완료" : "확인"}
      </button>
    </div>
  );
}

function ScoreFields({ form, patch, stopwatch, onCodingChange, afterCoding }) {
  const calc = calcScores(form);
  const used = landingCount(form);
  const missionPts = MISSIONS.reduce(
    (s, m) => s + (missionAllowed(m.key, form.category) ? n(form[m.key]) : 0),
    0
  );
  const landingPts = calc.lndBlueScore + calc.lndRedScore + calc.lndYellowScore;

  return (
    <>
      <Section step={2} title="Coding" icon={<Code size={18} />} tone="indigo"
        right={<ScoreChip tone="indigo">{n(form.coding)} / {CODING_MAX}</ScoreChip>}>
        <CodingInput
          value={form.coding}
          onChange={(v) => {
            patch({ coding: v });
            onCodingChange?.();
          }}
        />
      </Section>

      {afterCoding}

      <Section step={3} title="Flight Missions" icon={<Plane size={18} />} tone="emerald"
        subtitle={`Tap each checkpoint the drone cleared (+${MISSION_PTS} each).`}
        right={<ScoreChip tone="emerald">{missionPts} pts</ScoreChip>}>
        <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
          {MISSIONS.map((m, i) => (
            <MissionToggle
              key={m.key}
              index={i + 1}
              label={m.label}
              on={n(form[m.key]) === MISSION_PTS}
              disabled={!missionAllowed(m.key, form.category)}
              disabledNote={`Not used in ${form.category}`}
              onToggle={() =>
                patch({ [m.key]: n(form[m.key]) === MISSION_PTS ? 0 : MISSION_PTS })
              }
            />
          ))}
        </div>
      </Section>

      <Section step={4} title="Landing Zones" icon={<Medal size={18} />} tone="amber"
        subtitle={`Up to ${MAX_LANDINGS} landings in total.`}
        right={<ScoreChip tone="amber">{landingPts} pts</ScoreChip>}>
        <LandingMeter form={form} used={used} />
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mt-4">
          {LANDING_ZONES.map((z) => (
            <LandingCounter
              key={z.key}
              zone={z}
              count={n(form[z.key])}
              canAdd={used < MAX_LANDINGS}
              onChange={(c) => patch({ [z.key]: clamp(c, 0, MAX_LANDINGS) })}
            />
          ))}
        </div>
      </Section>

      <Section step={5} title="Lap Time & Result" icon={<Clock size={18} />} tone="rose"
        subtitle={
          stopwatch
            ? "Use the timer in the bottom bar. The time is filled in automatically when you tap STOP — adjust here only if needed."
            : undefined
        }
        right={stopwatch ? <TimerStatusChip status={stopwatch.status} /> : undefined}>
        <div className="flex flex-col md:flex-row md:items-end gap-4">
          <TimeInput form={form} patch={patch} />
          <DQToggle on={!!form.isDQ} onToggle={() => patch({ isDQ: !form.isDQ })} />
        </div>
      </Section>
    </>
  );
}

function TimerStatusChip({ status }) {
  const map = {
    idle: ["bg-slate-100 text-slate-500 ring-slate-200", "Timer ready"],
    running: ["bg-rose-50 text-rose-600 ring-rose-200 animate-pulse", "Timing…"],
    stopped: ["bg-emerald-50 text-emerald-700 ring-emerald-200", "Time recorded"],
  };
  const [cls, text] = map[status] || map.idle;
  return (
    <span className={`shrink-0 px-3 py-1.5 rounded-xl ring-1 text-sm font-black flex items-center gap-1.5 ${cls}`}>
      <Timer size={14} /> {text}
    </span>
  );
}

/** 하단 고정 바의 초시계 — 어느 위치로 스크롤해도 바로 누를 수 있음 */
function StopwatchControl({ sw, recorded, highlightStart = false }) {
  const running = sw.status === "running";
  const stopped = sw.status === "stopped";
  const box = running
    ? "bg-rose-50 ring-rose-200"
    : stopped
    ? "bg-emerald-50 ring-emerald-200"
    : "bg-slate-50 ring-slate-200";
  const txt = running ? "text-rose-600" : stopped ? "text-emerald-700" : "text-slate-700";
  const dot = running ? "bg-rose-500 animate-pulse" : stopped ? "bg-emerald-500" : "bg-slate-300";
  const btnBase =
    "h-14 md:h-16 rounded-2xl font-black flex items-center justify-center gap-2 transition active:scale-95";

  return (
    <div className="flex items-center gap-2 md:gap-3">
      <div
        className={`h-14 md:h-16 px-3 md:px-4 rounded-2xl ring-1 flex items-center gap-2.5 min-w-[9.5rem] md:min-w-[11rem] ${box}`}
        aria-live="off"
      >
        <span className={`w-2.5 h-2.5 rounded-full shrink-0 ${dot}`} />
        <div className="leading-none">
          <div className="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-1">
            {running ? "Timing" : stopped ? "Recorded" : "Lap time"}
          </div>
          <div className={`font-mono text-2xl md:text-3xl font-black tabular-nums ${txt}`}>
            {running ? formatMs(sw.elapsed) : recorded}
          </div>
        </div>
      </div>

      {running && (
        <button
          type="button"
          onClick={sw.stop}
          className={`${btnBase} flex-1 md:flex-none md:w-36 px-5 bg-rose-500 hover:bg-rose-600 text-white text-lg shadow-lg shadow-rose-200`}
        >
          <Square size={20} fill="currentColor" /> STOP
        </button>
      )}
      {sw.status === "idle" && (
        <button
          type="button"
          onClick={sw.start}
          className={`${btnBase} flex-1 md:flex-none md:w-36 px-5 bg-emerald-500 hover:bg-emerald-600 text-white text-lg shadow-lg shadow-emerald-200 ${
            highlightStart ? "animate-blink ring-4 ring-red-500 ring-offset-2" : ""
          }`}
        >
          <Play size={20} fill="currentColor" /> START
        </button>
      )}
      {stopped && (
        <>
          <button
            type="button"
            onClick={sw.undo}
            title="Stopped by mistake? Continue timing from the original start."
            className={`${btnBase} flex-1 md:flex-none px-3 bg-white border-2 border-slate-200 text-slate-600 text-sm hover:border-slate-300`}
          >
            <Undo2 size={18} /> Undo
          </button>
          <button
            type="button"
            onClick={() => {
              if (window.confirm("Reset the timer and lap time to 0:00.00?"))
                sw.reset();
            }}
            className={`${btnBase} flex-1 md:flex-none px-3 bg-white border-2 border-slate-200 text-slate-600 text-sm hover:border-slate-300`}
          >
            <RotateCcw size={18} /> Reset
          </button>
        </>
      )}
    </div>
  );
}

function SubmitBar({ total, form, isSubmitting, onSubmit, label, stopwatch, highlightStart = false }) {
  const timerRunning = stopwatch?.status === "running";
  const blocked = isSubmitting || timerRunning;
  return (
    <div
      className="fixed bottom-0 inset-x-0 z-40 bg-white/95 backdrop-blur border-t border-slate-200 shadow-[0_-10px_30px_-15px_rgba(15,23,42,0.3)]"
      style={{ paddingBottom: "env(safe-area-inset-bottom)" }}
    >
      <div className="max-w-4xl mx-auto px-3 md:px-6 py-2.5 md:py-3 flex flex-col md:flex-row md:items-center gap-2.5 md:gap-4">
        {stopwatch && (
          <StopwatchControl sw={stopwatch} recorded={formatTime(form)} highlightStart={highlightStart} />
        )}
        <div className="flex items-center gap-3 md:gap-4 flex-1 min-w-0">
        <div className="shrink-0">
          <div className="text-[11px] font-black uppercase tracking-widest text-slate-400">
            Total
          </div>
          <div className="text-4xl font-black text-blue-600 tabular-nums leading-none">
            {total}
            <span className="text-base text-slate-400 ml-1 font-bold">pts</span>
          </div>
        </div>
        {!stopwatch && (
          <div className="hidden sm:flex flex-col text-sm font-bold text-slate-500 gap-0.5">
            <span className="flex items-center gap-1.5">
              <Clock size={14} /> {formatTime(form)}
            </span>
            <span>Landings {landingCount(form)}/{MAX_LANDINGS}</span>
          </div>
        )}
        {form.isDQ && (
          <span className="px-2.5 py-1 rounded-lg bg-rose-100 text-rose-700 text-sm font-black">
            DQ
          </span>
        )}
        <button
          onClick={onSubmit}
          disabled={blocked}
          className={`ml-auto flex-1 md:flex-none md:min-w-[220px] h-14 md:h-16 px-3 rounded-2xl text-white text-base sm:text-lg md:text-xl font-black flex items-center justify-center gap-2 shadow-lg transition active:scale-[0.98] ${
            blocked
              ? "bg-slate-300 cursor-not-allowed shadow-none"
              : "bg-gradient-to-r from-blue-500 to-sky-500 hover:from-blue-600 hover:to-sky-600 shadow-sky-200"
          }`}
        >
          {isSubmitting ? (
            <>
              <Loader2 size={22} className="animate-spin" /> Submitting…
            </>
          ) : timerRunning ? (
            <>
              <Timer size={22} /> Stop timer first
            </>
          ) : (
            <>
              <CheckCircle2 size={24} /> {label}
            </>
          )}
        </button>
        </div>
      </div>
    </div>
  );
}

function RecentSubmissions({ records, judgeName }) {
  const name = String(judgeName || "").trim();
  const mine = useMemo(
    () =>
      name
        ? records
            .filter((r) => String(r.judgeName || "").trim() === name)
            .sort((a, b) => n(b.timestamp) - n(a.timestamp))
            .slice(0, 5)
        : [],
    [records, name]
  );
  if (!mine.length) return null;
  return (
    <section className="bg-white/70 rounded-3xl border border-slate-200 p-4 md:p-5">
      <h3 className="text-xs font-black uppercase tracking-widest text-slate-500 mb-2">
        Your recent submissions
      </h3>
      <ul className="divide-y divide-slate-100">
        {mine.map((r) => (
          <li key={r.id} className="py-2.5 flex items-center gap-2.5 text-sm">
            <CategoryBadge category={r.category} />
            <span className="font-mono font-bold">#{r.participantNumber}</span>
            <span className="font-semibold truncate">{r.participantName}</span>
            <span className="text-slate-400 shrink-0">{r.trial}</span>
            <span className="ml-auto font-black text-blue-600 tabular-nums shrink-0">
              {r.isDQ ? "DQ" : `${n(r.totalScore)} pts`}
            </span>
            <span className="text-slate-400 tabular-nums hidden sm:inline shrink-0">
              {formatClock(r.timestamp)}
            </span>
          </li>
        ))}
      </ul>
    </section>
  );
}

// ─────────────────────────────────────────────────────────────
// HQ
// ─────────────────────────────────────────────────────────────
function HQView({ records, onGoToJudge, onLock }) {
  const [editingId, setEditingId] = useState(null);
  const [catFilter, setCatFilter] = useState("All");
  const [mode, setMode] = useState("all"); // all | best
  const [sortBy, setSortBy] = useState("rank");
  const [query, setQuery] = useState("");
  const [toast, showToast] = useToast();

  const ranked = useMemo(
    () => rankByCategory(mode === "best" ? bestPerParticipant(records) : records),
    [records, mode]
  );

  const visible = useMemo(() => {
    let list = ranked;
    if (catFilter !== "All") list = list.filter((r) => r.category === catFilter);
    const q = query.trim().toLowerCase();
    if (q)
      list = list.filter((r) =>
        [r.participantName, r.participantNumber, r.judgeName].some((v) =>
          String(v ?? "").toLowerCase().includes(q)
        )
      );
    if (sortBy === "recent") list = [...list].sort((a, b) => n(b.timestamp) - n(a.timestamp));
    if (sortBy === "judge")
      list = [...list].sort(
        (a, b) =>
          String(a.judgeName || "").localeCompare(String(b.judgeName || "")) ||
          n(b.timestamp) - n(a.timestamp)
      );
    return list;
  }, [ranked, catFilter, query, sortBy]);

  const stats = useMemo(() => {
    const perCat = {};
    CATEGORIES.forEach((c) => (perCat[c] = 0));
    records.forEach((r) => {
      if (perCat[r.category] !== undefined) perCat[r.category] += 1;
    });
    return {
      total: records.length,
      participants: new Set(records.map(participantKey)).size,
      dq: records.filter((r) => r.isDQ).length,
      last: records.reduce((m, r) => Math.max(m, n(r.timestamp)), 0),
      perCat,
    };
  }, [records]);

  const editingRecord = records.find((r) => r.id === editingId);

  const toExcelRow = (r) => ({
    Rank: r.rank ?? "DQ",
    Category: r.category,
    "No.": r.participantNumber,
    Participant: r.participantName,
    Trial: r.trial,
    Coding: n(r.coding),
    Takeoff: n(r.takeoff),
    "Obstacle A": n(r.obsA),
    "Obstacle B": n(r.obsB),
    "Obstacle C": n(r.obsC),
    "Rescue Landing": n(r.resLnd),
    "Rescue Takeoff": n(r.resTko),
    "Blue (count)": n(r.lndBlueCount),
    "Red (count)": n(r.lndRedCount),
    "Yellow (count)": n(r.lndYellowCount),
    "Landing Pts": n(r.lndBlueScore) + n(r.lndRedScore) + n(r.lndYellowScore),
    Total: n(r.totalScore),
    Time: formatTime(r),
    DQ: r.isDQ ? "YES" : "NO",
    Judge: r.judgeName,
    Submitted: r.timestamp ? new Date(r.timestamp).toLocaleString() : "",
    Edited: r.editedAt ? new Date(r.editedAt).toLocaleString() : "",
  });

  const downloadExcel = () => {
    if (!ranked.length) return alert("There are no records to export.");
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ranked.map(toExcelRow)), "All");
    CATEGORIES.forEach((cat) => {
      const rows = ranked.filter((r) => r.category === cat);
      if (rows.length)
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.map(toExcelRow)), cat);
    });
    XLSX.writeFile(
      wb,
      `IDCC_DCC_Results_${mode === "best" ? "BestTrial" : "AllTrials"}_${localStamp()}.xlsx`
    );
  };

  const handleResetAll = async () => {
    const typed = window.prompt(
      "⚠️ This permanently deletes ALL score records and signatures.\nExport to Excel first.\n\nType RESET to confirm."
    );
    if (typed === null) return;
    if (typed.trim() !== "RESET") return alert("Reset cancelled.");
    try {
      await update(ref(db), { scores: null, signatures: null });
      showToast("All records deleted", "warn");
    } catch (e) {
      alert("Reset failed.\n" + e.message);
    }
  };

  const showGroupHeaders = catFilter === "All" && sortBy === "rank" && !query.trim();

  return (
    <div className="max-w-[1600px] mx-auto px-3 md:px-6 py-6 md:py-8">
      <Toast toast={toast} />

      <div className="flex flex-col lg:flex-row lg:items-end justify-between gap-4 mb-6">
        <div>
          <div className="text-xs font-black uppercase tracking-widest text-violet-600">
            HQ Admin
          </div>
          <h2 className="text-3xl md:text-4xl font-black tracking-tight">Live Results</h2>
          <p className="text-slate-500 font-medium mt-1">
            Updates automatically · ranks are calculated within each category
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          <ActionButton onClick={onGoToJudge} tone="sky" icon={<UserCheck size={18} />}>
            Judge console
          </ActionButton>
          <ActionButton onClick={downloadExcel} tone="emerald" solid icon={<Download size={18} />}>
            Export Excel
          </ActionButton>
          <ActionButton onClick={handleResetAll} tone="rose" icon={<Trash2 size={18} />}>
            Reset data
          </ActionButton>
          <ActionButton onClick={onLock} tone="slate" icon={<Lock size={18} />}>
            Lock
          </ActionButton>
        </div>
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 md:gap-4 mb-6">
        <StatCard tone="sky" label="Records" value={stats.total} />
        <StatCard tone="violet" label="Participants" value={stats.participants} />
        <StatCard tone="rose" label="Disqualified" value={stats.dq} />
        <StatCard tone="amber" label="Last submission" value={stats.last ? formatClock(stats.last) : "—"} />
      </div>

      <div className="bg-white rounded-3xl border border-slate-200 shadow-sm p-4 md:p-5 mb-4 space-y-4">
        <div className="flex flex-wrap gap-2">
          {["All", ...CATEGORIES].map((c) => {
            const active = catFilter === c;
            const count = c === "All" ? stats.total : stats.perCat[c];
            return (
              <button
                key={c}
                onClick={() => setCatFilter(c)}
                className={`h-11 px-4 rounded-2xl font-bold text-sm flex items-center gap-2 border-2 transition ${
                  active
                    ? "bg-violet-500 border-violet-500 text-white shadow-md shadow-violet-200"
                    : "bg-white border-slate-200 text-slate-600 hover:border-violet-300"
                }`}
              >
                {c}
                <span
                  className={`min-w-[1.5rem] h-6 px-1.5 rounded-full text-xs grid place-items-center ${
                    active ? "bg-white/25" : "bg-slate-100 text-slate-500"
                  }`}
                >
                  {count}
                </span>
              </button>
            );
          })}
        </div>
        <div className="flex flex-col md:flex-row gap-3 md:items-center">
          <div className="flex p-1 bg-slate-100 rounded-2xl w-full md:w-auto">
            {[
              ["all", "All trials"],
              ["best", "Best trial only"],
            ].map(([k, lbl]) => (
              <button
                key={k}
                onClick={() => setMode(k)}
                className={`flex-1 md:flex-none h-10 px-4 rounded-xl text-sm font-bold transition ${
                  mode === k ? "bg-white shadow text-slate-800" : "text-slate-500"
                }`}
              >
                {lbl}
              </button>
            ))}
          </div>
          <select
            value={sortBy}
            onChange={(e) => setSortBy(e.target.value)}
            className="h-12 px-4 rounded-2xl border-2 border-slate-200 bg-white font-bold text-slate-700 outline-none focus:border-violet-400"
          >
            <option value="rank">Sort: Rank (by category)</option>
            <option value="recent">Sort: Most recent</option>
            <option value="judge">Sort: Judge</option>
          </select>
          <div className="relative md:ml-auto md:w-80">
            <Search size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search name, No., judge"
              className="w-full h-12 pl-11 pr-4 rounded-2xl border-2 border-slate-200 bg-slate-50 focus:bg-white focus:border-violet-400 outline-none font-semibold"
            />
          </div>
        </div>
      </div>

      <div className="bg-white rounded-3xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full min-w-[1150px] text-left border-collapse">
            <thead className="bg-slate-50 text-slate-500 text-xs uppercase tracking-wider font-black border-b border-slate-200">
              <tr>
                <th className="p-3 text-center w-20">Rank</th>
                <th className="p-3">Category</th>
                <th className="p-3">No.</th>
                <th className="p-3">Participant</th>
                <th className="p-3 text-center">Trial</th>
                <th className="p-3 text-center">Coding</th>
                <th className="p-3 text-center">Missions</th>
                <th className="p-3 text-center">Landing</th>
                <th className="p-3 text-center bg-blue-50 text-blue-600">Total</th>
                <th className="p-3 text-center">Time</th>
                <th className="p-3">Judge</th>
                <th className="p-3 text-center w-16">Edit</th>
              </tr>
            </thead>
            <tbody>
              {visible.length === 0 && (
                <tr>
                  <td colSpan={12} className="p-12 text-center text-slate-400 font-semibold">
                    No records yet.
                  </td>
                </tr>
              )}
              {visible.map((r, i) => {
                const header =
                  showGroupHeaders && (i === 0 || visible[i - 1].category !== r.category);
                return (
                  <React.Fragment key={r.id}>
                    {header && (
                      <tr className="bg-gradient-to-r from-slate-50 to-white border-y border-slate-200">
                        <td colSpan={12} className="px-4 py-2.5">
                          <CategoryBadge category={r.category} full />
                        </td>
                      </tr>
                    )}
                    <tr
                      className={`border-b border-slate-100 transition-colors ${
                        r.isDQ ? "bg-rose-50/60" : "hover:bg-sky-50/60"
                      }`}
                    >
                      <td className="p-3 text-center">
                        <RankBadge rank={r.rank} />
                      </td>
                      <td className="p-3">
                        <CategoryBadge category={r.category} />
                      </td>
                      <td className="p-3 font-mono font-bold text-slate-600">{r.participantNumber}</td>
                      <td className="p-3 font-black text-slate-800 text-base">{r.participantName}</td>
                      <td className="p-3 text-center text-sm font-bold text-slate-500">{r.trial}</td>
                      <td className="p-3 text-center font-bold tabular-nums">{n(r.coding)}</td>
                      <td className="p-3">
                        <MissionDots record={r} />
                      </td>
                      <td className="p-3">
                        <LandingDots record={r} />
                      </td>
                      <td className="p-3 text-center bg-blue-50/50">
                        <span className={`text-xl font-black tabular-nums ${r.isDQ ? "text-rose-400 line-through" : "text-blue-600"}`}>
                          {n(r.totalScore)}
                        </span>
                      </td>
                      <td className="p-3 text-center font-mono font-bold text-slate-600">{formatTime(r)}</td>
                      <td className="p-3 text-sm text-slate-500 font-semibold">
                        {r.judgeName}
                        {r.editedAt && (
                          <span className="ml-1.5 text-[10px] font-black uppercase text-amber-600 bg-amber-50 px-1.5 py-0.5 rounded">
                            edited
                          </span>
                        )}
                      </td>
                      <td className="p-3 text-center">
                        <button
                          onClick={() => setEditingId(r.id)}
                          className="w-10 h-10 rounded-xl grid place-items-center mx-auto text-slate-400 hover:text-violet-600 hover:bg-violet-50 transition"
                          title="Edit record"
                        >
                          <Pencil size={18} />
                        </button>
                      </td>
                    </tr>
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {editingRecord && (
        <EditModal
          key={editingRecord.id}
          record={editingRecord}
          onClose={() => setEditingId(null)}
          showToast={showToast}
        />
      )}
    </div>
  );
}

function EditModal({ record, onClose, showToast }) {
  const sigRef = useRef(null);
  const [form, patch] = useScoreForm(() => ({
    judgeName: "",
    category: "Elementary",
    participantNumber: "",
    participantName: "",
    trial: "1st",
    ...EMPTY_SCORES,
    ...record,
  }));
  const [sigUrl, setSigUrl] = useState(null);
  const [sigState, setSigState] = useState("loading"); // loading | ok | error
  const [saving, setSaving] = useState(false);
  const calc = calcScores(form);

  // 서명: 신규 경로(signatures/{id}) → 구버전 내장 서명 순으로 조회
  useEffect(() => {
    let alive = true;
    (async () => {
      try {
        let s = (await get(ref(db, `signatures/${record.id}`))).val();
        if (!s) s = (await get(ref(db, `scores/${record.id}/signature`))).val();
        if (alive) {
          setSigUrl(s || null);
          setSigState("ok");
        }
      } catch (e) {
        console.error(e);
        if (alive) setSigState("error");
      }
    })();
    return () => {
      alive = false;
    };
  }, [record.id]);

  useEffect(() => {
    const onKey = (e) => e.key === "Escape" && onClose();
    window.addEventListener("keydown", onKey);
    document.body.style.overflow = "hidden";
    return () => {
      window.removeEventListener("keydown", onKey);
      document.body.style.overflow = "";
    };
  }, [onClose]);

  const handleSave = async () => {
    const rec = buildRecord(form);
    if (!rec.judgeName || !rec.participantNumber || !rec.participantName)
      return alert("Judge Name, Participant No. and Participant Name are required.");
    if (!window.confirm("Save changes to this record?")) return;

    setSaving(true);
    try {
      const pad = sigRef.current;
      const changed = pad && pad.isChanged();
      let signature = sigUrl;
      if (changed) signature = pad.isEmpty() ? null : pad.toDataURL();

      const updates = {
        // [수정] id 제외, 점수 재계산 값으로 레코드 전체 교체 (구버전 내장 서명은 분리 이동)
        [`scores/${record.id}`]: {
          ...rec,
          hasSignature: !!signature,
          timestamp: record.timestamp || Date.now(),
          editedAt: Date.now(),
        },
      };
      if (changed || sigState === "ok") updates[`signatures/${record.id}`] = signature;
      await update(ref(db), updates);
      showToast("Record updated");
      onClose();
    } catch (e) {
      alert("Save failed.\n" + e.message);
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!window.confirm(`Delete the record of No. ${record.participantNumber} (${record.participantName})?\nThis cannot be undone.`))
      return;
    try {
      await update(ref(db), { [`scores/${record.id}`]: null, [`signatures/${record.id}`]: null });
      showToast("Record deleted", "warn");
      onClose();
    } catch (e) {
      alert("Delete failed.\n" + e.message);
    }
  };

  return (
    <div className="fixed inset-0 z-[100] bg-slate-900/40 backdrop-blur-sm flex items-end sm:items-center justify-center sm:p-4">
      <div className="bg-slate-50 w-full max-w-3xl max-h-[95vh] sm:max-h-[92vh] rounded-t-3xl sm:rounded-3xl shadow-2xl flex flex-col overflow-hidden">
        <div className="flex items-center gap-3 px-5 py-4 bg-white border-b border-slate-200">
          <div className="w-10 h-10 rounded-2xl bg-violet-100 text-violet-600 grid place-items-center">
            <Pencil size={18} />
          </div>
          <div className="min-w-0">
            <h3 className="text-xl font-black leading-tight">Edit Record</h3>
            <p className="text-sm text-slate-500 font-semibold truncate">
              #{record.participantNumber} · {record.participantName} · {record.trial}
            </p>
          </div>
          <button
            onClick={onClose}
            className="ml-auto w-10 h-10 rounded-xl grid place-items-center text-slate-400 hover:text-slate-700 hover:bg-slate-100"
          >
            <X size={22} />
          </button>
        </div>

        <div className="overflow-y-auto p-4 md:p-6 space-y-5">
          <ParticipantFields form={form} patch={patch} />
          <ScoreFields form={form} patch={patch} />
          <Section step={6} title="Participant Signature" icon={<PenLine size={18} />} tone="slate">
            {sigState === "loading" ? (
              <div className="h-40 grid place-items-center text-slate-400 font-bold">
                <span className="flex items-center gap-2">
                  <Loader2 size={18} className="animate-spin" /> Loading signature…
                </span>
              </div>
            ) : (
              <>
                {sigState === "error" && (
                  <p className="mb-2 text-sm font-bold text-amber-600">
                    Could not load the existing signature. It will be kept unless you draw a new one.
                  </p>
                )}
                <SignatureBox ref={sigRef} height={170} initialDataURL={sigUrl} />
              </>
            )}
          </Section>
        </div>

        <div className="flex items-center gap-3 px-4 md:px-6 py-4 bg-white border-t border-slate-200">
          <div>
            <div className="text-[11px] font-black uppercase tracking-widest text-slate-400">Total</div>
            <div className="text-3xl font-black text-blue-600 tabular-nums leading-none">
              {calc.totalScore}
            </div>
          </div>
          <button
            onClick={handleDelete}
            className="ml-auto h-12 px-4 rounded-2xl border-2 border-rose-200 text-rose-600 font-bold hover:bg-rose-50 flex items-center gap-2"
          >
            <Trash2 size={18} /> <span className="hidden sm:inline">Delete</span>
          </button>
          <button
            onClick={handleSave}
            disabled={saving || sigState === "loading"}
            className="h-12 px-6 rounded-2xl bg-gradient-to-r from-violet-500 to-indigo-500 text-white font-black shadow-lg shadow-violet-200 flex items-center gap-2 disabled:opacity-50"
          >
            {saving ? <Loader2 size={18} className="animate-spin" /> : <Save size={18} />} Save
          </button>
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// Auth screens
// ─────────────────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [pw, setPw] = useState("");
  const [err, setErr] = useState(false);
  const submit = () => {
    if (!onLogin(pw)) {
      setErr(true);
      setPw("");
    }
  };
  return (
    <div className="min-h-screen grid place-items-center bg-gradient-to-br from-sky-100 via-white to-amber-100 p-6 text-slate-800">
      <div className="w-full max-w-sm bg-white rounded-[2rem] shadow-xl shadow-sky-100 border border-sky-50 p-8 text-center">
        <div className="mx-auto w-16 h-16 rounded-3xl bg-gradient-to-br from-sky-400 to-blue-600 text-white grid place-items-center shadow-lg shadow-sky-200 mb-5">
          <Plane size={30} />
        </div>
        <h1 className="text-2xl font-black">IDCC 2026</h1>
        <p className="text-sm font-bold text-sky-600 mb-7">DCC · Rescue Flight Scoring System</p>
        <input
          type="password"
          value={pw}
          autoFocus
          onChange={(e) => {
            setPw(e.target.value);
            setErr(false);
          }}
          onKeyDown={(e) => e.key === "Enter" && submit()}
          placeholder="Password"
          className={`w-full h-14 px-4 rounded-2xl border-2 bg-slate-50 text-center text-lg font-bold outline-none focus:bg-white transition ${
            err ? "border-rose-300 focus:border-rose-400" : "border-slate-200 focus:border-sky-400"
          }`}
        />
        <p className={`text-sm font-bold text-rose-600 h-6 mt-2 ${err ? "" : "invisible"}`}>
          Incorrect password. Please try again.
        </p>
        <button
          onClick={submit}
          className="w-full h-14 mt-2 rounded-2xl bg-gradient-to-r from-blue-500 to-sky-500 text-white font-black text-lg shadow-lg shadow-sky-200 hover:from-blue-600 hover:to-sky-600 transition"
        >
          Sign in
        </button>
      </div>
    </div>
  );
}

function HQAuth({ onAuth, onBack }) {
  const [pw, setPw] = useState("");
  const [err, setErr] = useState(false);
  const submit = () => {
    if (pw === HQ_PASSWORD) onAuth();
    else {
      setErr(true);
      setPw("");
    }
  };
  return (
    <div className="px-4 py-16">
      <div className="max-w-md mx-auto p-8 md:p-10 bg-white rounded-[2rem] shadow-xl shadow-violet-100 border border-violet-50 text-center">
        <div className="mx-auto w-16 h-16 rounded-3xl bg-gradient-to-br from-violet-400 to-indigo-500 text-white grid place-items-center shadow-lg shadow-violet-200 mb-5">
          <Lock size={28} />
        </div>
        <h3 className="text-2xl font-black">HQ Admin Access</h3>
        <p className="text-slate-500 font-medium mb-6">Enter the admin password to continue.</p>
        <input
          type="password"
          value={pw}
          autoFocus
          onChange={(e) => {
            setPw(e.target.value);
            setErr(false);
          }}
          onKeyDown={(e) => e.key === "Enter" && submit()}
          placeholder="••••••••"
          className={`w-full h-14 px-4 rounded-2xl border-2 bg-slate-50 text-center text-xl tracking-[0.4em] font-bold outline-none focus:bg-white ${
            err ? "border-rose-300" : "border-slate-200 focus:border-violet-400"
          }`}
        />
        <p className={`text-sm font-bold text-rose-600 h-6 mt-2 ${err ? "" : "invisible"}`}>
          Access denied.
        </p>
        <div className="flex gap-3 mt-2">
          <button
            onClick={onBack}
            className="h-14 px-5 rounded-2xl bg-slate-100 text-slate-600 font-bold flex items-center gap-1 hover:bg-slate-200"
          >
            <ChevronLeft size={18} /> Back
          </button>
          <button
            onClick={submit}
            className="flex-1 h-14 rounded-2xl bg-gradient-to-r from-violet-500 to-indigo-500 text-white font-black shadow-lg shadow-violet-200"
          >
            Unlock
          </button>
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// UI building blocks
// ─────────────────────────────────────────────────────────────
const TONES = {
  sky: { badge: "bg-sky-500", text: "text-sky-500", chip: "bg-sky-50 text-sky-700 ring-sky-200" },
  indigo: { badge: "bg-indigo-500", text: "text-indigo-500", chip: "bg-indigo-50 text-indigo-700 ring-indigo-200" },
  emerald: { badge: "bg-emerald-500", text: "text-emerald-500", chip: "bg-emerald-50 text-emerald-700 ring-emerald-200" },
  amber: { badge: "bg-amber-400", text: "text-amber-500", chip: "bg-amber-50 text-amber-700 ring-amber-200" },
  rose: { badge: "bg-rose-500", text: "text-rose-500", chip: "bg-rose-50 text-rose-700 ring-rose-200" },
  slate: { badge: "bg-slate-500", text: "text-slate-500", chip: "bg-slate-100 text-slate-700 ring-slate-200" },
  violet: { badge: "bg-violet-500", text: "text-violet-500", chip: "bg-violet-50 text-violet-700 ring-violet-200" },
};

function Section({ step, title, subtitle, icon, tone = "sky", right, children }) {
  const t = TONES[tone];
  return (
    <section className="bg-white rounded-3xl border border-slate-200/80 shadow-sm p-4 md:p-6">
      <div className="flex items-start gap-3 mb-4">
        <div className={`w-9 h-9 rounded-xl ${t.badge} text-white grid place-items-center font-black shrink-0 shadow-sm`}>
          {step}
        </div>
        <div className="min-w-0 flex-1">
          <h3 className="font-black text-lg leading-tight flex items-center gap-2">
            <span className={t.text}>{icon}</span>
            {title}
          </h3>
          {subtitle && <p className="text-sm text-slate-500 font-medium mt-0.5">{subtitle}</p>}
        </div>
        {right}
      </div>
      {children}
    </section>
  );
}

function ScoreChip({ tone, children }) {
  return (
    <span className={`shrink-0 px-3 py-1.5 rounded-xl ring-1 text-sm font-black tabular-nums ${TONES[tone].chip}`}>
      {children}
    </span>
  );
}

function FieldLabel({ children }) {
  return (
    <span className="block text-xs font-black uppercase tracking-widest text-slate-500 mb-1.5 ml-1">
      {children}
    </span>
  );
}

function TextField({ label, value, onChange, placeholder, className = "" }) {
  return (
    <label className={`block ${className}`}>
      <FieldLabel>{label}</FieldLabel>
      <input
        type="text"
        value={value ?? ""}
        onChange={(e) => onChange(e.target.value)}
        onFocus={selectAll}
        placeholder={placeholder}
        autoComplete="off"
        className="w-full h-14 px-4 rounded-2xl border-2 border-slate-200 bg-slate-50 focus:bg-white focus:border-sky-400 outline-none font-bold text-lg text-slate-800 placeholder:text-slate-300 placeholder:font-medium transition"
      />
    </label>
  );
}

function Segmented({ options, value, onChange, cols }) {
  return (
    <div className={`grid gap-2 ${cols}`}>
      {options.map((o) => (
        <button
          type="button"
          key={o}
          onClick={() => onChange(o)}
          className={`h-12 px-3 rounded-xl font-bold text-sm md:text-base border-2 transition ${
            value === o
              ? "bg-sky-500 border-sky-500 text-white shadow-md shadow-sky-200"
              : "bg-white border-slate-200 text-slate-600 hover:border-sky-300"
          }`}
        >
          {o}
        </button>
      ))}
    </div>
  );
}

function StepButton({ onClick, disabled, children, label }) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      aria-label={label}
      className="w-14 h-14 rounded-2xl bg-white border-2 border-slate-200 text-slate-700 grid place-items-center hover:border-slate-300 active:scale-95 transition disabled:opacity-30 disabled:active:scale-100"
    >
      {children}
    </button>
  );
}

function CodingInput({ value, onChange }) {
  const v = n(value);
  return (
    <div className="flex flex-col sm:flex-row sm:items-center gap-4">
      <div className="flex items-center gap-3">
        <StepButton label="Decrease" onClick={() => onChange(clamp(v - 1, 0, CODING_MAX))} disabled={v <= 0}>
          <Minus size={22} />
        </StepButton>
        <input
          type="text"
          inputMode="numeric"
          value={v}
          onFocus={selectAll}
          onChange={(e) => onChange(clamp(toInt(e.target.value), 0, CODING_MAX))}
          className="w-24 h-16 text-center text-3xl font-black tabular-nums rounded-2xl border-2 border-indigo-200 bg-indigo-50 text-indigo-700 outline-none focus:border-indigo-500"
        />
        <StepButton label="Increase" onClick={() => onChange(clamp(v + 1, 0, CODING_MAX))} disabled={v >= CODING_MAX}>
          <Plus size={22} />
        </StepButton>
      </div>
      <div className="flex flex-wrap gap-2">
        {[0, 5, 10, 15, 20].map((q) => (
          <button
            type="button"
            key={q}
            onClick={() => onChange(q)}
            className={`h-11 min-w-[3.25rem] px-3 rounded-xl font-bold border-2 transition ${
              v === q
                ? "bg-indigo-500 border-indigo-500 text-white"
                : "bg-white border-slate-200 text-slate-600 hover:border-indigo-300"
            }`}
          >
            {q}
          </button>
        ))}
      </div>
    </div>
  );
}

function MissionToggle({ index, label, on, disabled, disabledNote, onToggle }) {
  if (disabled)
    return (
      <div className="rounded-2xl border-2 border-dashed border-slate-200 bg-slate-50 p-4 min-h-[112px] flex flex-col justify-between text-slate-400">
        <div className="flex items-center justify-between">
          <span className="text-xs font-black">#{index}</span>
          <Ban size={18} />
        </div>
        <div>
          <div className="font-bold">{label}</div>
          <div className="text-xs font-semibold mt-0.5">{disabledNote}</div>
        </div>
      </div>
    );
  return (
    <button
      type="button"
      onClick={onToggle}
      aria-pressed={on}
      className={`rounded-2xl border-2 p-4 min-h-[112px] flex flex-col justify-between text-left transition active:scale-[0.97] ${
        on
          ? "bg-emerald-500 border-emerald-500 text-white shadow-lg shadow-emerald-200"
          : "bg-white border-slate-200 text-slate-700 hover:border-emerald-300 hover:bg-emerald-50/50"
      }`}
    >
      <div className="flex items-center justify-between">
        <span className={`text-xs font-black ${on ? "text-emerald-100" : "text-slate-400"}`}>#{index}</span>
        <span
          className={`w-8 h-8 rounded-full grid place-items-center ${
            on ? "bg-white text-emerald-600" : "border-2 border-slate-300 text-transparent"
          }`}
        >
          <Check size={18} strokeWidth={3} />
        </span>
      </div>
      <div>
        <div className="font-black text-base leading-tight">{label}</div>
        <div className={`text-sm font-bold mt-1 ${on ? "text-white" : "text-slate-400"}`}>
          {on ? `Cleared +${MISSION_PTS}` : "Tap if cleared"}
        </div>
      </div>
    </button>
  );
}

function LandingMeter({ form, used }) {
  const slots = [];
  LANDING_ZONES.forEach((z) => {
    for (let i = 0; i < n(form[z.key]); i++) slots.push(z.pad);
  });
  while (slots.length < MAX_LANDINGS) slots.push(null);
  return (
    <div className="flex flex-wrap items-center gap-3">
      <span className="text-sm font-bold text-slate-500">Landings used</span>
      <div className="flex gap-1.5">
        {slots.slice(0, MAX_LANDINGS).map((c, i) => (
          <span key={i} className={`w-9 h-3 rounded-full ${c || "bg-slate-200"}`} />
        ))}
      </div>
      <span className="text-sm font-black text-slate-700 tabular-nums">
        {used}/{MAX_LANDINGS}
      </span>
      {used >= MAX_LANDINGS && (
        <span className="text-xs font-black text-amber-700 bg-amber-100 px-2 py-0.5 rounded-lg">
          Max reached
        </span>
      )}
    </div>
  );
}

function LandingCounter({ zone, count, canAdd, onChange }) {
  return (
    <div className={`rounded-2xl ${zone.soft} ring-1 ${zone.ring} p-4`}>
      <div className="flex items-center gap-3">
        <div className={`w-12 h-12 rounded-full ${zone.pad} ring-4 ring-white shadow grid place-items-center text-white font-black`}>
          {zone.pts}
        </div>
        <div>
          <div className={`font-black ${zone.text}`}>{zone.name} Zone</div>
          <div className="text-xs font-semibold text-slate-500">{zone.pts} pts per landing</div>
        </div>
      </div>
      <div className="flex items-center justify-between mt-4">
        <StepButton label={`Remove ${zone.name}`} onClick={() => onChange(count - 1)} disabled={count <= 0}>
          <Minus size={22} />
        </StepButton>
        <div className="text-center">
          <div className="text-4xl font-black tabular-nums text-slate-800 leading-none">{count}</div>
          <div className={`text-sm font-bold ${zone.text} mt-1 tabular-nums`}>= {count * zone.pts} pts</div>
        </div>
        <StepButton label={`Add ${zone.name}`} onClick={() => onChange(count + 1)} disabled={!canAdd}>
          <Plus size={22} />
        </StepButton>
      </div>
    </div>
  );
}

function TimeInput({ form, patch }) {
  const parts = [
    { k: "min", label: "MIN", max: 59 },
    { k: "sec", label: "SEC", max: 59 },
    { k: "ms", label: "1/100 s", max: 99 },
  ];
  return (
    <div>
      <FieldLabel>Lap Time</FieldLabel>
      <div className="flex items-start gap-2">
        {parts.map((p, i) => (
          <React.Fragment key={p.k}>
            {i > 0 && (
              <span className="text-3xl font-black text-slate-300 leading-[4rem]">{i === 1 ? ":" : "."}</span>
            )}
            <div className="flex flex-col items-center">
              <input
                type="text"
                inputMode="numeric"
                value={i === 0 ? n(form[p.k]) : pad2(form[p.k])}
                onFocus={selectAll}
                onChange={(e) => patch({ [p.k]: clamp(toInt(e.target.value.slice(-2)), 0, p.max) })}
                className="w-20 h-16 text-center text-3xl font-black tabular-nums rounded-2xl border-2 border-slate-200 bg-slate-50 focus:bg-white focus:border-rose-400 outline-none"
              />
              <span className="text-[11px] font-black text-slate-400 mt-1 tracking-widest">{p.label}</span>
            </div>
          </React.Fragment>
        ))}
      </div>
    </div>
  );
}

function DQToggle({ on, onToggle }) {
  return (
    <button
      type="button"
      onClick={onToggle}
      aria-pressed={on}
      className={`md:ml-auto md:mb-5 flex items-center justify-center gap-3 px-6 h-16 rounded-2xl border-2 font-black text-lg transition ${
        on
          ? "bg-rose-500 border-rose-500 text-white shadow-lg shadow-rose-200"
          : "bg-white border-rose-200 text-rose-500 hover:bg-rose-50"
      }`}
    >
      <Ban size={22} />
      {on ? "Disqualified (DQ)" : "Mark as DQ"}
    </button>
  );
}

/** 서명 패드 — 화면 크기 변경 시 서명 보존, 기존 서명 이미지 불러오기 지원 */
const SignatureBox = forwardRef(function SignatureBox({ height = 200, initialDataURL = null }, outerRef) {
  const padRef = useRef(null);
  const boxRef = useRef(null);
  const widthRef = useRef(0);
  const strokesRef = useRef(null);
  const drawnRef = useRef(false);
  const changedRef = useRef(false);
  const showInitialRef = useRef(!!initialDataURL);
  const [width, setWidth] = useState(0);
  const [hasInk, setHasInk] = useState(!!initialDataURL);

  const clearPad = useCallback(() => {
    if (padRef.current) padRef.current.clear();
    strokesRef.current = null;
    drawnRef.current = false;
    showInitialRef.current = false;
    changedRef.current = true;
    setHasInk(false);
  }, []);

  useImperativeHandle(
    outerRef,
    () => ({
      isEmpty: () => !padRef.current || padRef.current.isEmpty(),
      isChanged: () => changedRef.current,
      toDataURL: () => padRef.current.toDataURL("image/png"),
      clear: () => {
        clearPad();
        changedRef.current = false;
      },
    }),
    [clearPad]
  );

  useLayoutEffect(() => {
    const el = boxRef.current;
    if (!el) return undefined;
    const measure = () => {
      const w = Math.floor(el.clientWidth);
      if (!w || w === widthRef.current) return;
      if (padRef.current && drawnRef.current) strokesRef.current = padRef.current.toData();
      widthRef.current = w;
      setWidth(w);
    };
    measure();
    let ro = null;
    if (typeof ResizeObserver !== "undefined") {
      ro = new ResizeObserver(measure);
      ro.observe(el);
    } else {
      window.addEventListener("resize", measure);
    }
    return () => {
      if (ro) ro.disconnect();
      else window.removeEventListener("resize", measure);
    };
  }, []);

  useEffect(() => {
    const pad = padRef.current;
    if (!pad || !width) return;
    if (strokesRef.current && strokesRef.current.length) {
      pad.fromData(strokesRef.current);
      strokesRef.current = null;
    } else if (showInitialRef.current && initialDataURL) {
      pad.fromDataURL(initialDataURL, { width, height });
    }
  }, [width, height, initialDataURL]);

  return (
    <div>
      <div
        ref={boxRef}
        className="relative w-full rounded-2xl border-2 border-dashed border-sky-200 bg-gradient-to-b from-white to-sky-50/50 overflow-hidden touch-none"
        style={{ height }}
      >
        <div className="absolute left-6 right-6 bottom-10 border-b-2 border-slate-200" />
        {!hasInk && (
          <div className="absolute inset-0 grid place-items-center text-slate-300 pointer-events-none">
            <div className="flex flex-col items-center gap-1">
              <PenLine size={28} />
              <span className="font-bold">Participant signs here</span>
            </div>
          </div>
        )}
        {width > 0 && (
          <div className="relative z-10">
            <SignatureCanvas
              ref={padRef}
              penColor="#0f172a"
              minWidth={1.2}
              maxWidth={3}
              clearOnResize={false}
              onBegin={() => {
                drawnRef.current = true;
                changedRef.current = true;
                setHasInk(true);
              }}
              canvasProps={{ width, height, className: "block cursor-crosshair" }}
            />
          </div>
        )}
      </div>
      <button
        type="button"
        onClick={clearPad}
        className="mt-2 inline-flex items-center gap-1.5 text-sm font-bold text-slate-500 hover:text-rose-600 transition"
      >
        <RotateCcw size={14} /> Clear signature
      </button>
    </div>
  );
});

function CategoryBadge({ category, full }) {
  return (
    <span
      className={`inline-flex px-2.5 py-1 rounded-lg text-xs font-black ring-1 whitespace-nowrap ${
        CATEGORY_STYLE[category] || "bg-slate-100 text-slate-600 ring-slate-200"
      }`}
    >
      {full ? category : CATEGORY_SHORT[category] || category || "—"}
    </span>
  );
}

function RankBadge({ rank }) {
  if (rank == null)
    return (
      <span className="inline-grid place-items-center h-9 px-2.5 rounded-xl bg-rose-100 text-rose-600 text-sm font-black">
        DQ
      </span>
    );
  const style =
    rank === 1
      ? "bg-gradient-to-br from-amber-300 to-amber-500 text-white shadow-md shadow-amber-200"
      : rank === 2
      ? "bg-gradient-to-br from-slate-200 to-slate-400 text-white shadow-md shadow-slate-200"
      : rank === 3
      ? "bg-gradient-to-br from-orange-200 to-orange-400 text-white shadow-md shadow-orange-200"
      : "bg-slate-100 text-slate-500";
  return (
    <span className={`inline-grid place-items-center w-9 h-9 rounded-xl text-base font-black ${style}`}>
      {rank}
    </span>
  );
}

function MissionDots({ record }) {
  return (
    <div className="flex justify-center gap-1">
      {MISSIONS.map((m) => {
        const allowed = missionAllowed(m.key, record.category);
        const ok = allowed && n(record[m.key]) === MISSION_PTS;
        return (
          <span
            key={m.key}
            title={`${m.label}: ${allowed ? n(record[m.key]) : "N/A"}`}
            className={`min-w-[2.25rem] h-7 px-1 rounded-lg text-[10px] font-black grid place-items-center ${
              !allowed
                ? "bg-slate-50 text-slate-300 border border-dashed border-slate-200"
                : ok
                ? "bg-emerald-500 text-white"
                : "bg-slate-100 text-slate-400"
            }`}
          >
            {m.short}
          </span>
        );
      })}
    </div>
  );
}

function LandingDots({ record }) {
  const pts = n(record.lndBlueScore) + n(record.lndRedScore) + n(record.lndYellowScore);
  return (
    <div className="flex items-center justify-center gap-2">
      {LANDING_ZONES.map((z) => (
        <span key={z.key} className="flex items-center gap-1 text-sm font-bold text-slate-600 tabular-nums">
          <span className={`w-3 h-3 rounded-full ${z.pad}`} />
          {n(record[z.key])}
        </span>
      ))}
      <span className="text-xs font-black text-slate-400 tabular-nums">({pts})</span>
    </div>
  );
}

const STAT_TONES = {
  sky: "from-sky-50 border-sky-100 text-sky-600",
  violet: "from-violet-50 border-violet-100 text-violet-600",
  rose: "from-rose-50 border-rose-100 text-rose-600",
  amber: "from-amber-50 border-amber-100 text-amber-600",
};

function StatCard({ tone, label, value }) {
  return (
    <div className={`rounded-3xl border bg-gradient-to-br to-white p-4 md:p-5 ${STAT_TONES[tone]}`}>
      <div className="text-xs font-black uppercase tracking-widest opacity-80">{label}</div>
      <div className="text-3xl md:text-4xl font-black text-slate-800 tabular-nums mt-1">{value}</div>
    </div>
  );
}

const ACTION_TONES = {
  sky: "border-sky-200 text-sky-600 hover:bg-sky-50",
  rose: "border-rose-200 text-rose-600 hover:bg-rose-50",
  slate: "border-slate-200 text-slate-600 hover:bg-slate-50",
  emerald: "border-emerald-500 bg-emerald-500 text-white hover:bg-emerald-600 shadow-lg shadow-emerald-200",
};

function ActionButton({ onClick, tone, solid, icon, children }) {
  return (
    <button
      onClick={onClick}
      className={`h-12 px-4 md:px-5 rounded-2xl border-2 font-bold text-sm flex items-center gap-2 transition whitespace-nowrap ${
        solid ? "" : "bg-white"
      } ${ACTION_TONES[tone]}`}
    >
      {icon}
      {children}
    </button>
  );
}
