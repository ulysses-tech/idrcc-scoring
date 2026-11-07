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
import {
  getDatabase,
  ref,
  push,
  onValue,
  update,
  get,
  set,
  remove,
  onDisconnect,
} from "firebase/database";
import {
  RosterPicker,
  RosterManager,
  DEVICE_ID,
  claimedByOther,
  FLIGHTS_PER_TRIAL,
  flightOf,
  nextRoundFor,
} from "./Roster";
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
/** 각 시기(1st·2nd)마다 2회 비행 → 총 4회. 비행 번호(flight)가 없는 예전 기록은 1회로 취급 */
const ATTEMPTS = TRIALS.flatMap((t) =>
  Array.from({ length: FLIGHTS_PER_TRIAL }, (_, i) => `${t}-${i + 1}`)
); // ["1st-1", "1st-2", "2nd-1", "2nd-2"]
const attemptLabel = (r) => `${r?.trial || "1st"}-${flightOf(r)}`;
const parseAttempt = (label) => {
  const [trial, f] = String(label).split("-");
  return { trial, flight: Number(f) || 1 };
};

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
    flight: flightOf(form),
  };
  Object.keys(EMPTY_SCORES).forEach((k) => {
    out[k] = k === "isDQ" ? !!form.isDQ : n(form[k]);
  });
  // 명단에서 가져온 부가 정보 (있을 때만 저장)
  PROFILE_FIELDS.forEach((k) => {
    const v = String(form[k] ?? "").trim();
    if (v) out[k] = v;
  });
  if (form.rosterId) out.rosterId = String(form.rosterId);
  // 라운드 연결 정보: 같은 라운드의 1·2차 비행은 roundId·서명(signatureId)을 공유
  if (form.roundId) out.roundId = String(form.roundId);
  if (form.signatureId) out.signatureId = String(form.signatureId);
  if (form.noFlight) out.noFlight = true;
  return { ...out, ...calcScores(out) };
}

// ─────────────────────────────────────────────────────────────
// 참가자 명단 저장소 (Firebase: dcc_meta/roster, dcc_meta/claims)
// ─────────────────────────────────────────────────────────────
const ROSTER_PATH = "dcc_meta";
const PROFILE_FIELDS = ["country", "affiliation", "grade", "gender"];
const WRITE_TIMEOUT_MS = 10000;
const withTimeout = (p) =>
  Promise.race([
    p,
    new Promise((_, reject) => setTimeout(() => reject(new Error("timeout")), WRITE_TIMEOUT_MS)),
  ]);
const str = (v) => String(v ?? "").trim();

const rosterFromObject = (obj) =>
  Object.entries(obj || {})
    .map(([id, v]) => ({ id, ...v }))
    .sort((a, b) => n(a.order) - n(b.order));

/** 명단 전체 교체 (빈 배열이면 비우기) */
function saveRosterEntries(entries) {
  const obj = {};
  entries.forEach((e, i) => {
    obj[`p${String(i + 1).padStart(4, "0")}`] = {
      order: n(e.order) || i + 1,
      bib: str(e.bib),
      name: str(e.name),
      category: str(e.category),
      group: str(e.group),
      ...Object.fromEntries(PROFILE_FIELDS.map((k) => [k, str(e[k])])),
    };
  });
  return withTimeout(update(ref(db, ROSTER_PATH), { roster: entries.length ? obj : null, claims: null }));
}

/** 다른 심판에게 '채점 중' 표시 — 기기 연결이 끊기면 서버가 자동으로 지움 */
async function claimRosterEntry(rosterId, judge) {
  const r = ref(db, `${ROSTER_PATH}/claims/${rosterId}`);
  await withTimeout(set(r, { judge: str(judge) || "Judge", device: DEVICE_ID, at: Date.now() }));
  onDisconnect(r).remove().catch(() => {});
}

function releaseRosterEntry(rosterId) {
  if (!db || !rosterId) return;
  const r = ref(db, `${ROSTER_PATH}/claims/${rosterId}`);
  onDisconnect(r).cancel().catch(() => {});
  remove(r).catch(() => {});
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
/** 순위 제외 기록: 실격(DQ) 또는 비행 안 함 */
const unranked = (r) => !!r.isDQ || !!r.noFlight;

function compareRecords(a, b) {
  if (unranked(a) !== unranked(b)) return unranked(a) ? 1 : -1;
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
      if (unranked(r)) {
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
  const reset = () => setSw({ status: "idle", startAt: 0, elapsed: 0 });
  const toggle = () => {
    if (sw.status === "running") stop();
    else start();
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
  const [legacySigs, setLegacySigs] = useState({}); // 예전 방식(기록 안에 저장된) 서명
  const [rosterState, setRosterState] = useState({ roster: [], claims: {}, error: null, loaded: false });
  const [hqAuthenticated, setHqAuthenticated] = useState(false);
  const connected = useConnection();

  // 참가자 명단 · 채점 중 표시 구독
  useEffect(() => {
    if (!db || !isLoggedIn) return undefined;
    return onValue(
      ref(db, ROSTER_PATH),
      (snap) => {
        const v = snap.val() || {};
        setRosterState({ roster: rosterFromObject(v.roster), claims: v.claims || {}, error: null, loaded: true });
      },
      (err) => {
        console.error("Roster sync error:", err);
        setRosterState({ roster: [], claims: {}, error: err?.code || "error", loaded: false });
      }
    );
  }, [isLoggedIn]);

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
        const legacy = {};
        setRecords(
          Object.entries(data).map(([id, v]) => {
            const { signature, ...rest } = v || {};
            if (signature) legacy[id] = signature;
            return { ...rest, id, hasLegacySignature: !!signature };
          })
        );
        setLegacySigs(legacy);
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
        <JudgeView
          records={records}
          connected={connected}
          onGoToHQ={goHQ}
          roster={rosterState.roster}
          claims={rosterState.claims}
          rosterLoaded={rosterState.loaded}
        />
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
          legacySigs={legacySigs}
          rosterState={rosterState}
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
// Judge — 라운드 단위 채점
//  1 라운드 = 코딩 + 연속 2회 비행(3분 안). 1차 비행은 심판 신호로, 2차 비행은 참가자가 바로 시작.
//  · 1차 STOP → 1차 기록 자동 저장(HQ에 즉시 표시) → 곧바로 2차 START 대기
//  · 2차 STOP → 1·2차 요약을 학생이 확인하고 서명 → 확정 시 2차 기록 저장(HQ 표시)
//  · 서명은 라운드당 1회(총 2회), 같은 라운드의 두 기록이 함께 사용
// ─────────────────────────────────────────────────────────────
const ROUND_KEY = "idcc_round_v2";
const ROUND_LABEL = { "1st": "1 라운드", "2nd": "2 라운드" };
const ROUND_OPTIONS = TRIALS.map((t) => ROUND_LABEL[t]);
const FLOW_OPTIONS = [
  { key: "round", label: "라운드별 진행 (전원 1R → 2R)" },
  { key: "participant", label: "참가자별 연속 (1R → 바로 2R)" },
];
const FLIGHT_FIELDS = [
  ...MISSIONS.map((m) => m.key),
  "lndBlueCount",
  "lndRedCount",
  "lndYellowCount",
  "min",
  "sec",
  "ms",
  "isDQ",
];

/** 비행 1회의 입력 상태 (status: ready → running → stopped → saved) */
const emptyFlight = () => ({
  ...Object.fromEntries(FLIGHT_FIELDS.map((k) => [k, EMPTY_SCORES[k]])),
  noFlight: false,
  status: "ready",
  recordId: null,
  savedAt: 0,
});

/** 새 라운드 — 심판 이름·부문·라운드는 이어서 사용 */
const freshRound = (prev = {}) => ({
  judgeName: prev.judgeName ?? (storage.get("idrcc_judge_name") || ""),
  category: prev.category ?? (storage.get("idrcc_category") || "Elementary"),
  trial: prev.trial ?? "1st",
  participantNumber: "",
  participantName: "",
  country: "",
  affiliation: "",
  grade: "",
  gender: "",
  rosterId: null,
  coding: 0,
  codingDone: false,
  readyAck: false,
  flights: [emptyFlight(), emptyFlight()],
  active: 0, // 0: 1차 비행 차례, 1: 2차 비행 차례, 2: 학생 확인·서명
});

function loadRound() {
  try {
    const r = JSON.parse(storage.get(ROUND_KEY) || "null");
    if (r && Array.isArray(r.flights) && r.flights.length === 2) {
      // 새로고침 직전 측정 중이던 비행은 초시계 상태가 이어서 관리
      return { ...freshRound(), ...r };
    }
  } catch {
    /* ignore corrupt state */
  }
  return freshRound();
}

/** 라운드 + 비행 i → 저장용 기록 */
const flightRecord = (round, i) =>
  buildRecord({
    ...round,
    ...round.flights[i],
    flight: i + 1,
    rosterId: round.rosterId || undefined,
    roundId: round.flights[0].recordId || undefined,
  });

/** 이 부문에서 쓰는 비행 단계 (순서대로) */
const stagesFor = (category) => MISSIONS.filter((m) => missionAllowed(m.key, category));

/** 통과한 단계 목록 (단계마다 개별 선택) */
const clearedStages = (flight, category) =>
  stagesFor(category).filter((m) => n(flight[m.key]) === MISSION_PTS);

const flightTotal = (round, i) => calcScores({ ...round, ...round.flights[i] }).totalScore;

function JudgeView({ records, connected, onGoToHQ, roster = [], claims = {}, rosterLoaded = false }) {
  const sigRef = useRef(null);
  const flightRefs = [useRef(null), useRef(null)];
  const signRef = useRef(null);
  const [toast, showToast] = useToast();
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [round, setRound] = useState(loadRound);
  const roundRef = useRef(round);
  roundRef.current = round;

  // 라운드 진행 상태를 기기에 저장 → 새로고침·화면 이동 후에도 이어서 채점
  useEffect(() => storage.set(ROUND_KEY, JSON.stringify(round)), [round]);
  useEffect(() => storage.set("idrcc_judge_name", round.judgeName), [round.judgeName]);
  useEffect(() => storage.set("idrcc_category", round.category), [round.category]);

  /** 라운드 공통 항목 수정 (부문이 바뀌면 그 부문에서 쓰지 않는 장애물 점수는 0) */
  const patchRound = useCallback((p) => {
    setRound((prev) => {
      const next = { ...prev, ...p };
      if (p.category && p.category !== prev.category) {
        next.flights = prev.flights.map((f) => {
          const g = { ...f };
          MISSIONS.forEach((m) => {
            if (!missionAllowed(m.key, next.category)) g[m.key] = 0;
          });
          return g;
        });
      }
      return next;
    });
  }, []);
  const patchFlight = useCallback((i, p) => {
    setRound((prev) => ({
      ...prev,
      flights: prev.flights.map((f, j) => (j === i ? { ...f, ...p } : f)),
    }));
  }, []);

  // ── 진행 방식 ──────────────────────────────────────────────
  // round: 라운드별 진행(전원 1라운드 → 전원 2라운드)
  // participant: 참가자별 연속 진행(1라운드 확정 → 같은 참가자 2라운드 바로 시작)
  const [flowMode, setFlowModeState] = useState(() =>
    storage.get("idcc_flow_mode") === "participant" ? "participant" : "round"
  );
  const setFlowMode = (v) => {
    setFlowModeState(v);
    storage.set("idcc_flow_mode", v);
  };
  const continuous = flowMode === "participant";

  // ── 참가자 명단 ───────────────────────────────────────────
  const [selecting, setSelecting] = useState(false);
  const [group, setGroupState] = useState(() => storage.get("idcc_roster_group") || "");
  const setGroup = (v) => {
    setGroupState(v);
    storage.set("idcc_roster_group", v);
  };
  const selectedId = round.rosterId;
  const selectedEntry = roster.find((r) => r.id === selectedId) || null;
  const selectedIdRef = useRef(null);
  selectedIdRef.current = selectedId;
  useEffect(() => () => releaseRosterEntry(selectedIdRef.current), []);

  // 선택한 참가자의 번호·이름을 직접 고치면 명단 연결을 끊음
  useEffect(() => {
    if (!selectedEntry) return;
    if (
      String(round.participantNumber).trim() !== String(selectedEntry.bib) ||
      String(round.participantName).trim() !== String(selectedEntry.name)
    ) {
      releaseRosterEntry(selectedEntry.id);
      patchRound({ rosterId: null, country: "", affiliation: "", grade: "", gender: "" });
    }
  }, [round.participantNumber, round.participantName, selectedEntry, patchRound]);

  // HQ에서 명단을 비우거나 이 참가자를 명단에서 뺀 경우:
  // 아직 비행 전이면 심판 화면도 비움, 비행 중·후면 입력은 남기고 명단 연결만 끊음
  useEffect(() => {
    const r = roundRef.current;
    if (!rosterLoaded || !r.rosterId || roster.some((x) => x.id === r.rosterId)) return;
    const started = r.flights.some((f) => f.status !== "ready") || sw.status === "running";
    if (started) {
      patchRound({ rosterId: null });
      showToast("이 참가자가 HQ 명단에서 빠졌습니다. 입력 중인 라운드는 그대로 유지됩니다", "warn");
    } else {
      sigRef.current?.clear();
      setRound((prev) => freshRound(prev));
      showToast("HQ에서 명단이 바뀌어 참가자 입력을 비웠습니다", "warn");
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [roster, rosterLoaded, round.rosterId]);

  // ── 초시계 ────────────────────────────────────────────────
  const onTimerStop = (ms) => {
    setRound((prev) => {
      const i = prev.active;
      if (i > 1) return prev;
      return {
        ...prev,
        flights: prev.flights.map((f, j) => (j === i ? { ...f, ...msToTimeFields(ms), status: "stopped" } : f)),
        active: i + 1,
      };
    });
  };
  const sw = useStopwatch(onTimerStop);
  const running = sw.status === "running";

  // 새로고침 후 초시계 상태가 없는데 '측정 중'으로 남은 비행은 START 대기로 되돌림
  useEffect(() => {
    const i = roundRef.current.active;
    if (i <= 1 && roundRef.current.flights[i]?.status === "running" && sw.status !== "running")
      patchFlight(i, { status: "ready" });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const scrollToEl = (el) => {
    try {
      el?.scrollIntoView({ behavior: "smooth", block: "start" });
    } catch {
      /* ignore */
    }
  };

  /** START — 현재 차례의 비행 측정 시작 */
  const startFlight = () => {
    const r = roundRef.current;
    if (r.active > 1 || sw.status === "running") return;
    if (!String(r.judgeName).trim() || !String(r.participantNumber).trim() || !String(r.participantName).trim()) {
      showToast("Judge Name · 참가번호 · 이름을 먼저 입력하세요", "error");
      window.scrollTo({ top: 0, behavior: "smooth" });
      return;
    }
    sw.start();
    patchFlight(r.active, { status: "running", noFlight: false });
  };
  const stopFlight = () => sw.stop();

  /** 실수로 STOP을 눌렀을 때: 같은 비행을 처음 시작 시각부터 이어서 측정 */
  const undoStop = () => {
    const r = roundRef.current;
    if (sw.status !== "stopped" || r.active < 1) return;
    const i = r.active - 1;
    if (r.flights[i].noFlight) return;
    sw.undo();
    setRound((prev) => ({
      ...prev,
      active: i,
      flights: prev.flights.map((f, j) => (j === i ? { ...f, status: "running" } : f)),
    }));
  };

  // PC: 스페이스바로 START/STOP
  const keyRef = useRef(null);
  keyRef.current = () => (running ? stopFlight() : startFlight());
  useEffect(() => {
    const onKey = (e) => {
      if (e.code !== "Space" || e.repeat) return;
      const t = e.target;
      const tag = String(t?.tagName || "").toLowerCase();
      if (["input", "textarea", "select", "button", "canvas"].includes(tag) || t?.isContentEditable) return;
      e.preventDefault();
      keyRef.current();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  // 비행이 끝나면 다음 칸으로 화면 이동
  const prevActive = useRef(round.active);
  useEffect(() => {
    if (round.active !== prevActive.current) {
      if (round.active === 1) scrollToEl(flightRefs[1].current);
      if (round.active === 2) scrollToEl(signRef.current);
    }
    prevActive.current = round.active;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [round.active]);

  // ── 1차 비행 자동 저장 · 이후 수정 자동 반영 ─────────────────
  const f1 = round.flights[0];
  const f1Rec = useMemo(() => (f1.status === "ready" || f1.status === "running" ? null : flightRecord(round, 0)), [round, f1.status]);
  const f1Json = f1Rec ? JSON.stringify(f1Rec) : "";
  const syncedRef = useRef({ id: null, json: "" });

  useEffect(() => {
    if (!db || !f1Rec || f1.status !== "stopped") return;
    // 처음 저장 (또는 Undo 후 다시 STOP)
    const key = f1.recordId || push(ref(db, "scores")).key;
    const savedAt = f1.savedAt || Date.now();
    const rec = { ...f1Rec, roundId: key, timestamp: savedAt };
    syncedRef.current = { id: key, json: JSON.stringify(flightRecord({ ...round, flights: [{ ...f1, recordId: key }, round.flights[1]] }, 0)) };
    patchFlight(0, { status: "saved", recordId: key, savedAt });
    const offline = connected === false;
    withTimeout(update(ref(db), { [`scores/${key}`]: rec }))
      .then(() => {
        if (!offline) showToast(`1차 비행 기록 저장 · ${rec.totalScore} pts → 2차 비행 START 대기`);
      })
      .catch((e) => {
        console.error("Flight 1 save failed:", e);
        showToast(offline ? "오프라인 — 연결되면 1차 기록이 자동 업로드됩니다" : "1차 기록 저장 실패 — 연결을 확인하세요", "warn");
      });
    if (offline) showToast("오프라인 — 연결되면 1차 기록이 자동 업로드됩니다", "warn");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [f1.status]);

  // 저장된 1차 기록을 고치면(단계·시간·코딩·참가자 정보) HQ 기록도 0.6초 뒤 자동 수정
  useEffect(() => {
    if (!db || f1.status !== "saved" || !f1.recordId || !f1Rec) return undefined;
    if (syncedRef.current.id === f1.recordId && syncedRef.current.json === f1Json) return undefined;
    const t = setTimeout(() => {
      update(ref(db, `scores/${f1.recordId}`), { ...f1Rec, roundId: f1.recordId, editedAt: Date.now() })
        .then(() => (syncedRef.current = { id: f1.recordId, json: f1Json }))
        .catch((e) => console.error("Flight 1 update failed:", e));
    }, 600);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [f1Json, f1.status, f1.recordId]);

  // ── 명단에서 참가자 선택 ───────────────────────────────────
  const roundRecordsOf = (bib, trial) =>
    records.filter((r) => String(r.participantNumber ?? "").trim() === String(bib).trim() && r.trial === trial);

  const isDirty = () => {
    const r = roundRef.current;
    return (
      r.flights.some((f) => f.status !== "ready") ||
      n(r.coding) > 0 ||
      sw.status === "running" ||
      (sigRef.current && !sigRef.current.isEmpty())
    );
  };

  const selectParticipant = async (entry) => {
    if (selecting) return;
    const r = roundRef.current;
    if (!String(r.judgeName).trim()) {
      showToast("먼저 Judge Name(심판 이름)을 입력하세요", "warn");
      return;
    }
    if (running) {
      showToast("비행 측정 중에는 참가자를 바꿀 수 없습니다. STOP을 먼저 누르세요", "error");
      return;
    }
    const claim = claims[entry.id];
    if (claimedByOther(claim) && !window.confirm(`${claim.judge || "다른"} 심판이 이 참가자를 채점 중입니다.\n그래도 선택할까요?`))
      return;

    // 참가자별 연속 진행이면 이 참가자가 아직 마치지 않은 첫 라운드부터
    const targetTrial = continuous ? nextRoundFor(entry.bib, records)?.trial || r.trial : r.trial;
    const prevRecs = roundRecordsOf(entry.bib, targetTrial);
    const has1 = prevRecs.find((x) => flightOf(x) === 1);
    const has2 = prevRecs.find((x) => flightOf(x) === 2);
    if (has2 && !window.confirm(`#${entry.bib} ${entry.name} 은(는) 이미 ${ROUND_LABEL[targetTrial]} 기록이 확정되었습니다.\n다시 채점할까요? (새 기록이 추가됩니다)`))
      return;
    const resume = has1 && !has2; // 1차만 저장되고 끊긴 라운드 → 2차부터 이어서
    if (isDirty() && !window.confirm("지금 입력 중인 라운드 내용(초시계·서명 포함)이 화면에서 지워집니다.\n이미 저장된 1차 기록은 HQ에 남습니다. 참가자를 바꿀까요?"))
      return;

    setSelecting(true);
    let claimed = false;
    if (connected) {
      try {
        await claimRosterEntry(entry.id, r.judgeName);
        claimed = true;
      } catch (e) {
        console.error("Claim failed:", e);
      }
    }
    setSelecting(false);
    if (!claimed) showToast("오프라인이라 '채점 중' 표시를 하지 못했습니다. 채점은 계속할 수 있습니다", "warn");
    if (selectedId && selectedId !== entry.id) releaseRosterEntry(selectedId);

    sw.reset();
    sigRef.current?.clear();
    const base = {
      ...freshRound(r),
      trial: targetTrial,
      participantNumber: String(entry.bib),
      participantName: String(entry.name),
      category: CATEGORIES.includes(entry.category) ? entry.category : r.category,
      rosterId: entry.id,
      ...Object.fromEntries(PROFILE_FIELDS.map((k) => [k, entry[k] || ""])),
    };
    if (resume) {
      const f = { ...emptyFlight(), status: "saved", recordId: has1.id, savedAt: n(has1.timestamp) || Date.now() };
      FLIGHT_FIELDS.forEach((k) => (f[k] = k === "isDQ" ? !!has1[k] : n(has1[k])));
      base.flights = [f, emptyFlight()];
      base.coding = n(has1.coding);
      base.codingDone = true;
      base.readyAck = true;
      base.active = 1;
      syncedRef.current = { id: has1.id, json: JSON.stringify(flightRecord(base, 0)) };
      showToast(`1차 비행 기록을 불러왔습니다 → ${ROUND_LABEL[targetTrial]} 2차 비행부터 이어서 진행`, "warn");
    }
    setRound(base);
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  /** 입력 지우기 — 참가자·코딩·비행 입력을 모두 비움 (심판 이름·부문·라운드는 유지) */
  const clearInputs = () => {
    const r = roundRef.current;
    if (running) return showToast("비행 측정 중에는 지울 수 없습니다. STOP을 먼저 누르세요", "error");
    const saved1 = r.flights[0].recordId && r.flights[0].status === "saved";
    if (
      !window.confirm(
        "심판 화면의 참가자·코딩·비행 입력과 서명을 모두 지울까요?" +
          (saved1 ? "\n\n이미 자동 저장된 1차 비행 기록은 HQ에 그대로 남습니다. (지우려면 HQ에서 삭제)" : "")
      )
    )
      return;
    if (r.rosterId) releaseRosterEntry(r.rosterId);
    sw.reset();
    sigRef.current?.clear();
    syncedRef.current = { id: null, json: "" };
    setRound((prev) => freshRound(prev));
    window.scrollTo({ top: 0, behavior: "smooth" });
    showToast("입력을 지웠습니다", "warn");
  };

  /** 라운드를 처음부터 다시 (저장된 1차 기록 삭제 여부 선택) */
  const restartRound = async () => {
    const r = roundRef.current;
    if (running) return showToast("측정 중에는 다시 시작할 수 없습니다. STOP을 먼저 누르세요", "error");
    if (!window.confirm("이 라운드의 비행 입력을 처음부터 다시 할까요?\n(참가자·코딩 점수는 유지됩니다)")) return;
    const id = r.flights[0].recordId;
    if (id && window.confirm("HQ에 저장된 1차 비행 기록도 삭제할까요?\n[확인] 삭제 · [취소] 남겨 둠")) {
      try {
        await withTimeout(update(ref(db), { [`scores/${id}`]: null }));
      } catch (e) {
        alert("1차 기록 삭제 실패.\n" + (e?.message || e));
      }
    }
    sw.reset();
    sigRef.current?.clear();
    setRound((prev) => ({ ...prev, flights: [emptyFlight(), emptyFlight()], active: 0 }));
  };

  /** 2차 비행을 하지 않은 경우 (시간 초과 등) */
  const skipFlight2 = () => {
    if (running || round.active !== 1) return;
    if (!window.confirm("2차 비행을 하지 않은 것으로 기록하고 학생 확인·서명으로 넘어갈까요?")) return;
    setRound((prev) => ({
      ...prev,
      active: 2,
      flights: prev.flights.map((f, j) => (j === 1 ? { ...emptyFlight(), noFlight: true, status: "stopped" } : f)),
    }));
  };

  // ── 학생 서명 · 라운드 확정 ────────────────────────────────
  const handleSubmit = async () => {
    if (!db) return alert("Database is not connected. Check the Firebase configuration.");
    const r = roundRef.current;
    if (running) return alert("비행 측정 중입니다. STOP을 먼저 누르세요.");
    if (r.active < 2) return alert("1차·2차 비행이 모두 끝난 뒤 확정할 수 있습니다.");
    if (!sigRef.current || sigRef.current.isEmpty()) {
      alert("학생 서명이 필요합니다.");
      scrollToEl(signRef.current);
      return;
    }
    const offline = connected === false;
    if (
      offline &&
      !window.confirm(
        "OFFLINE 상태입니다.\n\n기록은 이 기기에 보관되었다가 연결되면 자동 업로드됩니다.\nOnline 표시가 나올 때까지 이 화면을 닫거나 새로고침하지 마세요.\n\n계속할까요?"
      )
    )
      return;

    setIsSubmitting(true);
    try {
      const key1 = r.flights[0].recordId || push(ref(db, "scores")).key;
      const key2 = push(ref(db, "scores")).key;
      const roundId = key1;
      const withRound = { ...r, flights: [{ ...r.flights[0], recordId: key1 }, r.flights[1]] };
      const rec1 = flightRecord(withRound, 0);
      const rec2 = flightRecord(withRound, 1);
      const signature = sigRef.current.toDataURL();
      const writing = update(ref(db), {
        [`scores/${key1}`]: {
          ...rec1,
          roundId,
          signatureId: roundId,
          hasSignature: true,
          timestamp: r.flights[0].savedAt || Date.now(),
        },
        [`scores/${key2}`]: { ...rec2, roundId, signatureId: roundId, hasSignature: true, timestamp: Date.now() },
        [`signatures/${roundId}`]: signature,
      });
      if (offline) writing.catch((e) => console.error("Queued write failed:", e));
      else await withTimeout(writing);

      const best = Math.max(rec1.isDQ ? -1 : rec1.totalScore, rec2.isDQ || rec2.noFlight ? -1 : rec2.totalScore);
      showToast(
        offline
          ? "오프라인 저장 — 연결되면 자동 업로드됩니다"
          : `${ROUND_LABEL[r.trial]} 확정 · #${r.participantNumber} · 최고 ${Math.max(0, best)} pts`,
        offline ? "warn" : "success"
      );
      const ri = TRIALS.indexOf(r.trial);
      const goNextRound = continuous && ri >= 0 && ri < TRIALS.length - 1;
      sw.reset();
      sigRef.current.clear();
      syncedRef.current = { id: null, json: "" };
      if (goNextRound) {
        // 같은 참가자의 다음 라운드: 참가자·명단 연결·코딩 점수는 그대로, 비행·서명만 새로
        const nextTrial = TRIALS[ri + 1];
        setRound((prev) => ({
          ...freshRound(prev),
          trial: nextTrial,
          participantNumber: prev.participantNumber,
          participantName: prev.participantName,
          rosterId: prev.rosterId,
          coding: prev.coding,
          ...Object.fromEntries(PROFILE_FIELDS.map((k) => [k, prev[k] || ""])),
        }));
        if (!offline) showToast(`${ROUND_LABEL[r.trial]} 확정 → 같은 참가자 ${ROUND_LABEL[nextTrial]} 바로 진행 (코딩 점수 확인)`);
      } else {
        if (r.rosterId) releaseRosterEntry(r.rosterId);
        // 참가자별 연속 진행이면 다음 참가자는 1라운드부터
        setRound((prev) => ({ ...freshRound(prev), trial: continuous ? TRIALS[0] : prev.trial }));
      }
      window.scrollTo({ top: 0, behavior: "smooth" });
    } catch (e) {
      console.error(e);
      alert("저장 실패. 인터넷 연결을 확인하고 다시 시도하세요.\n\n" + (e?.message || e));
    } finally {
      setIsSubmitting(false);
    }
  };

  const totals = [flightTotal(round, 0), flightTotal(round, 1)];

  return (
    <div className="max-w-4xl mx-auto px-3 md:px-6 pt-5 md:pt-8 pb-56 md:pb-44">
      <Toast toast={toast} />

      <div className="flex items-end justify-between gap-3 mb-5">
        <div>
          <div className="text-xs font-black uppercase tracking-widest text-sky-600">Judge Console</div>
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
        {(roster.length > 0 || rosterLoaded) && (
          <RosterPicker
            roster={roster}
            claims={claims}
            records={records}
            trial={round.trial}
            flight={Math.min(2, round.active + 1)}
            selectedId={selectedId}
            group={group}
            setGroup={setGroup}
            onSelect={selectParticipant}
            busy={selecting}
            continuous={continuous}
          />
        )}

        <ParticipantFields
          form={round}
          patch={patchRound}
          roundMode
          onClear={clearInputs}
          flowMode={flowMode}
          onFlowMode={setFlowMode}
        />

        <Section step={2} title="Coding" icon={<Code size={18} />} tone="indigo"
          subtitle="라운드마다 1회 입력 · 이 라운드의 1·2차 비행 기록에 함께 들어갑니다"
          right={<ScoreChip tone="indigo">{n(round.coding)} / {CODING_MAX}</ScoreChip>}>
          <CodingInput value={round.coding} onChange={(v) => patchRound({ coding: v, codingDone: true })} />
        </Section>

        {round.codingDone && round.active === 0 && f1.status === "ready" && (
          <ReadyAlert acknowledged={round.readyAck} onAcknowledge={() => patchRound({ readyAck: true })} />
        )}

        {[0, 1].map((i) => (
          <div key={i} ref={flightRefs[i]} style={{ scrollMarginTop: "5.5rem" }}>
            <FlightPanel
              index={i}
              round={round}
              flight={round.flights[i]}
              isCurrent={round.active === i}
              locked={round.active < i}
              running={running && round.active === i}
              liveMs={sw.elapsed}
              total={totals[i]}
              onPatch={(p) => patchFlight(i, p)}
              onSkip={i === 1 ? skipFlight2 : undefined}
            />
          </div>
        ))}

        <div ref={signRef} style={{ scrollMarginTop: "5.5rem" }}>
          <RoundSignSection
            round={round}
            totals={totals}
            ready={round.active === 2}
            sigRef={sigRef}
            onRestart={restartRound}
          />
        </div>

        <RecentSubmissions records={records} judgeName={round.judgeName} />
      </div>

      <FlightBar
        round={round}
        sw={sw}
        totals={totals}
        isSubmitting={isSubmitting}
        onStart={startFlight}
        onStop={stopFlight}
        onUndo={undoStop}
        onSubmit={handleSubmit}
      />
    </div>
  );
}

/** 비행 1회 입력 칸 — 단계 버튼은 '마지막으로 통과한 단계'를 누르면 앞 단계까지 자동 포함 */
function FlightPanel({ index, round, flight, isCurrent, locked, running, liveMs, total, onPatch, onSkip }) {
  const stages = stagesFor(round.category);
  const cleared = clearedStages(flight, round.category);
  const used = landingCount(flight);
  const calc = calcScores({ ...round, ...flight });
  const landingPts = calc.lndBlueScore + calc.lndRedScore + calc.lndYellowScore;
  const editable = !locked && !flight.noFlight;

  /** 단계 하나만 켜고 끔 (다른 단계에는 영향 없음) */
  const toggleStage = (key) => onPatch({ [key]: n(flight[key]) === MISSION_PTS ? 0 : MISSION_PTS });

  const statusChip = flight.noFlight
    ? ["bg-slate-200 text-slate-600", "비행 안 함"]
    : flight.status === "running"
    ? ["bg-rose-500 text-white animate-pulse", "측정 중"]
    : flight.status === "saved"
    ? ["bg-emerald-100 text-emerald-800", "자동 저장됨 · HQ 표시"]
    : flight.status === "stopped"
    ? ["bg-amber-100 text-amber-900", index === 0 ? "저장 중…" : "종료 · 서명 대기"]
    : isCurrent
    ? ["bg-sky-100 text-sky-800", "START 대기"]
    : ["bg-slate-100 text-slate-500", "대기"];

  return (
    <section
      className={`rounded-3xl border-2 shadow-sm p-4 md:p-6 transition ${
        running
          ? "bg-rose-50/60 border-rose-400 shadow-rose-100"
          : isCurrent && !locked
          ? "bg-white border-sky-300"
          : "bg-white border-slate-200/80"
      } ${locked ? "opacity-50" : ""}`}
    >
      <div className="flex flex-wrap items-center gap-3 mb-4">
        <div
          className={`w-11 h-11 rounded-2xl grid place-items-center font-black text-white text-lg shrink-0 ${
            index === 0 ? "bg-emerald-500" : "bg-teal-500"
          }`}
        >
          {index + 1}차
        </div>
        <div className="min-w-0 flex-1">
          <h3 className="font-black text-xl leading-tight">
            {ROUND_LABEL[round.trial]} · {index + 1}차 비행
          </h3>
          <p className="text-sm text-slate-500 font-semibold">
            {index === 0 ? "심판 신호로 시작합니다" : "참가자가 신호 없이 바로 시작합니다 — 출발하면 바로 START"}
          </p>
        </div>
        <span className={`px-3 py-1.5 rounded-xl text-sm font-black whitespace-nowrap ${statusChip[0]}`}>{statusChip[1]}</span>
        <ScoreChip tone="emerald">{flight.noFlight || flight.status === "ready" ? "—" : `${total} pts`}</ScoreChip>
      </div>

      {locked ? (
        <p className="text-base font-bold text-slate-500 py-3">1차 비행이 끝나면 활성화됩니다.</p>
      ) : flight.noFlight ? (
        <div className="flex flex-wrap items-center gap-3 py-2">
          <p className="text-base font-bold text-slate-600">2차 비행을 하지 않은 것으로 기록됩니다.</p>
          <button
            type="button"
            onClick={() => onPatch({ noFlight: false, status: "ready" })}
            className="h-10 px-3 rounded-xl border-2 border-slate-200 text-sm font-bold text-slate-600 hover:border-slate-300"
          >
            취소
          </button>
        </div>
      ) : (
        <>
          {/* 비행 단계 */}
          <div className="flex items-baseline justify-between gap-2 mb-2">
            <FieldLabel>비행 임무 — 성공한 단계를 각각 누르세요 (다시 누르면 취소)</FieldLabel>
            <span className="text-sm font-black text-emerald-700 whitespace-nowrap">
              {cleared.length}개 성공 · {cleared.length * MISSION_PTS}점
            </span>
          </div>
          <div className={`grid gap-2 ${stages.length > 4 ? "grid-cols-3 sm:grid-cols-6" : "grid-cols-2 sm:grid-cols-4"}`}>
            {stages.map((m, i) => {
              const on = n(flight[m.key]) === MISSION_PTS;
              return (
                <button
                  key={m.key}
                  type="button"
                  disabled={!editable}
                  onClick={() => toggleStage(m.key)}
                  aria-pressed={on}
                  className={`relative min-h-[78px] rounded-2xl border-2 px-2 py-2 text-left transition active:scale-[0.97] ${
                    on
                      ? "bg-emerald-500 border-emerald-500 text-white"
                      : "bg-white border-slate-200 text-slate-700 hover:border-emerald-300"
                  }`}
                >
                  <span className={`block text-xs font-black ${on ? "text-emerald-100" : "text-slate-400"}`}>
                    {i + 1}단계 · +{MISSION_PTS}
                  </span>
                  <span className="block font-black leading-tight mt-1">{m.label}</span>
                  {on && <Check size={18} strokeWidth={3} className="absolute top-2 right-2" />}
                </button>
              );
            })}
          </div>

          {/* 착륙 */}
          <div className="mt-4 flex items-center justify-between gap-2">
            <FieldLabel>Landing Zones (최대 {MAX_LANDINGS}회)</FieldLabel>
            <span className="text-sm font-black text-amber-700">{landingPts}점</span>
          </div>
          <div className="grid grid-cols-3 gap-2">
            {LANDING_ZONES.map((z) => {
              const c = n(flight[z.key]);
              return (
                <div key={z.key} className={`rounded-2xl ${z.soft} ring-1 ${z.ring} p-2 flex items-center justify-between gap-1`}>
                  <button
                    type="button"
                    disabled={!editable || c <= 0}
                    onClick={() => onPatch({ [z.key]: c - 1 })}
                    className="w-11 h-11 rounded-xl bg-white border-2 border-slate-200 grid place-items-center disabled:opacity-30"
                    aria-label={`Remove ${z.name}`}
                  >
                    <Minus size={18} />
                  </button>
                  <div className="text-center leading-tight">
                    <div className={`text-xs font-black ${z.text}`}>
                      {z.name} {z.pts}
                    </div>
                    <div className="text-2xl font-black tabular-nums">{c}</div>
                  </div>
                  <button
                    type="button"
                    disabled={!editable || used >= MAX_LANDINGS}
                    onClick={() => onPatch({ [z.key]: c + 1 })}
                    className="w-11 h-11 rounded-xl bg-white border-2 border-slate-200 grid place-items-center disabled:opacity-30"
                    aria-label={`Add ${z.name}`}
                  >
                    <Plus size={18} />
                  </button>
                </div>
              );
            })}
          </div>

          {/* 시간 · DQ */}
          <div className="mt-4 flex flex-col md:flex-row md:items-end gap-4">
            {running ? (
              <div>
                <FieldLabel>Lap Time</FieldLabel>
                <div className="h-16 px-4 rounded-2xl bg-rose-500 text-white font-mono text-3xl font-black tabular-nums grid place-items-center">
                  {formatMs(liveMs)}
                </div>
              </div>
            ) : (
              <TimeInput form={flight} patch={onPatch} />
            )}
            <DQToggle on={!!flight.isDQ} onToggle={() => onPatch({ isDQ: !flight.isDQ })} />
            {onSkip && isCurrent && !running && (
              <button
                type="button"
                onClick={onSkip}
                className="md:mb-5 h-16 px-4 rounded-2xl border-2 border-slate-200 bg-white text-slate-600 font-bold hover:border-slate-300"
              >
                2차 비행 안 함
              </button>
            )}
          </div>
        </>
      )}
    </section>
  );
}

/** 라운드 요약(1·2차) + 학생 서명 */
function RoundSignSection({ round, totals, ready, sigRef, onRestart }) {
  const rows = round.flights.map((f, i) => {
    const c = clearedStages(f, round.category);
    const notFlown = f.noFlight || f.status === "ready";
    return {
      i,
      f,
      stage: f.noFlight
        ? "비행 안 함"
        : f.status === "ready"
        ? "아직 비행 전"
        : c.length
        ? `${c.length}개 · ${c.map((m) => m.label).join(", ")}`
        : "—",
      landing: notFlown ? "—" : `${landingCount(f)}회`,
      time: notFlown ? "—" : formatTime(f),
      total: notFlown ? null : totals[i],
    };
  });
  const valid = rows.filter((r) => r.total !== null && !r.f.isDQ);
  const best = valid.length
    ? valid.reduce((a, b) => (b.total > a.total || (b.total === a.total && timeToCs(b.f) < timeToCs(a.f)) ? b : a))
    : null;

  return (
    <section className={`rounded-3xl border-2 p-4 md:p-6 shadow-sm ${ready ? "bg-white border-violet-300" : "bg-white border-slate-200/80"}`}>
      <div className="flex flex-wrap items-center gap-3 mb-4">
        <div className="w-11 h-11 rounded-2xl bg-violet-500 text-white grid place-items-center shrink-0">
          <PenLine size={20} />
        </div>
        <div className="flex-1 min-w-0">
          <h3 className="font-black text-xl leading-tight">{ROUND_LABEL[round.trial]} 기록 확인 · 학생 서명</h3>
          <p className="text-sm text-slate-500 font-semibold">
            2차 비행이 끝나면 학생이 1·2차 기록을 확인하고 서명합니다. 확정하면 2차 기록이 HQ에 올라갑니다.
          </p>
        </div>
        <button
          type="button"
          onClick={onRestart}
          className="h-10 px-3 rounded-xl border-2 border-slate-200 text-sm font-bold text-slate-500 hover:text-rose-600 hover:border-rose-200 flex items-center gap-1.5"
        >
          <RotateCcw size={15} /> 라운드 다시
        </button>
      </div>

      <div className="overflow-x-auto rounded-2xl border border-slate-200">
        <table className="w-full min-w-[520px] text-left">
          <thead className="bg-slate-50 text-xs font-black uppercase tracking-wider text-slate-500">
            <tr>
              <th className="p-3">비행</th>
              <th className="p-3">성공 단계</th>
              <th className="p-3 text-center">착륙</th>
              <th className="p-3 text-center">시간</th>
              <th className="p-3 text-center">총점</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => {
              const isBest = best && best.i === r.i;
              return (
                <tr key={r.i} className={`border-t border-slate-100 ${isBest ? "bg-amber-50" : ""}`}>
                  <td className="p-3 font-black whitespace-nowrap">
                    {r.i + 1}차 {isBest && <span className="ml-1 text-amber-600">★ 최고</span>}
                  </td>
                  <td className="p-3 font-bold">{r.stage}</td>
                  <td className="p-3 text-center font-bold tabular-nums">{r.landing}</td>
                  <td className="p-3 text-center font-mono font-bold">{r.time}</td>
                  <td className="p-3 text-center">
                    {r.f.isDQ ? (
                      <span className="px-2 py-1 rounded-lg bg-rose-100 text-rose-700 text-sm font-black">DQ</span>
                    ) : (
                      <span className="text-2xl font-black text-blue-600 tabular-nums">{r.total ?? "—"}</span>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div className={`mt-4 ${ready ? "" : "pointer-events-none opacity-40"}`}>
        <FieldLabel>학생 서명 {ready ? "" : "— 2차 비행이 끝나면 서명할 수 있습니다"}</FieldLabel>
        <SignatureBox ref={sigRef} height={190} />
      </div>
    </section>
  );
}

/** 하단 고정 바 — 현재 비행의 START/STOP을 크게, 서명 단계에서는 확정 버튼 */
function FlightBar({ round, sw, totals, isSubmitting, onStart, onStop, onUndo, onSubmit }) {
  const running = sw.status === "running";
  const active = round.active;
  const label = active <= 1 ? `${active + 1}차 비행` : "서명 · 확정";
  const armed = !running && active <= 1 && (active === 1 || round.readyAck);
  const canUndo = !running && sw.status === "stopped" && active >= 1 && !round.flights[active - 1]?.noFlight;
  const shownTime = running
    ? formatMs(sw.elapsed)
    : active <= 1
    ? formatTime(round.flights[active])
    : formatTime(round.flights[1].noFlight ? round.flights[0] : round.flights[1]);
  const bestTotal = Math.max(
    ...round.flights.map((f, i) => (f.status === "ready" || f.noFlight || f.isDQ ? 0 : totals[i]))
  );

  return (
    <div
      className="fixed bottom-0 inset-x-0 z-40 bg-white/95 backdrop-blur border-t border-slate-200 shadow-[0_-10px_30px_-15px_rgba(15,23,42,0.3)]"
      style={{ paddingBottom: "env(safe-area-inset-bottom)" }}
    >
      <div className="max-w-4xl mx-auto px-3 md:px-6 py-2.5 md:py-3 flex flex-col md:flex-row md:items-center gap-2.5 md:gap-4">
        <div className="flex items-center gap-2 md:gap-3 flex-1 min-w-0">
          <div
            className={`h-16 px-3 md:px-4 rounded-2xl ring-1 flex items-center gap-2.5 min-w-[10rem] ${
              running ? "bg-rose-50 ring-rose-200" : "bg-slate-50 ring-slate-200"
            }`}
          >
            <span className={`w-2.5 h-2.5 rounded-full shrink-0 ${running ? "bg-rose-500 animate-pulse" : "bg-slate-300"}`} />
            <div className="leading-none">
              <div className="text-[11px] font-black uppercase tracking-widest text-slate-500 mb-1">
                {ROUND_LABEL[round.trial]} · {label}
              </div>
              <div className={`font-mono text-2xl md:text-3xl font-black tabular-nums ${running ? "text-rose-600" : "text-slate-700"}`}>
                {shownTime}
              </div>
            </div>
          </div>

          {running ? (
            <button
              type="button"
              onClick={onStop}
              className="flex-1 h-16 px-4 rounded-2xl bg-rose-600 hover:bg-rose-700 text-white text-xl md:text-2xl font-black flex items-center justify-center gap-2 shadow-lg shadow-rose-300 active:scale-95"
            >
              <Square size={22} fill="currentColor" /> {active + 1}차 STOP
            </button>
          ) : active <= 1 ? (
            <button
              type="button"
              onClick={onStart}
              className={`flex-1 h-16 px-4 rounded-2xl bg-emerald-500 hover:bg-emerald-600 text-white text-xl md:text-2xl font-black flex items-center justify-center gap-2 shadow-lg shadow-emerald-200 active:scale-95 ${
                armed ? "animate-blink ring-4 ring-red-500 ring-offset-2" : ""
              }`}
            >
              <Play size={22} fill="currentColor" /> {active + 1}차 START
            </button>
          ) : null}

          {canUndo && (
            <button
              type="button"
              onClick={onUndo}
              title="STOP을 잘못 눌렀다면, 처음 시작 시각부터 이어서 측정합니다"
              className="h-16 px-3 rounded-2xl bg-white border-2 border-slate-200 text-slate-600 text-sm font-bold flex flex-col items-center justify-center leading-tight"
            >
              <Undo2 size={18} /> {active}차 계속
            </button>
          )}
        </div>

        {active === 2 && (
          <div className="flex items-center gap-3">
            <div className="shrink-0">
              <div className="text-[11px] font-black uppercase tracking-widest text-slate-400">라운드 최고</div>
              <div className="text-3xl font-black text-blue-600 tabular-nums leading-none">
                {bestTotal}
                <span className="text-sm text-slate-400 ml-1">pts</span>
              </div>
            </div>
            <button
              onClick={onSubmit}
              disabled={isSubmitting}
              className="flex-1 md:min-w-[240px] h-16 px-4 rounded-2xl text-white text-lg font-black flex items-center justify-center gap-2 shadow-lg transition active:scale-[0.98] bg-gradient-to-r from-violet-500 to-blue-500 hover:from-violet-600 hover:to-blue-600 shadow-violet-200 disabled:from-slate-300 disabled:to-slate-300"
            >
              {isSubmitting ? <Loader2 size={22} className="animate-spin" /> : <CheckCircle2 size={22} />}
              서명 완료 · 라운드 확정
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

function ParticipantFields({ form, patch, showJudge = true, roundMode = false, onClear, flowMode, onFlowMode }) {
  return (
    <Section
      step={1}
      title="Participant"
      icon={<Users size={18} />}
      tone="sky"
      right={
        onClear && (
          <button
            type="button"
            onClick={onClear}
            className="shrink-0 h-10 px-3 rounded-xl border-2 border-rose-200 text-rose-600 text-sm font-bold flex items-center gap-1.5 hover:bg-rose-50"
          >
            <Trash2 size={15} /> 입력 지우기
          </button>
        )
      }
    >
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
      <div className="grid gap-4 mt-4">
        <div>
          <FieldLabel>Category</FieldLabel>
          <Segmented
            options={CATEGORIES}
            value={form.category}
            onChange={(v) => patch({ category: v })}
            cols="grid-cols-2 sm:grid-cols-4"
          />
        </div>
        {roundMode ? (
          <>
            {onFlowMode && (
              <div>
                <FieldLabel>진행 방식</FieldLabel>
                <Segmented
                  options={FLOW_OPTIONS.map((o) => o.label)}
                  value={(FLOW_OPTIONS.find((o) => o.key === flowMode) || FLOW_OPTIONS[0]).label}
                  onChange={(v) => onFlowMode(FLOW_OPTIONS.find((o) => o.label === v)?.key || "round")}
                  cols="grid-cols-1 sm:grid-cols-2"
                />
                <p className="mt-1.5 ml-1 text-xs font-semibold text-slate-500">
                  {flowMode === "participant"
                    ? "1라운드를 확정하면 같은 참가자의 2라운드로 바로 넘어갑니다. 2라운드까지 끝나면 다음 참가자."
                    : "전원이 1라운드를 마친 뒤 2라운드를 진행합니다. 확정하면 다음 참가자로 넘어갑니다."}
                </p>
              </div>
            )}
            <div>
              <FieldLabel>Round (라운드마다 {FLIGHTS_PER_TRIAL}회 비행)</FieldLabel>
              <Segmented
                options={ROUND_OPTIONS}
                value={ROUND_LABEL[form.trial]}
                onChange={(v) => patch({ trial: TRIALS[ROUND_OPTIONS.indexOf(v)] || "1st" })}
                cols="grid-cols-2"
              />
            </div>
          </>
        ) : (
          <div>
            <FieldLabel>Round · Flight</FieldLabel>
            <Segmented
              options={ATTEMPTS}
              value={attemptLabel(form)}
              onChange={(v) => patch(parseAttempt(v))}
              cols="grid-cols-4"
            />
          </div>
        )}
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
            <span className="text-slate-400 shrink-0">{attemptLabel(r)}</span>
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
function HQView({ records, legacySigs = {}, rosterState, onGoToJudge, onLock }) {
  const [editingId, setEditingId] = useState(null);
  const [viewSig, setViewSig] = useState(null); // 크게 볼 서명의 기록

  // 서명 이미지 구독 (HQ 화면에서만 내려받음 — 심판 기기 부담 없음)
  const [sigMap, setSigMap] = useState({});
  useEffect(() => {
    if (!db) return undefined;
    return onValue(
      ref(db, "signatures"),
      (snap) => setSigMap(snap.val() || {}),
      (err) => console.error("Signature sync error:", err)
    );
  }, []);
  /** 기록의 서명 — 같은 라운드의 1·2차 기록은 서명 1개(signatureId)를 함께 사용 */
  const sigKeyOf = (r) => r.signatureId || r.id;
  const sigOf = (r) => sigMap[sigKeyOf(r)] || legacySigs[r.id] || null;

  /** 기록 1건 삭제 (서명 포함) */
  const handleDelete = async (r) => {
    if (
      !window.confirm(
        `#${r.participantNumber} ${r.participantName} (${attemptLabel(r)}) 기록을 삭제할까요?\n서명도 함께 지워지며 되돌릴 수 없습니다.`
      )
    )
      return;
    try {
      // 서명은 같은 라운드의 다른 기록이 쓰지 않을 때만 함께 삭제
      const k = sigKeyOf(r);
      const shared = records.some((x) => x.id !== r.id && sigKeyOf(x) === k);
      const upd = { [`scores/${r.id}`]: null };
      if (!shared) upd[`signatures/${k}`] = null;
      await withTimeout(update(ref(db), upd));
      showToast("Record deleted", "warn");
    } catch (e) {
      alert("Delete failed.\n" + (e?.message || e));
    }
  };
  const [catFilter, setCatFilter] = useState("All");
  const [mode, setMode] = useState("best"); // best: 참가자별 4회 중 최고 기록 | all: 비행 기록 전체
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

  /** 비행 1회 기록 → 엑셀 한 줄 */
  const toFlightRow = (r) => ({
    Category: r.category,
    "No.": r.participantNumber,
    Participant: r.participantName,
    Country: r.country || "",
    Affiliation: r.affiliation || "",
    Grade: r.grade || "",
    Gender: r.gender || "",
    Round: ROUND_LABEL[r.trial] || r.trial,
    Flight: flightOf(r),
    "Round-Flight": attemptLabel(r),
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
    Total: r.noFlight ? "" : n(r.totalScore),
    Time: r.noFlight ? "" : formatTime(r),
    DQ: r.isDQ ? "YES" : "NO",
    "No Flight": r.noFlight ? "YES" : "",
    Judge: r.judgeName,
    Signature: sigOf(r) ? "Y" : "N",
    Submitted: r.timestamp ? new Date(r.timestamp).toLocaleString() : "",
    Edited: r.editedAt ? new Date(r.editedAt).toLocaleString() : "",
  });

  const downloadExcel = () => {
    if (!records.length) return alert("There are no records to export.");
    // 참가자별 4회 기록 (같은 비행이 여러 번 저장된 경우 가장 최근 것)
    const byPerson = new Map();
    records.forEach((r) => {
      const k = participantKey(r);
      const m = byPerson.get(k) || {};
      const a = attemptLabel(r);
      if (!m[a] || n(r.timestamp) > n(m[a].timestamp)) m[a] = r;
      byPerson.set(k, m);
    });
    const cell = (r) => (!r ? "" : r.noFlight ? "No flight" : r.isDQ ? "DQ" : n(r.totalScore));

    const finalRanked = rankByCategory(bestPerParticipant(records));
    const finalRows = finalRanked.map((b) => {
      const m = byPerson.get(participantKey(b)) || {};
      return {
        Rank: b.rank ?? (b.noFlight ? "-" : "DQ"),
        Category: b.category,
        "No.": b.participantNumber,
        Participant: b.participantName,
        Country: b.country || "",
        Affiliation: b.affiliation || "",
        ...Object.fromEntries(ATTEMPTS.map((a) => [`${a} Total`, cell(m[a])])),
        ...Object.fromEntries(ATTEMPTS.map((a) => [`${a} Time`, m[a] && !m[a].noFlight ? formatTime(m[a]) : ""])),
        "Final (Best)": b.noFlight ? "" : n(b.totalScore),
        "Best Time": b.noFlight ? "" : formatTime(b),
        "Best Flight": attemptLabel(b),
        "Flights Recorded": Object.keys(m).length,
      };
    });

    const allFlights = [...records].sort(
      (a, b) =>
        CATEGORIES.indexOf(a.category) - CATEGORIES.indexOf(b.category) ||
        String(a.participantNumber).localeCompare(String(b.participantNumber), undefined, { numeric: true }) ||
        ATTEMPTS.indexOf(attemptLabel(a)) - ATTEMPTS.indexOf(attemptLabel(b)) ||
        n(a.timestamp) - n(b.timestamp)
    );

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(finalRows), "Final (Best of 4)");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(allFlights.map(toFlightRow)), "All Flights");
    CATEGORIES.forEach((cat) => {
      const rows = finalRows.filter((r) => r.Category === cat);
      if (rows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), cat);
    });
    XLSX.writeFile(wb, `IDCC_DCC_Results_${localStamp()}.xlsx`);
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

      <RosterManager
        roster={rosterState?.roster || []}
        claims={rosterState?.claims || {}}
        records={records}
        error={rosterState?.error}
        onSave={saveRosterEntries}
        onClear={() => saveRosterEntries([])}
        showToast={showToast}
      />

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
              ["best", "최종 순위 (4회 중 최고)"],
              ["all", "전체 비행 기록"],
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
          <table className="w-full min-w-[1320px] text-left border-collapse">
            <thead className="bg-slate-50 text-slate-500 text-xs uppercase tracking-wider font-black border-b border-slate-200">
              <tr>
                <th className="p-3 text-center w-20">Rank</th>
                <th className="p-3">Category</th>
                <th className="p-3">No.</th>
                <th className="p-3">Participant</th>
                <th className="p-3 text-center">Trial-Flight</th>
                <th className="p-3 text-center">Coding</th>
                <th className="p-3 text-center">Missions</th>
                <th className="p-3 text-center">Landing</th>
                <th className="p-3 text-center bg-blue-50 text-blue-600">Total</th>
                <th className="p-3 text-center">Time</th>
                <th className="p-3">Judge</th>
                <th className="p-3 text-center">Signature</th>
                <th className="p-3 text-center w-28">Edit · Delete</th>
              </tr>
            </thead>
            <tbody>
              {visible.length === 0 && (
                <tr>
                  <td colSpan={13} className="p-12 text-center text-slate-400 font-semibold">
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
                        <td colSpan={13} className="px-4 py-2.5">
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
                        <RankBadge rank={r.rank} noFlight={!!r.noFlight && !r.isDQ} />
                      </td>
                      <td className="p-3">
                        <CategoryBadge category={r.category} />
                      </td>
                      <td className="p-3 font-mono font-bold text-slate-600">{r.participantNumber}</td>
                      <td className="p-3">
                        <div className="font-black text-slate-800 text-base">{r.participantName}</div>
                        {(r.affiliation || r.country) && (
                          <div className="text-xs font-semibold text-slate-400">
                            {[r.country, r.affiliation].filter(Boolean).join(" · ")}
                          </div>
                        )}
                      </td>
                      <td className="p-3 text-center text-sm font-bold text-slate-500 whitespace-nowrap">{attemptLabel(r)}</td>
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
                      <td className="p-3 text-center font-mono font-bold text-slate-600">
                        {r.noFlight ? <span className="font-sans text-xs font-black text-slate-400">No flight</span> : formatTime(r)}
                      </td>
                      <td className="p-3 text-sm text-slate-500 font-semibold">
                        {r.judgeName}
                        {r.editedAt && (
                          <span className="ml-1.5 text-[10px] font-black uppercase text-amber-600 bg-amber-50 px-1.5 py-0.5 rounded">
                            edited
                          </span>
                        )}
                      </td>
                      <td className="p-3 text-center">
                        {sigOf(r) ? (
                          <button
                            onClick={() => setViewSig(r)}
                            className="inline-block w-28 h-12 rounded-xl border-2 border-slate-200 bg-white hover:border-sky-400 overflow-hidden align-middle"
                            title="서명 크게 보기"
                            aria-label={`${r.participantName} 서명 크게 보기`}
                          >
                            <img src={sigOf(r)} alt="" className="w-full h-full object-contain" />
                          </button>
                        ) : (
                          <span className="text-sm font-bold text-slate-300">없음</span>
                        )}
                      </td>
                      <td className="p-3">
                        <div className="flex justify-center gap-1.5">
                          <button
                            onClick={() => setEditingId(r.id)}
                            className="w-10 h-10 rounded-xl grid place-items-center border-2 border-slate-200 text-slate-500 hover:text-violet-600 hover:border-violet-300 hover:bg-violet-50 transition"
                            title="Edit record"
                            aria-label="Edit record"
                          >
                            <Pencil size={17} />
                          </button>
                          <button
                            onClick={() => handleDelete(r)}
                            className="w-10 h-10 rounded-xl grid place-items-center border-2 border-slate-200 text-slate-500 hover:text-rose-600 hover:border-rose-300 hover:bg-rose-50 transition"
                            title="Delete record"
                            aria-label="Delete record"
                          >
                            <Trash2 size={17} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {viewSig && <SignatureViewer record={viewSig} src={sigOf(viewSig)} onClose={() => setViewSig(null)} />}

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

/** HQ: 서명 크게 보기 */
function SignatureViewer({ record, src, onClose }) {
  useEffect(() => {
    const onKey = (e) => e.key === "Escape" && onClose();
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);
  const when = record.timestamp ? new Date(record.timestamp).toLocaleString("ko-KR") : "";
  return (
    <div className="fixed inset-0 z-[120] bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4" onClick={onClose}>
      <div
        role="dialog"
        aria-modal="true"
        aria-label="Signature"
        className="w-full max-w-2xl bg-white rounded-3xl shadow-2xl overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between gap-3">
          <div className="min-w-0">
            <h3 className="text-xl font-black text-slate-800 truncate">
              #{record.participantNumber} {record.participantName} · 서명
            </h3>
            <p className="text-sm font-semibold text-slate-500 tabular-nums">
              {attemptLabel(record)} · 총점 {n(record.totalScore)}점{when ? ` · 제출 ${when}` : ""}
            </p>
          </div>
          <button onClick={onClose} className="w-10 h-10 shrink-0 grid place-items-center rounded-xl hover:bg-slate-100 text-slate-500" aria-label="Close">
            <X size={22} />
          </button>
        </div>
        <div className="p-6 bg-slate-50">
          {src ? (
            <img src={src} alt={`${record.participantName} 서명`} className="w-full max-h-[60vh] object-contain bg-white rounded-2xl border-2 border-slate-200" />
          ) : (
            <p className="text-center text-slate-400 font-bold py-10">서명이 없습니다.</p>
          )}
        </div>
      </div>
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
        let s = (await get(ref(db, `signatures/${record.signatureId || record.id}`))).val();
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
      if (changed || sigState === "ok") updates[`signatures/${record.signatureId || record.id}`] = signature;
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
      // 같은 라운드의 다른 기록이 함께 쓰는 서명은 남김
      const upd = { [`scores/${record.id}`]: null };
      if (!record.signatureId || record.signatureId === record.id) {
        const others = (await get(ref(db, "scores"))).val() || {};
        const k = record.signatureId || record.id;
        const shared = Object.entries(others).some(([id, v]) => id !== record.id && (v?.signatureId || id) === k);
        if (!shared) upd[`signatures/${k}`] = null;
      }
      await update(ref(db), upd);
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
              #{record.participantNumber} · {record.participantName} · {attemptLabel(record)}
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

function RankBadge({ rank, noFlight = false }) {
  if (rank == null)
    return (
      <span
        className={`inline-grid place-items-center h-9 px-2.5 rounded-xl text-sm font-black ${
          noFlight ? "bg-slate-100 text-slate-500" : "bg-rose-100 text-rose-600"
        }`}
      >
        {noFlight ? "—" : "DQ"}
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
