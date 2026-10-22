// =============================================================================
//  참가자 명단 (드론 추격전과 같은 방식)
//   - RosterPicker : 심판 화면 — '다음 순서' 참가자를 눌러 번호·이름·부문 자동 입력
//   - RosterManager: HQ 화면  — 엑셀 명단 불러오기 · 양식 받기 · 진행 현황
//  DCC는 1차·2차 시기가 있으므로 '완료' 여부를 심판이 고른 시기(Trial) 기준으로 계산합니다.
// =============================================================================
import React, { useMemo, useRef, useState } from "react";
import {
  Users,
  ChevronRight,
  Search,
  CheckCircle2,
  Loader2,
  Upload,
  FileSpreadsheet,
  Trash2,
  X,
  AlertTriangle,
  UserCheck,
  ListOrdered,
} from "lucide-react";
import { parseRosterFile, downloadRosterTemplate } from "./rosterFile";

const local = {
  get: (k) => {
    try {
      return localStorage.getItem(k);
    } catch {
      return null;
    }
  },
  set: (k, v) => {
    try {
      localStorage.setItem(k, v);
    } catch {
      /* ignore */
    }
  },
};

/** '채점 중' 표시가 이 시간보다 오래되면 무시 (태블릿이 꺼진 경우 대비) */
export const CLAIM_TTL_MS = 20 * 60 * 1000;

/** 이 기기 고유 ID (같은 이름의 심판이 있어도 구분) */
export const DEVICE_ID = (() => {
  let id = local.get("idcc_device_id");
  if (!id) {
    id = Math.random().toString(36).slice(2, 10);
    local.set("idcc_device_id", id);
  }
  return id;
})();

/** 다른 기기의 유효한 '채점 중' 표시인지 */
export const claimedByOther = (claim) =>
  Boolean(claim && claim.device !== DEVICE_ID && Date.now() - (Number(claim.at) || 0) < CLAIM_TTL_MS);

const CATEGORY_KO = {
  Elementary: "초등",
  "Middle School": "중등",
  "High School": "고등",
  General: "일반",
};

/** 부문 · 국가 · 소속 · 학년 · 성별 한 줄 */
export const profileLine = (p) =>
  [
    CATEGORY_KO[p.category] || p.category,
    p.country,
    p.affiliation,
    p.grade && `${p.grade}${/^\d+$/.test(String(p.grade)) ? "학년" : ""}`,
    p.gender,
  ]
    .filter(Boolean)
    .join(" · ");

/** 채점 기록에서 '참가번호|시기' 집합 만들기 */
const doneKeysOf = (records) =>
  new Set(records.map((r) => `${String(r.participantNumber ?? "").trim()}|${r.trial}`));

/** 명단 항목의 상태: done | busy | mine | wait */
const statusOf = (entry, { doneKeys, trial, claims, selectedId }) => {
  if (entry.id === selectedId) return "mine";
  if (entry.bib && doneKeys.has(`${entry.bib}|${trial}`)) return "done";
  if (claimedByOther(claims[entry.id])) return "busy";
  return "wait";
};

function StatusChip({ status, claim, trial }) {
  const map = {
    done: ["bg-emerald-100 text-emerald-800", `✓ ${trial} 완료`],
    busy: ["bg-amber-100 text-amber-900", `채점 중${claim?.judge ? ` · ${claim.judge}` : ""}`],
    mine: ["bg-sky-500 text-white", "선택됨"],
    wait: ["bg-slate-100 text-slate-600", "대기"],
  };
  const [cls, text] = map[status];
  return <span className={`shrink-0 text-xs font-black px-2.5 py-1 rounded-lg whitespace-nowrap ${cls}`}>{text}</span>;
}

// -----------------------------------------------------------------------------
//  심판용: 다음 순서 / 명단에서 선택
// -----------------------------------------------------------------------------
export function RosterPicker({ roster, claims, records, trial, selectedId, group, setGroup, onSelect, busy }) {
  const [showList, setShowList] = useState(false);
  const [query, setQuery] = useState("");

  const doneKeys = useMemo(() => doneKeysOf(records), [records]);
  const groups = useMemo(() => [...new Set(roster.map((r) => r.group).filter(Boolean))].sort(), [roster]);
  const inGroup = group && groups.includes(group) ? roster.filter((r) => r.group === group) : roster;
  const ctx = { doneKeys, trial, claims, selectedId };

  const doneCount = inGroup.filter((r) => r.bib && doneKeys.has(`${r.bib}|${trial}`)).length;
  const next = inGroup.find((r) => statusOf(r, ctx) === "wait");
  const selected = roster.find((r) => r.id === selectedId);

  const q = query.trim().toLowerCase();
  const listed = q
    ? inGroup.filter((r) => `${r.bib} ${r.name} ${r.affiliation} ${r.country}`.toLowerCase().includes(q))
    : inGroup;

  return (
    <section className="bg-white rounded-3xl border-2 border-sky-200 shadow-sm overflow-hidden">
      {/* 머리줄: 진행 현황 · 조 */}
      <div className="px-4 md:px-6 py-3.5 bg-sky-50 border-b border-sky-100 flex flex-wrap items-center gap-3">
        <div className="flex items-center gap-2 mr-auto">
          <Users size={20} className="text-sky-600" />
          <h3 className="text-lg font-black text-slate-800">참가자 명단</h3>
          <span className="text-sm font-bold text-slate-500 tabular-nums">
            {trial} 완료 {doneCount} / {inGroup.length}
          </span>
        </div>
        {groups.length > 0 && (
          <label className="flex items-center gap-2 text-xs font-black text-slate-500">
            조 · 코스
            <select
              value={group}
              onChange={(e) => setGroup(e.target.value)}
              className="h-10 px-3 rounded-xl border-2 border-slate-200 bg-white text-base font-bold text-slate-800 outline-none focus:border-sky-400"
            >
              <option value="">전체</option>
              {groups.map((g) => (
                <option key={g} value={g}>
                  {g}
                </option>
              ))}
            </select>
          </label>
        )}
      </div>

      <div className="p-4 md:p-6">
        {selected ? (
          <div className="flex flex-wrap items-center gap-4">
            <UserCheck size={28} className="text-sky-500 shrink-0" />
            <div className="min-w-0 flex-1">
              <div className="text-xs font-black text-sky-600">
                채점 중인 참가자 · 순서 {selected.order} · {trial}
              </div>
              <div className="text-2xl font-black text-slate-800 truncate">
                <span className="tabular-nums">#{selected.bib}</span> {selected.name}
              </div>
              {profileLine(selected) && (
                <div className="text-sm font-semibold text-slate-500">{profileLine(selected)}</div>
              )}
            </div>
            <button
              type="button"
              onClick={() => setShowList((v) => !v)}
              className="h-11 px-4 rounded-2xl border-2 border-slate-200 text-slate-700 font-bold hover:border-sky-400"
            >
              다른 참가자 선택
            </button>
          </div>
        ) : next ? (
          <div className="flex flex-wrap items-center gap-4">
            <div className="min-w-0 flex-1">
              <div className="text-xs font-black text-slate-500">
                다음 순서 · {next.order}번째{next.group ? ` · ${next.group}조` : ""} · {trial}
              </div>
              <div className="text-3xl font-black text-slate-800 truncate">
                <span className="tabular-nums text-sky-600">#{next.bib}</span> {next.name}
              </div>
              {profileLine(next) && <div className="text-base font-semibold text-slate-500">{profileLine(next)}</div>}
            </div>
            <button
              type="button"
              disabled={busy}
              onClick={() => onSelect(next)}
              className="h-16 px-6 rounded-2xl bg-gradient-to-r from-blue-500 to-sky-500 hover:from-blue-600 hover:to-sky-600 disabled:from-slate-300 disabled:to-slate-300 text-white text-lg font-black flex items-center gap-2 shadow-lg shadow-sky-200 active:scale-[0.98]"
            >
              {busy ? <Loader2 size={22} className="animate-spin" /> : <ChevronRight size={24} />}
              이 참가자 채점
            </button>
          </div>
        ) : (
          <p className="text-lg font-bold text-emerald-700 flex items-center gap-2">
            <CheckCircle2 size={22} />
            {inGroup.length
              ? `이 명단의 모든 참가자가 ${trial} 채점을 마쳤거나 다른 심판이 채점 중입니다.`
              : "명단이 비어 있습니다."}
          </p>
        )}

        <button
          type="button"
          onClick={() => setShowList((v) => !v)}
          className="mt-4 h-10 px-3 -ml-3 rounded-xl inline-flex items-center gap-1.5 text-sm font-bold text-sky-600 hover:bg-sky-50"
        >
          <ListOrdered size={16} /> {showList ? "명단 닫기" : "명단에서 직접 찾기 (번호·이름)"}
        </button>

        {showList && (
          <div className="mt-2">
            <div className="relative">
              <Search size={18} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="search"
                inputMode="search"
                autoFocus
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder="참가번호 또는 이름"
                className="w-full h-12 pl-10 pr-3 rounded-2xl border-2 border-slate-200 bg-slate-50 text-lg font-bold outline-none focus:bg-white focus:border-sky-400"
              />
            </div>
            <ul className="mt-2 max-h-80 overflow-y-auto divide-y divide-slate-100 border border-slate-200 rounded-2xl">
              {listed.length === 0 && <li className="p-4 text-slate-500 font-semibold">찾는 참가자가 없습니다.</li>}
              {listed.map((r) => {
                const st = statusOf(r, ctx);
                return (
                  <li key={r.id}>
                    <button
                      type="button"
                      disabled={busy || st === "mine"}
                      onClick={() => {
                        onSelect(r);
                        setShowList(false);
                        setQuery("");
                      }}
                      className={`w-full text-left px-4 py-3 flex items-center gap-3 hover:bg-sky-50 disabled:hover:bg-transparent ${
                        st === "done" ? "opacity-60" : ""
                      }`}
                    >
                      <span className="w-8 text-sm font-bold text-slate-400 tabular-nums">{r.order}</span>
                      <span className="w-16 font-black text-slate-800 tabular-nums">#{r.bib}</span>
                      <span className="flex-1 min-w-0">
                        <span className="block font-bold text-slate-800 truncate">{r.name}</span>
                        {profileLine(r) && (
                          <span className="block text-xs font-semibold text-slate-500 truncate">{profileLine(r)}</span>
                        )}
                      </span>
                      <StatusChip status={st} claim={claims[r.id]} trial={trial} />
                    </button>
                  </li>
                );
              })}
            </ul>
          </div>
        )}
      </div>
    </section>
  );
}

// -----------------------------------------------------------------------------
//  HQ용: 엑셀 명단 불러오기 · 양식 받기 · 진행 현황
// -----------------------------------------------------------------------------
export function RosterManager({ roster, claims, records, error, onSave, onClear, showToast }) {
  const fileRef = useRef(null);
  const [preview, setPreview] = useState(null); // { entries, warnings, fileName }
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [open, setOpen] = useState(false);

  const doneKeys = useMemo(() => doneKeysOf(records), [records]);
  const done1 = roster.filter((r) => doneKeys.has(`${r.bib}|1st`)).length;
  const done2 = roster.filter((r) => doneKeys.has(`${r.bib}|2nd`)).length;
  const inProgress = roster.filter((r) => claimedByOther(claims[r.id])).length;

  const onFile = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setLoading(true);
    try {
      const result = await parseRosterFile(file);
      setPreview({ ...result, fileName: file.name });
    } catch (err) {
      showToast(err?.message || "파일을 읽지 못했습니다. 엑셀(.xlsx) 또는 CSV 파일인지 확인하세요.", "error");
    } finally {
      setLoading(false);
    }
  };

  const confirmSave = async () => {
    setSaving(true);
    try {
      await onSave(preview.entries);
      showToast(`참가자 ${preview.entries.length}명 저장 · 모든 심판 화면에 바로 나타납니다`);
      setPreview(null);
    } catch (err) {
      showToast(saveFailMessage(err), "error");
    } finally {
      setSaving(false);
    }
  };

  const clearAll = async () => {
    if (!roster.length) return;
    if (!window.confirm(`참가자 명단 ${roster.length}명을 비울까요?\n(채점 기록은 지워지지 않습니다.)`)) return;
    try {
      await onClear();
      showToast("명단을 비웠습니다", "warn");
    } catch (err) {
      showToast(saveFailMessage(err), "error");
    }
  };

  return (
    <section className="bg-white rounded-3xl border border-slate-200 shadow-sm overflow-hidden mb-6">
      <div className="px-4 md:px-5 py-4 flex flex-wrap items-center gap-3">
        <div className="flex items-center gap-2 mr-auto">
          <Users size={22} className="text-sky-600" />
          <h3 className="text-lg font-black text-slate-800">참가자 명단</h3>
          {roster.length > 0 ? (
            <span className="text-sm font-bold text-slate-500 tabular-nums">
              {roster.length}명 · 1st 완료 {done1} · 2nd 완료 {done2} · 채점 중 {inProgress}
            </span>
          ) : (
            <span className="text-sm font-semibold text-slate-400">아직 없음 — 엑셀로 불러오세요</span>
          )}
        </div>
        <button
          type="button"
          onClick={() => downloadRosterTemplate()}
          className="h-11 px-4 rounded-2xl border-2 border-slate-200 bg-white text-slate-700 font-bold flex items-center gap-2 hover:border-sky-400"
        >
          <FileSpreadsheet size={18} /> 양식 받기
        </button>
        <button
          type="button"
          onClick={() => fileRef.current?.click()}
          disabled={loading || !!error}
          className="h-11 px-4 rounded-2xl bg-sky-500 hover:bg-sky-600 disabled:bg-slate-300 text-white font-bold flex items-center gap-2 shadow-md shadow-sky-200"
        >
          {loading ? <Loader2 size={18} className="animate-spin" /> : <Upload size={18} />} 엑셀 명단 불러오기
        </button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={onFile} />
        {roster.length > 0 && (
          <>
            <button
              type="button"
              onClick={() => setOpen((v) => !v)}
              className="h-11 px-4 rounded-2xl border-2 border-slate-200 bg-white text-slate-700 font-bold hover:border-sky-400"
            >
              {open ? "명단 접기" : "명단 보기"}
            </button>
            <button
              type="button"
              onClick={clearAll}
              title="명단 비우기"
              aria-label="명단 비우기"
              className="h-11 w-11 grid place-items-center rounded-2xl border-2 border-rose-200 text-rose-600 hover:bg-rose-50"
            >
              <Trash2 size={18} />
            </button>
          </>
        )}
      </div>

      {error && (
        <div className="mx-4 md:mx-5 mb-4 rounded-2xl bg-rose-50 border border-rose-200 p-4 text-sm font-semibold text-rose-700">
          <p className="flex items-center gap-2 font-black">
            <AlertTriangle size={16} /> 명단 저장 공간에 접근할 수 없습니다 ({error})
          </p>
          <p className="mt-1">
            Firebase 콘솔 → Realtime Database → 규칙에 <code className="font-mono bg-white px-1 rounded">dcc_meta</code> 항목을
            추가해야 합니다. (README 참고)
          </p>
        </div>
      )}

      {open && roster.length > 0 && (
        <div className="border-t border-slate-200 max-h-96 overflow-auto">
          <table className="w-full text-left text-sm min-w-[720px]">
            <thead className="bg-slate-50 text-slate-500 text-xs uppercase font-black sticky top-0">
              <tr>
                <th className="py-2.5 px-3">순서</th>
                <th className="py-2.5 px-3">번호</th>
                <th className="py-2.5 px-3">이름</th>
                <th className="py-2.5 px-3">부문 · 국가 · 소속 · 학년 · 성별</th>
                <th className="py-2.5 px-3">조</th>
                <th className="py-2.5 px-3">1st</th>
                <th className="py-2.5 px-3">2nd</th>
              </tr>
            </thead>
            <tbody>
              {roster.map((r) => (
                <tr key={r.id} className="border-t border-slate-100">
                  <td className="py-2 px-3 tabular-nums text-slate-500">{r.order}</td>
                  <td className="py-2 px-3 font-black tabular-nums">{r.bib}</td>
                  <td className="py-2 px-3 font-bold">{r.name}</td>
                  <td className="py-2 px-3 text-slate-500">{profileLine(r)}</td>
                  <td className="py-2 px-3">{r.group}</td>
                  {["1st", "2nd"].map((t) => (
                    <td key={t} className="py-2 px-3">
                      <StatusChip
                        status={statusOf(r, { doneKeys, trial: t, claims, selectedId: null })}
                        claim={claims[r.id]}
                        trial={t}
                      />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {preview && (
        <div className="fixed inset-0 z-[130] bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4" onClick={() => setPreview(null)}>
          <div
            role="dialog"
            aria-modal="true"
            aria-label="명단 미리보기"
            className="w-full max-w-3xl max-h-[90vh] flex flex-col bg-white rounded-3xl shadow-2xl overflow-hidden"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between gap-3">
              <div className="min-w-0">
                <h3 className="text-xl font-black text-slate-800">명단 미리보기 · {preview.entries.length}명</h3>
                <p className="text-sm font-semibold text-slate-500 truncate">{preview.fileName}</p>
              </div>
              <button onClick={() => setPreview(null)} className="w-10 h-10 grid place-items-center rounded-xl hover:bg-slate-100" aria-label="Close">
                <X size={22} />
              </button>
            </div>
            <div className="p-6 overflow-y-auto">
              {preview.warnings.length > 0 && (
                <div className="mb-4 rounded-2xl bg-amber-50 border border-amber-200 p-3 text-sm text-amber-900 font-semibold">
                  <p className="flex items-center gap-2 font-black">
                    <AlertTriangle size={16} /> 확인이 필요한 줄 {preview.warnings.length}건
                  </p>
                  <ul className="mt-1 list-disc pl-5 max-h-28 overflow-y-auto">
                    {preview.warnings.map((w, i) => (
                      <li key={i}>{w}</li>
                    ))}
                  </ul>
                </div>
              )}
              <div className="overflow-x-auto border border-slate-200 rounded-2xl">
                <table className="w-full text-left text-sm min-w-[600px]">
                  <thead className="bg-slate-50 font-black text-slate-500 text-xs uppercase">
                    <tr>
                      <th className="py-2.5 px-3">순서</th>
                      <th className="py-2.5 px-3">번호</th>
                      <th className="py-2.5 px-3">이름</th>
                      <th className="py-2.5 px-3">부문 · 국가 · 소속 · 학년 · 성별</th>
                      <th className="py-2.5 px-3">조</th>
                    </tr>
                  </thead>
                  <tbody>
                    {preview.entries.slice(0, 50).map((r, i) => (
                      <tr key={i} className="border-t border-slate-100">
                        <td className="py-2 px-3 tabular-nums">{r.order}</td>
                        <td className="py-2 px-3 font-black tabular-nums">{r.bib}</td>
                        <td className="py-2 px-3 font-bold">{r.name}</td>
                        <td className="py-2 px-3 text-slate-500">{profileLine(r)}</td>
                        <td className="py-2 px-3">{r.group}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {preview.entries.length > 50 && (
                <p className="mt-2 text-sm text-slate-500 font-semibold">… 외 {preview.entries.length - 50}명</p>
              )}
              {roster.length > 0 && (
                <p className="mt-4 text-sm font-bold text-rose-600">
                  저장하면 기존 명단 {roster.length}명이 이 명단으로 바뀝니다. 채점 기록은 그대로 유지됩니다.
                </p>
              )}
            </div>
            <div className="px-6 py-4 border-t border-slate-200 flex justify-end gap-2">
              <button onClick={() => setPreview(null)} className="h-12 px-5 rounded-2xl font-bold text-slate-600 hover:bg-slate-100">
                취소
              </button>
              <button
                onClick={confirmSave}
                disabled={saving}
                className="h-12 px-6 rounded-2xl bg-sky-500 hover:bg-sky-600 disabled:bg-slate-300 text-white font-black flex items-center gap-2"
              >
                {saving && <Loader2 size={18} className="animate-spin" />} {preview.entries.length}명 저장
              </button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}

/** 저장 실패 안내 문구 */
export function saveFailMessage(err) {
  const t = String(err?.code || err?.message || "").toLowerCase();
  if (t.includes("timeout")) return "서버 응답이 늦습니다. 인터넷 연결을 확인한 뒤 다시 시도하세요.";
  if (t.includes("permission")) return "저장 권한이 없습니다. Firebase 규칙에 dcc_meta 항목을 추가했는지 확인하세요.";
  if (t.includes("offline")) return "인터넷 연결이 끊겨 있습니다. Online 표시를 확인한 뒤 다시 시도하세요.";
  return "저장 실패 — 인터넷 연결을 확인하세요.";
}
