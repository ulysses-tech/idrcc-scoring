// =============================================================================
//  참가자 명단 엑셀(.xlsx / .xls / .csv) 읽기 · 빈 양식 만들기  (드론 추격전과 같은 방식)
//  - 첫 줄(제목 줄)의 이름으로 칸을 찾습니다. 칸 순서는 자유, 한글·영어 모두 인식.
//  - DCC 전용: '부문' 칸(초등·중등·고등·일반)을 읽어 Category를 자동 선택합니다.
// =============================================================================
import * as XLSX from "xlsx";

/** 제목 줄 이름 → 항목. 공백·대소문자·괄호는 무시하고 비교합니다. */
const HEADER_ALIASES = {
  order: ["순서", "순번", "출전순서", "no", "order", "seq"],
  bib: ["참가번호", "배번", "배번호", "번호", "선수번호", "bib", "bibno", "number", "id"],
  name: ["이름", "성명", "선수명", "참가자", "참가자명", "팀명", "name", "participant", "player"],
  category: ["부문", "종별", "구분", "부", "참가부문", "category", "division", "level"],
  country: ["국가", "나라", "국적", "country", "nation", "nationality"],
  affiliation: ["소속", "학교", "기관", "소속기관", "팀", "affiliation", "school", "team", "organization", "org"],
  grade: ["학년", "나이", "grade", "year", "age"],
  gender: ["성별", "gender", "sex"],
  group: ["조", "코스", "그룹", "경기장", "group", "course", "heat", "field"],
};

const norm = (s) => String(s ?? "").toLowerCase().replace(/[\s()[\]·._\-/]/g, "");
const ALIAS_LOOKUP = Object.fromEntries(
  Object.entries(HEADER_ALIASES).flatMap(([key, names]) => names.map((n) => [norm(n), key]))
);

export const TEMPLATE_HEADERS = ["순서", "참가번호", "이름", "부문", "국가", "소속", "학년", "성별", "조"];

/** 부문 표기 → 앱의 Category 이름 ('' = 비어 있음, null = 알 수 없음) */
export function toCategory(value) {
  const s = norm(value);
  if (!s) return "";
  if (/초등|elem|primary/.test(s)) return "Elementary";
  if (/중등|중학|middle|junior/.test(s)) return "Middle School";
  if (/고등|high|senior/.test(s)) return "High School";
  if (/일반|대학|성인|general|adult|open|univ/.test(s)) return "General";
  return null;
}

/**
 * 파일을 읽어 { entries, warnings } 를 돌려줍니다.
 * entries: [{ order, bib, name, category, country, affiliation, grade, gender, group }]
 */
export async function parseRosterFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: false, raw: false });
  if (!rows.length) throw new Error("파일에 내용이 없습니다.");

  // 제목 줄 찾기: 위에서 10줄 안에서 '이름'이나 '참가번호'가 있는 첫 줄
  const headerIdx = rows
    .slice(0, 10)
    .findIndex((r) => r.some((c) => ["name", "bib"].includes(ALIAS_LOOKUP[norm(c)])));
  if (headerIdx < 0) throw new Error("제목 줄을 찾지 못했습니다. 첫 줄에 '참가번호', '이름' 같은 제목을 넣어 주세요.");

  const colMap = {};
  rows[headerIdx].forEach((c, i) => {
    const key = ALIAS_LOOKUP[norm(c)];
    if (key && colMap[key] === undefined) colMap[key] = i;
  });
  if (colMap.name === undefined && colMap.bib === undefined) {
    throw new Error("'참가번호' 또는 '이름' 칸이 필요합니다.");
  }

  const warnings = [];
  const entries = [];
  rows.slice(headerIdx + 1).forEach((r, i) => {
    const line = headerIdx + i + 2;
    const get = (k) => (colMap[k] === undefined ? "" : String(r[colMap[k]] ?? "").trim());
    const rawCategory = get("category");
    const category = toCategory(rawCategory);
    const e = {
      order: get("order"),
      bib: get("bib"),
      name: get("name"),
      category: category || "",
      country: get("country"),
      affiliation: get("affiliation"),
      grade: get("grade"),
      gender: get("gender"),
      group: get("group"),
    };
    if (!e.bib && !e.name) return; // 빈 줄
    if (!e.name) warnings.push(`${line}번째 줄: 이름이 비어 있습니다.`);
    if (!e.bib) warnings.push(`${line}번째 줄: 참가번호가 비어 있습니다.`);
    if (category === null)
      warnings.push(`${line}번째 줄: 부문 '${rawCategory}'을(를) 알 수 없습니다. (초등·중등·고등·일반 중 하나로 적어 주세요)`);
    entries.push(e);
  });

  // 순서 칸이 없거나 숫자가 아니면 엑셀의 줄 순서를 그대로 사용
  const hasOrder = entries.length && entries.every((e) => e.order !== "" && !Number.isNaN(Number(e.order)));
  entries.forEach((e, i) => {
    e.order = hasOrder ? Number(e.order) : i + 1;
  });
  entries.sort((a, b) => a.order - b.order);

  // 참가번호 중복 확인
  const seen = new Map();
  entries.forEach((e) => {
    if (!e.bib) return;
    if (seen.has(e.bib)) warnings.push(`참가번호 ${e.bib} 이(가) 중복됩니다: ${seen.get(e.bib)}, ${e.name}`);
    else seen.set(e.bib, e.name);
  });

  if (!entries.length) throw new Error("명단에 참가자가 없습니다.");
  return { entries, warnings };
}

/** 빈 양식(.xlsx) 내려받기 — 예시 4줄 포함 */
export function downloadRosterTemplate(filename = "DCC_명단양식.xlsx") {
  const data = [
    TEMPLATE_HEADERS,
    [1, "101", "(예시) 홍길동", "초등", "Korea", "OO초등학교", "6", "남", "A"],
    [2, "102", "(예시) Bat-Erdene", "중등", "Mongolia", "MNUE Lab School", "8", "M", "A"],
    [3, "103", "(예시) Nguyen An", "고등", "Vietnam", "Hanoi STEAM Club", "11", "F", "B"],
    [4, "104", "(예시) 김드론", "일반", "Korea", "OO대학교", "", "여", "B"],
  ];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [{ wch: 6 }, { wch: 10 }, { wch: 22 }, { wch: 8 }, { wch: 12 }, { wch: 22 }, { wch: 6 }, { wch: 6 }, { wch: 6 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "명단");
  XLSX.writeFile(wb, filename);
}
