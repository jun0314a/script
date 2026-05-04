"use client";

import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── 타입 ────────────────────────────────────────────────────────────────────

interface Module {
  cf: string;
  상위모듈: string;
  모듈코드: string;
  모듈명: string;
  발화목록: string[];
  중요: boolean;
}

interface Branch {
  cf: string;
  모듈명: string;
  응답유형: string;
  다음모듈: string;
  비고: string;
}

interface CustomImportant {
  제목: string;
  내용: string;
}

interface ParsedData {
  modules: Module[];
  branches: Branch[];
  customImportant: CustomImportant[];
}

interface SavedScript {
  id: string;
  모듈명: string;
  모듈코드: string;
  cf: string;
  발화: string;
  savedAt: string;
}

interface CustomerInfo {
  customerName: string;
  gender: string;
  age: string;
  consultantName: string;
  renewalType: string;
  coverageRange: string;
  coverageAmount: string;
  coveragePeriod: string;
  하위보험사: string;
  유병력자: string;
  생손보사: string;
  ci: string;
  우체국: string;
  실손: string;
}

interface CustomerProfile extends CustomerInfo {
  id: string;
  profileName: string;
  savedAt: string;
}

interface LocalImage {
  name: string;
  src: string;
}

// ─── 이미지 맵 ────────────────────────────────────────────────────────────────
// public/images/ 폴더에 이미지 파일을 넣고 아래에 등록하세요.
// 모듈코드는 상담 화면 파란 헤더에서 확인할 수 있어요.
//
// 예시:
// "MD 1": [{ name: "보장비교표", src: "/images/md1-chart.png" }],
// "MD 2": [
//   { name: "갱신형 그래프", src: "/images/renewal-graph.png" },
//   { name: "보장금액 비교", src: "/images/amount-compare.png" },
// ],

const MODULE_IMAGES: Record<string, LocalImage[]> = {
    "MD 3-1": [{ name: "신원증빙", src: "/images/sjcard.png",}],
    "MD 2-2-1": [{ name: "보장범위", src: "/images/뇌혈관질환.png" },
                 { name : "보장범위", src : "/images/허혈성심장질환.png"}],
};

// ─── 상수 ────────────────────────────────────────────────────────────────────

const CF_LABELS: Record<string, string> = {
  "CF1. 상담 가치 형성(구체화)": "CF1 상담 가치 형성",
  "CF2. 문제 인식": "CF2 문제 인식",
  "CF3. 솔루션 확정": "CF3 솔루션 확정",
  "CF4. 청약 완료": "CF4 청약 완료",
  "CF5. 상담 마무리": "CF5 상담 마무리",
};

const INPUT_CLS =
  "w-full rounded-lg border border-gray-300 px-3 py-2 text-sm text-black focus:outline-none focus:ring-2 focus:ring-blue-400 bg-white";

// ─── localStorage 헬퍼 ───────────────────────────────────────────────────────

function lsGet<T>(key: string, fallback: T): T {
  if (typeof window === "undefined") return fallback;
  try { return JSON.parse(localStorage.getItem(key) || "null") ?? fallback; } catch { return fallback; }
}
function lsSet(key: string, value: unknown) {
  try { localStorage.setItem(key, JSON.stringify(value)); } catch {}
}

// ─── 이름 치환 ────────────────────────────────────────────────────────────────

function applyReplacements(text: string, info: CustomerInfo): string {
  const { customerName, consultantName } = info;
  let r = text;
  if (consultantName) {
    r = r.replace(/보험전문가\s*\[이름\]/g, `보험전문가 ${consultantName}`);
    r = r.replace(/보험전문가\s*○{2,3}/g, `보험전문가 ${consultantName}`);
    r = r.replace(/보험전문가\s*O{2,3}/g, `보험전문가 ${consultantName}`);
    r = r.replace(/\[이름\](입니다|이에요|예요|이야)/g, `${consultantName}$1`);
  }
  if (customerName) {
    r = r.replace(/○{2,3}님/g, `${customerName}님`);
    r = r.replace(/O{2,3}님/g, `${customerName}님`);
    r = r.replace(/0{2}님/g, `${customerName}님`);
    r = r.replace(/\[이름\]님/g, `${customerName}님`);
    r = r.replace(/고객님/g, `${customerName}님`);
  }
  return r;
}

// ─── 엑셀 파싱 ────────────────────────────────────────────────────────────────

function parseExcelFromBuffer(buffer: ArrayBuffer): ParsedData {
  const data = new Uint8Array(buffer);
  const wb = XLSX.read(data, { type: "array" });

  const scriptSheet = wb.Sheets[wb.SheetNames[2]];
  const scriptRows = XLSX.utils.sheet_to_json<Record<string, string>>(scriptSheet, { defval: "" });
  const RESERVED_COLS = ["CF", "상위 모듈", "모듈 코드", "모듈명", "중요"];
  const scriptCols = Object.keys(scriptRows[0] || {}).filter((k) => !RESERVED_COLS.includes(k));
  const modules: Module[] = scriptRows
    .filter((r) => r["모듈명"])
    .map((r) => {
      const rawImportant = r["중요"]?.toString().trim() ?? "";
      const 중요 = rawImportant !== "" &&
        !["n", "no", "false", "아니오"].includes(rawImportant.toLowerCase());
      return {
        cf: r["CF"] || "",
        상위모듈: r["상위 모듈"] || "",
        모듈코드: r["모듈 코드"] || "",
        모듈명: r["모듈명"] || "",
        발화목록: scriptCols.map((c) => r[c]?.toString().trim()).filter(Boolean),
        중요,
      };
    })
    .filter((m) => m.발화목록.length > 0);

  const branchSheet = wb.Sheets[wb.SheetNames[0]];
  const branchRows = XLSX.utils.sheet_to_json<Record<string, string>>(branchSheet, { defval: "" });
  const branches: Branch[] = branchRows
    .filter((r) => r["현재 모듈명"] && r["고객 응답 유형"] && r["다음 모듈"])
    .map((r) => ({
      cf: r["CF"] || "",
      모듈명: r["현재 모듈명"],
      응답유형: r["고객 응답 유형"],
      다음모듈: r["다음 모듈"],
      비고: r["비고"] || "",
    }));

  const customSheet = wb.Sheets["중요발화"];
  const customImportant: CustomImportant[] = customSheet
    ? XLSX.utils.sheet_to_json<Record<string, string>>(customSheet, { defval: "" })
        .filter((r) => r["내용"]?.toString().trim())
        .map((r) => ({
          제목: r["제목"]?.toString().trim() || "",
          내용: r["내용"]?.toString().trim(),
        }))
    : [];

  return { modules, branches, customImportant };
}

function parseExcel(file: File): Promise<ParsedData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try { resolve(parseExcelFromBuffer(e.target!.result as ArrayBuffer)); }
      catch { reject(new Error("파일을 읽는 중 오류가 발생했습니다.")); }
    };
    reader.onerror = () => reject(new Error("파일 읽기 실패"));
    reader.readAsArrayBuffer(file);
  });
}

// ─── 모듈 필터링 ──────────────────────────────────────────────────────────────

const FIELD_MODULE_PATTERNS: { field: keyof CustomerInfo; patterns: string[] }[] = [
  { field: "renewalType",    patterns: ["갱신형"] },
  { field: "coverageRange",  patterns: ["보장범위"] },
  { field: "coverageAmount", patterns: ["보장금액"] },
  { field: "coveragePeriod", patterns: ["보장기간", "납입기간"] },
  { field: "하위보험사",       patterns: ["하위보험사", "하위 보험사"] },
  { field: "유병력자",         patterns: ["유병력자"] },
  { field: "생손보사",         patterns: ["생보사", "손보사", "생/손보사"] },
  { field: "ci",             patterns: ["CI"] },
  { field: "우체국",           patterns: ["우체국"] },
];

const 실손_세대 = ["1세대", "2세대", "3세대", "4세대"];

function shouldShowModule(module: Module, info: CustomerInfo): boolean {
  for (const { field, patterns } of FIELD_MODULE_PATTERNS) {
    const isRelated = patterns.some((p) => module.모듈명.includes(p));
    if (isRelated && !info[field]) return false;
  }
  if (module.모듈명.includes("실손")) {
    if (!info.실손) return false;
    const moduleGen = 실손_세대.find((g) => module.모듈명.includes(g));
    if (moduleGen && moduleGen !== info.실손) return false;
  }
  return true;
}

function findNextVisible(fromModule: Module, modules: Module[], info: CustomerInfo): Module | undefined {
  let idx = modules.findIndex((m) => m.모듈코드 === fromModule.모듈코드 && m.모듈명 === fromModule.모듈명);
  if (idx === -1) idx = modules.findIndex((m) => m.모듈명 === fromModule.모듈명);
  for (let i = idx + 1; i < modules.length; i++) {
    if (shouldShowModule(modules[i], info)) return modules[i];
  }
  return undefined;
}

function findModule(targetName: string, modules: Module[]): Module | undefined {
  const name = targetName.includes("→") ? targetName.split("→").pop()!.trim() : targetName.trim();
  return modules.find((m) => m.모듈명 === name) ||
    modules.find((m) => m.모듈명.includes(name) || name.includes(m.모듈명));
}

function getBranches(moduleName: string, branches: Branch[]): Branch[] {
  return branches.filter(
    (b) => b.모듈명 === moduleName || moduleName.includes(b.모듈명) || b.모듈명.includes(moduleName)
  );
}

function pickScript(module: Module, info: CustomerInfo, edits: Record<string, string[]>): string {
  const lines = edits[module.모듈코드] ?? module.발화목록;
  if (!lines.length) return "";
  return applyReplacements(lines[Math.floor(Math.random() * lines.length)], info);
}

function formatDate(iso: string): string {
  const d = new Date(iso);
  return `${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

// ─── 컴포넌트 ─────────────────────────────────────────────────────────────────

export default function Home() {
  // ── 파일 ──
  const [parsedData,   setParsedData]   = useState<ParsedData | null>(null);
  const [fileName,     setFileName]     = useState("");
  const [fileError,    setFileError]    = useState("");
  const [dragging,     setDragging]     = useState(false);
  const [excelLoading, setExcelLoading] = useState(true);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // ── 고객 정보 ──
  const [customerName,   setCustomerName]   = useState("");
  const [gender,         setGender]         = useState("");
  const [age,            setAge]            = useState("");
  const [consultantName, setConsultantName] = useState("");
  const [renewalType,    setRenewalType]    = useState("");
  const [coverageRange,  setCoverageRange]  = useState("");
  const [coverageAmount, setCoverageAmount] = useState("");
  const [coveragePeriod, setCoveragePeriod] = useState("");
  const [하위보험사, set하위보험사] = useState("");
  const [유병력자,   set유병력자]   = useState("");
  const [생손보사,   set생손보사]   = useState("");
  const [ci,         setCi]         = useState("");
  const [우체국,     set우체국]     = useState("");
  const [실손,       set실손]       = useState("");

  const info: CustomerInfo = {
    customerName, gender, age, consultantName,
    renewalType, coverageRange, coverageAmount, coveragePeriod,
    하위보험사, 유병력자, 생손보사, ci, 우체국, 실손,
  };

  // ── 상담 흐름 ──
  const [started,       setStarted]       = useState(false);
  const [currentModule, setCurrentModule] = useState<Module | null>(null);
  const [currentScript, setCurrentScript] = useState("");
  const [history,       setHistory]       = useState<{ module: Module; script: string }[]>([]);
  const [done,          setDone]          = useState(false);

  // ── 홈 탭 ──
  const [homeTab, setHomeTab] = useState<"상담" | "저장">("상담");

  // ── 로컬 데이터 ──
  const [savedScripts,     setSavedScripts]     = useState<SavedScript[]>([]);
  const [customerProfiles, setCustomerProfiles] = useState<CustomerProfile[]>([]);
  const [scriptEdits,      setScriptEdits]      = useState<Record<string, string[]>>({});

  // ── 프로필 UI ──
  const [profileName,       setProfileName]       = useState("");
  const [showProfileSave,   setShowProfileSave]   = useState(false);
  const [selectedProfileId, setSelectedProfileId] = useState("");

  // ── 스크립트 편집 ──
  const [editingModule, setEditingModule] = useState<string | null>(null);
  const [editLines,     setEditLines]     = useState<string[]>([]);

  // ── 검색 ──
  const [searchQuery, setSearchQuery] = useState("");

  // ── 이미지 ──
  const [lightboxImage,  setLightboxImage]  = useState<LocalImage | null>(null);
  const [copiedImageSrc, setCopiedImageSrc] = useState<string | null>(null);

  // ── 초기 로드 ──
  useEffect(() => {
    setSavedScripts(lsGet<SavedScript[]>("savedScripts", []));
    setCustomerProfiles(lsGet<CustomerProfile[]>("customerProfiles", []));
    setScriptEdits(lsGet<Record<string, string[]>>("scriptEdits", {}));

    // public/scripts.xlsx 자동 로드
    fetch("/scripts.xlsx")
      .then((res) => {
        if (!res.ok) throw new Error("없음");
        return res.arrayBuffer();
      })
      .then((buf) => {
        const data = parseExcelFromBuffer(buf);
        setParsedData(data);
        setFileName("scripts.xlsx");
      })
      .catch(() => {/* 파일 없음 → 수동 업로드 */})
      .finally(() => setExcelLoading(false));
  }, []);

  // ── 저장된 발화 ──
  function handleSave() {
    if (!currentModule || !currentScript) return;
    if (savedScripts.some((s) => s.발화 === currentScript && s.모듈명 === currentModule.모듈명)) return;
    const next = [
      { id: `${Date.now()}`, 모듈명: currentModule.모듈명, 모듈코드: currentModule.모듈코드, cf: currentModule.cf, 발화: currentScript, savedAt: new Date().toISOString() },
      ...savedScripts,
    ];
    setSavedScripts(next);
    lsSet("savedScripts", next);
  }
  function handleDeleteSaved(id: string) {
    const next = savedScripts.filter((s) => s.id !== id);
    setSavedScripts(next); lsSet("savedScripts", next);
  }
  function handleClearAll() {
    if (!confirm("저장된 발화를 모두 삭제할까요?")) return;
    setSavedScripts([]); lsSet("savedScripts", []);
  }
  const isAlreadySaved = !!currentModule &&
    savedScripts.some((s) => s.발화 === currentScript && s.모듈명 === currentModule.모듈명);

  // ── 고객 프로필 ──
  function handleSaveProfile() {
    if (!profileName.trim()) return;
    const next = [
      { id: `${Date.now()}`, profileName: profileName.trim(), savedAt: new Date().toISOString(), ...info },
      ...customerProfiles,
    ];
    setCustomerProfiles(next); lsSet("customerProfiles", next);
    setProfileName(""); setShowProfileSave(false);
  }
  function handleLoadProfile() {
    const p = customerProfiles.find((p) => p.id === selectedProfileId);
    if (!p) return;
    setCustomerName(p.customerName); setGender(p.gender); setAge(p.age);
    setConsultantName(p.consultantName); setRenewalType(p.renewalType);
    setCoverageRange(p.coverageRange); setCoverageAmount(p.coverageAmount);
    setCoveragePeriod(p.coveragePeriod); set하위보험사(p.하위보험사);
    set유병력자(p.유병력자); set생손보사(p.생손보사); setCi(p.ci);
    set우체국(p.우체국); set실손(p.실손);
  }
  function handleDeleteProfile(id: string) {
    const next = customerProfiles.filter((p) => p.id !== id);
    setCustomerProfiles(next); lsSet("customerProfiles", next);
    if (selectedProfileId === id) setSelectedProfileId("");
  }

  // ── 스크립트 편집 ──
  function startEditing(module: Module) {
    setEditingModule(module.모듈코드);
    setEditLines([...(scriptEdits[module.모듈코드] ?? module.발화목록)]);
  }
  function saveEdits() {
    if (!currentModule) return;
    const cleaned = editLines.filter((l) => l.trim());
    const next = { ...scriptEdits, [currentModule.모듈코드]: cleaned };
    setScriptEdits(next); lsSet("scriptEdits", next);
    setEditingModule(null);
  }
  function resetModuleEdits() {
    if (!currentModule) return;
    const next = { ...scriptEdits };
    delete next[currentModule.모듈코드];
    setScriptEdits(next); lsSet("scriptEdits", next);
    setEditingModule(null);
  }
  const isEditing = !!currentModule && editingModule === currentModule.모듈코드;

  // ── 검색 ──
  const searchResults = searchQuery.trim() && parsedData
    ? parsedData.modules.filter((m) =>
        m.모듈명.includes(searchQuery) ||
        (scriptEdits[m.모듈코드] ?? m.발화목록).some((line) => line.includes(searchQuery))
      )
    : [];

  // ── 이미지 복사 ──
  async function handleCopyImage(src: string) {
    try {
      const res = await fetch(src);
      const blob = await res.blob();
      await navigator.clipboard.write([new ClipboardItem({ [blob.type]: blob })]);
      setCopiedImageSrc(src);
      setTimeout(() => setCopiedImageSrc(null), 2000);
    } catch {
      window.open(src, "_blank");
    }
  }

  // ── 파일 ──
  const handleFile = useCallback(async (file: File) => {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
      setFileError("엑셀 파일(.xlsx, .xls)만 업로드 가능합니다."); return;
    }
    setFileError("");
    try {
      const data = await parseExcel(file);
      setParsedData(data); setFileName(file.name);
      setStarted(false); setDone(false);
    } catch (e) { setFileError((e as Error).message); }
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault(); setDragging(false);
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  }, [handleFile]);

  // ── 상담 흐름 ──
  function handleStart(fromModule?: Module) {
    if (!parsedData) return;
    const first = fromModule ?? parsedData.modules.find((m) => shouldShowModule(m, info));
    if (!first) return;
    setCurrentModule(first);
    setCurrentScript(pickScript(first, info, scriptEdits));
    setHistory([]); setDone(false); setStarted(true); setSearchQuery("");
  }
  function handleResponse(branch: Branch) {
    if (!parsedData || !currentModule) return;
    setHistory((h) => [...h, { module: currentModule, script: currentScript }]);
    const nextName = branch.다음모듈;
    if (nextName.includes("종료") || nextName.includes("CF1부터") || currentModule.모듈명 === "마무리 인사") {
      setDone(true); return;
    }
    const next = findModule(nextName, parsedData.modules);
    const target = next && shouldShowModule(next, info) ? next
      : next ? findNextVisible(next, parsedData.modules, info)
      : findNextVisible(currentModule, parsedData.modules, info);
    if (target) { setCurrentModule(target); setCurrentScript(pickScript(target, info, scriptEdits)); }
    else setDone(true);
  }
  function handleNext() {
    if (!parsedData || !currentModule) return;
    setHistory((h) => [...h, { module: currentModule, script: currentScript }]);
    const next = findNextVisible(currentModule, parsedData.modules, info);
    if (next) { setCurrentModule(next); setCurrentScript(pickScript(next, info, scriptEdits)); }
    else setDone(true);
  }
  function handleBack() {
    if (history.length === 0) return;
    const prev = history[history.length - 1];
    setHistory((h) => h.slice(0, -1));
    setCurrentModule(prev.module); setCurrentScript(prev.script); setDone(false);
  }
  function handleReroll() {
    if (!currentModule) return;
    setCurrentScript(pickScript(currentModule, info, scriptEdits));
  }

  const currentBranches = parsedData && currentModule
    ? getBranches(currentModule.모듈명, parsedData.branches) : [];
  const currentModuleImages = currentModule ? (MODULE_IMAGES[currentModule.모듈코드] ?? []) : [];

  return (
    <main className="min-h-screen py-10 px-4">
      <div className="max-w-2xl mx-auto">

        <div className="mb-8 text-center">
          <h1 className="text-2xl font-bold text-black mb-1">보험 상담 스크립트</h1>
          <p className="text-sm text-black">고객 반응에 따라 다음 발화를 안내해드립니다.</p>
        </div>

        {/* ── 준비 화면 ── */}
        {!started && (
          <>
            <div className="flex gap-1 mb-4 bg-gray-200 rounded-xl p-1">
              <button onClick={() => setHomeTab("상담")}
                className={`flex-1 py-2 rounded-lg text-sm font-semibold transition-colors ${homeTab === "상담" ? "bg-white text-black shadow-sm" : "text-gray-500 hover:text-black"}`}>
                상담 시작
              </button>
              <button onClick={() => setHomeTab("저장")}
                className={`flex-1 py-2 rounded-lg text-sm font-semibold transition-colors ${homeTab === "저장" ? "bg-white text-black shadow-sm" : "text-gray-500 hover:text-black"}`}>
                저장된 발화
                {savedScripts.length > 0 && (
                  <span className="ml-1.5 inline-flex items-center justify-center w-5 h-5 rounded-full bg-blue-500 text-white text-xs font-bold">
                    {savedScripts.length}
                  </span>
                )}
              </button>
            </div>

            {/* ── 상담 탭 ── */}
            {homeTab === "상담" && (
              <>
                <div className="bg-white rounded-2xl border border-gray-300 p-6 mb-4">
                  <div className="flex items-center justify-between mb-4">
                    <h2 className="text-sm font-semibold text-black uppercase tracking-wide">상담 정보 입력</h2>
                    <div className="flex items-center gap-2">
                      {customerProfiles.length > 0 && (
                        <div className="flex items-center gap-1.5">
                          <select value={selectedProfileId} onChange={(e) => setSelectedProfileId(e.target.value)}
                            className="text-xs border border-gray-300 rounded-lg px-2 py-1.5 text-black bg-white focus:outline-none">
                            <option value="">프로필 선택</option>
                            {customerProfiles.map((p) => (
                              <option key={p.id} value={p.id}>{p.profileName}</option>
                            ))}
                          </select>
                          <button onClick={handleLoadProfile} disabled={!selectedProfileId}
                            className="text-xs bg-gray-100 border border-gray-300 rounded-lg px-2 py-1.5 text-black hover:bg-gray-200 disabled:opacity-40 disabled:cursor-not-allowed">
                            불러오기
                          </button>
                          {selectedProfileId && (
                            <button onClick={() => handleDeleteProfile(selectedProfileId)}
                              className="text-xs text-red-400 hover:text-red-600" title="프로필 삭제">✕</button>
                          )}
                        </div>
                      )}
                      <button onClick={() => setShowProfileSave((v) => !v)}
                        className="text-xs text-blue-600 hover:text-blue-800 border border-blue-200 rounded-lg px-2 py-1.5 bg-blue-50 whitespace-nowrap">
                        프로필 저장
                      </button>
                    </div>
                  </div>

                  {showProfileSave && (
                    <div className="flex gap-2 mb-4 p-3 bg-blue-50 rounded-xl border border-blue-100">
                      <input type="text" placeholder="프로필 이름 (예: 홍길동_25.01.01)"
                        value={profileName} onChange={(e) => setProfileName(e.target.value)}
                        onKeyDown={(e) => e.key === "Enter" && handleSaveProfile()}
                        className="flex-1 rounded-lg border border-gray-300 px-3 py-1.5 text-sm text-black focus:outline-none focus:ring-2 focus:ring-blue-400 bg-white" />
                      <button onClick={handleSaveProfile} disabled={!profileName.trim()}
                        className="text-sm bg-blue-600 text-white rounded-lg px-3 py-1.5 hover:bg-blue-700 disabled:opacity-40 disabled:cursor-not-allowed">
                        저장
                      </button>
                      <button onClick={() => { setShowProfileSave(false); setProfileName(""); }}
                        className="text-sm text-gray-500 hover:text-black">취소</button>
                    </div>
                  )}

                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">고객 이름</label>
                      <input type="text" placeholder="홍길동" value={customerName}
                        onChange={(e) => setCustomerName(e.target.value)} className={INPUT_CLS} />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">설계사 이름</label>
                      <input type="text" placeholder="내 이름" value={consultantName}
                        onChange={(e) => setConsultantName(e.target.value)} className={INPUT_CLS} />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">성별</label>
                      <select value={gender} onChange={(e) => setGender(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="남성">남성</option>
                        <option value="여성">여성</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">나이</label>
                      <input type="text" placeholder="예: 35세" value={age}
                        onChange={(e) => setAge(e.target.value)} className={INPUT_CLS} />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">갱신형</label>
                      <select value={renewalType} onChange={(e) => setRenewalType(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="갱신형">갱신형</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">보장범위</label>
                      <select value={coverageRange} onChange={(e) => setCoverageRange(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="좁음">좁음</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">보장금액</label>
                      <select value={coverageAmount} onChange={(e) => setCoverageAmount(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="부족">부족</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">보장기간</label>
                      <select value={coveragePeriod} onChange={(e) => setCoveragePeriod(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="짧음">짧음</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">하위보험사</label>
                      <select value={하위보험사} onChange={(e) => set하위보험사(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="있음">있음</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">유병력자 보험</label>
                      <select value={유병력자} onChange={(e) => set유병력자(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="보유">보유</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">생/손보사</label>
                      <select value={생손보사} onChange={(e) => set생손보사(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="문제 있음">문제 있음</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">CI</label>
                      <select value={ci} onChange={(e) => setCi(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="보유">보유</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">우체국</label>
                      <select value={우체국} onChange={(e) => set우체국(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="보유">보유</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-black mb-1">실손</label>
                      <select value={실손} onChange={(e) => set실손(e.target.value)} className={INPUT_CLS}>
                        <option value="">선택 안함</option>
                        <option value="1세대">1세대</option>
                        <option value="2세대">2세대</option>
                        <option value="3세대">3세대</option>
                        <option value="4세대">4세대</option>
                      </select>
                    </div>
                  </div>
                </div>

                {/* 검색 */}
                {parsedData && (
                  <div className="bg-white rounded-2xl border border-gray-300 p-4 mb-4">
                    <div className="flex items-center gap-2">
                      <span className="text-gray-400">🔍</span>
                      <input type="text" placeholder="모듈명 또는 발화 내용 검색..."
                        value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)}
                        className="flex-1 text-sm text-black focus:outline-none bg-transparent placeholder-gray-400" />
                      {searchQuery && (
                        <button onClick={() => setSearchQuery("")} className="text-gray-400 hover:text-black text-sm">✕</button>
                      )}
                    </div>
                    {searchResults.length > 0 && (
                      <div className="mt-3 space-y-2 max-h-64 overflow-y-auto">
                        {searchResults.map((m, i) => (
                          <button key={`search-${i}-${m.모듈코드}`} onClick={() => handleStart(m)}
                            className="w-full text-left rounded-xl border border-gray-200 px-3 py-2.5 hover:border-blue-400 hover:bg-blue-50 transition-colors">
                            <p className="text-xs font-semibold text-black">{m.모듈명}</p>
                            <p className="text-xs text-gray-400 mt-0.5 truncate">
                              {(scriptEdits[m.모듈코드] ?? m.발화목록)[0]}
                            </p>
                          </button>
                        ))}
                      </div>
                    )}
                    {searchQuery.trim() && searchResults.length === 0 && (
                      <p className="mt-3 text-xs text-gray-400">검색 결과가 없습니다.</p>
                    )}
                  </div>
                )}

                {/* 파일 업로드 */}
                <div className="bg-white rounded-2xl border border-gray-300 p-6 mb-6">
                  <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden"
                    onChange={(e) => { if (e.target.files?.[0]) handleFile(e.target.files[0]); e.target.value = ""; }} />

                  {excelLoading ? (
                    <div className="flex items-center justify-center gap-2 py-8 text-sm text-gray-400">
                      <svg className="animate-spin w-4 h-4 text-blue-500" fill="none" viewBox="0 0 24 24">
                        <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                        <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z" />
                      </svg>
                      스크립트 불러오는 중…
                    </div>
                  ) : parsedData ? (
                    <div className="flex items-center justify-between">
                      <div>
                        <p className="text-sm font-semibold text-black">{fileName}</p>
                        <p className="text-xs text-gray-400 mt-0.5">
                          발화 {parsedData.modules.length}개 · 분기 {parsedData.branches.length}개 로드됨
                        </p>
                      </div>
                      <button onClick={() => fileInputRef.current?.click()}
                        className="text-xs text-blue-600 hover:text-blue-800 border border-blue-200 rounded-lg px-3 py-1.5 bg-blue-50 whitespace-nowrap">
                        엑셀 교체
                      </button>
                    </div>
                  ) : (
                    <>
                      <h2 className="text-sm font-semibold text-black uppercase tracking-wide mb-4">엑셀 파일 업로드</h2>
                      <div
                        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
                        onDragLeave={() => setDragging(false)}
                        onDrop={handleDrop}
                        onClick={() => fileInputRef.current?.click()}
                        className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-colors
                          ${dragging ? "border-blue-400 bg-blue-50" : "border-gray-300 hover:border-gray-400 hover:bg-gray-50"}`}
                      >
                        <p className="text-sm text-black mb-1">엑셀 파일을 드래그하거나 클릭해서 업로드</p>
                        <p className="text-xs text-gray-400">.xlsx, .xls 지원 (시트1: 분기, 시트3: 발화)</p>
                      </div>
                    </>
                  )}

                  {fileError && <p className="mt-3 text-sm text-red-500">{fileError}</p>}
                  <button onClick={() => handleStart()} disabled={!parsedData || excelLoading}
                    className="mt-4 w-full py-3 rounded-lg text-sm font-semibold transition-colors
                      bg-blue-600 text-white hover:bg-blue-700 disabled:bg-gray-200 disabled:text-black disabled:cursor-not-allowed">
                    {excelLoading ? "로딩 중…" : parsedData ? "상담 시작" : "파일을 먼저 업로드해주세요"}
                  </button>
                </div>
              </>
            )}

            {/* ── 저장된 발화 탭 ── */}
            {homeTab === "저장" && (
              <div className="space-y-4 mb-6">
                {parsedData && parsedData.modules.some((m) => m.중요) && (
                  <div className="bg-white rounded-2xl border border-yellow-300 p-6">
                    <div className="flex items-center gap-2 mb-4">
                      <span className="text-yellow-400 text-sm">★</span>
                      <h2 className="text-sm font-semibold text-black">엑셀 중요 발화</h2>
                      <span className="ml-auto text-xs text-gray-400">파일에서 자동으로 불러온 중요 발화입니다.</span>
                    </div>
                    <div className="space-y-3">
                      {parsedData.modules.filter((m) => m.중요).map((m, i) => (
                        <div key={`important-${i}-${m.모듈코드}`} className="rounded-xl border border-yellow-100 bg-yellow-50 p-4">
                          <div className="flex items-center gap-2 mb-2">
                            <span className="text-xs font-semibold text-yellow-700">{m.모듈명}</span>
                            {m.모듈코드 && <span className="text-xs text-gray-400">{m.모듈코드}</span>}
                          </div>
                          <div className="space-y-2">
                            {(scriptEdits[m.모듈코드] ?? m.발화목록).map((line, i) => (
                              <p key={i} className="text-sm text-black leading-relaxed whitespace-pre-wrap bg-white rounded-lg p-3 border border-yellow-100">
                                {applyReplacements(line, info)}
                              </p>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {parsedData && parsedData.customImportant.length > 0 && (
                  <div className="bg-white rounded-2xl border border-blue-200 p-6">
                    <div className="flex items-center gap-2 mb-4">
                      <span className="text-blue-400 text-sm">📋</span>
                      <h2 className="text-sm font-semibold text-black">직접 작성한 중요 발화</h2>
                      <span className="ml-auto text-xs text-gray-400">엑셀 &quot;중요발화&quot; 시트</span>
                    </div>
                    <div className="space-y-3">
                      {parsedData.customImportant.map((item, i) => (
                        <div key={i} className="rounded-xl border border-blue-100 bg-blue-50 p-4">
                          {item.제목 && <p className="text-xs font-semibold text-blue-700 mb-2">{item.제목}</p>}
                          <p className="text-sm text-black leading-relaxed whitespace-pre-wrap bg-white rounded-lg p-3 border border-blue-100">
                            {applyReplacements(item.내용, info)}
                          </p>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                <div className="bg-white rounded-2xl border border-gray-300 p-6">
                  <div className="flex items-center justify-between mb-4">
                    <h2 className="text-sm font-semibold text-black">직접 저장한 발화</h2>
                    {savedScripts.length > 0 && (
                      <button onClick={handleClearAll} className="text-xs text-red-400 hover:text-red-600 transition-colors">전체 삭제</button>
                    )}
                  </div>
                  {savedScripts.length === 0 ? (
                    <div className="py-10 text-center">
                      <p className="text-3xl mb-3">★</p>
                      <p className="text-sm text-gray-400">아직 저장된 발화가 없습니다.</p>
                      <p className="text-xs text-gray-400 mt-1">상담 중 ★ 버튼을 눌러 발화를 저장하세요.</p>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      {savedScripts.map((s) => (
                        <div key={s.id} className="rounded-xl border border-gray-200 p-4">
                          <div className="flex items-start justify-between gap-2 mb-2">
                            <div>
                              <span className="text-xs font-semibold text-blue-600">{s.모듈명}</span>
                              {s.모듈코드 && <span className="ml-2 text-xs text-gray-400">{s.모듈코드}</span>}
                            </div>
                            <div className="flex items-center gap-2 flex-shrink-0">
                              <span className="text-xs text-gray-400">{formatDate(s.savedAt)}</span>
                              <button onClick={() => handleDeleteSaved(s.id)}
                                className="text-gray-300 hover:text-red-400 transition-colors text-base leading-none">✕</button>
                            </div>
                          </div>
                          <p className="text-sm text-black leading-relaxed whitespace-pre-wrap bg-gray-50 rounded-lg p-3">{s.발화}</p>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            )}
          </>
        )}

        {/* ── 상담 진행 화면 ── */}
        {started && !done && currentModule && (
          <>
            <div className="flex items-center gap-2 mb-4">
              {history.length > 0 && (
                <button onClick={handleBack}
                  className="text-sm text-black bg-white border border-gray-300 rounded-lg px-3 py-1.5 hover:bg-gray-100 transition-colors whitespace-nowrap">
                  ← 이전 모듈
                </button>
              )}
              <button onClick={() => { setStarted(false); setDone(false); setHistory([]); setEditingModule(null); }}
                className="text-sm text-black bg-white border border-gray-300 rounded-lg px-3 py-1.5 hover:bg-gray-100 transition-colors whitespace-nowrap">
                ⌂ 첫 화면
              </button>
              {history.length > 0 && (
                <p className="text-xs text-black truncate">
                  {history.map((h) => h.module.모듈명).join(" › ")} › <span className="font-semibold">{currentModule.모듈명}</span>
                </p>
              )}
            </div>

            <div className="bg-white rounded-2xl border border-gray-300 overflow-hidden mb-4">
              <div className="bg-blue-600 px-5 py-3 flex items-center justify-between">
                <div>
                  <p className="text-xs text-blue-100">{CF_LABELS[currentModule.cf] ?? currentModule.cf}</p>
                  <p className="text-base font-bold text-white">{currentModule.모듈명}</p>
                  <p className="text-xs text-blue-200">{currentModule.모듈코드}</p>
                </div>
                {(customerName || gender || age) && (
                  <div className="text-right text-xs text-blue-100 space-y-0.5">
                    {customerName && <p>{customerName}님</p>}
                    {gender && age && <p>{gender} · {age}</p>}
                    {gender && !age && <p>{gender}</p>}
                    {!gender && age && <p>{age}</p>}
                  </div>
                )}
              </div>

              <div className="p-5">
                <div className="flex items-center justify-between gap-3 mb-3">
                  <p className="text-xs font-semibold text-black">발화 스크립트</p>
                  <div className="flex items-center gap-3 flex-shrink-0">
                    <button onClick={handleSave} disabled={isAlreadySaved}
                      title={isAlreadySaved ? "이미 저장됨" : "발화 저장"}
                      className={`text-sm transition-colors ${isAlreadySaved ? "text-yellow-400 cursor-default" : "text-gray-300 hover:text-yellow-400"}`}>★</button>
                    <button onClick={() => isEditing ? setEditingModule(null) : startEditing(currentModule)}
                      className={`text-xs transition-colors ${isEditing ? "text-blue-600 font-semibold" : "text-gray-400 hover:text-blue-600"}`}>
                      ✏️ {isEditing ? "취소" : "편집"}
                    </button>
                    {!isEditing && (
                      <button onClick={handleReroll} className="text-xs text-blue-600 hover:text-blue-800 whitespace-nowrap">↻ 다시 뽑기</button>
                    )}
                  </div>
                </div>

                {isEditing ? (
                  <div className="space-y-2">
                    {editLines.map((line, i) => (
                      <div key={i} className="flex gap-2">
                        <textarea value={line} rows={3} onChange={(e) => {
                            const next = [...editLines]; next[i] = e.target.value; setEditLines(next);
                          }}
                          className="flex-1 rounded-lg border border-gray-300 px-3 py-2 text-sm text-black focus:outline-none focus:ring-2 focus:ring-blue-400 bg-white resize-none" />
                        <button onClick={() => setEditLines(editLines.filter((_, j) => j !== i))}
                          className="text-gray-300 hover:text-red-400 text-lg self-start mt-1">✕</button>
                      </div>
                    ))}
                    <button onClick={() => setEditLines([...editLines, ""])}
                      className="text-xs text-blue-600 hover:text-blue-800">+ 발화 추가</button>
                    <div className="flex gap-2 mt-2">
                      <button onClick={saveEdits}
                        className="flex-1 py-2 rounded-lg bg-blue-600 text-white text-sm font-semibold hover:bg-blue-700">저장</button>
                      <button onClick={resetModuleEdits}
                        className="px-3 py-2 rounded-lg border border-gray-300 text-sm text-gray-500 hover:text-black">원본으로</button>
                    </div>
                  </div>
                ) : (
                  <p className="text-sm text-black leading-relaxed whitespace-pre-wrap bg-gray-50 rounded-lg p-4">
                    {currentScript}
                    {scriptEdits[currentModule.모듈코드] && (
                      <span className="ml-2 text-xs text-blue-400">(수정됨)</span>
                    )}
                  </p>
                )}
              </div>
            </div>

            {/* 이미지 섹션 */}
            {currentModuleImages.length > 0 && (
              <div className="bg-white rounded-2xl border border-gray-300 overflow-hidden mb-4">
                <div className="px-5 py-3 border-b border-gray-100">
                  <p className="text-xs font-semibold text-black">
                    🖼 참고 이미지 <span className="text-gray-400">({currentModuleImages.length}장)</span>
                  </p>
                </div>
                <div className="p-4 flex gap-3 overflow-x-auto">
                  {currentModuleImages.map((img, i) => (
                    <div key={i} className="relative flex-shrink-0 group">
                      <img src={img.src} alt={img.name} onClick={() => setLightboxImage(img)}
                        className="h-28 w-auto rounded-xl border border-gray-200 cursor-pointer hover:opacity-90 transition-opacity object-cover" />
                      <button onClick={() => handleCopyImage(img.src)}
                        className="absolute bottom-6 left-1/2 -translate-x-1/2 opacity-0 group-hover:opacity-100 transition-opacity bg-black/60 text-white text-xs rounded-lg px-2 py-1 whitespace-nowrap">
                        {copiedImageSrc === img.src ? "✓ 복사됨" : "복사"}
                      </button>
                      <p className="text-xs text-gray-400 mt-1 truncate max-w-24">{img.name}</p>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div className="bg-white rounded-2xl border border-gray-300 p-5">
              {currentBranches.length > 0 && (
                <>
                  <p className="text-xs font-semibold text-black mb-3">고객 반응은?</p>
                  <div className="space-y-2 mb-3">
                    {currentBranches.map((b, i) => (
                      <button key={i} onClick={() => handleResponse(b)}
                        className="w-full text-left rounded-xl border border-gray-200 px-4 py-3 hover:border-blue-400 hover:bg-blue-50 transition-colors group">
                        <p className="text-sm font-semibold text-black group-hover:text-blue-700">{b.응답유형}</p>
                        <p className="text-xs text-black mt-0.5">
                          → {b.다음모듈}
                          {b.비고 && <span className="ml-2 text-gray-400">({b.비고})</span>}
                        </p>
                      </button>
                    ))}
                  </div>
                  <hr className="border-gray-200 mb-3" />
                </>
              )}
              <button onClick={handleNext}
                className="w-full py-3 rounded-xl bg-blue-600 text-white text-sm font-semibold hover:bg-blue-700 transition-colors">
                다음 모듈로 →
              </button>
            </div>
          </>
        )}

        {/* ── 상담 완료 화면 ── */}
        {started && done && (
          <div className="bg-white rounded-2xl border border-gray-300 p-10 text-center">
            <p className="text-2xl mb-2">🎉</p>
            <p className="text-lg font-bold text-black mb-1">상담 완료</p>
            {customerName && <p className="text-sm text-black mb-6">{customerName}님과의 상담이 마무리되었습니다.</p>}
            <div className="flex gap-3 justify-center">
              <button onClick={handleBack}
                className="px-5 py-2 rounded-lg border border-gray-300 text-sm font-medium text-black bg-white hover:bg-gray-100">
                ← 이전으로
              </button>
              <button onClick={() => { setStarted(false); setDone(false); setHistory([]); }}
                className="px-5 py-2 rounded-lg bg-blue-600 text-white text-sm font-semibold hover:bg-blue-700">
                새 상담 시작
              </button>
            </div>
          </div>
        )}

      </div>

      {/* 라이트박스 */}
      {lightboxImage && (
        <div onClick={() => setLightboxImage(null)}
          className="fixed inset-0 z-50 bg-black/80 flex items-center justify-center p-4">
          <div className="relative max-w-4xl max-h-full" onClick={(e) => e.stopPropagation()}>
            <img src={lightboxImage.src} alt={lightboxImage.name}
              className="max-w-full max-h-[85vh] rounded-xl object-contain" />
            <div className="flex items-center justify-center gap-3 mt-2">
              <p className="text-white text-sm opacity-70">{lightboxImage.name}</p>
              <button onClick={() => handleCopyImage(lightboxImage.src)}
                className="text-xs bg-white/20 hover:bg-white/30 text-white rounded-lg px-3 py-1.5 transition-colors">
                {copiedImageSrc === lightboxImage.src ? "✓ 복사됨" : "📋 복사"}
              </button>
            </div>
            <button onClick={() => setLightboxImage(null)}
              className="absolute -top-3 -right-3 w-8 h-8 rounded-full bg-white text-black text-sm font-bold flex items-center justify-center hover:bg-gray-200">
              ✕
            </button>
          </div>
        </div>
      )}

    </main>
  );
}
