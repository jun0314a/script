"use client";

import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── 타입 ────────────────────────────────────────────────────────────────────

interface Module {
  cf: string;
  상위모듈: string;
  모듈코드: string;
  모듈명: string;
  발화목록: string[];
}

interface Branch {
  cf: string;
  모듈명: string;
  응답유형: string;
  다음모듈: string;
  비고: string;
}

interface ParsedData {
  modules: Module[];
  branches: Branch[];
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
}

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

function parseExcel(file: File): Promise<ParsedData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target!.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });

        // ── 시트3: 발화 스크립트 ──
        const scriptSheet = wb.Sheets[wb.SheetNames[2]];
        const scriptRows = XLSX.utils.sheet_to_json<Record<string, string>>(scriptSheet, { defval: "" });
        const scriptCols = Object.keys(scriptRows[0] || {}).filter(
          (k) => !["CF", "상위 모듈", "모듈 코드", "모듈명"].includes(k)
        );
        const modules: Module[] = scriptRows
          .filter((r) => r["모듈명"])
          .map((r) => ({
            cf: r["CF"] || "",
            상위모듈: r["상위 모듈"] || "",
            모듈코드: r["모듈 코드"] || "",
            모듈명: r["모듈명"] || "",
            발화목록: scriptCols.map((c) => r[c]?.toString().trim()).filter(Boolean),
          }))
          .filter((m) => m.발화목록.length > 0);

        // ── 시트1: 분기 테이블 ──
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

        resolve({ modules, branches });
      } catch {
        reject(new Error("파일을 읽는 중 오류가 발생했습니다."));
      }
    };
    reader.onerror = () => reject(new Error("파일 읽기 실패"));
    reader.readAsArrayBuffer(file);
  });
}

// ─── 모듈 검색 (유사 이름 매칭) ──────────────────────────────────────────────

function findModule(targetName: string, modules: Module[]): Module | undefined {
  // "이탈 방어 → 니즈 선택" 형태면 → 이후 부분을 실제 목적지로 사용
  const name = targetName.includes("→")
    ? targetName.split("→").pop()!.trim()
    : targetName.trim();

  return (
    modules.find((m) => m.모듈명 === name) ||
    modules.find((m) => m.모듈명.includes(name) || name.includes(m.모듈명))
  );
}

function getBranches(moduleName: string, branches: Branch[]): Branch[] {
  return branches.filter(
    (b) => b.모듈명 === moduleName || moduleName.includes(b.모듈명) || b.모듈명.includes(moduleName)
  );
}

function pickScript(module: Module, info: CustomerInfo): string {
  const line = module.발화목록[Math.floor(Math.random() * module.발화목록.length)];
  return applyReplacements(line, info);
}

// ─── 컴포넌트 ─────────────────────────────────────────────────────────────────

export default function Home() {
  // 파일 & 파싱
  const [parsedData, setParsedData] = useState<ParsedData | null>(null);
  const [fileName, setFileName] = useState("");
  const [fileError, setFileError] = useState("");
  const [dragging, setDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // 고객 정보
  const [customerName, setCustomerName] = useState("");
  const [gender, setGender] = useState("");
  const [age, setAge] = useState("");
  const [consultantName, setConsultantName] = useState("");
  const [renewalType, setRenewalType] = useState("");
  const [coverageRange, setCoverageRange] = useState("");
  const [coverageAmount, setCoverageAmount] = useState("");
  const [coveragePeriod, setCoveragePeriod] = useState("");

  const info: CustomerInfo = {
    customerName, gender, age, consultantName,
    renewalType, coverageRange, coverageAmount, coveragePeriod,
  };

  // 상담 진행 상태
  const [started, setStarted] = useState(false);
  const [currentModule, setCurrentModule] = useState<Module | null>(null);
  const [currentScript, setCurrentScript] = useState("");
  const [history, setHistory] = useState<{ module: Module; script: string }[]>([]);
  const [done, setDone] = useState(false);

  // ── 파일 업로드 ──
  const handleFile = useCallback(async (file: File) => {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
      setFileError("엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.");
      return;
    }
    setFileError("");
    try {
      const data = await parseExcel(file);
      setParsedData(data);
      setFileName(file.name);
      setStarted(false);
      setDone(false);
    } catch (e) {
      setFileError((e as Error).message);
    }
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragging(false);
      if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
    },
    [handleFile]
  );

  // ── 상담 시작 ──
  function handleStart() {
    if (!parsedData) return;
    const first = parsedData.modules[0];
    if (!first) return;
    setCurrentModule(first);
    setCurrentScript(pickScript(first, info));
    setHistory([]);
    setDone(false);
    setStarted(true);
  }

  // ── 고객 응답 선택 → 다음 모듈로 ──
  function handleResponse(branch: Branch) {
    if (!parsedData || !currentModule) return;

    // 현재 모듈을 히스토리에 저장
    setHistory((h) => [...h, { module: currentModule, script: currentScript }]);

    const nextName = branch.다음모듈;

    // "상담 종료" 계열 키워드면 완료 처리
    if (
      nextName.includes("종료") ||
      nextName.includes("CF1부터") ||
      currentModule.모듈명 === "마무리 인사"
    ) {
      setDone(true);
      return;
    }

    const next = findModule(nextName, parsedData.modules);
    if (!next) {
      // 매칭되는 모듈이 없으면 순서상 다음 모듈로
      const idx = parsedData.modules.findIndex((m) => m.모듈명 === currentModule.모듈명);
      const fallback = parsedData.modules[idx + 1];
      if (fallback) {
        setCurrentModule(fallback);
        setCurrentScript(pickScript(fallback, info));
      } else {
        setDone(true);
      }
      return;
    }

    setCurrentModule(next);
    setCurrentScript(pickScript(next, info));
  }

  // ── 분기 없을 때 다음 모듈 순서 진행 ──
  function handleNext() {
    if (!parsedData || !currentModule) return;
    setHistory((h) => [...h, { module: currentModule, script: currentScript }]);
    const idx = parsedData.modules.findIndex((m) => m.모듈명 === currentModule.모듈명);
    const next = parsedData.modules[idx + 1];
    if (next) {
      setCurrentModule(next);
      setCurrentScript(pickScript(next, info));
    } else {
      setDone(true);
    }
  }

  // ── 뒤로 가기 ──
  function handleBack() {
    if (history.length === 0) return;
    const prev = history[history.length - 1];
    setHistory((h) => h.slice(0, -1));
    setCurrentModule(prev.module);
    setCurrentScript(prev.script);
    setDone(false);
  }

  // ── 발화 다시 뽑기 ──
  function handleReroll() {
    if (!currentModule) return;
    setCurrentScript(pickScript(currentModule, info));
  }

  // ── 현재 모듈의 분기 목록 ──
  const currentBranches = parsedData && currentModule
    ? getBranches(currentModule.모듈명, parsedData.branches)
    : [];

  // ─── 렌더 ───────────────────────────────────────────────────────────────────

  return (
    <main className="min-h-screen py-10 px-4">
      <div className="max-w-2xl mx-auto">

        {/* ── 헤더 ── */}
        <div className="mb-8 text-center">
          <h1 className="text-2xl font-bold text-black mb-1">보험 상담 스크립트</h1>
          <p className="text-sm text-black">고객 반응에 따라 다음 발화를 안내해드립니다.</p>
        </div>

        {/* ── 준비 화면 ── */}
        {!started && (
          <>
            {/* 고객 & 설계사 정보 */}
            <div className="bg-white rounded-2xl border border-gray-300 p-6 mb-4">
              <h2 className="text-sm font-semibold text-black uppercase tracking-wide mb-4">상담 정보 입력</h2>
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
                    <option value="비갱신형">비갱신형</option>
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-medium text-black mb-1">보장범위</label>
                  <select value={coverageRange} onChange={(e) => setCoverageRange(e.target.value)} className={INPUT_CLS}>
                    <option value="">선택 안함</option>
                    <option value="기본형">기본형</option>
                    <option value="표준형">표준형</option>
                    <option value="광범위형">광범위형</option>
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-medium text-black mb-1">보장금액</label>
                  <select value={coverageAmount} onChange={(e) => setCoverageAmount(e.target.value)} className={INPUT_CLS}>
                    <option value="">선택 안함</option>
                    <option value="소액">소액</option>
                    <option value="중액">중액</option>
                    <option value="고액">고액</option>
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-medium text-black mb-1">보장기간</label>
                  <select value={coveragePeriod} onChange={(e) => setCoveragePeriod(e.target.value)} className={INPUT_CLS}>
                    <option value="">선택 안함</option>
                    <option value="10년">10년</option>
                    <option value="20년">20년</option>
                    <option value="30년">30년</option>
                    <option value="80세 만기">80세 만기</option>
                    <option value="90세 만기">90세 만기</option>
                    <option value="100세 만기">100세 만기</option>
                  </select>
                </div>
              </div>
            </div>

            {/* 파일 업로드 */}
            <div className="bg-white rounded-2xl border border-gray-300 p-6 mb-6">
              <h2 className="text-sm font-semibold text-black uppercase tracking-wide mb-4">엑셀 파일 업로드</h2>
              <div
                onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
                onDragLeave={() => setDragging(false)}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
                className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-colors
                  ${dragging ? "border-blue-400 bg-blue-50" : "border-gray-300 hover:border-gray-400 hover:bg-gray-50"}`}
              >
                <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden"
                  onChange={(e) => { if (e.target.files?.[0]) handleFile(e.target.files[0]); }} />
                {fileName ? (
                  <>
                    <p className="text-sm font-semibold text-black mb-1">{fileName}</p>
                    <p className="text-xs text-black">
                      발화 {parsedData?.modules.length}개 · 분기 {parsedData?.branches.length}개 로드됨 — 클릭하면 다시 업로드
                    </p>
                  </>
                ) : (
                  <>
                    <p className="text-sm text-black mb-1">엑셀 파일을 드래그하거나 클릭해서 업로드</p>
                    <p className="text-xs text-black">.xlsx, .xls 지원 (시트1: 분기, 시트3: 발화)</p>
                  </>
                )}
              </div>
              {fileError && <p className="mt-3 text-sm text-red-500">{fileError}</p>}

              <button
                onClick={handleStart}
                disabled={!parsedData}
                className="mt-5 w-full py-3 rounded-lg text-sm font-semibold transition-colors
                  bg-blue-600 text-white hover:bg-blue-700 disabled:bg-gray-200 disabled:text-black disabled:cursor-not-allowed"
              >
                {parsedData ? "상담 시작" : "파일을 먼저 업로드해주세요"}
              </button>
            </div>
          </>
        )}

        {/* ── 상담 진행 화면 ── */}
        {started && !done && currentModule && (
          <>
            {/* 상단 바 */}
            <div className="flex items-center justify-between mb-4">
              <button
                onClick={history.length > 0 ? handleBack : () => setStarted(false)}
                className="text-sm text-black bg-white border border-gray-300 rounded-lg px-3 py-1.5 hover:bg-gray-100 transition-colors"
              >
                ← {history.length > 0 ? "이전 모듈" : "처음으로"}
              </button>

              {/* 진행 히스토리 요약 */}
              {history.length > 0 && (
                <p className="text-xs text-black">
                  {history.map((h) => h.module.모듈명).join(" › ")} › <span className="font-semibold">{currentModule.모듈명}</span>
                </p>
              )}
            </div>

            {/* 모듈 카드 */}
            <div className="bg-white rounded-2xl border border-gray-300 overflow-hidden mb-4">
              {/* 모듈 헤더 */}
              <div className="bg-blue-600 px-5 py-3 flex items-center justify-between">
                <div>
                  <p className="text-xs text-blue-100">{CF_LABELS[currentModule.cf] ?? currentModule.cf}</p>
                  <p className="text-base font-bold text-white">{currentModule.모듈명}</p>
                  <p className="text-xs text-blue-200">{currentModule.모듈코드}</p>
                </div>
                {/* 고객 정보 배지 */}
                {(customerName || gender || age) && (
                  <div className="text-right text-xs text-blue-100 space-y-0.5">
                    {customerName && <p>{customerName}님</p>}
                    {gender && age && <p>{gender} · {age}</p>}
                    {gender && !age && <p>{gender}</p>}
                    {!gender && age && <p>{age}</p>}
                  </div>
                )}
              </div>

              {/* 발화 */}
              <div className="p-5">
                <div className="flex items-start justify-between gap-3 mb-1">
                  <p className="text-xs font-semibold text-black">발화 스크립트</p>
                  <button
                    onClick={handleReroll}
                    title="다른 발화로 다시 뽑기"
                    className="text-xs text-blue-600 hover:text-blue-800 whitespace-nowrap flex-shrink-0"
                  >
                    ↻ 다시 뽑기
                  </button>
                </div>
                <p className="text-sm text-black leading-relaxed whitespace-pre-wrap bg-gray-50 rounded-lg p-4">
                  {currentScript}
                </p>
              </div>
            </div>

            {/* 고객 응답 선택 */}
            <div className="bg-white rounded-2xl border border-gray-300 p-5">
              <p className="text-xs font-semibold text-black mb-3">
                {currentBranches.length > 0 ? "고객 반응은?" : "다음 단계"}
              </p>

              {currentBranches.length > 0 ? (
                <div className="space-y-2">
                  {currentBranches.map((b, i) => (
                    <button
                      key={i}
                      onClick={() => handleResponse(b)}
                      className="w-full text-left rounded-xl border border-gray-200 px-4 py-3 hover:border-blue-400 hover:bg-blue-50 transition-colors group"
                    >
                      <p className="text-sm font-semibold text-black group-hover:text-blue-700">
                        {b.응답유형}
                      </p>
                      <p className="text-xs text-black mt-0.5">
                        → {b.다음모듈}
                        {b.비고 && <span className="ml-2 text-gray-400">({b.비고})</span>}
                      </p>
                    </button>
                  ))}
                </div>
              ) : (
                <button
                  onClick={handleNext}
                  className="w-full py-3 rounded-xl bg-blue-600 text-white text-sm font-semibold hover:bg-blue-700 transition-colors"
                >
                  다음 모듈로 →
                </button>
              )}
            </div>
          </>
        )}

        {/* ── 상담 완료 화면 ── */}
        {started && done && (
          <div className="bg-white rounded-2xl border border-gray-300 p-10 text-center">
            <p className="text-2xl mb-2">🎉</p>
            <p className="text-lg font-bold text-black mb-1">상담 완료</p>
            {customerName && (
              <p className="text-sm text-black mb-6">{customerName}님과의 상담이 마무리되었습니다.</p>
            )}
            <div className="flex gap-3 justify-center">
              <button
                onClick={handleBack}
                className="px-5 py-2 rounded-lg border border-gray-300 text-sm font-medium text-black bg-white hover:bg-gray-100"
              >
                ← 이전으로
              </button>
              <button
                onClick={() => { setStarted(false); setDone(false); setHistory([]); }}
                className="px-5 py-2 rounded-lg bg-blue-600 text-white text-sm font-semibold hover:bg-blue-700"
              >
                새 상담 시작
              </button>
            </div>
          </div>
        )}

      </div>
    </main>
  );
}
