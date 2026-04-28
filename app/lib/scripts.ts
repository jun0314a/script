export type AgeGroup = "20대" | "30대" | "40대" | "50대" | "60대 이상";
export type Gender = "남성" | "여성";
export type HasInsurance = "있음" | "없음";

export interface ScriptSection {
  title: string;
  lines: string[];
}

export interface ConsultationScript {
  opening: ScriptSection;
  needsFinding: ScriptSection;
  proposal: ScriptSection;
  closing: ScriptSection;
}

function getOpening(gender: Gender, hasInsurance: HasInsurance): ScriptSection {
  const suffix = gender === "남성" ? "고객님" : "고객님";
  return {
    title: "1단계 — 인사 및 소개",
    lines: [
      `안녕하세요, ${suffix}. 저는 ○○보험 설계사 [이름]입니다.`,
      "오늘 잠깐 시간 내주셔서 감사합니다.",
      hasInsurance === "있음"
        ? "현재 가입하신 보험이 있다고 하셨는데, 오늘은 기존 보장 내용을 함께 점검하고 혹시 빠진 부분이 있는지 살펴보는 시간으로 진행하겠습니다."
        : "오늘은 고객님께 꼭 필요한 보장이 무엇인지 함께 알아보는 시간으로 진행하겠습니다.",
      "편하게 말씀해 주시면 됩니다.",
    ],
  };
}

function getNeedsFinding(
  age: AgeGroup,
  gender: Gender,
  hasInsurance: HasInsurance
): ScriptSection {
  const lines: string[] = [];

  if (hasInsurance === "있음") {
    lines.push("현재 어떤 보험에 가입되어 계신지 여쭤봐도 될까요?");
    lines.push("보험료는 한 달에 어느 정도 내고 계세요?");
    lines.push("혹시 보장 내용을 직접 확인해 보신 적이 있으신가요?");
  } else {
    lines.push("현재 따로 가입하신 보험은 없으신 거죠?");
    lines.push("혹시 보험 가입을 생각해 보신 적은 있으셨나요?");
  }

  if (age === "20대") {
    lines.push("지금 직장이나 학교 생활은 어떻게 하고 계세요?");
    lines.push("혹시 부모님께서 가입해 주신 보험이 있으신지도 확인이 필요할 것 같습니다.");
  } else if (age === "30대") {
    lines.push("결혼하셨거나 자녀가 계신가요?");
    lines.push("가족분들 보장도 함께 챙기고 싶으신 부분이 있으실 것 같아서요.");
  } else if (age === "40대") {
    lines.push("혹시 건강검진 결과에서 주의 소견을 받으신 게 있으신가요?");
    lines.push("요즘 몸 관리하시는 데 특별히 신경 쓰시는 부분이 있으세요?");
  } else if (age === "50대") {
    lines.push("가족 중에 큰 병을 앓으셨던 분이 계신가요?");
    lines.push("노후 준비나 의료비 걱정이 크신 편이세요?");
  } else {
    lines.push("지금 건강 상태는 어떠세요? 정기적으로 병원을 다니시는 편인가요?");
    lines.push("의료비가 부담이 되실 때 가장 걱정되시는 부분이 어떤 건가요?");
  }

  if (gender === "여성") {
    lines.push("여성 특화 질환이나 여성암에 대한 보장도 함께 살펴보면 좋을 것 같습니다.");
  }

  return {
    title: "2단계 — 니즈 파악 (질문)",
    lines,
  };
}

function getProposal(
  age: AgeGroup,
  gender: Gender,
  hasInsurance: HasInsurance
): ScriptSection {
  const lines: string[] = [];

  if (hasInsurance === "있음") {
    lines.push("말씀해 주신 내용을 바탕으로 현재 보장에서 빠진 부분을 정리해 드릴게요.");
    lines.push("보통 기존 보험에서 놓치기 쉬운 부분이 [실손 의료비 / 진단금 / 수술비] 인데요,");
    lines.push("고객님 경우에는 특히 [해당 항목]이 미흡하실 수 있습니다.");
    lines.push("이 부분을 채워드릴 수 있는 상품으로 [상품명]을 추천드립니다.");
  } else {
    lines.push("처음 보험을 준비하실 때는 너무 많은 걸 한 번에 넣으려다 보면 보험료 부담이 커질 수 있어요.");
    lines.push("가장 우선적으로 챙겨야 할 3가지는 실손 의료비, 암 진단금, 입원비입니다.");
    lines.push("이 3가지를 기본으로 하되, 고객님 상황에 맞게 구성해 드릴게요.");
  }

  if (age === "20대") {
    lines.push("20대는 보험료가 가장 저렴한 시기입니다. 지금 가입하시면 같은 보장을 훨씬 낮은 보험료로 유지하실 수 있어요.");
  } else if (age === "30대") {
    lines.push("30대는 가족 보장도 중요한 시기입니다. 개인 보장과 함께 가족 관련 특약도 살펴보겠습니다.");
  } else if (age === "40대") {
    lines.push("40대부터는 성인병 및 3대 질병 진단비 보장이 더욱 중요해집니다. 이 부분을 중심으로 제안드릴게요.");
  } else if (age === "50대") {
    lines.push("50대에는 보험료가 올라가는 만큼 꼭 필요한 보장 위주로 효율적으로 구성하는 게 중요합니다.");
  } else {
    lines.push("노후에는 의료비가 가장 큰 부담이 될 수 있습니다. 실손과 간병 관련 보장을 중심으로 설명드릴게요.");
  }

  if (gender === "여성") {
    lines.push("여성분들께는 유방암·자궁암 등 여성 특화 진단비 특약도 함께 추천드립니다.");
  } else {
    lines.push("남성분들께는 뇌혈관·심장 질환 진단비 특약을 함께 추천드립니다.");
  }

  lines.push("예상 보험료는 월 약 [금액]원 수준이고, 고객님 상황에 맞게 조정도 가능합니다.");

  return {
    title: "3단계 — 상품 제안",
    lines,
  };
}

function getClosing(hasInsurance: HasInsurance): ScriptSection {
  return {
    title: "4단계 — 마무리 및 다음 단계 안내",
    lines: [
      "오늘 말씀드린 내용을 정리해서 서면으로 보내드릴 수 있습니다.",
      hasInsurance === "있음"
        ? "현재 보험 증권을 가져오시면 더 정확하게 분석해 드릴 수 있습니다."
        : "가입 의향이 있으시다면 건강 고지 사항 몇 가지만 확인하면 바로 진행 가능합니다.",
      "궁금하신 점이나 걱정되시는 부분이 있으시면 편하게 여쭤봐 주세요.",
      "오늘 시간 내주셔서 감사합니다. 좋은 결정 하시도록 최선을 다해 도와드리겠습니다.",
    ],
  };
}

export function generateScript(
  age: AgeGroup,
  gender: Gender,
  hasInsurance: HasInsurance
): ConsultationScript {
  return {
    opening: getOpening(gender, hasInsurance),
    needsFinding: getNeedsFinding(age, gender, hasInsurance),
    proposal: getProposal(age, gender, hasInsurance),
    closing: getClosing(hasInsurance),
  };
}
