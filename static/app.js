const jdInput = document.querySelector("#jdInput");
const customInput = document.querySelector("#customInput");
const resumeInput = document.querySelector("#resumeInput");
const sectorSelect = document.querySelector("#sector");
const analyzeBtn = document.querySelector("#analyzeBtn");
const webBtn = document.querySelector("#webBtn");
const downloadBtn = document.querySelector("#downloadBtn");
const downloadEnBtn = document.querySelector("#downloadEnBtn");
const translationStatus = document.querySelector("#translationStatus");
const resumeFile = document.querySelector("#resumeFile");
const jdCount = document.querySelector("#jdCount");
const sectorBadge = document.querySelector("#sectorBadge");
const keywords = document.querySelector("#keywords");
const adviceList = document.querySelector("#adviceList");
const resumePreview = document.querySelector("#resumePreview");
const downloadLink = document.querySelector("#downloadLink");
const downloadEnLink = document.querySelector("#downloadEnLink");
const previewLang = document.querySelector("#previewLang");
const materialFiles = document.querySelector("#materialFiles");
const materialList = document.querySelector("#materialList");
const toast = document.querySelector("#toast");

const sampleJD = `岗位方向：律师事务所 / 金融机构 / 国央企法务
请粘贴真实 JD，例如：
1. 负责合同审核、法律检索、合规风险识别及法律文书起草；
2. 参与投融资、并购、债务重组、资本市场项目尽职调查；
3. 具备英文读写能力、材料归纳能力、Excel/Word 熟练；
4. 有律所、金融机构、政府机关或国央企实习经验优先。`;

let latestResume = null;
let resumeOverride = null;
let uploadedMaterials = [];
let englishResumeOverride = null;

jdInput.value = sampleJD;
updateCount();
renderEmpty();
analyze();
checkTranslationStatus();

jdInput.addEventListener("input", () => {
  updateCount();
  resumeOverride = null;
  englishResumeOverride = null;
});
customInput.addEventListener("input", () => {
  resumeOverride = null;
  englishResumeOverride = null;
});

analyzeBtn.addEventListener("click", analyze);
webBtn.addEventListener("click", runWebResearch);
sectorSelect.addEventListener("change", analyze);
downloadBtn.addEventListener("click", generateDocx);
downloadEnBtn.addEventListener("click", generateEnglishDocx);
materialFiles.addEventListener("change", uploadMaterials);
previewLang.addEventListener("change", async () => {
  if (!latestResume) return;
  if (previewLang.value === "en") {
    await translatePreview();
  } else {
    renderResume(latestResume);
  }
});

resumeFile.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  if (file.name.endsWith(".txt")) {
    resumeInput.value = await file.text();
    resumeOverride = null;
    showToast("已读取文本简历，正在生成预览");
    analyze();
    return;
  }
  if (!file.name.endsWith(".docx")) {
    showToast("目前支持 .docx 和 .txt");
    return;
  }
  const data = await readFileAsDataURL(file);
  const res = await fetch("/api/extract", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: file.name, data }),
  });
  const json = await res.json();
  if (json.text) {
    resumeInput.value = json.text;
    resumeOverride = null;
    showToast("已读取 Word 简历，正在生成预览");
    analyze();
  } else {
    showToast("读取失败");
  }
});

function updateCount() {
  jdCount.textContent = `${jdInput.value.trim().length} 字`;
}

async function analyze() {
  setBusy(analyzeBtn, true, "分析中");
  try {
    const payload = collectPayload();
    const json = await postJSON("/api/analyze", payload);
    latestResume = json.resume;
    resumeOverride = json.resume;
    englishResumeOverride = null;
    renderAdvice(json.resume);
    if (json.answer) appendAdviceCard(json.mode === "model" ? "AI 改写状态" : "生成状态", json.answer);
    renderResume(json.resume);
    downloadLink.hidden = true;
    showToast(json.mode === "model" ? "修改意见已生成" : "修改意见已生成（本地规则）");
  } catch (error) {
    showToast(error.message || "分析失败");
  } finally {
    setBusy(analyzeBtn, false, "生成修改意见");
  }
}

async function runWebResearch() {
  setBusy(webBtn, true, "检索中");
  try {
    const payload = collectPayload();
    const json = await postJSON("/api/web-research", payload);
    latestResume = json.resume;
    englishResumeOverride = null;
    resumeOverride = json.resume;
    renderAdvice(json.resume);
    appendAdviceCard(json.resume.mode === "model" ? "联网检索 + AI 改写状态" : "联网检索状态", json.answer || "已联网检索并完成改写");
    renderResume(json.resume);
    showToast(json.resume.mode === "model" ? "联网检索后已用 AI 改写" : "联网检索完成（本地规则）");
  } catch (error) {
    appendAdviceCard("联网检索状态", error.message || "联网检索失败");
    showToast(error.message || "联网检索失败");
  } finally {
    setBusy(webBtn, false, "联网检索并改写");
  }
}

async function generateDocx() {
  setBusy(downloadBtn, true, "生成中");
  try {
    const json = await postJSON("/api/generate", collectPayload());
    latestResume = json.resume;
    renderAdvice(json.resume);
    renderResume(json.resume);
    downloadLink.href = json.url;
    downloadLink.hidden = false;
    downloadLink.textContent = "下载中文";
    showToast("中文 Word 简历已生成");
  } catch (error) {
    showToast(error.message || "生成失败");
  } finally {
    setBusy(downloadBtn, false, "生成中文简历");
  }
}

async function generateEnglishDocx() {
  setBusy(downloadEnBtn, true, "生成中");
  try {
    const json = await postJSON("/api/generate-en", collectPayload());
    if (json.sourceResume) latestResume = json.sourceResume;
    englishResumeOverride = json.resume;
    previewLang.value = "en";
    renderAdvice(latestResume || json.resume);
    renderResume(englishResumeOverride);
    downloadEnLink.href = json.url;
    downloadEnLink.hidden = false;
    downloadEnLink.textContent = "下载英文";
    showToast(json.translationMode === "model" ? "英文 Word 已用模型翻译生成" : "英文 Word 已用内置专业翻译生成");
  } catch (error) {
    showToast(error.message || "生成失败");
  } finally {
    setBusy(downloadEnBtn, false, "生成英文简历");
  }
}

async function checkTranslationStatus() {
  try {
    const res = await fetch("/api/translation-status");
    const json = await res.json();
    if (json.configured) {
      translationStatus.textContent = "AI 改写引擎已连接";
      translationStatus.classList.add("connected");
    } else {
      translationStatus.textContent = "AI 改写引擎未配置（请联系管理员）";
      translationStatus.classList.remove("connected");
    }
  } catch (error) {
    translationStatus.textContent = "AI 改写引擎状态未知";
  }
}

function deepSeekPrompt(baseResume, payload) {
  return {
    JD: payload.jd || "",
    修改要求: payload.custom || "",
    必须执行修改要求: Boolean((payload.custom || "").trim()),
    原简历文本: (payload.resumeText || "").slice(0, 12000),
    上传补充材料: (payload.uploads || []).map((item) => item.text || "").join("\n").slice(0, 12000),
    当前基础简历JSON: baseResume,
    事实规则: "不得虚构经历、客户名称、项目名称或数据。若原简历文本非空，只能使用该候选人的事实，不得混入不属于该候选人的默认/模板经历。修改要求优先级最高，只要不违背事实，必须写进resume正文。",
    输出格式: {
      summary: "一句话说明如何执行JD和修改要求",
      resume: {
        profile: { name: "姓名", contact: "联系方式", education: [["学校", "专业/学位", "时间", ["细节"]],], honors: "荣誉" },
        experiences: [{ org: "机构", role: "岗位", date: "时间", bullets: [["标签：", "正文"]] }],
        competitions: [["标题", "时间", "内容"]],
        research: [["标题", "时间", "内容"]],
        opinions: [{ title: "手动修改要求如何执行", body: "说明具体改到了哪些正文条目" }],
      },
    },
  };
}

function extractJSON(text) {
  const clean = String(text || "").replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim();
  try {
    return JSON.parse(clean);
  } catch (error) {
    const start = clean.indexOf("{");
    const end = clean.lastIndexOf("}");
    if (start >= 0 && end > start) return JSON.parse(clean.slice(start, end + 1));
    throw error;
  }
}

async function translatePreview() {
  if (!latestResume) return;
  if (englishResumeOverride) {
    renderResume(englishResumeOverride);
    return;
  }
  try {
    const json = await postJSON("/api/translate-en", { resume: latestResume });
    englishResumeOverride = json.resume;
    renderResume(englishResumeOverride);
    showToast(json.mode === "model" ? "英文预览已用模型翻译" : "英文预览已用内置专业翻译");
  } catch (error) {
    renderResume(latestResume);
    showToast(error.message || "英文翻译失败，已使用兜底预览");
  }
}

function collectPayload() {
  return {
    jd: jdInput.value,
    custom: customInput.value,
    resumeText: resumeInput.value,
    sector: sectorSelect.value,
    uploads: uploadedMaterials,
    resumeOverride,
  };
}

async function uploadMaterials(event) {
  const files = Array.from(event.target.files || []);
  if (!files.length) return;
  setMaterialStatus("正在读取上传材料...");
  for (const file of files) {
    try {
      const data = await readFileAsDataURL(file);
      const item = await postJSON("/api/upload", { name: file.name, data });
      uploadedMaterials.push(item);
    } catch (error) {
      uploadedMaterials.push({ name: file.name, text: `[读取失败：${file.name}]` });
    }
  }
  renderMaterials();
  showToast("补充材料已加入");
}

async function postJSON(url, payload) {
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const json = await res.json();
  if (!res.ok) throw new Error(json.error || "请求失败");
  return json;
}

function renderAdvice(data) {
  sectorBadge.textContent = data.sectorName || "已识别";
  keywords.innerHTML = "";
  const terms = data.terms && data.terms.length ? data.terms : ["未命中具体词"];
  terms.forEach((term) => {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = term;
    keywords.appendChild(chip);
  });

  adviceList.innerHTML = "";
  data.opinions.forEach((item) => {
    appendAdviceCard(item.title, item.body);
  });
  renderRewriteDraft(data);
  if (data.library && data.library.matches && data.library.matches.length) {
    renderLibraryMaterials(data.library);
  }
  if (data.webResearch && data.webResearch.results && data.webResearch.results.length) {
    renderWebResearch(data.webResearch);
  }
}

function appendAdviceCard(titleText, bodyText) {
  const card = document.createElement("article");
  card.className = "advice-card";
  card.innerHTML = `<h3>${escapeHTML(titleText)}</h3><p>${escapeHTML(bodyText)}</p>`;
  adviceList.appendChild(card);
}

function renderRewriteDraft(data) {
  const card = document.createElement("article");
  card.className = "advice-card rewrite-draft";
  const parts = [];

  if (data.experiences && data.experiences.length) {
    parts.push(`<h4>实习 / 工作经历</h4>`);
    data.experiences.slice(0, 5).forEach((exp) => {
      parts.push(`<div class="draft-entry"><div class="draft-title">${escapeHTML(exp.org)}｜${escapeHTML(exp.role)}<span>${escapeHTML(exp.date || "")}</span></div>`);
      (exp.bullets || []).forEach(([label, text]) => {
        parts.push(`<div class="draft-bullet"><b>${escapeHTML(label)}</b>${escapeHTML(text)}</div>`);
      });
      parts.push(`</div>`);
    });
  }

  if (data.competitions && data.competitions.length) {
    parts.push(`<h4>竞赛经历</h4>`);
    data.competitions.slice(0, 3).forEach(([title, date, text]) => {
      parts.push(`<div class="draft-entry"><div class="draft-title">${escapeHTML(title)}<span>${escapeHTML(date || "")}</span></div><div class="draft-line">${escapeHTML(text || "")}</div></div>`);
    });
  }

  if (data.research && data.research.length) {
    parts.push(`<h4>科研 / 项目经历</h4>`);
    data.research.slice(0, 3).forEach(([title, date, text]) => {
      parts.push(`<div class="draft-entry"><div class="draft-title">${escapeHTML(title)}<span>${escapeHTML(date || "")}</span></div><div class="draft-line">${escapeHTML(text || "")}</div></div>`);
    });
  }

  card.innerHTML = `<h3>可直接替换进简历的内容</h3>${parts.join("") || "<p>暂无可展示的改写内容。</p>"}`;
  adviceList.appendChild(card);
}

function renderResume(data) {
  if (previewLang.value === "en") {
    renderEnglishResume(data);
    return;
  }
  const profile = data.profile;
  const html = [];
  html.push(`<article class="resume-sheet">`);
  html.push(`<div class="resume-name">${escapeHTML(profile.name)}</div>`);
  html.push(`<div class="resume-contact">${escapeHTML(profile.contact)}</div>`);

  html.push(section("教育背景"));
  profile.education.forEach(([school, major, date, details]) => {
    html.push(row(`<strong>${escapeHTML(school)}</strong><span>　${escapeHTML(major)}</span>`, date));
    details.forEach((detail) => html.push(detailLine(detail)));
  });
  html.push(detailLine(profile.honors));

  html.push(section("实习经历"));
  data.experiences.forEach((exp) => {
    html.push(row(`<strong>${escapeHTML(exp.org)}</strong><span> | ${escapeHTML(exp.role)}</span>`, exp.date));
    exp.bullets.forEach(([label, text]) => {
      html.push(`<div class="resume-bullet"><b>${escapeHTML(label)}</b>${escapeHTML(text)}</div>`);
    });
  });

  html.push(section("竞赛经历"));
  data.competitions.forEach(([title, date, text]) => {
    html.push(row(`<strong>${escapeHTML(title)}</strong>`, date));
    html.push(detailLine(text));
  });

  html.push(section("科研经历"));
  data.research.forEach(([title, date, text]) => {
    html.push(row(`<strong>${escapeHTML(title)}</strong>`, date));
    html.push(detailLine(text));
  });

  html.push(section("技能与兴趣"));
  html.push(detailLine("技能：法律检索、案例检索、Microsoft Office（Excel、Word、PPT）"));
  html.push(detailLine("语言：英语（IELTS 7.0）、普通话（母语）"));
  html.push(`</article>`);
  resumePreview.innerHTML = html.join("");
}

const orgEn = {};

const bulletEn = {
  "新能源行业研究：": ["Green industry research: ", "Researched wind power, geothermal energy, new-energy pipeline networks and SAF-related sectors; summarized policy context, business models and project economics based on project materials."],
  "投资尽调支持：": ["Investment due diligence: ", "Organized corporate, equity, operations, litigation, environmental, social/resettlement and financial due diligence materials for green private-equity projects; identified compliance, operational and exit risks."],
  "投决材料整理：": ["Investment committee materials: ", "Assisted in organizing investment proposals, risk assessment reports, investment committee/general manager/board materials, investment agreements, shareholders' agreements and ancillary agreements; extracted key terms and decision points."],
  "绿色投资项目材料：": ["Green investment materials: ", "Organized investment proposals, risk assessments, internal approval documents and due diligence materials for wind, geothermal and new-energy pipeline projects."],
  "合规与决策支持：": ["Compliance and decision support: ", "Assisted in reviewing project initiation, investment committee, board and internal approval materials; summarized legal, financial, environmental and social due diligence risk points."],
  "绿色基金项目尽调：": ["Green fund due diligence: ", "Organized corporate, transaction, environmental, legal and financial due diligence materials for new-energy investment projects; summarized compliance and transaction risks."],
  "结构性产品文件：": ["Structured products documentation: ", "Supported the Financial Regulation team on structured product documents for foreign banks; assisted in drafting and checking Final Terms and reviewed equity-linked notes, including fixed coupons, autocall mechanics, worst-of basket, observation-date trigger levels, barriers and redemption calculations."],
  "金融监管跟踪：": ["Financial regulatory tracking: ", "Tracked regulatory updates and HKEX announcements for foreign banks; summarized potential implications for structured note issuance, disclosure and compliance arrangements."],
  "条款逻辑核查：": ["Clause logic review: ", "Reviewed highly technical clauses for foreign-bank matters, focusing on non-reliance, no fiduciary, unsecured, no floor, Calculation Agent, selling restrictions and cross-border compliance issues beyond formatting."],
  "合规检索与规则适用：": ["Compliance research and rule application: ", "Researched pharmaceutical regulation, internet-service pre-approval, technology import/export and SEP liability boundaries; distilled approval/filing paths, liability boundaries and practical risk points."],
  "客户咨询 Memo 写作：": ["Client memorandum drafting: ", "Worked directly with a partner on client consultations; independently prepared memorandum drafts organizing regulatory rules into issue lists, applicable rules and risk notes."],
  "交易材料风险核查：": ["Transaction risk review: ", "Reviewed negative information, IP ownership/use restrictions and third-party consent/change-of-control clauses in M&A and investment matters; summarized issues to be confirmed before closing."],
  "数据合规与技术出口管制：": ["Data compliance and export controls: ", "Researched AI outbound business and technology export-control issues; analyzed catalog classifications, restricted-export rules and approval paths; prepared a client-facing memorandum draft."],
  "互联网与药监合规：": ["Internet and pharmaceutical compliance: ", "Researched pre-approval requirements for education-related internet services, pharmaceutical regulation and SEP damages-liability boundaries; summarized regulatory risks."],
  "跨境争议证据审阅：": ["Cross-border dispute evidence review: ", "Reviewed transaction documents, correspondence and meeting records; identified evidence on put option exercise, transaction communications and loss issues."],
  "交易与监管支持：": ["Transaction and regulatory support: ", "Reviewed negative information, IP ownership/use restrictions and third-party consent/change-of-control clauses; summarized pre-closing issues."],
  "交易尽调支持：": ["Transaction due diligence: ", "Reviewed negative information, IP ownership, third-party consent and change-of-control clauses in investment-related materials; summarized pre-closing issues and transaction risks."],
  "M&A 文件核查：": ["M&A document review: ", "Assisted in organizing transaction documents, resolutions and closing-stage materials; checked executed versions, exhibit lists and consistency of key clauses."],
  "Regulatory 研究：": ["Regulatory research: ", "Researched AI outbound technology export controls, pharmaceutical compliance, internet-service approvals and SEP liability boundaries; prepared structured memo drafts."],
  "跨境投融资争议：": ["Cross-border investment dispute: ", "Reviewed transaction communications, meeting records and correspondence; extracted evidence on equity repurchase, investment exit and loss issues."],
  "证据审阅与整理：": ["Evidence review: ", "Reviewed transaction documents, correspondence and meeting records in a major cross-border investment dispute; identified key evidence and assisted in preparing evidence indexes."],
  "法律检索与文书写作：": ["Legal research and drafting: ", "Worked directly with a partner on two client memoranda; researched pharmaceutical compliance, AI export controls, SEP-related damages and regulatory risk boundaries."],
  "M&A 尽职调查：": ["M&A due diligence: ", "Checked negative information, IP ownership/use restrictions and third-party consent clauses; translated patent materials and summarized risk points."],
  "商标争议文书：": ["Trademark proceedings: ", "Drafted over 60 submissions for invalidation, opposition and non-use cancellation proceedings, covering prior rights, similarity, bad faith and non-use issues."],
  "近似与恶意分析：": ["Similarity and bad-faith analysis: ", "Compared marks, goods/services subclasses, filing portfolios and prior decisions; organized evidence indexes and procedural materials."],
  "英文客户沟通：": ["English client communication: ", "Prepared over 10 English watch opinions and 40+ client emails summarizing CNIPA decisions, case progress and follow-up strategies."],
  "民事检察监督：": ["Civil procuratorial supervision: ", "Supported civil litigation, trial-procedure supervision and enforcement supervision matters; mapped procedural paths from first instance to enforcement."],
  "法律文书撰写：": ["Legal document drafting: ", "Assisted in drafting 40+ review reports, litigation-support statements, procuratorial recommendations and protest documents."],
  "商事争议检索：": ["Commercial dispute research: ", "Researched professional lending, early termination, liquidated damages/lost profits, negotiable-instrument disputes and unjust enrichment; prepared research memoranda."],
  "企业尽调支持：": ["Due diligence support: ", "Assisted on a debt restructuring matter by organizing corporate registration, shareholding, operations and litigation information."],
};

const schoolEn = {
  "北京大学": "Peking University",
  "清华大学": "Tsinghua University",
  "中国人民大学": "Renmin University of China",
  "中国政法大学": "China University of Political Science and Law",
  "复旦大学": "Fudan University",
  "上海交通大学": "Shanghai Jiao Tong University",
  "武汉大学": "Wuhan University",
  "华东政法大学": "East China University of Political Science and Law",
  "意大利博洛尼亚大学，公派交流项目": "Alma Mater Studiorum - University of Bologna | Exchange Program",
};

const genericLabelEn = {
  "数据与技术合规：": "Data and technology compliance: ",
  "金融与交易支持：": "Finance and transaction support: ",
  "合规与风险识别：": "Compliance and risk identification: ",
  "研究分析：": "Research and analysis: ",
  "文书写作：": "Legal drafting: ",
  "英文与跨境支持：": "English and cross-border support: ",
  "工作内容：": "Work scope: ",
  "原始经历摘要：": "Experience summary: ",
};

function hasChinese(text) {
  return /[\u4e00-\u9fff]/.test(String(text || ""));
}

function enName(name) {
  if (!hasChinese(name)) return name || "Candidate";
  return name || "Candidate";
}

function replaceTerms(text, pairs) {
  let result = String(text || "");
  pairs.forEach(([zh, en]) => {
    result = result.split(zh).join(en);
  });
  return result;
}

function translateOrg(text) {
  if (orgEn[text]) return orgEn[text][0];
  if (!hasChinese(text)) return text || "Organization";
  let result = text || "";
  Object.entries(schoolEn).forEach(([zh, en]) => {
    result = result.split(zh).join(en);
  });
  return replaceTerms(result, [["律师事务所", "Law Firm"], ["检察院", "Procuratorate"], ["法院", "Court"], ["银行", "Bank"], ["证券", "Securities"], ["基金", "Fund"], ["投资", "Investment"], ["公司", "Company"], ["大学", "University"], ["学院", "School"]]);
}

function translateRole(text) {
  if (!hasChinese(text)) return text || "Intern";
  return replaceTerms(text, [["数据合规实习生", "Data Compliance Intern"], ["法务实习生", "Legal Intern"], ["法律实习生", "Legal Intern"], ["律师助理", "Legal Assistant"], ["实习生", "Intern"], ["项目核心成员", "Core Member"], ["负责人", "Lead Member"], ["成员", "Member"]]);
}

function translateEducation(school, major) {
  const schoolText = schoolEn[school] || translateOrg(school);
  const majorText = replaceTerms(major || "", [["民商法学硕", "LL.M. in Civil and Commercial Law"], ["民商法", "Civil and Commercial Law"], ["法学本科", "LL.B. / Law"], ["法学", "Law"], ["法律", "Law"], ["硕士", "Master's Program"], ["本科", "Bachelor's Program"], ["专业", ""]]);
  return majorText ? `${schoolText} | ${majorText}` : schoolText;
}

function translateDetail(text) {
  if (!hasChinese(text)) return (text || "").replaceAll("：", ": ");
  return replaceTerms(text, [["2024年通过国家司法考试", "Passed PRC Legal Professional Qualification Examination (2024)"], ["GPA：", "GPA: "], ["综合排名", "Comprehensive Ranking"], ["专业排名", "Major Ranking"], ["专必均分", "Major Required Courses Average"], ["雅思", "IELTS "], ["通过国家司法考试", "Passed PRC Legal Professional Qualification Examination"], ["荣誉奖项", "Honors"], ["校级", "University-level "], ["竞赛一等奖学金", "First Prize Competition Scholarship"], ["连续两年学业奖学金", "Academic Scholarship for Two Consecutive Years"], ["帕特森基金会奖学金", "Patterson Foundation Scholarship"], ["优秀毕业生", "Outstanding Graduate"], ["三好学生", "Outstanding Student"], ["奖学金", "Scholarship"], ["法律检索", "Legal Research"], ["案例检索", "Case Research"], ["普通话", "Mandarin"], ["母语", "Native"]]).replaceAll("：", ": ").replaceAll("；", "; ").replaceAll("，", ", ").replaceAll("、", "; ").replace(/\s+/g, " ").trim();
}

function fallbackBullet(label, text) {
  if (bulletEn[label]) return bulletEn[label];
  if (!hasChinese(text)) return [genericLabelEn[label] || label.replace("：", ": "), text || ""];
  const combined = `${label}${text}`;
  const parts = [];
  if (/个人信息|数据出境|数据合规|隐私|网络安全/.test(combined)) parts.push("researched personal information protection, data export compliance and cybersecurity requirements");
  if (/生成式人工智能|AIGC|AI|算法|技术出口管制|出口管制/i.test(combined)) parts.push("analyzed AI-related regulatory issues and technology export-control requirements");
  if (/法律检索|法规|监管|政策|案例|司法解释/.test(combined)) parts.push("conducted legal and regulatory research and summarized applicable rules");
  if (/合同|条款|审查|审核/.test(combined)) parts.push("reviewed contract clauses and identified legal risk points");
  if (/尽调|工商|股权|涉诉|负面信息|知识产权权属/.test(combined)) parts.push("organized due diligence materials on corporate registration, shareholding, litigation and IP issues");
  if (/备忘录|memo|报告|文书|起草|撰写|邮件/i.test(combined)) parts.push("prepared memoranda, reports and drafting materials for lawyer or client review");
  if (/金融|投资|基金|投融资|估值|行业研究|投决/.test(combined)) parts.push("supported financial, investment and industry research work");
  if (/争议|诉讼|仲裁|证据|庭审/.test(combined)) parts.push("reviewed dispute materials and organized evidence and procedural issues");
  if (/商标|专利|著作权|知识产权|侵权/.test(combined)) parts.push("analyzed intellectual property issues including trademarks, patents and infringement risks");
  if (/英文|英语|翻译|跨境|境外/.test(combined)) parts.push("translated materials and supported English or cross-border communications");
  const sentence = `${(parts.length ? parts : ["organized project materials, summarized key issues and supported legal research or drafting work"]).join("; ")}.`;
  return [genericLabelEn[label] || "Work scope: ", sentence.charAt(0).toUpperCase() + sentence.slice(1)];
}

function translateItem(title, text) {
  if (!hasChinese(`${title}${text}`)) return [title || "", text || ""];
  const titleEn = replaceTerms(title || "", [["全国法科学生模拟立法大赛", "National Legislative Simulation Competition"], ["模拟立法", "Legislative Simulation"], ["模拟法庭", "Moot Court"], ["项目核心成员", "Core Member"], ["独著", "Sole Author"], ["论文", "Research Paper"], ["研究", "Research"]]);
  return [titleEn, fallbackBullet("研究分析：", text || title)[1]];
}

function renderEnglishResume(data) {
  const html = [];
  html.push(`<article class="resume-sheet en">`);
  html.push(`<div class="resume-name">${escapeHTML(data.profile?.enName || enName(data.profile?.name))}</div>`);
  html.push(`<div class="resume-contact">${escapeHTML(data.profile?.contact || "")}</div>`);

  html.push(section("Education & Honors"));
  (data.profile?.education || []).forEach(([school, major, date, details]) => {
    html.push(row(`<strong>${escapeHTML(translateEducation(school, major))}</strong>`, date));
    (details || []).forEach((detail) => html.push(detailLine(translateDetail(detail))));
  });
  if (data.profile?.honors) html.push(detailLine(translateDetail(data.profile.honors)));

  html.push(section("Internship Experience"));
  data.experiences.forEach((exp) => {
    const [org, role] = orgEn[exp.org] || [translateOrg(exp.org), translateRole(exp.role)];
    html.push(row(`<strong>${escapeHTML(org)}</strong>`, exp.date));
    html.push(detailLine(role));
    exp.bullets.forEach(([label, text]) => {
      const [enLabel, enText] = fallbackBullet(label, text);
      html.push(`<div class="resume-bullet"><b>${escapeHTML(enLabel)}</b>${escapeHTML(enText)}</div>`);
    });
  });

  if (data.competitions?.length) {
    html.push(section("Competition Experience"));
    data.competitions.forEach(([title, date, text]) => {
      const [enTitle, enText] = translateItem(title, text);
      html.push(row(`<strong>${escapeHTML(enTitle)}</strong>`, date));
      if (enText) html.push(detailLine(enText));
    });
  }
  if (data.research?.length) {
    html.push(section("Research Experience"));
    data.research.forEach(([title, date, text]) => {
      const [enTitle, enText] = translateItem(title, text);
      html.push(row(`<strong>${escapeHTML(enTitle)}</strong>`, date));
      if (enText) html.push(detailLine(enText));
    });
  }
  html.push(section("Skills & Interests"));
  html.push(detailLine(data.externalParsed ? (translateDetail(data.profile?.honors) || "Skills: Legal Research; Microsoft Office; drafting and document review.") : "Skills: PRC Legal Professional Qualification (2024); Legal Research; Microsoft Office; HeinOnline; Westlaw"));
  if (!data.externalParsed) {
    html.push(detailLine("Language: IELTS 7.5"));
    html.push(detailLine("Interests: Singing (Top Ten Singer Award); Swimming (Deep-Water Swimming Certificate); Piano (Grade 10)"));
  }
  html.push(`</article>`);
  resumePreview.innerHTML = html.join("");
}

function section(title) {
  return `<div class="resume-section">${escapeHTML(title)}</div>`;
}

function row(left, date) {
  return `<div class="resume-row"><div>${left}</div><div class="resume-date">${escapeHTML(date)}</div></div>`;
}

function detailLine(text) {
  return `<div class="resume-detail">${escapeHTML(text)}</div>`;
}

function renderEmpty() {
  adviceList.innerHTML = `<div class="empty-state">粘贴 JD 后生成修改意见</div>`;
  resumePreview.innerHTML = `<div class="empty-state">修改后的简历会出现在这里</div>`;
}

function renderMaterials() {
  if (!uploadedMaterials.length) {
    setMaterialStatus("已连接桌面资料库");
    return;
  }
  materialList.innerHTML = uploadedMaterials
    .map((item) => `<div>已上传：${escapeHTML(item.name)}${item.text ? "（已提取文本）" : "（图片/暂未提取文字）"}</div>`)
    .join("");
}

function renderLibraryMaterials(library) {
  const title = document.createElement("article");
  title.className = "advice-card";
  const items = library.matches
    .slice(0, 8)
    .map((item) => `• ${escapeHTML(item.name)}（${escapeHTML((item.tags || []).join("、"))}）`)
    .join("<br>");
  title.innerHTML = `<h3>可调用的桌面材料</h3><p>共扫描 ${library.count} 个文件，当前 JD/修改要求命中：<br>${items}</p>`;
  adviceList.appendChild(title);
}

function renderWebResearch(research) {
  const title = document.createElement("article");
  title.className = "advice-card";
  const items = research.results
    .slice(0, 6)
    .map((item) => {
      const link = `<a href="${escapeAttr(item.url)}" target="_blank" rel="noreferrer">${escapeHTML(item.title)}</a>`;
      return `• ${link}<br><span>${escapeHTML(item.snippet || "")}</span>`;
    })
    .join("<br>");
  title.innerHTML = `<h3>联网检索结果</h3><p>检索式：${escapeHTML(research.query || "")}</p><p>${items}</p>`;
  adviceList.appendChild(title);
}

function setMaterialStatus(text) {
  materialList.textContent = text;
}

function setBusy(button, busy, text) {
  button.disabled = busy;
  button.textContent = text;
}

function showToast(message) {
  toast.textContent = message;
  toast.hidden = false;
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => {
    toast.hidden = true;
  }, 1800);
}

function readFileAsDataURL(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function escapeHTML(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttr(value) {
  return escapeHTML(value).replaceAll("`", "&#096;");
}
