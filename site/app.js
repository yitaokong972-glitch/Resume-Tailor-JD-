/* JD 智能简历定制台 · 纯前端版
 * - 直接在浏览器调用 OpenAI 兼容接口（已验证 DeepSeek 支持 CORS）
 * - 密钥仅存 localStorage，不离开本机
 * - 用 docx 库在浏览器生成中 / 英文 Word 下载
 */
(function () {
  "use strict";

  // ---------- 配置（localStorage）----------
  const LS = {
    key: "rt_apiKey",
    base: "rt_apiBase",
    model: "rt_apiModel",
  };
  const cfg = {
    key: localStorage.getItem(LS.key) || "",
    base: localStorage.getItem(LS.base) || "https://api.deepseek.com",
    model: localStorage.getItem(LS.model) || "deepseek-chat",
  };

  // ---------- DOM ----------
  const $ = (id) => document.getElementById(id);
  const el = {
    settingsBtn: $("settingsBtn"),
    settingsPanel: $("settingsPanel"),
    apiKey: $("apiKey"),
    apiBase: $("apiBase"),
    apiModel: $("apiModel"),
    connStatus: $("connStatus"),
    jd: $("jdInput"),
    jdCount: $("jdCount"),
    custom: $("customInput"),
    resumeInput: $("resumeInput"),
    resumeFile: $("resumeFile"),
    analyzeBtn: $("analyzeBtn"),
    downloadBtn: $("downloadBtn"),
    downloadEnBtn: $("downloadEnBtn"),
    actionHint: $("actionHint"),
    adviceList: $("adviceList"),
    sectorBadge: $("sectorBadge"),
    previewLang: $("previewLang"),
    resumePreview: $("resumePreview"),
    toast: $("toast"),
  };

  // ---------- 状态 ----------
  let currentZh = null; // 中文简历 JSON
  let currentEn = null; // 英文简历 JSON
  let busy = false;

  // ---------- 工具 ----------
  function toast(msg, ms = 2600) {
    el.toast.textContent = msg;
    el.toast.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(() => (el.toast.hidden = true), ms);
  }
  function setBusy(on, label) {
    busy = on;
    [el.analyzeBtn, el.downloadBtn, el.downloadEnBtn].forEach((b) => (b.disabled = on));
    if (on && label) el.actionHint.textContent = label;
  }
  function extractJson(text) {
    text = (text || "").trim();
    if (text.startsWith("```")) {
      text = text.replace(/^```(?:json)?\s*/, "").replace(/\s*```$/, "");
    }
    try {
      return JSON.parse(text);
    } catch (e) {
      const s = text.indexOf("{");
      const end = text.lastIndexOf("}");
      if (s !== -1 && end > s) return JSON.parse(text.slice(s, end + 1));
      throw new Error("模型未返回合法 JSON");
    }
  }

  // ---------- 调用模型（浏览器直连，已验证 CORS）----------
  async function callLLM(messages, temperature = 0.35) {
    if (!cfg.key) throw new Error("请先在右上角「模型设置」填写 API Key");
    const url = cfg.base.replace(/\/+$/, "") + "/chat/completions";
    const body = {
      model: cfg.model,
      messages,
      temperature,
      response_format: { type: "json_object" },
    };
    const resp = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + cfg.key,
      },
      body: JSON.stringify(body),
    });
    if (!resp.ok) {
      let detail = "";
      try {
        detail = (await resp.json()).error?.message || "";
      } catch (_) {}
      throw new Error("API " + resp.status + (detail ? "：" + detail : ""));
    }
    const data = await resp.json();
    const content = data.choices?.[0]?.message?.content;
    if (!content) throw new Error("模型返回为空");
    return content;
  }

  async function testConnection() {
    try {
      await callLLM([{ role: "user", content: "ping" }], 0);
      el.connStatus.textContent = "● 已连接";
      el.connStatus.className = "conn-status ok";
    } catch (e) {
      el.connStatus.textContent = "○ 未连接 / 失败";
      el.connStatus.className = "conn-status err";
    }
  }

  // ---------- 提示词（移植自后端）----------
  function buildZhMessages() {
    const system =
      "你是一个严谨的中文简历改写器。必须直接生成一份可进入 Word 的一页简历 JSON，不是只给建议。" +
      "严格根据 JD 和用户修改要求现场选择经历、重写 bullet。用户修改要求的优先级高于你的默认判断；只要不违背事实、不构成虚构，就必须执行。" +
      "不得虚构经历、客户名称、项目名称或数据；可以在已提供经历范围内重组、删减、改写表达。" +
      "如果用户提供了原始简历文本，必须以该原始简历为唯一事实来源，禁止使用任何不属于该候选人的人物、学校、机构、经历。" +
      "bullet 用 STAR 逻辑但不写 Situation/Task 字样，每条要具体、专业、可落地。" +
      "必须在 resume.profile 中填写 enName 字段，为姓名拼音（姓在前、空格分隔、每词首字母大写，如 张三→Zhang San、李四→Li Si）。" +
      "输出必须是 JSON 对象，不要 Markdown。";
    const user = {
      JD: el.jd.value || "",
      修改要求: el.custom.value || "",
      必须执行修改要求: !!(el.custom.value || "").trim(),
      用户原始简历文本: el.resumeInput.value.slice(0, 12000),
      事实来源规则:
        "若“用户原始简历文本”非空，只能使用该文本中的事实；禁止混入不属于该候选人的人物、学校、机构、经历。若为空，可基于 JD 生成通用占位并提示用户补充。",
      输出JSON格式: {
        summary: "一句话说明改写策略",
        resume: {
          profile: {
            name: "候选人姓名",
            enName: "Li Si",
            contact: "电话 | 邮箱 | 链接",
            education: [{ school: "学校", major: "专业/学位", date: "时间", details: ["成绩/荣誉/课程"] }],
            honors: "荣誉奖项",
          },
          experiences: [{ org: "机构名", role: "岗位名", date: "时间", bullets: [["标签：", "正文"]] }],
          competitions: [["标题", "时间", "内容"]],
          research: [["标题", "时间", "内容"]],
          opinions: [{ title: "改写说明", body: "如何根据 JD 与要求改写" }],
        },
      },
      输出质量要求:
        "opinions 中必须单列一条“手动修改要求如何执行”，说明修改要求具体改进到了哪些经历；resume 必须体现这些改动。",
    };
    return [
      { role: "system", content: system },
      { role: "user", content: JSON.stringify(user, null, 0) },
    ];
  }

  function buildEnMessages(zhResume) {
    const system =
      "You are a professional resume translator. Convert the provided Chinese resume JSON into a polished English one-page resume. " +
      "Do not invent, exaggerate, delete dates, change schools/employers, or add clients/projects. " +
      "Return JSON only with the same schema: profile, experiences, competitions, research. " +
      "Keep all dates and numbers exactly. Use concise, polished English resume language. " +
      "profile.name should be the English name (pinyin).";
    const user = { resume: zhResume };
    return [
      { role: "system", content: system },
      { role: "user", content: JSON.stringify(user, null, 0) },
    ];
  }

  // ---------- 归一化 ----------
  function normBullets(bullets) {
    const out = [];
    (bullets || []).forEach((b) => {
      let label = "",
        text = "";
      if (Array.isArray(b) && b.length >= 2) {
        label = String(b[0] || "").trim();
        text = String(b[1] || "").trim();
      } else if (b && typeof b === "object") {
        label = String(b.label || "").trim();
        text = String(b.text || "").trim();
      }
      if (label && !label.endsWith("：") && !label.endsWith(":")) label += label.includes(":") ? "" : "：";
      if (label && text) out.push([label, text]);
    });
    return out;
  }
  function normEducation(edu) {
    return (edu || []).map((e) => {
      if (e && typeof e === "object")
        return {
          school: String(e.school || "").trim(),
          major: String(e.major || "").trim(),
          date: String(e.date || "").trim(),
          details: Array.isArray(e.details) ? e.details.map((d) => String(d).trim()).filter(Boolean) : [],
        };
      if (Array.isArray(e) && e.length >= 4)
        return {
          school: String(e[0] || "").trim(),
          major: String(e[1] || "").trim(),
          date: String(e[2] || "").trim(),
          details: [String(e[3] || "")].filter(Boolean),
        };
      return null;
    }).filter(Boolean);
  }
  function normTuples(arr) {
    return (arr || [])
      .map((it) => {
        if (Array.isArray(it) && it.length >= 3)
          return [String(it[0] || "").trim(), String(it[1] || "").trim(), String(it[2] || "").trim()];
        if (it && typeof it === "object")
          return [String(it.title || "").trim(), String(it.date || "").trim(), String(it.detail || "").trim()];
        return null;
      })
      .filter((x) => x && (x[0] || x[2]));
  }
  function normResume(data) {
    const r = data?.resume && typeof data.resume === "object" ? data.resume : data;
    const profile = r.profile && typeof r.profile === "object" ? r.profile : {};
    const experiences = (r.experiences || [])
      .map((exp) => {
        if (!exp || typeof exp !== "object") return null;
        const bullets = normBullets(exp.bullets);
        if (!bullets.length) return null;
        return {
          org: String(exp.org || "").trim(),
          role: String(exp.role || "").trim(),
          date: String(exp.date || "").trim(),
          bullets: bullets.slice(0, 3),
        };
      })
      .filter(Boolean)
      .slice(0, 5);
    return {
      profile: {
        name: String(profile.name || "").trim(),
        enName: String(profile.enName || "").trim(),
        contact: String(profile.contact || "").trim(),
        honors: String(profile.honors || "").trim(),
        education: normEducation(profile.education),
      },
      experiences,
      competitions: normTuples(r.competitions).slice(0, 3),
      research: normTuples(r.research).slice(0, 3),
      opinions: Array.isArray(r.opinions)
        ? r.opinions.map((o) => ({ title: String(o?.title || "改写说明"), body: String(o?.body || "") }))
        : [],
      summary: typeof data?.summary === "string" ? data.summary : "",
    };
  }

  // ---------- 渲染 ----------
  function renderAdvice(resume) {
    const items = [];
    if (resume.summary) items.push({ title: "改写策略", body: resume.summary });
    (resume.opinions || []).forEach((o) => items.push(o));
    if (!items.length) {
      el.adviceList.innerHTML = '<p class="empty-tip">未生成改写说明。</p>';
      return;
    }
    el.adviceList.innerHTML = items
      .map(
        (o) =>
          '<div class="advice-item"><div class="advice-title">' +
          esc(o.title) +
          '</div><div class="advice-body">' +
          esc(o.body) +
          "</div></div>"
      )
      .join("");
  }
  function renderResume(resume, lang) {
    const isEn = lang === "en";
    const name = isEn ? resume.profile.enName || resume.profile.name : resume.profile.name;
    if (!name) {
      el.resumePreview.innerHTML = '<p class="empty-tip">暂无内容，请先生成。</p>';
      return;
    }
    let html = '<div class="rp-name">' + esc(name) + "</div>";
    if (resume.profile.contact)
      html += '<div class="rp-contact">' + esc(resume.profile.contact) + "</div>";

    const section = (t, rows) => {
      if (!rows) return "";
      return '<div class="rp-section"><div class="rp-h">' + esc(t) + "</div>" + rows + "</div>";
    };
    if (resume.profile.education?.length)
      html += section(
        isEn ? "Education" : "教育背景",
        resume.profile.education
          .map(
            (e) =>
              '<div class="rp-line"><span>' +
              esc(e.school + (e.major ? "  " + e.major : "")) +
              '</span><span class="rp-right">' +
              esc(e.date) +
              '</span></div>' +
              (e.details || []).map((d) => '<div class="rp-detail">' + esc(d) + "</div>").join("")
          )
          .join("")
      );
    if (resume.experiences?.length)
      html += section(
        isEn ? "Experience" : "实习经历",
        resume.experiences
          .map(
            (x) =>
              '<div class="rp-line"><span>' +
              esc((x.org || "") + (x.role ? "  |  " + x.role : "")) +
              '</span><span class="rp-right">' +
              esc(x.date) +
              '</span></div>' +
              (x.bullets || [])
                .map((b) => '<div class="rp-bullet"><b>' + esc(b[0]) + "</b>" + esc(b[1]) + "</div>")
                .join("")
          )
          .join("")
      );
    if (resume.competitions?.length)
      html += section(
        isEn ? "Competitions" : "竞赛经历",
        resume.competitions
          .map(
            (c) =>
              '<div class="rp-line"><span>' +
              esc(c[0]) +
              '</span><span class="rp-right">' +
              esc(c[1]) +
              '</span></div><div class="rp-detail">' +
              esc(c[2]) +
              "</div>"
          )
          .join("")
      );
    if (resume.research?.length)
      html += section(
        isEn ? "Research" : "科研经历",
        resume.research
          .map(
            (c) =>
              '<div class="rp-line"><span>' +
              esc(c[0]) +
              '</span><span class="rp-right">' +
              esc(c[1]) +
              '</span></div><div class="rp-detail">' +
              esc(c[2]) +
              "</div>"
          )
          .join("")
      );
    html += section(isEn ? "Skills & Interests" : "技能与兴趣", resume.profile.honors
      ? '<div class="rp-detail">' + esc(resume.profile.honors) + "</div>"
      : '<div class="rp-detail">' + (isEn ? "Add your skills, languages and certificates." : "补充语言、证书、软件等技能。") + "</div>");

    el.resumePreview.innerHTML = html;
  }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  // ---------- 流程 ----------
  async function generateZh() {
    if (busy) return;
    setBusy(true, "正在调用模型改写…");
    try {
      const text = await callLLM(buildZhMessages(), 0.35);
      currentZh = normResume(extractJson(text));
      currentEn = null;
      el.sectorBadge.textContent = "已生成（中文）";
      renderAdvice(currentZh);
      el.previewLang.value = "zh";
      renderResume(currentZh, "zh");
      el.actionHint.textContent = "已生成中文预览，可下载中文 / 英文简历。";
      toast("中文简历已生成");
    } catch (e) {
      toast("生成失败：" + e.message);
    } finally {
      setBusy(false);
    }
  }
  async function generateEn() {
    if (!currentZh) {
      toast("请先生成中文简历，再生成英文版");
      return;
    }
    if (busy) return;
    setBusy(true, "正在翻译为英文…");
    try {
      const text = await callLLM(buildEnMessages(currentZh), 0.15);
      currentEn = normResume({ resume: extractJson(text) });
      el.sectorBadge.textContent = "已生成（英文）";
      el.previewLang.value = "en";
      renderResume(currentEn, "en");
      toast("英文简历已生成");
    } catch (e) {
      toast("英文生成失败：" + e.message);
    } finally {
      setBusy(false);
    }
  }

  // ---------- docx 生成（浏览器端）----------
  function buildDocx(resume, lang) {
    const D = window.docx;
    const isEn = lang === "en";
    const name = isEn ? resume.profile.enName || resume.profile.name : resume.profile.name;
    const ACCENT = "1F3A5F";
    const children = [];
    const fontHead = isEn ? "Times New Roman" : "黑体";
    const fontBody = isEn ? "Times New Roman" : "宋体";
    const H = (t) =>
      new D.Paragraph({
        spacing: { before: 160, after: 40 },
        border: { bottom: { color: ACCENT, style: D.BorderStyle.SINGLE, size: 6 } },
        children: [new D.TextRun({ text: t, bold: true, size: 22, color: ACCENT, font: fontHead })],
      });
    const line = (left, right) =>
      new D.Paragraph({
        spacing: { after: 20 },
        tabStops: [{ type: D.TabStopType.RIGHT, position: 9020 }],
        children: [
          new D.TextRun({ text: left, bold: true, size: 18, font: fontBody }),
          right ? new D.TextRun({ text: "\t" + right, size: 18, font: fontBody }) : null,
        ].filter(Boolean),
      });
    const plain = (t) =>
      new D.Paragraph({ spacing: { after: 20 }, indent: { left: 280 }, children: [new D.TextRun({ text: t, size: 18, font: fontBody })] });
    const bullet = (label, text) =>
      new D.Paragraph({
        spacing: { after: 16 },
        indent: { left: 360, hanging: 200 },
        children: [new D.TextRun({ text: "• " + label, bold: true, size: 18, font: fontBody }), new D.TextRun({ text: text, size: 18, font: fontBody })],
      });

    // name + contact
    children.push(
      new D.Paragraph({ alignment: D.AlignmentType.CENTER, spacing: { after: 30 }, children: [new D.TextRun({ text: name, bold: true, size: 32, font: fontHead })] })
    );
    if (resume.profile.contact)
      children.push(
        new D.Paragraph({ alignment: D.AlignmentType.CENTER, spacing: { after: 80 }, children: [new D.TextRun({ text: resume.profile.contact, size: 18, color: "595959", font: fontBody })] })
      );

    if (resume.profile.education?.length) {
      children.push(H(isEn ? "Education" : "教育背景"));
      resume.profile.education.forEach((e) => {
        children.push(line(e.school + (e.major ? "  " + e.major : ""), e.date));
        (e.details || []).forEach((d) => children.push(plain(d)));
      });
    }
    if (resume.experiences?.length) {
      children.push(H(isEn ? "Experience" : "实习经历"));
      resume.experiences.forEach((x) => {
        children.push(line((x.org || "") + (x.role ? "  |  " + x.role : ""), x.date));
        (x.bullets || []).forEach((b) => children.push(bullet(b[0], b[1])));
      });
    }
    if (resume.competitions?.length) {
      children.push(H(isEn ? "Competitions" : "竞赛经历"));
      resume.competitions.forEach((c) => {
        children.push(line(c[0], c[1]));
        children.push(plain(c[2]));
      });
    }
    if (resume.research?.length) {
      children.push(H(isEn ? "Research" : "科研经历"));
      resume.research.forEach((c) => {
        children.push(line(c[0], c[1]));
        children.push(plain(c[2]));
      });
    }
    children.push(H(isEn ? "Skills & Interests" : "技能与兴趣"));
    children.push(plain(resume.profile.honors || (isEn ? "Add your skills, languages and certificates." : "补充语言、证书、软件等技能。")));

    const doc = new D.Document({
      sections: [{ properties: { page: { margin: { top: 720, bottom: 680, left: 900, right: 900 } } }, children }],
    });
    return D.Packer.toBlob(doc);
  }
  async function downloadDocx(lang) {
    const resume = lang === "en" ? currentEn : currentZh;
    if (!resume) {
      toast(lang === "en" ? "请先生成英文简历" : "请先生成中文简历");
      return;
    }
    setBusy(true, "正在生成 Word…");
    try {
      const blob = await buildDocx(resume, lang);
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = (lang === "en" ? "JD_Tailored_Resume_EN_" : "JD定制简历_") + Date.now() + ".docx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 1000);
      toast("已下载 " + (lang === "en" ? "英文" : "中文") + " Word");
    } catch (e) {
      toast("下载失败：" + e.message);
    } finally {
      setBusy(false);
    }
  }

  // ---------- 文件解析 ----------
  async function parseResumeFile(file) {
    if (!file) return;
    const ext = file.name.toLowerCase().split(".").pop();
    try {
      if (ext === "txt") {
        el.resumeInput.value = await file.text();
        toast("已读取 .txt 简历");
      } else if (ext === "docx") {
        if (!window.mammoth) throw new Error("解析库未加载");
        const arrayBuffer = await file.arrayBuffer();
        const res = await window.mammoth.extractRawText({ arrayBuffer });
        el.resumeInput.value = res.value;
        toast("已读取 .docx 简历");
      } else {
        toast("仅支持 .txt / .docx，请用粘贴方式提供其他格式");
      }
    } catch (e) {
      toast("读取文件失败：" + e.message);
    }
  }

  // ---------- 事件 ----------
  function saveCfg() {
    cfg.key = el.apiKey.value.trim();
    cfg.base = el.apiBase.value.trim() || "https://api.deepseek.com";
    cfg.model = el.apiModel.value.trim() || "deepseek-chat";
    localStorage.setItem(LS.key, cfg.key);
    localStorage.setItem(LS.base, cfg.base);
    localStorage.setItem(LS.model, cfg.model);
  }
  function init() {
    el.apiKey.value = cfg.key;
    el.apiBase.value = cfg.base;
    el.apiModel.value = cfg.model;
    el.settingsBtn.addEventListener("click", () => {
      el.settingsPanel.hidden = !el.settingsPanel.hidden;
      if (!el.settingsPanel.hidden) testConnection();
    });
    [el.apiKey, el.apiBase, el.apiModel].forEach((i) => i.addEventListener("change", () => { saveCfg(); testConnection(); }));
    el.jd.addEventListener("input", () => (el.jdCount.textContent = el.jd.value.length + " 字"));
    el.analyzeBtn.addEventListener("click", generateZh);
    el.downloadBtn.addEventListener("click", () => downloadDocx("zh"));
    el.downloadEnBtn.addEventListener("click", generateEn);
    el.resumeFile.addEventListener("change", (e) => parseResumeFile(e.target.files[0]));
    el.previewLang.addEventListener("change", () => {
      const lang = el.previewLang.value;
      const r = lang === "en" ? currentEn : currentZh;
      if (r) renderResume(r, lang);
      else el.resumePreview.innerHTML = '<p class="empty-tip">' + (lang === "en" ? "请先生成中文，再点「下载英文简历」生成英文版。" : "请先生成。") + "</p>";
    });
    if (cfg.key) testConnection();
  }
  init();
})();
