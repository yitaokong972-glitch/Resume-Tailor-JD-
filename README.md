# Resume Tailor · JD 智能简历定制台

一个开箱即用的本地 Web 应用：上传你的简历 + 目标岗位 JD，由大语言模型按 JD 改写、润色简历正文，
并一键导出排版统一的中 / 英文 Word 文档。无需注册、无需联网账号，模型密钥只在服务端配置一次。

> 面向求职者、实习生的简历优化工具：把"我有哪些经历"和"岗位要什么"对齐，输出可直接投递的定制简历。

---

## 🌐 在线使用

本仓库提供两种方式，按你的场景选：

- **方式 A · 分享版（推荐，别人免 Key 直接用）**：把后端部署到云平台（Railway / Render），你的 DeepSeek Key 装在服务端环境变量，**任何人打开链接即可直接用 DeepSeek 改写，无需填 Key**。这是"发给别人也能用"的版本。部署步骤见下方「部署到云平台」。
- **方式 B · 自填 Key 版（纯前端，已上线）**：https://yitaokong972-glitch.github.io/Resume-Tailor-JD-/
  > 密钥只存在你自己的浏览器 localStorage，适合只想自己用、不想把 Key 暴露给服务端的情况。源码在仓库 `site/` 目录。

[![Deploy on Railway](https://railway.app/button.svg)](https://railway.app/new/template?url=https://github.com/yitaokong972-glitch/Resume-Tailor-JD-)

---

## ✨ 功能特性

- **按 JD 智能改写**：根据岗位描述与你的补充要求，重写简历 bullet、突出匹配经历、补全缺失能力。
- **中 / 英双语导出**：内置专业中译英规则，导出排版统一、可直接投递的 `.docx`。
- **多方向模板**：律所 / 法务、金融 / 投融资、国央企 / 合规等方向有不同的 bullet 表述策略。
- **联网检索增强**（可选）：结合公开检索补充行业 / 岗位语境，再交给模型改写。
- **纯本地兜底**：未配置模型密钥时，使用内置规则生成，保证服务始终可用。
- **零前端密钥**：API Key 只在服务端配置，网页端不收集、不显示任何密钥。
- **隐私优先**：用户上传的简历仅在本地服务端处理，只在改写时发送给你所配置的模型服务。

---

## 🧱 技术栈

| 层 | 技术 |
| --- | --- |
| 后端 | Python 3.9+ · 标准库 `http.server`（零额外 Web 框架） |
| 模型接入 | OpenAI 兼容 Chat Completions 接口（可接入 DeepSeek / OpenAI / 通义千问 / 智谱 GLM / 本地 Ollama 等任意兼容服务） |
| 文档生成 | `python-docx` · `lxml` |
| 解析 | `pdfplumber`（PDF）· `python-docx`（Word） |
| 前端 | 原生 HTML / CSS / JavaScript（无构建步骤） |

---

## 🚀 快速开始

### 1. 克隆与安装依赖

```bash
git clone <your-repo-url> resume-tailor
cd resume-tailor
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. 配置模型密钥

任选其一（推荐方式一）：

**方式一 · 环境变量**

```bash
export LLM_API_KEY="sk-xxxxxxxx"
export LLM_API_BASE="https://api.deepseek.com"   # 可替换为任意 OpenAI 兼容地址
export LLM_MODEL="deepseek-chat"
```

**方式二 · 密钥文件**

在项目根目录创建 `.llm_key`，内容仅为你的 API Key 一行：

```bash
echo "sk-xxxxxxxx" > .llm_key
```

> 也兼容旧文件名 `.deepseek_key`。无论哪种方式，密钥文件都已被 `.gitignore` 忽略，不会进入仓库。

### 3. 启动

```bash
python server.py
```

打开浏览器访问：

- 本机：`http://127.0.0.1:8765/`
- 同局域网其他设备：用本机局域网 IP，如 `http://192.168.x.x:8765/`

### 4. 使用

1. 在「JD」框粘贴目标岗位描述；
2. 在「原简历」框粘贴简历文本，或上传简历 / 补充材料（支持 Word / PDF / 文本）；
3. 点「生成修改意见」查看 AI 改写预览；
4. 点「生成中文简历 / 英文简历」下载排版好的 Word。

---

## ⚙️ 配置项（环境变量）

| 变量 | 说明 | 默认值 |
| --- | --- | --- |
| `LLM_API_KEY` | 模型服务 API Key | 空（未配置则走本地规则兜底） |
| `LLM_API_BASE` | OpenAI 兼容接口地址 | `https://api.deepseek.com` |
| `LLM_MODEL` | 模型名 | `deepseek-chat` |

兼容旧变量名 `DEEPSEEK_API_KEY` / `DEEPSEEK_MODEL`，便于从旧部署迁移。

---

## 🔒 隐私说明

- 用户上传的简历、JD、补充材料**仅在运行本服务的本机处理**，不会上传到任何第三方存储；
- 仅在"调用模型改写"时，相关内容会按你的配置发送给 `LLM_API_BASE` 指向的模型服务；
- 未配置密钥时，服务使用内置规则在本地生成，完全不联网；
- 密钥文件（`.llm_key`）与运行时目录（`uploads/`、`generated/`）均已加入 `.gitignore`，不会进入仓库。

---

## 📁 目录结构

```
resume-tailor/
├── server.py            # 后端：HTTP 服务 + 模型改写 + Word 生成
├── static/
│   ├── index.html       # 页面结构
│   ├── app.js           # 前端交互逻辑
│   └── styles.css       # 样式
├── materials/           # 可选共享素材库（空目录即可）
├── requirements.txt     # Python 依赖
├── .env.example         # 配置样例
├── .gitignore
├── LICENSE
└── README.md
```

---

## 🌐 部署到云平台（获取稳定公网 URL）

本项目是标准 Python 后端（持有你的模型 Key），已内置 `Procfile` 与 `$PORT` 支持，可部署到任意支持 `Procfile` 的 PaaS，获得 7×24 固定公网地址；**任何人打开链接即可直接用 DeepSeek，无需自己填 Key**。

> **⚠️ 关键前置（解决 "Deploy from GitHub repo 找不到本仓库"）**：Railway / Render 默认看不到你的私有仓库，因为 **GitHub App 还没被授权访问它**。必须先在本机浏览器做一步授权：
> 进入 **GitHub → 右上角头像 → Settings → Integrations → Applications**（或 *Authorized OAuth Apps*）→ 找到 **Railway**（Render 同理）→ **Repository access** → 选择 **All repositories**（或仅勾选 `Resume-Tailor-JD-`）→ **Save**。授权后回到平台刷新，仓库就会出现。
> 此外请确认 Railway / Render 登录的 GitHub 账号就是 `yitaokong972-glitch` 本人。

### Railway（推荐，有免费额度）

1. 完成上面的 GitHub App 授权；
2. 注册 https://railway.app ，用 GitHub 登录；
3. **New Project → Deploy from GitHub repo** → 选中本仓库（授权后应可见）；
4. **Variables** 添加：`LLM_API_KEY=你的key`、`LLM_API_BASE=https://api.deepseek.com`、`LLM_MODEL=deepseek-chat`；
5. 部署完成后 Railway 自动分配 `https://xxx.up.railway.app` 固定地址，把上面的「方式 A」地址替换成它即可。

### Render（有免费额度）

1. 注册 https://render.com ，用 GitHub 登录；
2. **New → Web Service** → 关联本仓库；
3. **Build Command**：`pip install -r requirements.txt`；**Start Command**：`python server.py`；
4. **Environment** 添加上述三个变量；
5. 部署完成后获得 `https://xxx.onrender.com` 固定地址。

> 部署后服务默认读取 `LLM_API_KEY` 环境变量连接模型；模型调用从云平台服务器直接访问 `LLM_API_BASE`，不经过任何本机代理，连接稳定。

---

## 📄 许可证

[MIT](./LICENSE) — 可自由用于学习、二次开发与分发。
