# Mail-Sieve (Hotmail / Outlook 邮箱获取与深度取证审查面板)

**Mail-Sieve** 是一套专为批量采购 Hotmail / Outlook 邮箱打造的两步自动化**卡密提取、收件箱深度取证审查与业务可用性分类管理面板**。零第三方 Python 依赖（纯 Python 3.11 标准库实现），开箱即用。

---

## ✨ 核心功能（两步自动化工作流）

### 1️⃣ 第一步：从指定收货邮箱自动获取购买的邮箱信息
支持 4 种接入方式自动提取购买的 `email----password----client_id----refresh_token` 卡密及订单批次号：
- **📨 指定 IMAP 收货邮箱拉取**：支持 QQ 邮箱、Foxmail、163、126、Gmail、Outlook 等，自动扫描收件箱发货邮件正文及 `.txt` / `.csv` 附件。
- **🔑 指定 Outlook (OAuth2) 收货邮箱拉取**：粘贴收货 Outlook 邮箱卡密，通过 Microsoft OAuth2 / Live IMAP 免密拉取发货邮件。
- **☁️ Cloudflare 临时邮箱拉取**：对接 Cloudflare Worker 临时邮箱 API 自动拉取收件。
- **📋 批量粘贴 / 本地工作区同步**：支持粘贴整封发货邮件、上传文件，或一键与 `gptcli-2api` 邮箱池（`用于注册的邮箱.txt`、`outlook-accounts.txt`、`turb.sqlite3`）双向同步。

### 2️⃣ 第二步：多线程并发收件箱审查与二手翻新取证
通过 `Microsoft Live OAuth2 + IMAP (XOAUTH2)` 与 `Microsoft Graph API` 并发连接每个邮箱，深度扫描收件箱历史邮件：
- **🚨 号商改别名翻新（Alias-Recycled）检测**：
  - 自动比对历史邮件 `To` / `Delivered-To` / `X-Original-To` 邮件头及正文中的前身别名；
  - 自动核查微软安全通知（`连接到 Microsoft 帐户的新应用` / `Thunderbird connected to...`）中的掩码账号（如 `ox**1@outlook.com`）与当前邮箱是否不匹配。
- **🚨 封号与风控锁号黑历史检测**：
  - 自动检出 `OpenAI - Access Deactivated`（封号退信）与 `Important: Action Required to Secure Your ChatGPT Account (locked)`（风控锁号邮件）。
- **🔎 13+ 类常见互联网与 AI 服务注册指纹识别**：
  - 自动识别 `OpenAI / ChatGPT`、`xAI / Grok`、`Anthropic / Claude`、`Cursor`、`X (Twitter)`、`Apple ID`、`1040.com (美国报税)`、`Uber Eats`、`Google`、`GitHub`、`Discord`、`Telegram`、`HK Ticketing` 等历史注册与验证码邮件。
- **🎯 五级自动分类与“能否用于注册目标业务”精准判定**：
  - `🟢 全新纯净 (clean_ready)`：**可注册**（0 第三方注册痕迹）
  - `🟡 仅注册过其他服务 (other_services_only)`：**可注册目标业务**（未注册过目标业务且无封号/改名黑历史；亦支持切换“严苛纯白模式”）
  - `🟠 已注册目标业务 (target_registered)`：**不可注册**
  - `🔴 封号黑历史 / 别名翻新 (blacklisted_recycled)`：**不可注册**
  - `⚫ 凭证失效 / 死号 (auth_failed)`：**不可注册**（`refresh_token` 失效或邮箱未初始化收件箱）
- **📑 一键生成号商售后维权取证报告**：自动按批次汇总每个问题邮箱的原始邮件时间线、主题、发件人及前身别名证据，导出 Markdown 报告供退换维权。

---

## 🚀 快速部署

### 方式一：Docker Compose 一键部署（推荐）

```bash
git clone https://github.com/lzh3083/mail-sieve.git
cd mail-sieve
docker compose up -d --build
```

启动后访问：`http://<服务器IP>:8899`

### 方式二：Python 直接运行（零外部依赖）

仅需系统自带的 `Python 3.10+`：

```bash
python3 hotmail_manager_server.py --host 0.0.0.0 --port 8899
```

如需启动时自动同步同级 `gptcli-2api` 邮箱池并立即执行全量审查：

```bash
python3 hotmail_manager_server.py --host 0.0.0.0 --port 8899 --auto-init-audit
```
