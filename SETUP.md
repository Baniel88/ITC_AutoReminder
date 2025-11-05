# ITC AutoReminder 项目配置指南

## 项目简介

ITC AutoReminder 是一个自动化工具，用于：
- 自动登录 ITC 系统下载合规报告
- 分析待审核和已撤销的任务
- 通过邮件和 Teams 发送提醒通知

## 快速开始

### 1. 环境要求

- **操作系统**：Windows 10/11
- **Python**：3.8 或更高版本
- **浏览器**：Google Chrome（最新版本）
- **邮件客户端**：Microsoft Outlook（用于发送邮件）

### 2. 一键初始化

运行初始化脚本，自动创建必要的文件夹和安装依赖：

```batch
setup.bat
```

该脚本会自动：
- 创建虚拟环境
- 安装所有 Python 依赖
- 创建必要的文件夹结构
- 生成配置文件模板

### 3. 手动配置（如需要）

如果自动初始化失败，可以手动执行以下步骤：

#### 3.1 创建虚拟环境

```batch
python -m venv .venv
.venv\Scripts\activate
```

#### 3.2 安装依赖

```batch
pip install -r requirements.txt
```

#### 3.3 创建文件夹结构

```
ITC_AutoReminder/
├── ITC report/          # 报表存储目录（自动生成）
│   ├── RawData/        # 原始数据
│   └── Reminder/       # 提醒邮件HTML
├── ChromeProfile_ITC_9233/  # Chrome用户配置（首次运行自动创建）
└── .venv/              # Python虚拟环境
```

## 配置文件说明

### 1. email_config.json

邮件配置文件，控制邮件发送行为和收件人列表。

**示例配置**：

```json
{
    "reports": {
        "Pending review任务提醒": {
            "recipients": [],
            "cc": ["your.email@pg.com"],
            "cc1": {
                "LGDC": ["dc.contact1@pg.com"],
                "TCDC": ["dc.contact2@pg.com"]
            }
        },
        "Revoked状态任务提醒": {
            "recipients": [],
            "cc": ["your.email@pg.com"],
            "cc1": {
                "LGDC": ["dc.contact1@pg.com"],
                "TCDC": ["dc.contact2@pg.com"]
            }
        }
    },
    "system_config": {
        "EMAIL_ENABLED": true,
        "AUTO_SEND_EMAIL": false,
        "MAX_REMAINING_DAYS_FOR_REPORT": 19,
        "URGENCY_LEVELS": {
            "非常紧急": 2,
            "紧急": 4,
            "常规": 10
        },
        "EMAIL_SUBJECT_PENDING": "Pending review任务提醒",
        "EMAIL_SUBJECT_REVOKED": "Revoked状态任务提醒"
    }
}
```

**配置说明**：

- `EMAIL_ENABLED`: 是否启用邮件功能（true/false）
- `AUTO_SEND_EMAIL`: 是否自动发送（false=预览后手动发送，true=直接发送）
- `recipients`: 主要收件人列表
- `cc`: 抄送人列表
- `cc1`: 按DC分组的抄送人（根据任务Category自动添加）

### 2. teams_config.json

Teams 通知配置文件（可选）。

**示例配置**：

```json
{
    "enabled": true,
    "default_webhook": "itc_notifications",
    "webhooks": {
        "itc_notifications": "https://pgone.webhook.office.com/webhookb2/YOUR_WEBHOOK_URL_HERE"
    },
    "team_members": {
        "itc_team": [
            "your.email@pg.com"
        ]
    },
    "notification_rules": {
        "urgent_issues": {
            "webhook": "itc_notifications",
            "mention_team": "itc_team"
        },
        "normal_issues": {
            "webhook": "itc_notifications",
            "mention_team": null
        }
    }
}
```

**配置说明**：

- `enabled`: 是否启用 Teams 通知（true/false）
- `webhooks`: Teams Webhook URL 配置
- `team_members`: 团队成员邮箱列表（用于 @提醒）
- `notification_rules`: 通知规则配置

## 运行方式

### 方式一：直接运行主程序

```batch
python BatRun_ITCreport_downloader_rev1.py
```

### 方式二：使用批处理文件

```batch
Run_OneDriveITC.bat
```

### 方式三：定时任务

使用 Windows 任务计划程序，设置每天自动运行：

1. 打开"任务计划程序"
2. 创建基本任务
3. 选择触发器（例如每天上午 9:00）
4. 操作：启动程序 → 选择 `Run_OneDriveITC.bat`
5. 完成设置

## 首次运行

### 1. Chrome 持久化登录

首次运行时，程序会自动：
1. 启动 Chrome 浏览器（使用持久化配置文件）
2. 打开 ITC 登录页面
3. **手动登录一次**（输入用户名和密码，完成认证）
4. 登录成功后，Chrome 会保存登录状态

**后续运行**：
- 程序会自动复用已保存的登录状态
- 无需再次输入用户名和密码
- 直接下载报表

### 2. 邮件发送测试

建议先设置 `AUTO_SEND_EMAIL: false`，这样：
- 邮件会在 Outlook 中打开预览
- 可以检查收件人和内容
- 手动点击发送

确认无误后，可以改为 `AUTO_SEND_EMAIL: true` 实现全自动。

## 常见问题

### Q1: Chrome 无法启动或端口冲突

**解决方法**：
```batch
python chrome_session_cleanup.py
```
该脚本会清理残留的 Chrome 进程和端口占用。

### Q2: 邮件发送失败

**检查清单**：
- ✅ 确认 Outlook 已安装并配置好邮箱
- ✅ 检查 `email_config.json` 中的邮箱地址格式
- ✅ 确认没有其他程序占用 Outlook

### Q3: 报表下载失败

**可能原因**：
- Chrome 登录状态过期 → 重新手动登录一次
- ITC 系统维护 → 稍后重试
- 网络连接问题 → 检查网络

### Q4: 找不到配置文件

首次运行需要：
1. 复制 `email_config.json.example` → `email_config.json`
2. 复制 `teams_config.json.example` → `teams_config.json`
3. 修改配置文件中的邮箱地址和 Webhook URL

## 文件夹说明

### 需要提交到 Git 的文件

- ✅ 所有 `.py` 脚本文件
- ✅ `requirements.txt`（依赖列表）
- ✅ `SETUP.md`（本文档）
- ✅ `setup.bat`（初始化脚本）
- ✅ `*.example` 配置文件模板
- ✅ README 和文档文件

### 不需要提交的文件（已在 .gitignore 中）

- ❌ `ITC report/`（运行时生成的报表）
- ❌ `ChromeProfile_ITC_*/`（Chrome 用户配置）
- ❌ `chrome_debug_profile_*/`（Chrome 调试配置）
- ❌ `.venv/`（Python 虚拟环境）
- ❌ `email_config.json`（包含真实邮箱地址）
- ❌ `teams_config.json`（包含真实 Webhook URL）
- ❌ `__pycache__/`（Python 缓存）
- ❌ `chromedriver.exe`（会自动下载）

## 项目迁移

### 迁移到新环境

1. **克隆代码库**：
   ```bash
   git clone https://github.com/Baniel88/ITC_AutoReminder.git
   cd ITC_AutoReminder
   ```

2. **运行初始化脚本**：
   ```batch
   setup.bat
   ```

3. **配置邮件和 Teams**：
   - 复制 `email_config.json.example` → `email_config.json`
   - 复制 `teams_config.json.example` → `teams_config.json`
   - 修改配置文件中的邮箱和 Webhook

4. **首次登录**：
   ```batch
   python BatRun_ITCreport_downloader_rev1.py
   ```
   在打开的 Chrome 中手动登录 ITC 系统

5. **测试运行**：
   确认邮件和 Teams 通知正常工作

## 技术支持

如有问题，请联系：
- 开发者：liang.wq.1@pg.com
- 项目文档：查看 `开发说明文档.md`
- GitHub Issues：提交问题到代码仓库

## 更新日志

- **v1.0** - 初始版本，支持自动下载和邮件提醒
- **v2.0** - 添加 Teams 集成和持久化 Chrome 登录
- **v3.0** - 优化端口管理和多实例支持
