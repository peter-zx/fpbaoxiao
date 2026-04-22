# 报销费用填写工具

**Expense & Reimbursement Tool** — 报销费用记录、管理与 Excel 导出工具。

---

## 功能特点

- **两种模板**：费用模板（不计入销售成本）+ 报销模板（计入销售订单成本）
- **截图上传**：上传凭证截图，自动嵌入导出的 Excel
- **数据持久化**：本地 JSON 存储，支持多用户数据隔离
- **Excel 导出**：一键导出，支持 xlsxwriter / Excel COM 双引擎
- **视图切换**：记录支持按天 / 周 / 月 / 年 分组展示
- **用户系统**：注册 / 登录，数据按用户隔离

---

## 快速开始

### 本地运行

```bash
# 安装依赖
pip install -r requirements.txt

# 启动服务
python main.py
# 或双击 start.bat（Windows）

# 访问
http://localhost:8765
```

### 目录结构

```
baoxiao/
├── main.py              # 主入口
├── server.py            # HTTP 服务器（含 API 路由）
├── app/
│   ├── config.py        # 环境配置
│   ├── store.py         # JSON 数据存储
│   ├── server.py        # API 实现
│   └── excel_export.py  # Excel 导出引擎
├── static/
│   └── index.html       # 前端页面（含 HTML/CSS/JS）
├── data/                 # 数据文件（自动创建）
│   └── data.json
└── requirements.txt
```

### 云端部署（Docker）

```bash
# 在服务器上
cd ~/baoxiao
docker build -t baoxiao .
docker run -d --name baoxiao \
  --memory="400m" \
  --memory-swap="800m" \
  --restart=always \
  -p 8765:8765 \
  baoxiao
```

---

## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/` | 首页 |
| GET | `/static/<file>` | 静态文件 |
| GET | `/api/load` | 加载当前用户数据 |
| POST | `/api/save` | 保存当前用户数据 |
| POST | `/api/register` | 注册用户 |
| POST | `/api/login` | 用户登录 |
| POST | `/api/logout` | 退出登录 |
| GET | `/api/export/excel` | 导出 Excel（需 ?type=expense\|reimburse） |
| GET | `/health` | 健康检查 |

---

## 数据结构

### 费用记录

```json
{
  "createdAt": "2026/04/22 10:51:54",
  "time": "2026-04-22",
  "product": "办公用品",
  "related": "项目A",
  "reason": "日常办公",
  "amount": 150.00,
  "hasTicket": "有票",
  "ticketEntity": "某某公司",
  "image": "data:image/png;base64,..."
}
```

### 用户数据结构

```json
{
  "users": [{ "username": "...", "password": "..." }],
  "expense": [{ ... }],
  "reimburse": [{ ... }]
}
```

---

## 技术栈

- **前端**：原生 HTML / CSS / JavaScript（单文件，~60KB）
- **后端**：Python HTTP 服务器（无第三方框架依赖）
- **Excel 导出**：xlsxwriter（跨平台优先）→ Excel COM（Windows 精确控制）
- **部署**：Docker

---

## 作者

aigc创意人竹相左边
