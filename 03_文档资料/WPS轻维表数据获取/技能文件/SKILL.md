---
name: wps-lighttable-api
version: "1.0"
description: >
  WPS 轻维表 AirScript Webhook 数据获取技能。
  当用户提到以下需求时自动触发：
  - 获取 WPS 轻维表数据
  - 读取 WPS 多维表数据
  - 拉取轻维表 / 多维表
  - WPS AirScript 调用
  - Webhook 调用 WPS
  - KSDrive API 数据读取
  - WPS 云文档数据获取
  - 从 WPS 拉数据
  - WPS 数据自动化
author: WorkBuddy AI（华东区域公司实战整理）
created: 2026-03-24
---

# WPS 轻维表数据获取技能

## 技能说明

这个技能封装了通过 WPS AirScript Webhook 自动获取轻维表数据的完整流程。
由于 WPS 端的操作（进入 AirScript、生成令牌等）只能由人工完成，
技能采用「AI 生成脚本 + 人工配置 WPS + AI 调用 Webhook」的人机协作模式。

---

## 执行流程（SOP）

### 第一步：AI 询问配置信息

向用户收集以下信息：
1. **WPS 文件 URL**（如 `https://365.kdocs.cn/l/cpKK6tCXUzEV`）
2. **目标表名称**（如「未出库数据」）或已知的 SheetId
3. **需要读取的字段名**（如不清楚，先运行探查脚本）
4. **是否已有 AirScript 令牌和 Webhook 地址**（有就直接跳到第四步）

---

### 第二步：AI 生成 AirScript 脚本

根据收集到的信息，生成以下两类脚本之一：

#### A. 探查脚本（不知道 SheetId 时先用这个）

```javascript
// 探查脚本：列出文件中所有多维表
let sheets = Array.from(Application.Sheet.GetSheets());
let result = sheets.map(s => ({ id: s.id, name: s.name, count: s.recordsCount }));
return result;
```

#### B. 全量读取脚本模板（确认 SheetId 后使用）

```javascript
// AirScript 1.0 全量分页读取脚本
// 修改项：TARGET_SHEET_ID 和下方 return 的字段名

const TARGET_SHEET_ID = __SHEET_ID__; // 替换为实际 SheetId

try {
  let allRecords = [];
  let offset = null;

  do {
    let params = { SheetId: TARGET_SHEET_ID, PageSize: 100 };
    if (offset !== null) params.Offset = String(offset); // 注意：必须是字符串

    let result  = Application.Record.GetRecords(params);
    let records = Array.from(result.records || []); // 注意：必须用 Array.from()
    allRecords  = allRecords.concat(records);
    offset = result.offset ? String(result.offset) : null;

  } while (offset);

  // 提取字段（按实际字段名修改）
  let summary = allRecords.map(r => {
    let f = r.fields;
    return {
      id: r.id,
      // __FIELDS__ // 在此处列出需要的字段
    };
  });

  return { success: true, total: allRecords.length, data: summary };

} catch (e) {
  return { success: false, error: e.message };
}
```

---

### 第三步：指导用户在 WPS 配置（人工操作）

给用户以下操作指引：

```
请按以下步骤操作 WPS：

1. 打开文件：[填入文件 URL]
2. 顶部菜单 → 工具 → AirScript（⚠️ 必须选 1.0 版本）
3. 如需跨文件读取：右侧「服务」→ 开启「KSDrive」
4. 新建/打开脚本，将上面的脚本粘贴进去，按 Ctrl+S 保存
5. 点击右上角「更多」→「脚本令牌」→「生成令牌」，复制令牌
6. 在浏览器地址栏找 script_id 参数（V2-xxxx 格式），复制脚本 ID

完成后把以下信息告诉 AI：
- 令牌字符串
- 脚本 ID（或完整的 Webhook URL）
```

**Webhook URL 格式：**
```
https://365.kdocs.cn/api/v3/ide/file/{文件ID}/script/{脚本ID}/sync_task
```

其中文件 ID 是 WPS 链接末尾的字母数字组合（如 `cpKK6tCXUzEV`）。

---

### 第四步：AI 调用 Webhook 拉取数据

用户提供令牌和地址后，AI 自动执行以下 PowerShell 命令：

```powershell
$response = Invoke-RestMethod `
  -Uri    "{WEBHOOK_URL}" `
  -Method POST `
  -Headers @{"Content-Type"="application/json"; "AirScript-Token"="{TOKEN}"} `
  -Body   '{"Context":{"argv":{}}}' `
  -TimeoutSec 120

$result = $response.data.result
Write-Host "状态: $($result.success), 总条数: $($result.total)"
$result.data | ConvertTo-Json -Depth 5 | Out-File "wps_data_raw.json" -Encoding utf8
```

读取保存的数据（Python）：
```python
import json
with open('wps_data_raw.json', encoding='utf-8-sig') as f:  # 注意 utf-8-sig
    data = json.load(f)
```

---

### 第五步：AI 分析数据

数据拉取成功后，根据用户需求进行：
- 统计汇总（按部门/业务员/时间等维度）
- 超期/异常预警分析
- 生成可视化 HTML 报告
- 推送到企业微信群
- 导出 Excel 文件
- 配置定时自动化

---

## 踩坑提醒（每次调用前检查）

| 坑 | 解决方案 |
|----|---------|
| AirScript 版本错 | 必须选 **1.0 版本** |
| `GetSheets()` 字段为空 | 必须用 `Array.from()` 包裹 |
| `GetRecords()` Offset 报错 | Offset 必须是**字符串**，用 `String(offset)` |
| Python 读 JSON 乱码 | 用 `encoding='utf-8-sig'` |
| Webhook 超时 | 设置 TimeoutSec ≥ 120 |
| 找不到 SheetId | 先运行探查脚本 |

---

## 本次实战配置（华东区域公司）

下次调用时，以下配置可直接复用：

- **文件 URL**：`https://365.kdocs.cn/l/cpKK6tCXUzEV`
- **文件 ID**：`cpKK6tCXUzEV`
- **脚本 ID**：`V2-4bXnBOROXI2Woup5wsxXsC`
- **令牌**：`1vZUVJ1UCD2FFP4MXCILtD`（有效至约 2026-09-21）
- **Webhook URL**：`https://365.kdocs.cn/api/v3/ide/file/cpKK6tCXUzEV/script/V2-4bXnBOROXI2Woup5wsxXsC/sync_task`

已知 SheetId 对照表：

| 表名 | SheetId | 记录数 |
|------|---------|--------|
| 未出库数据 | 2989 | 428 |
| 未开票业务检查 | 2990 | 1092 |
| 出库及审核流程 | 2991 | 5534 |
| 已审核出库记录表 | 2994 | 1563 |
| 采购订单登记 | 2996 | 533 |
| 供应商台账 | 2992 | 41 |
| 工程客户台账 | 3002 | 131 |

---

## 技能文件位置

```
C:\Users\28141\.workbuddy\skills\wps-lighttable-api\
├── SKILL.md          # 本文件（技能定义 + SOP）
└── scripts\
    ├── fetch_data.ps1         # PowerShell 调用脚本
    └── airscript_template.js  # AirScript 模板
```
