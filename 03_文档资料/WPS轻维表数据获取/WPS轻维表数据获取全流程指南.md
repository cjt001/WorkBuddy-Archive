# WPS 轻维表数据获取全流程指南

> 版本：v1.0 | 整理日期：2026-03-24 | 基于实战案例「未出库数据分析」整理

---

## 一、整体架构说明

```
WPS 轻维表（云端数据）
        ↓  AirScript 脚本（读取多维表数据）
        ↓  发布为 Webhook 接口
        ↓  HTTP POST 调用（WorkBuddy / Python / PowerShell 均可）
        ↓  获取 JSON 数据
        ↓  本地分析处理（Python / 可视化 / 推送通知）
```

---

## 二、WPS 端操作流程（人工配置，一次性）

### 步骤 1：打开目标文件

访问文件链接，例如：`https://365.kdocs.cn/l/cpKK6tCXUzEV`

确保文件是「多维表格」类型（右上角显示表格图标），普通 ET 在线表格不支持 AirScript 多维表 API。

---

### 步骤 2：进入 AirScript 编辑器

1. 顶部菜单：**工具 → AirScript**
2. **⚠️ 关键：必须选择 1.0 版本**（右上角版本切换，2.0 版本 API 不同）
3. 编辑器打开后，点击左上角「新建脚本」或直接在已有脚本上编辑

---

### 步骤 3：开启 KSDrive 云文档 API 服务

1. 在 AirScript 编辑器右侧面板，找到「**服务**」选项卡
2. 找到「**KSDrive（云文档 API）**」，点击开启
3. 开启后可在脚本中使用 `KSDrive.openFile(url)` 跨文件读取

> 注意：如果只读取当前文件，不需要开启 KSDrive，直接用 `Application.Record.GetRecords()` 即可。跨文件（读取其他文档）才需要 KSDrive。

---

### 步骤 4：写入数据读取脚本

将以下模板粘贴到编辑器中（按需修改 `SheetId` 和字段名）：

```javascript
// ======= WPS 轻维表数据读取脚本模板 =======
// 适用：读取同文件内的多维表数据，通过 Webhook 返回

const TARGET_SHEET_ID = 2989; // 替换为目标多维表的 SheetId

try {
  let allRecords = [];
  let offset = null;
  
  // 分页读取全量数据（每页最多100条，自动翻页）
  do {
    let params = { SheetId: TARGET_SHEET_ID, PageSize: 100 };
    if (offset !== null) params.Offset = String(offset);
    
    let result = Application.Record.GetRecords(params);
    let records = Array.from(result.records || []);
    allRecords = allRecords.concat(records);
    offset = result.offset ? String(result.offset) : null;
    
  } while (offset);
  
  // 提取需要的字段（按实际字段名修改）
  let summary = allRecords.map(r => {
    let f = r.fields;
    return {
      id: r.id,
      单据号: f['单据号'],
      业务部门: f['业务部门'],
      业务员: f['业务员'],
      合作单位: f['合作单位'],
      品名: f['品名'],
      重量: f['重量'],
      提单金额万: f['提单金额（万）'],
      日期: f['日期'],
      当前间隔天数: f['当前间隔天数'],
      超期未出库天数: f['超期未出库天数(超五天)'],
      审核: f['审核']
    };
  });
  
  return { success: true, total: allRecords.length, data: summary };

} catch (e) {
  return { success: false, error: e.message };
}
```

**如何查找 SheetId？** 在 AirScript 中先运行以下探查脚本：

```javascript
// 探查脚本：列出所有多维表的 ID 和名称
let sheets = Array.from(Application.Sheet.GetSheets());
let result = sheets.map(s => ({ id: s.id, name: s.name, count: s.recordsCount }));
return result;
```

---

### 步骤 5：发布脚本并获取令牌

1. 脚本写好后，点击右上角「**更多**」按钮
2. 选择「**脚本令牌**」→「**生成令牌**」
3. 复制令牌字符串（格式类似：`1vZUVJ1UCD2FFP4MXCILtD`）

> ⚠️ 令牌有效期 **180天**，到期需要重新生成。令牌相当于访问密码，不要泄露。

---

### 步骤 6：获取 Webhook 地址

Webhook 地址格式固定：

```
https://365.kdocs.cn/api/v3/ide/file/{文件ID}/script/{脚本ID}/sync_task
```

其中：
- **文件 ID**：WPS 文件链接中的字母数字部分，如 `cpKK6tCXUzEV`
- **脚本 ID**：在 AirScript 编辑器的浏览器地址栏中找 `script_id` 参数，格式为 `V2-xxxxxxxx`

**本次案例的 Webhook 地址：**
```
https://365.kdocs.cn/api/v3/ide/file/cpKK6tCXUzEV/script/V2-4bXnBOROXI2Woup5wsxXsC/sync_task
```

---

## 三、WorkBuddy / 程序端调用流程

### 方式 A：WorkBuddy HTTP API 节点调用

在 WorkBuddy 工作流中添加「HTTP API 调用」节点：

| 配置项 | 填写内容 |
|--------|---------|
| 请求方式 | POST |
| URL | `https://365.kdocs.cn/api/v3/ide/file/{文件ID}/script/{脚本ID}/sync_task` |
| 请求头 | `Content-Type: application/json` 和 `AirScript-Token: {你的令牌}` |
| 请求体格式 | JSON |
| 请求体内容 | `{"Context":{"argv":{}}}` |
| 超时时间 | 建议设为 120 秒（数据量大时脚本执行时间较长） |

返回数据结构：
```json
{
  "data": {
    "result": {
      "success": true,
      "total": 428,
      "data": [ {...}, {...} ]
    }
  }
}
```

---

### 方式 B：PowerShell 调用（已验证）

```powershell
$webhookUrl = "https://365.kdocs.cn/api/v3/ide/file/cpKK6tCXUzEV/script/V2-4bXnBOROXI2Woup5wsxXsC/sync_task"
$token = "1vZUVJ1UCD2FFP4MXCILtD"

$response = Invoke-RestMethod `
  -Uri $webhookUrl `
  -Method POST `
  -Headers @{
    "Content-Type" = "application/json"
    "AirScript-Token" = $token
  } `
  -Body '{"Context":{"argv":{}}}' `
  -TimeoutSec 120

# 提取数据
$result = $response.data.result
Write-Host "获取成功: $($result.success), 总条数: $($result.total)"

# 保存到文件（注意用 utf-8-sig 编码，Python 读取时用 encoding='utf-8-sig'）
$result.data | ConvertTo-Json -Depth 5 | Out-File "output_data.json" -Encoding utf8
```

---

### 方式 C：Python 调用

```python
import requests
import json

webhook_url = "https://365.kdocs.cn/api/v3/ide/file/cpKK6tCXUzEV/script/V2-4bXnBOROXI2Woup5wsxXsC/sync_task"
token = "1vZUVJ1UCD2FFP4MXCILtD"

headers = {
    "Content-Type": "application/json",
    "AirScript-Token": token
}
body = {"Context": {"argv": {}}}

response = requests.post(webhook_url, headers=headers, json=body, timeout=120)
result = response.json()

data = result["data"]["result"]["data"]
total = result["data"]["result"]["total"]
print(f"获取成功，共 {total} 条记录")

# 保存数据
with open("output_data.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
```

---

## 四、踩坑记录（实战总结）

| 问题 | 原因 | 解决方案 |
|------|------|---------|
| `GetSheets()` 返回空对象 `{}` | 直接遍历结果，没有用 `Array.from()` | 必须用 `Array.from(Application.Sheet.GetSheets())` |
| `GetRecords()` 报错 | Offset 参数类型错误 | Offset 必须是**字符串**，用 `String(offset)` 转换 |
| Python 读取 JSON 乱码 | PowerShell `Out-File` 默认输出 UTF-8 BOM | Python 读取用 `encoding='utf-8-sig'` |
| Webhook 超时 | 脚本读取大量数据时执行慢 | 设置 TimeoutSec 120 以上 |
| 找不到脚本 ID | 从链接里找 `script_id` | 看浏览器地址栏，格式为 `V2-xxxxxxxxx` |
| AirScript API 报错 | 选了 2.0 版本 | 必须切换到 **1.0 版本** |

---

## 五、本次案例：未出库数据分析

### 背景

- **文件**：华东区域公司运营多维表（`cpKK6tCXUzEV`）
- **目标表**：「未出库数据」，SheetId = 2989
- **数据量**：428 条记录

### 分析结果摘要

| 指标 | 数值 |
|------|------|
| 未出库总金额 | 4271.36 万元 |
| 未出库总重量 | 17,570.3 吨 |
| 超期条数（超5天） | 86 条 |
| 未审核单据 | 42 条 |

**各公司分布：**

| 公司 | 条数 | 金额（万） |
|------|------|-----------|
| 南京公司 | 125 | 1526.96 |
| 昆山公司 | 124 | 1111.23 |
| 杭州公司 | 94 | 920.11 |
| 合肥公司 | 85 | 713.06 |

**超期业务员 TOP3：** 李馨华（37条）、伍舟畅（18条）、张翼鹏（14条）

---

## 六、扩展应用场景

1. **定时自动化**：结合 WorkBuddy 定时任务，每天早上 8 点自动拉取数据并推送超期预警到企业微信
2. **多表联动**：同时读取多个 SheetId 的数据，进行跨表分析
3. **数据写回**：用 `Application.Record.CreateRecords()` 把分析结果写回轻维表
4. **可视化报告**：拉取数据后用 Python + Plotly 生成 HTML 可视化报告

---

*文件路径：`C:\Users\28141\WorkBuddy\20260324211245\WPS轻维表数据获取全流程指南.md`*
