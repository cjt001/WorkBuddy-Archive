# WPS 轻维表数据拉取流程整理

> 整理日期：2026-03-25  
> 适用平台：【新高达】采购订单协同平台 / 【新高达】物流协同平台

---

## 一、基础信息

| 项目 | 说明 |
|------|------|
| 通用令牌 | `1vZUVJ1UCD2FFP4MXCILtD`（有效至约 2026-09-21） |
| AirScript 版本 | 1.0（必须选 1.0，不支持其他版本） |
| Offset 参数类型 | 必须是字符串，初始用 `null`，翻页用 `String(result.offset)` |
| 返回值包装 | 必须用 `Array.from()` 包裹 GetRecords() 结果 |

---

## 二、文件一：【新高达】采购订单协同平台

- **URL**：`https://365.kdocs.cn/l/ck0HLA8yfG4E`
- **Webhook**：`https://365.kdocs.cn/api/v3/ide/file/ck0HLA8yfG4E/script/V2-6yXwVgLiMnCakbaxDuqiXV/sync_task`

| 表名 | SheetId | 记录数（参考） |
|------|---------|--------------|
| 采购订单登记 | 3 | 491 |
| 采购订单报表 | 4 | — |
| 供应商台账 | 5 | — |

### 2.1 采购订单登记表 - AirScript 脚本（全字段）

```javascript
// 采购订单登记 - 全量读取脚本
// SheetId: 3，预计 491 条
const TARGET_SHEET_ID = 3;

try {
  let allRecords = [];
  let offset = null;

  do {
    let params = { SheetId: TARGET_SHEET_ID, PageSize: 100 };
    if (offset !== null) params.Offset = String(offset);

    let result = Application.Record.GetRecords(params);
    let records = Array.from(result.records || []);
    allRecords = allRecords.concat(records);
    offset = (result.offset !== undefined && result.offset !== null)
               ? String(result.offset)
               : null;
  } while (offset !== null);

  let data = allRecords.map(r => {
    let f = r.fields;
    return {
      id: r.id,
      日期: f['日期'],
      采购订单号: f['采购订单号'],
      合作单位: f['合作单位'],
      项目名称: f['项目名称'],
      中类: f['中类'],
      小类: f['小类'],
      品名: f['品名'],
      材质: f['材质'],
      规格: f['规格'],
      产地: f['产地'],
      数量: f['数量'],
      重量: f['重量'],
      仓库: f['仓库'],
      业务部门: f['业务部门'],
      业务员: f['业务员'],
      订单完结: f['订单完结'],
      审核日期: f['审核日期'],
      货齐日期: f['货齐日期'],
      备注: f['备注'],
    };
  });

  return { success: true, total: allRecords.length, data: data };

} catch (e) {
  return { success: false, error: e.message, stack: e.stack };
}
```

### 2.2 Webhook 调用（PowerShell）

```powershell
$wh  = "https://365.kdocs.cn/api/v3/ide/file/ck0HLA8yfG4E/script/V2-6yXwVgLiMnCakbaxDuqiXV/sync_task"
$tok = "1vZUVJ1UCD2FFP4MXCILtD"

$resp = Invoke-RestMethod `
    -Uri $wh `
    -Method POST `
    -Headers @{ "Content-Type" = "application/json"; "AirScript-Token" = $tok } `
    -Body '{"Context":{"argv":{}}}' `
    -TimeoutSec 120

# 保存到本地 JSON
$resp.data.result | ConvertTo-Json -Depth 10 | `
    Out-File "wps_xingaoda_caigou.json" -Encoding utf8

Write-Host "成功: $($resp.data.result.success), 条数: $($resp.data.result.total)"
```

---

## 三、文件二：【新高达】物流协同平台

- **URL**：`https://365.kdocs.cn/l/ccHhW7rnkiaC`
- **Webhook**：`https://365.kdocs.cn/api/v3/ide/file/ccHhW7rnkiaC/script/V2-1DuEkuWfDWr2JKam6Bi47X/sync_task`

| 表名 | SheetId | 记录数（参考） |
|------|---------|--------------|
| 物流提单明细表 | 3 | 474 |
| 物流订单清单 | 4 | 121 |
| 承运单位入场前安全生产检查表 | 5 | 8 |
| 各单位链接清单 | 7 | 4 |

### 3.1 物流订单清单 - AirScript 脚本（全字段）

```javascript
// 物流订单清单 - 全量读取脚本
// SheetId: 4，预计 121 条
const TARGET_SHEET_ID = 4;

try {
  let allRecords = [];
  let offset = null;

  do {
    let params = { SheetId: TARGET_SHEET_ID, PageSize: 100 };
    if (offset !== null) params.Offset = String(offset);

    let result = Application.Record.GetRecords(params);
    let records = Array.from(result.records || []);
    allRecords = allRecords.concat(records);
    offset = (result.offset !== undefined && result.offset !== null)
               ? String(result.offset)
               : null;
  } while (offset !== null);

  let data = allRecords.map(r => {
    let f = r.fields;
    return {
      id: r.id,
      日期: f['日期'],
      单据号: f['单据号'],
      合作单位: f['合作单位'],
      项目名称: f['项目名称'],
      中类: f['中类'],
      仓库: f['仓库'],
      提货方式: f['提货方式'],
      承运单位: f['承运单位'],
      业务部门: f['业务部门'],
      业务员: f['业务员'],
      备注: f['备注'],
      签单上传时间: f['签单上传时间'],
      判断上传签单: f['判断上传签单'],
      超期天数: f['超期天数'],
      全部单据已上传: f['全部单据已上传'],
    };
  });

  return { success: true, total: allRecords.length, data: data };

} catch (e) {
  return { success: false, error: e.message, stack: e.stack };
}
```

### 3.2 Webhook 调用（PowerShell）

```powershell
$wh  = "https://365.kdocs.cn/api/v3/ide/file/ccHhW7rnkiaC/script/V2-1DuEkuWfDWr2JKam6Bi47X/sync_task"
$tok = "1vZUVJ1UCD2FFP4MXCILtD"

$resp = Invoke-RestMethod `
    -Uri $wh `
    -Method POST `
    -Headers @{ "Content-Type" = "application/json"; "AirScript-Token" = $tok } `
    -Body '{"Context":{"argv":{}}}' `
    -TimeoutSec 120

# 保存到本地 JSON
$resp.data.result | ConvertTo-Json -Depth 10 | `
    Out-File "wps_wuliu_qingdan.json" -Encoding utf8

Write-Host "成功: $($resp.data.result.success), 条数: $($resp.data.result.total)"
```

---

## 四、踩坑记录

| 问题 | 原因 | 解决方案 |
|------|------|---------|
| Offset 传数字报错 | GetRecords() 要求 Offset 必须是字符串 | 初始用 `null`，翻页用 `String(result.offset)` |
| GetRecords() 返回值不可直接用 | 返回的是类数组对象，不是原生数组 | 必须用 `Array.from()` 包裹 |
| Python 读 JSON 乱码 | PowerShell Out-File 写出带 BOM 的 UTF-8 | 用 `encoding='utf-8-sig'` 读取 |
| AirScript 版本不兼容 | 新版本 API 有差异 | 必须选 **AirScript 1.0** 版本 |

---

## 五、数据拉取快速流程（每次使用步骤）

1. 打开对应的 WPS 文件，进入 AirScript 编辑器
2. 选择 **AirScript 1.0** 版本
3. 将对应脚本粘贴到编辑器，**Ctrl+S 保存**（不用手动运行）
4. 在本地执行对应的 PowerShell 命令调用 Webhook
5. 数据自动保存为本地 JSON 文件
6. 用 Python 读取时指定 `encoding='utf-8-sig'`

---

*文档由 WorkBuddy 自动整理 | 2026-03-25*
