// ===========================================
// WPS AirScript 1.0 - 全量数据读取模板
// ===========================================
// 使用说明：
// 1. 修改 TARGET_SHEET_ID 为目标多维表的 SheetId
// 2. 修改 fieldMapping 中的字段名（中文字段名要与多维表完全一致）
// 3. 粘贴到 WPS AirScript 1.0 编辑器，保存后生成令牌

const TARGET_SHEET_ID = 2989; // ← 修改这里

try {
  let allRecords = [];
  let offset = null;

  // 分页读取（自动处理所有页）
  do {
    let params = { SheetId: TARGET_SHEET_ID, PageSize: 100 };
    if (offset !== null) params.Offset = String(offset); // 必须是字符串！

    let result  = Application.Record.GetRecords(params);
    let records = Array.from(result.records || []);      // 必须用 Array.from()!
    allRecords  = allRecords.concat(records);
    offset = result.offset ? String(result.offset) : null;

  } while (offset);

  // 字段提取（按需修改字段名）
  let data = allRecords.map(r => {
    let f = r.fields;
    return {
      id: r.id,
      // ↓ 在此列出需要的字段（字段名必须与多维表完全一致）
      单据号:     f['单据号'],
      单据类型:   f['单据类型'],
      业务部门:   f['业务部门'],
      业务员:     f['业务员'],
      合作单位:   f['合作单位'],
      品名:       f['品名'],
      规格:       f['规格'],
      重量:       f['重量'],
      提单金额万: f['提单金额（万）'],
      日期:       f['日期'],
      到货日期:   f['到货日期'],
      当前间隔天数:     f['当前间隔天数'],
      超期未出库天数:   f['超期未出库天数(超五天)'],
      审核: f['审核'],
      货齐: f['货齐'],
      供应商: f['供应商'],
      仓库: f['仓库']
    };
  });

  return { success: true, total: allRecords.length, data: data };

} catch (e) {
  return { success: false, error: e.message, stack: e.stack };
}

// ===========================================
// 探查脚本：不知道 SheetId 时先运行这个
// ===========================================
/*
let sheets = Array.from(Application.Sheet.GetSheets());
return sheets.map(s => ({
  id:    s.id,
  name:  s.name,
  count: s.recordsCount
}));
*/
