# WPS 多维表格完整开发指南

> 文档来源：WPS开放平台 https://365.kdocs.cn/3rd/open/documents/app-integration-dev/guide/dbsheet/
> 整理日期：2026-04-06

---

## 目录

1. [能力体系总览](#能力体系总览)
2. [AirScript 开发指南](#airscript-开发指南)
   - [AirScript 简介](#airscript-简介)
   - [快速入门](#快速入门)
   - [脚本令牌（APIToken）](#脚本令牌apitoken)
   - [高级服务 API](#高级服务-api)
   - [内置基础类型](#内置基础类型)
   - [脚本经典案例](#脚本经典案例)
3. [内嵌 SDK 开发指南](#内嵌-sdk-开发指南)
   - [SDK 简介](#sdk-简介)
   - [快速入门（1.x / 3.x）](#快速入门1x--3x)
   - [配置参数](#配置参数)
   - [事件监听](#事件监听)
4. [API 参考文档](#api-参考文档)
   - [Application](#application)
   - [Sheet / Sheets](#sheet--sheets)
   - [View / Views](#view--views)
   - [Record / Records](#record--records)
   - [Field / Fields](#field--fields)
   - [FieldDescriptor](#fielddescriptor)
   - [RecordRange](#recordrange)

---

## 能力体系总览

WPS 多维表格面向开发者提供三种主要开放方式：

| 开放方式 | 适用场景 | 说明 |
|---|---|---|
| **在线脚本 AirScript** | 轻量级自动化、数据处理、对接外部系统 | 在多维表格内直接编写 JS 脚本，支持外部触发 |
| **内嵌 SDK** | 将多维表格嵌入企业自有系统/网页 | 提供 WebOffice SDK 1.x（iframe）和 3.x（微前端）两个版本 |
| **多维表格插件** | 扩展表格功能，提供自定义视图/面板 | 通过插件体系在表格内嵌入自定义前端应用 |

---

## AirScript 开发指南

### AirScript 简介

AirScript 是 WPS 多维表格提供的轻量级脚本应用开发平台。它基于 JavaScript 语言，让用户可以直接在多维表格内编写脚本，实现数据处理自动化、与外部系统对接等功能。

**核心特性：**

- 运行环境为**服务端（Node.js 类）环境**，不是浏览器环境
- 支持**同步 API 调用**（无需 async/await，所有 API 均可同步调用）
- 内置多种高级服务：网络请求、云文档操作、邮件发送、数据库访问
- 支持通过"脚本令牌"从外部系统触发执行

**与浏览器 JS 的主要区别：**

- 没有 `window`、`document` 等浏览器 API
- 没有 DOM 操作
- 可以直接同步调用异步操作（如网络请求）

---

### 快速入门

#### 第一步：打开脚本编辑器

在多维表格中，点击右上角"工具"→"脚本"，进入 AirScript 编辑器。

#### 第二步：编写你的第一个脚本

```javascript
// 获取当前激活的表格
const sheet = Application.Sheets.ActiveSheet;

// 获取所有记录
const records = sheet.Records.GetAll();

// 遍历记录并打印字段值
records.forEach(record => {
  const name = record.GetCellValue('姓名');
  const age = record.GetCellValue('年龄');
  console.log(`姓名: ${name}, 年龄: ${age}`);
});
```

#### 第三步：运行脚本

点击编辑器右上角"运行"按钮，脚本将立即执行，结果显示在下方的控制台中。

#### 基础示例：写入数据

```javascript
// 向表格中添加一条新记录
const sheet = Application.Sheets.ActiveSheet;
const records = sheet.Records;

// 添加记录（字段名: 值）
records.Add({
  '姓名': '张三',
  '年龄': 28,
  '部门': '技术部'
});

console.log('记录添加成功');
```

#### 基础示例：修改记录

```javascript
const sheet = Application.Sheets.ActiveSheet;
const records = sheet.Records.GetAll();

// 修改第一条记录
if (records.length > 0) {
  records[0].SetCellValue('状态', '已完成');
  console.log('记录修改成功');
}
```

---

### 脚本令牌（APIToken）

脚本令牌（APIToken）是用于**从外部系统触发 AirScript 脚本执行**的凭证机制。通过脚本令牌，外部系统可以调用 WPS 多维表格中的脚本，实现与企业内部系统的集成。

#### 创建脚本令牌

1. 在 AirScript 编辑器中，点击"令牌管理"
2. 点击"新建令牌"
3. 填写令牌名称、描述，选择关联的脚本
4. 点击"确认"，系统生成令牌字符串

**注意：令牌只显示一次，请妥善保存。**

#### 使用脚本令牌调用脚本

##### 同步执行接口

外部系统发送 HTTP 请求，等待脚本执行完成后返回结果。

**请求方式：**`POST`

**接口地址：**
```
https://www.kdocs.cn/api/v3/office/file/{fileId}/script/token/run
```

**请求头：**
```
Content-Type: application/json
Authorization: Bearer {access_token}
```

**请求体：**
```json
{
  "token": "your_script_token_here",
  "params": {
    "key1": "value1",
    "key2": "value2"
  }
}
```

**响应：**
```json
{
  "result": 0,
  "data": {
    "output": "脚本的返回值或console输出"
  }
}
```

##### 异步执行接口

外部系统发送 HTTP 请求，立即得到任务 ID，之后轮询查询执行结果。

**触发执行：**
```
POST https://www.kdocs.cn/api/v3/office/file/{fileId}/script/token/async_run
```

**查询执行结果：**
```
GET https://www.kdocs.cn/api/v3/office/file/{fileId}/script/task/{taskId}/result
```

**响应（执行中）：**
```json
{
  "result": 0,
  "data": {
    "status": "running",
    "taskId": "xxxx"
  }
}
```

**响应（执行完成）：**
```json
{
  "result": 0,
  "data": {
    "status": "success",
    "output": "脚本返回值"
  }
}
```

#### 在脚本中接收外部参数

```javascript
// 脚本内通过 Context.Params 获取外部传入的参数
const params = Context.Params;
const userId = params.userId;
const action = params.action;

console.log(`收到外部请求，用户ID: ${userId}, 操作: ${action}`);

// 根据参数执行不同逻辑
if (action === 'add') {
  const sheet = Application.Sheets.ActiveSheet;
  sheet.Records.Add({ '用户ID': userId, '状态': '新增' });
}

// 返回结果给外部系统
return { success: true, message: '操作完成' };
```

---

### 高级服务 API

AirScript 提供多种高级服务，用于与外部系统交互。

#### 1. 网络 API（HTTP）

通过 `Network` 服务发送 HTTP 请求。

```javascript
// GET 请求
const response = Network.Fetch('https://api.example.com/users');
const data = JSON.parse(response.Content);
console.log(data);

// POST 请求（发送 JSON）
const postResponse = Network.Fetch('https://api.example.com/create', {
  Method: 'POST',
  Headers: {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer token123'
  },
  Body: JSON.stringify({
    name: '张三',
    age: 28
  })
});
console.log(postResponse.StatusCode); // 200
console.log(postResponse.Content);    // 响应体字符串

// POST 请求（发送 Form 表单）
const formResponse = Network.Fetch('https://api.example.com/upload', {
  Method: 'POST',
  Headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  Body: 'key1=value1&key2=value2'
});
```

**Network.Fetch 参数说明：**

| 参数 | 类型 | 说明 |
|---|---|---|
| url | string | 请求地址 |
| Method | string | HTTP方法：GET/POST/PUT/DELETE等 |
| Headers | object | 请求头键值对 |
| Body | string | 请求体（字符串） |
| Timeout | number | 超时时间（毫秒），默认 30000 |

**Response 对象属性：**

| 属性 | 类型 | 说明 |
|---|---|---|
| StatusCode | number | HTTP 状态码 |
| Content | string | 响应体字符串 |
| Headers | object | 响应头 |

---

#### 2. 云文档 API（KSDrive）

通过 `KSDrive` 服务操作 WPS 云文档（金山文档）。

```javascript
// 获取文件列表
const files = KSDrive.GetFiles({
  ParentId: 'root',  // 根目录
  Type: 'all'        // 所有类型
});

files.forEach(file => {
  console.log(`${file.Name} - ${file.Type} - ${file.Id}`);
});

// 读取文档内容
const content = KSDrive.GetDocContent('file_id_here');
console.log(content);

// 创建文件夹
const folder = KSDrive.CreateFolder({
  Name: '新建文件夹',
  ParentId: 'root'
});
console.log('文件夹ID:', folder.Id);

// 上传文件
const uploadResult = KSDrive.UploadFile({
  FileName: 'test.txt',
  Content: 'Hello, World!',
  ParentId: 'root'
});
console.log('上传成功，文件ID:', uploadResult.FileId);
```

---

#### 3. 邮件 API（SMTP）

通过 `SMTP` 服务发送邮件（需要配置 SMTP 服务器信息）。

```javascript
// 发送邮件
const result = SMTP.SendMail({
  Host: 'smtp.example.com',
  Port: 465,
  SSL: true,
  Username: 'sender@example.com',
  Password: 'email_password',
  From: 'sender@example.com',
  To: ['receiver@example.com'],
  CC: ['cc@example.com'],        // 可选
  Subject: '邮件主题',
  Body: '<h1>这是邮件正文</h1><p>支持 HTML 格式</p>',
  IsHtml: true                   // 是否为 HTML 邮件
});

if (result.Success) {
  console.log('邮件发送成功');
} else {
  console.log('发送失败:', result.Error);
}
```

**SMTP.SendMail 参数说明：**

| 参数 | 类型 | 必填 | 说明 |
|---|---|---|---|
| Host | string | 是 | SMTP 服务器地址 |
| Port | number | 是 | SMTP 端口（通常 465 或 587） |
| SSL | boolean | 否 | 是否启用 SSL/TLS |
| Username | string | 是 | SMTP 账户名 |
| Password | string | 是 | SMTP 账户密码 |
| From | string | 是 | 发件人邮箱 |
| To | string[] | 是 | 收件人邮箱列表 |
| CC | string[] | 否 | 抄送邮箱列表 |
| BCC | string[] | 否 | 密送邮箱列表 |
| Subject | string | 是 | 邮件主题 |
| Body | string | 是 | 邮件正文 |
| IsHtml | boolean | 否 | 是否为 HTML 格式，默认 false |
| Attachments | object[] | 否 | 附件列表 |

---

#### 4. 数据库 API（SQL）

通过 `DB` 服务连接外部数据库执行 SQL 查询。支持 MySQL、PostgreSQL 等主流数据库。

```javascript
// 连接 MySQL 数据库并查询
const db = DB.Connect({
  Type: 'mysql',
  Host: 'db.example.com',
  Port: 3306,
  Database: 'mydb',
  Username: 'user',
  Password: 'password'
});

// 执行查询
const result = db.Query('SELECT * FROM users WHERE status = ?', ['active']);
result.Rows.forEach(row => {
  console.log(`用户: ${row.name}, 邮箱: ${row.email}`);
});

// 执行更新
const updateResult = db.Execute(
  'UPDATE users SET status = ? WHERE id = ?',
  ['inactive', 123]
);
console.log('影响行数:', updateResult.AffectedRows);

// 执行插入
const insertResult = db.Execute(
  'INSERT INTO logs (action, created_at) VALUES (?, NOW())',
  ['脚本执行']
);
console.log('新记录ID:', insertResult.LastInsertId);

// 使用事务
db.BeginTransaction();
try {
  db.Execute('UPDATE accounts SET balance = balance - 100 WHERE id = 1');
  db.Execute('UPDATE accounts SET balance = balance + 100 WHERE id = 2');
  db.Commit();
  console.log('转账成功');
} catch (e) {
  db.Rollback();
  console.log('转账失败，已回滚:', e.message);
}
```

**支持的数据库类型：**

| Type | 数据库 | 默认端口 |
|---|---|---|
| mysql | MySQL | 3306 |
| postgresql | PostgreSQL | 5432 |
| sqlserver | SQL Server | 1433 |

---

### 内置基础类型

AirScript 内置了多种数据类型，用于在脚本中处理多维表格的数据。

#### 日期时间类型

```javascript
// 创建日期
const date = new Date(2026, 3, 6); // 月份从0开始，3表示4月
const today = new Date();          // 当前日期时间

// 日期格式化
const dateStr = today.toLocaleDateString('zh-CN'); // 2026/4/6
const timeStr = today.toLocaleTimeString('zh-CN'); // 下午1:00:00

// 日期计算
const tomorrow = new Date(today.getTime() + 24 * 60 * 60 * 1000);
```

#### 附件类型

多维表格的附件字段返回附件对象数组：

```javascript
const record = sheet.Records.GetAll()[0];
const attachments = record.GetCellValue('附件字段');

attachments.forEach(att => {
  console.log(`文件名: ${att.Name}`);
  console.log(`文件大小: ${att.Size} bytes`);
  console.log(`文件URL: ${att.Url}`);
  console.log(`文件类型: ${att.MimeType}`);
});
```

#### 用户类型

人员字段返回用户对象：

```javascript
const record = sheet.Records.GetAll()[0];
const assignee = record.GetCellValue('负责人');

// 单个用户
if (assignee) {
  console.log(`用户名: ${assignee.Name}`);
  console.log(`用户ID: ${assignee.Id}`);
  console.log(`邮箱: ${assignee.Email}`);
}

// 多个用户（多人字段）
const members = record.GetCellValue('协作人');
members.forEach(user => {
  console.log(user.Name);
});
```

#### 选项类型（单选/多选）

```javascript
const record = sheet.Records.GetAll()[0];

// 单选字段
const status = record.GetCellValue('状态');
console.log(status); // 返回选项名称字符串，如 "进行中"

// 多选字段
const tags = record.GetCellValue('标签');
console.log(tags); // 返回字符串数组，如 ["技术", "重要"]
```

---

### 脚本经典案例

#### 案例1：数据同步 - 从外部 API 同步数据到表格

```javascript
// 从外部接口获取数据并同步到多维表格
function syncDataFromAPI() {
  const sheet = Application.Sheets.ActiveSheet;
  
  // 调用外部 API
  const response = Network.Fetch('https://api.example.com/products');
  const products = JSON.parse(response.Content);
  
  // 获取现有记录，建立索引（以产品ID为键）
  const existingRecords = sheet.Records.GetAll();
  const recordMap = {};
  existingRecords.forEach(record => {
    const id = record.GetCellValue('产品ID');
    if (id) recordMap[id] = record;
  });
  
  // 遍历外部数据，更新或新增
  products.forEach(product => {
    if (recordMap[product.id]) {
      // 更新现有记录
      const record = recordMap[product.id];
      record.SetCellValue('产品名称', product.name);
      record.SetCellValue('价格', product.price);
      record.SetCellValue('库存', product.stock);
    } else {
      // 新增记录
      sheet.Records.Add({
        '产品ID': product.id,
        '产品名称': product.name,
        '价格': product.price,
        '库存': product.stock
      });
    }
  });
  
  console.log(`同步完成，共处理 ${products.length} 条数据`);
  return { success: true, count: products.length };
}

syncDataFromAPI();
```

#### 案例2：定期报告 - 汇总数据并发送邮件

```javascript
function generateAndSendReport() {
  const sheet = Application.Sheets.ActiveSheet;
  const records = sheet.Records.GetAll();
  
  // 统计各状态数量
  const stats = { '待处理': 0, '进行中': 0, '已完成': 0, '已取消': 0 };
  let totalAmount = 0;
  
  records.forEach(record => {
    const status = record.GetCellValue('状态');
    const amount = record.GetCellValue('金额') || 0;
    if (stats.hasOwnProperty(status)) {
      stats[status]++;
    }
    totalAmount += amount;
  });
  
  // 构建邮件内容
  const today = new Date().toLocaleDateString('zh-CN');
  const htmlContent = `
    <h2>日报 - ${today}</h2>
    <table border="1" style="border-collapse:collapse;width:100%">
      <tr style="background:#f0f0f0">
        <th>状态</th><th>数量</th>
      </tr>
      ${Object.entries(stats).map(([k, v]) => `<tr><td>${k}</td><td>${v}</td></tr>`).join('')}
    </table>
    <p>合计金额：<strong>¥${totalAmount.toLocaleString()}</strong></p>
  `;
  
  // 发送邮件
  SMTP.SendMail({
    Host: 'smtp.company.com',
    Port: 465,
    SSL: true,
    Username: 'report@company.com',
    Password: 'your_password',
    From: 'report@company.com',
    To: ['manager@company.com'],
    Subject: `每日报告 ${today}`,
    Body: htmlContent,
    IsHtml: true
  });
  
  console.log('报告已发送');
}

generateAndSendReport();
```

#### 案例3：数据校验与清洗

```javascript
function validateAndCleanData() {
  const sheet = Application.Sheets.ActiveSheet;
  const records = sheet.Records.GetAll();
  
  let errorCount = 0;
  const errors = [];
  
  records.forEach((record, index) => {
    const name = record.GetCellValue('姓名');
    const phone = record.GetCellValue('手机号');
    const email = record.GetCellValue('邮箱');
    
    // 校验手机号
    if (phone && !/^1[3-9]\d{9}$/.test(phone)) {
      errors.push(`第${index + 1}行：手机号格式错误 - ${phone}`);
      record.SetCellValue('数据状态', '格式错误');
      errorCount++;
    }
    
    // 校验邮箱
    if (email && !/^[\w-]+@[\w-]+\.[a-z]{2,}$/.test(email)) {
      errors.push(`第${index + 1}行：邮箱格式错误 - ${email}`);
      record.SetCellValue('数据状态', '格式错误');
      errorCount++;
    }
    
    // 清理姓名（去除首尾空格）
    if (name && name !== name.trim()) {
      record.SetCellValue('姓名', name.trim());
    }
  });
  
  if (errors.length > 0) {
    console.log('发现以下错误：');
    errors.forEach(e => console.log(e));
  }
  
  console.log(`校验完成，共 ${records.length} 条记录，${errorCount} 条有误`);
  return { total: records.length, errors: errorCount };
}

validateAndCleanData();
```

---

## 内嵌 SDK 开发指南

### SDK 简介

WPS 多维表格提供内嵌 SDK，让开发者可以将多维表格**嵌入到自己的网页或系统中**。

提供两个版本：

| 版本 | 集成方式 | 推荐场景 |
|---|---|---|
| **WebOffice SDK 1.x** | `<iframe>` 嵌入 | 简单嵌入，快速集成，需要隔离 |
| **WebOffice SDK 3.x** | 微前端（JS注入） | 深度集成，性能更好，可访问内部 API |

---

### 快速入门（1.x / 3.x）

#### SDK 1.x 快速入门（iframe 模式）

```html
<!DOCTYPE html>
<html>
<head>
  <title>多维表格嵌入示例</title>
  <script src="https://qncdn.wpscdn.cn/weboffice/sdk/v1/index.js"></script>
</head>
<body>
  <div id="container" style="width:100%;height:800px;"></div>
  
  <script>
    // 初始化 SDK
    const instance = WebOfficeSDK.init({
      appId: 'your_app_id',       // 应用 ID
      token: 'your_access_token', // 访问令牌
      fileId: 'your_file_id',     // 文件 ID
      
      // 嵌入容器
      mount: document.getElementById('container'),
      
      // 挂载完成回调
      onSuccess: function(officeInstance) {
        console.log('多维表格加载成功');
        // officeInstance 即为 Application 对象
        const app = officeInstance.Application;
        const sheet = app.Sheets.ActiveSheet;
        console.log('当前表格名称:', sheet.Name);
      },
      
      // 错误回调
      onError: function(error) {
        console.error('加载失败:', error);
      }
    });
  </script>
</body>
</html>
```

#### SDK 3.x 快速入门（微前端模式）

```html
<!DOCTYPE html>
<html>
<head>
  <title>多维表格嵌入示例（3.x）</title>
</head>
<body>
  <div id="container" style="width:100%;height:800px;"></div>
  
  <script type="module">
    import WebOfficeSDK from 'https://qncdn.wpscdn.cn/weboffice/sdk/v3/index.esm.js';
    
    const instance = await WebOfficeSDK.init({
      appId: 'your_app_id',
      token: 'your_access_token',
      fileId: 'your_file_id',
      mount: document.getElementById('container'),
    });
    
    // 等待 Application 就绪
    await instance.ready();
    
    const app = instance.Application;
    const sheet = app.Sheets.ActiveSheet;
    console.log('当前表格:', sheet.Name);
    
    // 3.x 支持更丰富的 API
    const records = sheet.Records.GetAll();
    console.log(`共 ${records.length} 条记录`);
  </script>
</body>
</html>
```

---

### 配置参数

`WebOfficeSDK.init()` 支持以下配置参数：

| 参数 | 类型 | 必填 | 说明 |
|---|---|---|---|
| appId | string | 是 | 在 WPS 开放平台注册的应用 ID |
| token | string | 是 | 用户访问令牌（OAuth 2.0 Access Token） |
| fileId | string | 是 | 要嵌入的多维表格文件 ID |
| mount | HTMLElement | 是 | 挂载容器 DOM 节点 |
| lang | string | 否 | 界面语言，默认 `zh-CN`，支持 `en-US` |
| theme | string | 否 | 主题，支持 `light`（默认）、`dark` |
| readonly | boolean | 否 | 是否只读模式，默认 `false` |
| toolbar | boolean | 否 | 是否显示工具栏，默认 `true` |
| showSidePanel | boolean | 否 | 是否显示侧边面板，默认 `true` |
| onSuccess | function | 否 | 加载成功回调，参数为 officeInstance |
| onError | function | 否 | 加载失败回调，参数为 Error 对象 |
| onMessage | function | 否 | 消息事件回调（用于接收表格内发出的消息） |

**示例（带完整配置）：**

```javascript
const instance = WebOfficeSDK.init({
  appId: 'my_app_id',
  token: 'user_access_token',
  fileId: 'file_id_123',
  mount: document.getElementById('container'),
  lang: 'zh-CN',
  theme: 'light',
  readonly: false,
  toolbar: true,
  showSidePanel: false,
  
  onSuccess: function(app) {
    console.log('加载成功');
    // 可以在这里调用 API 进行初始化操作
    initMyApp(app);
  },
  
  onError: function(err) {
    console.error('加载失败', err.message);
    alert('文档加载失败，请刷新重试');
  }
});
```

---

### 事件监听

SDK 支持监听多维表格内部的各种事件，让外部页面可以响应表格的变化。

#### 监听记录变化

```javascript
// 使用 1.x SDK
const instance = WebOfficeSDK.init({ ... });
instance.onSuccess(function(app) {
  
  // 监听记录新增
  app.Sheets.ActiveSheet.Records.On('AddRecord', function(event) {
    console.log('新增了一条记录:', event.RecordId);
    // 刷新外部页面的数据统计
    updateStats();
  });
  
  // 监听记录修改
  app.Sheets.ActiveSheet.Records.On('UpdateRecord', function(event) {
    console.log('记录被修改:', event.RecordId, '修改字段:', event.FieldName);
  });
  
  // 监听记录删除
  app.Sheets.ActiveSheet.Records.On('DeleteRecord', function(event) {
    console.log('记录被删除:', event.RecordId);
  });
  
});
```

#### 监听视图切换

```javascript
app.Sheets.ActiveSheet.Views.On('ActivateView', function(event) {
  console.log('切换到视图:', event.ViewName);
});
```

#### 监听表格切换

```javascript
app.Sheets.On('ActivateSheet', function(event) {
  console.log('切换到表格:', event.SheetName);
});
```

#### 取消监听

```javascript
// 保存监听器引用
const handler = function(event) {
  console.log('事件:', event);
};

// 注册监听
app.Sheets.ActiveSheet.Records.On('AddRecord', handler);

// 取消监听
app.Sheets.ActiveSheet.Records.Off('AddRecord', handler);
```

---

## API 参考文档

### Application

`Application` 是多维表格 API 的根对象，通过它可以访问所有其他对象。

#### 访问方式

```javascript
// 在 AirScript 中
const app = Application;

// 在 SDK 中（1.x）
instance.onSuccess(function(app) {
  // app 即为 Application
});

// 在 SDK 中（3.x）
await instance.ready();
const app = instance.Application;
```

#### 属性

| 属性 | 类型 | 说明 |
|---|---|---|
| Sheets | Sheets | 所有表格对象的集合 |
| ActiveSheet | Sheet | 当前激活的表格 |
| Name | string | 文件名称 |
| FileId | string | 文件 ID |

#### 方法

```javascript
// 获取文件基本信息
const info = Application.GetFileInfo();
console.log(info.Name, info.FileId);

// 刷新文件（从服务器重新加载）
Application.Refresh();
```

---

### Sheet / Sheets

#### Sheets 集合

```javascript
const sheets = Application.Sheets;

// 获取所有表格
const allSheets = sheets.GetAll();
allSheets.forEach(sheet => {
  console.log(sheet.Name, sheet.Id);
});

// 按名称获取表格
const sheet = sheets.GetByName('Sheet1');

// 按索引获取表格（从1开始）
const firstSheet = sheets.Item(1);

// 获取当前激活的表格
const activeSheet = sheets.ActiveSheet;

// 新建表格
const newSheet = sheets.Add({ Name: '新表格' });

// 删除表格（按名称或ID）
sheets.Delete('Sheet2');

// 监听表格切换事件
sheets.On('ActivateSheet', function(event) {
  console.log('切换到:', event.SheetName);
});
```

#### Sheet 对象

```javascript
const sheet = Application.Sheets.ActiveSheet;

// 属性
console.log(sheet.Name);     // 表格名称
console.log(sheet.Id);       // 表格 ID
console.log(sheet.Index);    // 在文件中的位置索引（从1开始）

// 重命名
sheet.Name = '新名称';

// 访问子对象
const views = sheet.Views;    // Views 集合
const records = sheet.Records; // Records 集合
const fields = sheet.Fields;   // Fields 集合

// 获取字段描述（字段元信息）
const fieldDescriptors = sheet.FieldDescriptors.GetAll();

// 激活此表格（切换到该表格）
sheet.Activate();

// 复制表格
const copiedSheet = sheet.Copy({ Name: '副本' });

// 删除此表格
sheet.Delete();
```

---

### View / Views

#### Views 集合

```javascript
const sheet = Application.Sheets.ActiveSheet;
const views = sheet.Views;

// 获取所有视图
const allViews = views.GetAll();
allViews.forEach(view => {
  console.log(`视图: ${view.Name}, 类型: ${view.Type}`);
});

// 获取当前激活视图
const activeView = views.ActiveView;

// 按名称获取视图
const gridView = views.GetByName('表格视图');

// 新建视图
const newView = views.Add({
  Name: '我的甘特图',
  Type: 'gantt'  // 视图类型：grid/gallery/gantt/calendar/kanban/form
});

// 删除视图
views.Delete('旧视图');
```

#### View 对象

```javascript
const view = Application.Sheets.ActiveSheet.Views.ActiveView;

// 属性
console.log(view.Name);  // 视图名称
console.log(view.Id);    // 视图 ID
console.log(view.Type);  // 视图类型（grid/gallery/gantt/calendar/kanban/form）

// 激活此视图
view.Activate();

// 设置筛选条件
view.Filter.Set([
  {
    FieldName: '状态',
    Operator: 'is',
    Value: '进行中'
  }
]);

// 清除筛选
view.Filter.Clear();

// 设置排序
view.Sort.Set([
  {
    FieldName: '创建时间',
    Order: 'desc'  // asc / desc
  }
]);

// 设置分组
view.Group.Set([
  {
    FieldName: '部门',
    Order: 'asc'
  }
]);
```

**视图类型（Type）说明：**

| Type | 名称 |
|---|---|
| grid | 表格视图（默认） |
| gallery | 画册视图 |
| gantt | 甘特图视图 |
| calendar | 日历视图 |
| kanban | 看板视图 |
| form | 表单视图 |

---

### Record / Records

#### Records 集合

```javascript
const sheet = Application.Sheets.ActiveSheet;
const records = sheet.Records;

// 获取所有记录
const allRecords = records.GetAll();
console.log(`共 ${allRecords.length} 条记录`);

// 按 ID 获取记录
const record = records.GetById('record_id_here');

// 添加单条记录
const newRecord = records.Add({
  '姓名': '李四',
  '年龄': 30,
  '部门': '产品部'
});
console.log('新记录ID:', newRecord.Id);

// 批量添加记录（性能更好）
const batchResult = records.BatchAdd([
  { '姓名': '王五', '年龄': 25 },
  { '姓名': '赵六', '年龄': 32 },
  { '姓名': '钱七', '年龄': 28 }
]);
console.log('批量添加成功:', batchResult.length, '条');

// 删除记录（按 ID）
records.Delete('record_id_here');

// 批量删除
records.BatchDelete(['id1', 'id2', 'id3']);

// 查询记录（按条件筛选）
const filteredRecords = records.Find({
  Conditions: [
    { FieldName: '状态', Operator: 'is', Value: '已完成' },
    { FieldName: '优先级', Operator: 'is', Value: '高' }
  ],
  Logic: 'AND'  // AND / OR
});

// 事件监听
records.On('AddRecord', function(e) { console.log('新增:', e.RecordId); });
records.On('UpdateRecord', function(e) { console.log('更新:', e.RecordId); });
records.On('DeleteRecord', function(e) { console.log('删除:', e.RecordId); });
```

#### Record 对象

```javascript
const record = Application.Sheets.ActiveSheet.Records.GetAll()[0];

// 属性
console.log(record.Id);      // 记录 ID（唯一标识）
console.log(record.Index);   // 在表格中的行序号（从1开始）

// 获取字段值
const name = record.GetCellValue('姓名');
const age = record.GetCellValue('年龄');
const status = record.GetCellValue('状态');

// 设置字段值
record.SetCellValue('姓名', '新名字');
record.SetCellValue('年龄', 35);
record.SetCellValue('状态', '已完成');

// 批量设置多个字段值（性能更好，减少网络请求）
record.SetCellValues({
  '姓名': '批量更新',
  '年龄': 40,
  '备注': '批量修改'
});

// 获取记录的所有字段值（返回 {字段名: 值} 对象）
const allValues = record.GetAllCellValues();
console.log(allValues);

// 删除此记录
record.Delete();
```

---

### Field / Fields

#### Fields 集合

```javascript
const sheet = Application.Sheets.ActiveSheet;
const fields = sheet.Fields;

// 获取所有字段
const allFields = fields.GetAll();
allFields.forEach(field => {
  console.log(`字段: ${field.Name}, 类型: ${field.Type}`);
});

// 按名称获取字段
const nameField = fields.GetByName('姓名');

// 按 ID 获取字段
const field = fields.GetById('field_id_here');

// 新建字段
const newField = fields.Add({
  Name: '备注',
  Type: 'text'  // 字段类型
});

// 删除字段
fields.Delete('备注');
```

**字段类型（Type）说明：**

| Type | 字段类型 | 说明 |
|---|---|---|
| text | 文本 | 单行/多行文本 |
| number | 数字 | 整数或小数 |
| select | 单选 | 从预设选项中选一个 |
| multi_select | 多选 | 从预设选项中选多个 |
| date | 日期 | 日期/日期时间 |
| checkbox | 复选框 | true/false |
| person | 人员 | 选择用户 |
| attachment | 附件 | 上传文件 |
| url | 超链接 | URL 链接 |
| email | 邮件 | 邮件地址 |
| phone | 电话 | 电话号码 |
| formula | 公式 | 基于其他字段计算 |
| auto_number | 自动编号 | 系统自动生成序号 |
| created_time | 创建时间 | 记录创建时间（只读） |
| modified_time | 修改时间 | 记录修改时间（只读） |
| created_by | 创建人 | 记录创建者（只读） |
| modified_by | 修改人 | 最后修改者（只读） |
| lookup | 引用 | 引用其他表格的字段值 |
| rollup | 汇总 | 汇总关联表格的数据 |
| link | 关联 | 关联其他表格的记录 |
| rating | 评分 | 星级评分 |
| progress | 进度 | 百分比进度条 |
| currency | 货币 | 带货币符号的数字 |
| duration | 时长 | 时间长度 |
| location | 地理位置 | 经纬度或地址 |

#### Field 对象

```javascript
const field = Application.Sheets.ActiveSheet.Fields.GetByName('状态');

// 属性
console.log(field.Name);     // 字段名称
console.log(field.Id);       // 字段 ID
console.log(field.Type);     // 字段类型
console.log(field.Index);    // 字段列顺序（从1开始）
console.log(field.Required); // 是否必填

// 重命名字段
field.Name = '任务状态';

// 获取单选/多选字段的选项列表
if (field.Type === 'select' || field.Type === 'multi_select') {
  const options = field.Options;
  options.forEach(opt => {
    console.log(`选项: ${opt.Name}, 颜色: ${opt.Color}`);
  });
  
  // 新增选项
  field.Options.Add({ Name: '新选项', Color: '#ff6b6b' });
  
  // 删除选项
  field.Options.Delete('旧选项');
}

// 删除字段
field.Delete();
```

---

### FieldDescriptor

`FieldDescriptor` 是字段的**完整元数据描述**对象，包含字段类型的所有详细配置信息。

与 `Field` 对象的区别：
- `Field`：用于读写字段基本信息（名称、类型）和操作
- `FieldDescriptor`：提供字段的完整配置描述，通常用于分析字段结构

```javascript
const sheet = Application.Sheets.ActiveSheet;
const fieldDescriptors = sheet.FieldDescriptors.GetAll();

fieldDescriptors.forEach(desc => {
  console.log('字段名:', desc.Name);
  console.log('字段ID:', desc.Id);
  console.log('字段类型:', desc.Type);
  console.log('是否必填:', desc.Required);
  console.log('是否只读:', desc.Readonly);
  
  // 对于选项类字段，查看选项配置
  if (desc.Type === 'select' || desc.Type === 'multi_select') {
    console.log('选项列表:', desc.Options);
  }
  
  // 对于关联字段，查看关联配置
  if (desc.Type === 'link') {
    console.log('关联表格ID:', desc.LinkedSheetId);
    console.log('是否双向关联:', desc.IsBidirectional);
  }
  
  // 对于公式字段，查看公式内容
  if (desc.Type === 'formula') {
    console.log('公式内容:', desc.Formula);
    console.log('结果类型:', desc.ResultType);
  }
  
  // 对于数字字段，查看格式配置
  if (desc.Type === 'number') {
    console.log('小数位数:', desc.DecimalPlaces);
    console.log('千分位分隔符:', desc.UseThousandSeparator);
  }
});
```

**FieldDescriptor 常用属性：**

| 属性 | 类型 | 说明 |
|---|---|---|
| Id | string | 字段唯一 ID |
| Name | string | 字段名称 |
| Type | string | 字段类型 |
| Required | boolean | 是否必填 |
| Readonly | boolean | 是否只读 |
| Description | string | 字段描述/备注 |
| DefaultValue | any | 字段默认值 |
| Options | array | 选项列表（select/multi_select 类型） |
| Formula | string | 公式内容（formula 类型） |
| LinkedSheetId | string | 关联表格 ID（link 类型） |
| DecimalPlaces | number | 小数位数（number/currency 类型） |

---

### RecordRange

`RecordRange` 用于**选择和操作一批记录**，支持多种选择器语法，是批量操作记录的高效方式。

#### 创建 RecordRange

```javascript
const sheet = Application.Sheets.ActiveSheet;
const records = sheet.Records;

// 方法1：按索引范围（行号，从1开始）
const range1 = records.Range('1:10');    // 第1到第10行
const range2 = records.Range('5:');     // 第5行到最后
const range3 = records.Range(':5');     // 第1行到第5行

// 方法2：按记录 ID
const range4 = records.Range('id:rec001');
const range5 = records.Range('id:rec001,rec002,rec003');

// 方法3：按条件筛选
const range6 = records.Range({
  Conditions: [
    { FieldName: '状态', Operator: 'is', Value: '待处理' }
  ]
});

// 方法4：所有记录
const allRange = records.Range('*');

// 方法5：通过数组指定多行
const range7 = records.Range([1, 3, 5, 7]); // 第1、3、5、7行
```

#### 条件筛选操作符

| Operator | 说明 | 示例值 |
|---|---|---|
| is | 等于 | `'已完成'` |
| is_not | 不等于 | `'已取消'` |
| contains | 包含 | `'关键词'` |
| not_contains | 不包含 | `'关键词'` |
| is_empty | 为空 | `null` |
| is_not_empty | 不为空 | `null` |
| gt | 大于 | `100` |
| gte | 大于等于 | `100` |
| lt | 小于 | `100` |
| lte | 小于等于 | `100` |
| between | 在范围内 | `[10, 100]` |

#### 操作 RecordRange

```javascript
const sheet = Application.Sheets.ActiveSheet;
const records = sheet.Records;

// 获取满足条件的记录范围
const pendingRange = records.Range({
  Conditions: [
    { FieldName: '状态', Operator: 'is', Value: '待处理' },
    { FieldName: '优先级', Operator: 'is', Value: '高' }
  ],
  Logic: 'AND'
});

// 获取范围内的所有记录
const pendingRecords = pendingRange.GetAll();
console.log(`共 ${pendingRecords.length} 条待处理高优先级记录`);

// 批量更新范围内记录的某个字段
pendingRange.SetCellValue('负责人', '张三');

// 批量更新多个字段
pendingRange.SetCellValues({
  '处理状态': '已分配',
  '分配时间': new Date().toISOString()
});

// 删除范围内的所有记录
// pendingRange.Delete();  // 谨慎操作！

// 遍历范围内的记录
pendingRange.ForEach(function(record) {
  const name = record.GetCellValue('任务名称');
  const priority = record.GetCellValue('优先级');
  console.log(`任务: ${name}, 优先级: ${priority}`);
});

// 统计范围内记录数
const count = pendingRange.Count();
console.log('记录数:', count);
```

#### 组合使用示例

```javascript
// 综合示例：批量处理指定条件的记录
function batchProcessRecords() {
  const sheet = Application.Sheets.ActiveSheet;
  
  // 选择30天前创建的、状态为"待处理"的记录
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  
  const expiredRange = sheet.Records.Range({
    Conditions: [
      { FieldName: '状态', Operator: 'is', Value: '待处理' },
      { 
        FieldName: '创建时间', 
        Operator: 'lt', 
        Value: thirtyDaysAgo.toISOString() 
      }
    ],
    Logic: 'AND'
  });
  
  const count = expiredRange.Count();
  
  if (count === 0) {
    console.log('没有需要处理的过期记录');
    return;
  }
  
  console.log(`发现 ${count} 条过期待处理记录，准备批量更新...`);
  
  // 批量标记为"已过期"
  expiredRange.SetCellValues({
    '状态': '已过期',
    '备注': `系统于 ${new Date().toLocaleDateString('zh-CN')} 自动标记为过期`
  });
  
  console.log(`已将 ${count} 条记录标记为过期`);
  return count;
}

batchProcessRecords();
```

---

## 附录：常见问题与最佳实践

### Q: AirScript 中可以使用 npm 包吗？

不支持直接 `import` npm 包。但 AirScript 内置了常用功能（网络请求、数据库等），可以通过高级服务 API 满足大部分需求。

### Q: 如何处理大量数据的性能问题？

1. 优先使用批量操作方法（`BatchAdd`、`BatchDelete`、`SetCellValues`）
2. 使用 `RecordRange` 而不是逐条遍历修改
3. 避免在循环中逐条调用 API，尽量先收集数据再批量提交

```javascript
// 不推荐：逐条更新（慢）
records.forEach(record => {
  record.SetCellValue('状态', '已处理'); // 每次都发请求
});

// 推荐：使用 RecordRange 批量更新（快）
sheet.Records.Range('*').SetCellValue('状态', '已处理');
```

### Q: SDK 嵌入时如何处理用户权限？

SDK 的 `token` 参数对应用户的 OAuth 访问令牌，多维表格会根据该用户的实际权限来决定其在嵌入界面中可执行的操作（查看/编辑/管理）。确保向 SDK 传递正确的用户令牌，而非管理员令牌。

### Q: 如何在嵌入页面中监听表格数据变化并更新外部 UI？

```javascript
// 在 SDK 初始化成功后注册监听器
instance.onSuccess(function(app) {
  const sheet = app.Sheets.ActiveSheet;
  
  // 监听记录变化，更新外部统计面板
  sheet.Records.On('AddRecord', function() {
    refreshExternalDashboard();
  });
  
  sheet.Records.On('UpdateRecord', function(e) {
    if (e.FieldName === '状态') {
      refreshStatusSummary();
    }
  });
});
```

### Q: 脚本令牌和 OAuth Token 有什么区别？

| | 脚本令牌（Script Token） | OAuth Access Token |
|---|---|---|
| 用途 | 触发特定脚本执行 | 代表用户身份访问文件 |
| 授权范围 | 仅执行指定脚本 | 由 OAuth 授权范围决定 |
| 有效期 | 永久（可手动撤销） | 通常有效期较短（需刷新） |
| 使用场景 | 外部系统 Webhook、定时任务 | SDK 嵌入、API 调用 |

---

*文档整理完毕。如需最新内容，请访问 [WPS 开放平台](https://365.kdocs.cn/3rd/open/documents/app-integration-dev/guide/dbsheet/dbsheet-standard) 查阅官方文档。*
