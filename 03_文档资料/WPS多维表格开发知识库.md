# WPS 多维表格开发知识库

> 更新时间：2026-04-02

## 一、官方文档入口

- 主站：https://365.kdocs.cn/3rd/open/documents/app-integration-dev/guide/dbsheet/dbsheet-standard
- 开放平台：https://365.kdocs.cn/

## 二、三大开发方式

| 方式 | 说明 |
|------|------|
| **AirScript** | 在线脚本，服务端运行 JavaScript |
| **SDK** | 内嵌使用（iframe/微前端） |
| **插件** | 自定义仪表盘、视图、记录卡片 |

## 三、AirScript 文档

- 简介：AirScript-instro
- 快速入门：AirScript-quickstart
- 脚本令牌：AirScript-apitoken-instro
- 内置基础类型：AirScript-build-in
- 经典案例：AirScript-demo

## 四、内嵌 SDK 文档

- 简介：weboffice-instro
- 快速入门：weboffice-quickstart

## 五、API 文档体系

| 分类 | 说明 |
|------|------|
| 数据表 | 创建/删除/管理数据表 |
| 视图 | 视图操作 |
| 记录 | 增删改查记录 |
| 字段 | 字段管理 |
| 排序/筛选/分组 | 数据处理 |
| 评论 | 评论功能 |

## 六、核心字段类型（17种）

文本 | 多行文本 | 日期 | 单选 | 多选 | 数字 | 评分 | 公式 | 级联 | 单向关联 | 货币 | 百分比 | 身份证 | 电话 | 邮箱 | 时间 | 附件

## 七、踩坑经验

- Offset 参数必须是字符串类型：let offset = '0' 而非 let offset = 0
- 表名匹配用固定 ID，不用模糊查找（sheets.find() 不靠谱）
- sheets.find() 模糊匹配容易匹配到错误的表，建议直接用固定 ID
