# Hair Analysis Service

Spring Boot 服务用于计算美容门店美发来源及地拓来源的数据指标，并写入 `analysis_from` 表，同时支持 Excel 导出。

## 主要功能

- 连接两个 MySQL 数据源：
  - `mr_xgly_bb`：写入分析结果以及读取门店基础信息。
  - `boka_data_center`：读取项目消费、卖卡、充值等明细数据。
- 调用博卡接口查询会员卡列表和消费记录（内置签名逻辑）。
- 计算美发体验人数/业绩、成交人数/业绩（根据 49 元体验价规则、首开首充判定等）并写入 `analysis_from`。
- 通过 REST API 返回分析结果或导出 Excel 文件。

## 构建与运行

```bash
cd hair-analysis-service
mvn spring-boot:run
```

默认端口为 `8080`，可在 `application.yml` 中调整。

> **注意：** `application.yml` 中的数据库及接口凭证支持通过环境变量覆盖（如 `MR_DB_URL`、`BOKA_APP_KEY` 等）。生产环境请务必配置真实的 `BOKA_APP_KEY`。

## 接口说明

### 1. 计算分析并写库

- **URL**: `POST /api/analysis`
- **请求体**:
  ```json
  {
    "storeId": "046",
    "month": "2025-05"
  }
  ```
- **响应**: `AnalysisResult` JSON。

### 2. 计算并导出 Excel

- **URL**: `POST /api/analysis/export`
- **请求体**: 同上。
- **响应**: `application/octet-stream`，附件名称形如 `analysis-046.xlsx`。

## 代码结构

- `config/`：多数据源及 MyBatis 配置。
- `client/`：博卡接口客户端及 DTO。
- `mapper/`：MyBatis Mapper 接口及 XML。
- `model/`：实体与 DTO。
- `service/`：核心业务计算、Excel 导出。
- `controller/`：REST 接口。

## 后续扩展

- 根据业务口径完善地拓来源统计逻辑。
- 引入缓存/批处理以减少接口调用次数。
- 增加单元测试覆盖复杂计算场景。
