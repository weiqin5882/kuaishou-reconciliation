# 快手订单对账与利润统计系统

一个基于 Flask 的网页版订单比对工具，用于对比快手官方订单表与客服手动统计表，自动发现漏单、重复单、错误单，并计算销售额、成本与利润。

## 功能特性

- ✅ 网页版界面，浏览器访问
- ✅ 支持上传两个 Excel 文件（官方表 + 客服表）
- ✅ 智能列映射，自动识别 + 手动调整
- ✅ 数据预处理：去空格、去特殊字符、统一格式、自动去重
- ✅ 订单状态过滤（仅保留"交易成功"和"已发货"）
- ✅ 订单匹配对比，以订单号为唯一键
- ✅ 漏单检测（官方有客服无 / 客服有官方无）
- ✅ 自动计算利润 = 销售额 - 成本
- ✅ 亏损订单红色标记
- ✅ 底部汇总统计
- ✅ 导出标准 Excel 报表
- ✅ 支持 10 万条数据量
- ✅ 完善的容错机制

## 技术栈

- 后端：Python 3.8+ / Flask
- 数据处理：pandas, openpyxl
- 前端：HTML5 + JavaScript (原生)
- Excel 导出：openpyxl

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 启动服务

```bash
python app.py
```

### 3. 访问系统

打开浏览器访问：`http://localhost:5000`

## 使用流程

1. **上传官方订单表**：选择从快手后台导出的 Excel 文件
2. **上传客服统计表**：选择客服手动整理的 Excel 文件
3. **配置列映射**：确认系统自动识别的列名，或手动调整
4. **点击开始对比**：系统自动处理并生成结果
5. **查看结果**：预览对比结果，亏损订单标红
6. **导出报表**：下载整理后的标准 Excel 报表

## 数据格式说明

### 官方订单表
- 列名相对固定
- 必须包含：订单号、订单状态、商品名称、销售金额
- 可选：成本（如没有需在客服表中补充）

### 客服手动统计表
- 列名可能不固定
- 需要手动指定：订单号、商品名称、销售额、成本 对应的列
- 支持容错：空格、特殊字符、重复订单号等

## 部署说明

### 使用 Gunicorn 部署（生产环境）

```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

### Docker 部署

```bash
docker build -t kuaishou-reconciliation .
docker run -p 5000:5000 kuaishou-reconciliation
```

## 文件说明

```
kuaishou-reconciliation/
├── app.py                 # Flask 主应用
├── requirements.txt       # Python 依赖
├── templates/
│   └── index.html        # 主页面
├── static/
│   ├── css/
│   │   └── style.css     # 样式文件
│   └── js/
│       └── main.js       # 前端交互逻辑
├── uploads/              # 上传文件临时存储
└── exports/              # 导出文件存储
```

## 许可证

MIT License
