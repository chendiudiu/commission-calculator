# 🍺 提成计算器

门店点餐订单报表自动统计分析工具

## 功能特点

- 📂 支持批量上传多个CSV文件
- 🏪 自动从文件名提取门店名称
- 📊 12项提成统计项目
- 🎨 现代精酿酒吧风格界面
- 📥 Excel报表导出

## 统计项目

| 项目 | 说明 |
|------|------|
| 桶装精酿 | 招牌原浆精酿，排除福袋 |
| 瓦猫猫听装精酿 | 招牌精酿瓦猫猫的酒，除以12 |
| 鸡尾酒套餐 | 特调鸡尾酒单品，排除12杯和半价 |
| 小吃套餐(59元) | 排除49元二销 |
| 小吃套餐(79元) | 小吃B套餐 |
| 小吃套餐(99元) | 小吃C套餐 |
| 1升装精酿双拼套餐 | 套餐下单品1L，除以2 |
| 瓦猫猫二销套餐 | 所属套餐含二销，除以24 |
| 小吃二销套餐 | 59元小吃套餐实付49元 |
| 奔富 | 商品名称含奔富 |
| 点歌 | 商品名称含点歌 |

## macOS 版本

直接运行:
```bash
./dist/提成计算器
```

## Windows 版本构建

### 1. 安装 Python
下载并安装 Python 3.8+: https://www.python.org/downloads/

### 2. 安装依赖
```bash
pip install pandas openpyxl pyinstaller
```

### 3. 构建
```bash
pyinstaller --onefile --windowed --name 提成计算器 commission_calculator.py
```

## 目录结构

```
commission_calculator/
├── commission_calculator.py   # 主程序
├── 提成计算器_windows.spec   # Windows 构建配置
└── dist/                    # 构建输出
    └── 提成计算器           # macOS 可执行文件
```
