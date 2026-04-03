# 提成计算器 - 构建说明

## macOS 版本
已构建完成:
- 位置: `dist/提成计算器`
- 格式: macOS 可执行文件 (arm64)

直接运行:
```bash
./dist/提成计算器
```

## Windows 版本构建

在 Windows 电脑上执行以下步骤:

### 1. 安装 Python
下载并安装 Python 3.8+: https://www.python.org/downloads/

### 2. 安装依赖
打开命令提示符 (CMD) 或 PowerShell, 运行:
```bash
pip install pandas openpyxl pyinstaller
```

### 3. 构建 Windows 版本
```bash
cd 提成计算器
pyinstaller --onefile --windowed --name 提成计算器 commission_calculator.py
```

或使用 spec 文件:
```bash
pyinstaller 提成计算器_windows.spec
```

### 4. 找到可执行文件
构建完成后, 可执行文件在:
- 单文件模式: `dist/提成计算器/提成计算器.exe`
- 目录模式: `dist/提成计算器/`

## 目录结构
```
commission_calculator/
├── commission_calculator.py   # 主程序
├── 提成计算器_windows.spec   # Windows 构建配置
├── README.md                 # 本说明文件
└── dist/                    # 构建输出
    ├── 提成计算器           # macOS 可执行文件
    └── 提成计算器.app       # macOS 应用
```

## 使用说明
1. 启动程序
2. 点击"选择CSV文件"按钮
3. 选择一个或多个点餐订单报表CSV文件
4. 点击"计算提成"
5. 查看结果
6. 点击"导出Excel"保存报表
