# 运输数据分析看板

这是一个基于Streamlit的交互式数据看板，用于分析SYD、ADL和BNE三个城市的运输数据。

## 功能特点

看板提供以下分析指标：

1. 运单总量走势
2. 总收入走势
3. 活跃派送员数量走势
4. 包裹平均重量走势
5. 收件城市比例
6. 商家比例

数据可以按两种维度查看：
- 按日维度
- 按周维度

## 安装步骤

1. 确保已安装Python 3.8或更高版本
2. 克隆本仓库到本地

## 使用方法

### 方法一：使用脚本文件（推荐）

#### Windows用户：
直接双击运行 `run_dashboard.bat` 文件，它会自动安装依赖并启动应用。

#### Linux/macOS用户：
```bash
chmod +x run_dashboard.sh  # 添加执行权限
./run_dashboard.sh         # 运行脚本
```

### 方法二：手动安装和运行

1. 安装所需依赖：

```bash
pip install -r requirements.txt
```

2. 运行Streamlit应用：

```bash
python -m streamlit run dashboard.py
```

3. 在浏览器中打开应用（通常会自动打开）
4. 选择数据源：
   - 上传本地Excel文件
   - 或从Supabase云存储获取数据
5. 使用侧边栏选择城市
6. 使用选项卡切换日维度和周维度视图
7. 根据需要调整日期范围或周数

## 配置Supabase云存储

为了使用Supabase云存储功能，您需要配置凭证。有两种方式：

### 1. 使用环境变量

1. 复制`.env.example`为`.env`
2. 编辑文件并填入您的Supabase凭证
3. 或直接在系统中设置以下环境变量：
   - `SUPABASE_URL`
   - `SUPABASE_KEY`
   - `SUPABASE_BUCKET`

### 2. 使用Streamlit密钥

如果您在Streamlit Cloud上部署，可以使用其密钥管理功能：

1. 在Streamlit应用设置中添加以下密钥：
   - `SUPABASE_URL`
   - `SUPABASE_KEY`
   - `SUPABASE_BUCKET`

## 数据文件要求

上传的Excel文件应包含以下工作表：
- SYD
- ADL
- BNE

每个工作表应包含以下列：
- 签收时间
- 运单号
- 收派员
- 计费
- 商家上传重量
- 收件城市 