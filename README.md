# 偷偷开源的排版工具

一键格式化毕业论文。双击 exe，选文件，点"开始格式化"，搞定。

默认配置为 **华南农业大学 2024 本科毕业论文** 格式规范，支持通过 YAML 配置文件适配其他学校。

## 功能

- GUI 界面，所有参数可视化调整
- 页面设置（页边距、装订线、页眉页脚）
- 字体字号（正文、各级标题、图表题、脚注）
- 自动生成封面和声明页（也支持上传自定义封面）
- 特殊标题格式化（摘要、目录、参考文献、致谢）
- 标题自动编号修正
- 题序与标题间距规范化
- 图表题注格式化 + 编号连续性检查
- 参考文献悬挂缩进
- 三线表格式
- 目录生成与字体修正（需 Word COM）
- 页码设置（前置罗马数字 + 正文阿拉伯数字）
- 引用逗号间距自动修正
- 配置保存/加载，方便分享给同校同学

## 下载使用

### 方式一：下载 exe（推荐）

从 [Releases](../../releases) 下载 `thesis-format.exe`，双击运行。

### 方式二：从源码运行

```bash
pip install -r requirements.txt
python thesis_format_cli.py
```

## 配置

默认使用华南农业大学 2024 规范。如需修改：

1. 点击 GUI 中的「保存配置」导出 `thesis_config.yaml`
2. 修改需要调整的参数
3. 点击「加载配置」导入修改后的文件

也可以命令行指定：

```bash
thesis-format.exe --input 论文.docx --config 我的学校.yaml
```

配置文件示例见 `defaults/scau_2024.yaml`。

## 打包 exe

```bash
build_exe.bat
```

需要安装 PyInstaller：`pip install pyinstaller`

## 文件说明

| 文件 | 说明 |
|------|------|
| `thesis_format_cli.py` | GUI 入口 |
| `thesis_format_2024.py` | 格式化核心引擎 |
| `thesis_config.py` | 配置加载器 + 内置默认值 |
| `word_postprocess.py` | Word COM 后处理（更新目录） |
| `preprocess_txt_to_md.py` | txt 预处理转 Markdown |
| `defaults/scau_2024.yaml` | 默认配置文件 |
| `defaults/scau_logo.png` | 学校 Logo |

## License

GPL-3.0
