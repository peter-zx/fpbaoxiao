# 报销费用填写工具

一个简单易用的报销费用填写工具，支持图片嵌入Excel表格，方便财务部门审核。

## 功能特点

- 🖼️ **图片嵌入Excel** - 真正将图片嵌入到Excel单元格中，非外部链接
- 📊 **自动计算合计** - 自动汇总报销金额
- 💾 **数据本地保存** - 数据保存在本地，保护隐私
- 📤 **一键导出** - 导出带图片的Excel文件，可直接提交财务
- 🎯 **开箱即用** - 无需配置，双击启动即可使用

## 技术栈

| 技术 | 用途 |
|------|------|
| Python 3 | 后端服务器 |
| openpyxl | Excel读写 + 图片嵌入 |
| PIL (Pillow) | 图片处理 |
| HTML/CSS/JS | 前端界面 |

## 快速开始

### 1. 安装依赖

```bash
pip install openpyxl pillow
```

### 2. 启动服务

双击 `start.bat` 或命令行运行：

```bash
python server.py
```

### 3. 打开浏览器

访问 http://localhost:8765

## 使用说明

1. 点击「添加报销」按钮
2. 填写报销信息（日期、项目、金额、说明）
3. 上传相关图片（发票、凭证等）
4. 点击「保存」
5. 所有记录填写完成后，点击「导出Excel（含图片）」
6. 打开导出的Excel文件，图片已嵌入在对应单元格中

## 目录结构

```
fpbaoxiao/
├── server.py              # Python后端服务器
├── index.html             # 前端界面
├── start.bat              # Windows启动脚本
├── 2026年启用报销与费用填写.xlsx  # 模板文件（可选）
├── .gitignore             # Git忽略配置
├── README.md              # 项目说明
└── requirements.txt       # Python依赖
```

## 核心技术说明

### 图片嵌入Excel的实现原理

这是本项目最核心的技术点，解决了纯前端无法将图片嵌入Excel的问题。

#### 技术路径

```
用户上传图片 → Base64解码 → PIL处理 → BytesIO内存流 → openpyxl.Image → 锚定单元格
```

#### 关键代码

```python
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
import io
import base64

# 1. 解码Base64图片数据
image_data = base64.b64decode(base64_string)

# 2. 用PIL打开并处理
pil_img = Image.open(io.BytesIO(image_data))

# 3. 缩放图片（可选）
pil_img = pil_img.resize((150, 150))

# 4. 转回内存流（关键：不写临时文件）
buf = io.BytesIO()
pil_img.save(buf, format='PNG')
buf.seek(0)  # 重置指针

# 5. 创建Excel图片对象并锚定到单元格
xl_img = XLImage(buf)
xl_img.anchor = 'E3'  # 锚定到E3单元格

# 6. 添加到工作表
ws.add_image(xl_img)
```

#### 为什么用内存流？

| 方案 | 问题 |
|------|------|
| 写临时文件 | 文件名扩展名问题、路径转义、需清理 |
| **内存流 (BytesIO)** | 无临时文件、无路径问题、更快更可靠 |

### 为什么纯前端做不到？

- SheetJS (xlsx.js) 主要处理单元格数据，不擅长图片嵌入
- 浏览器安全限制，无法直接操作文件系统
- 需要后端配合处理图片的二进制数据

## 常见问题

**Q: 图片显示不正常？**
A: 确保安装了 Pillow 库：`pip install pillow`

**Q: 导出的Excel打不开？**
A: 检查是否有足够的磁盘空间，以及文件是否被其他程序占用

**Q: 如何修改默认端口？**
A: 编辑 `server.py`，修改 `PORT = 8765` 为其他端口

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！
