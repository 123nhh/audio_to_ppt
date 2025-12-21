# 🎵 音乐转PPT工具套装

一套强大的音乐可视化工具，可将音频文件自动转换为精美的PPT演示文稿，并支持批量合并与音频同步。

## ✨ 功能特性

### 🎨 核心功能
- ✅ **智能歌词提取**：自动从音频文件提取歌词并智能清洗
- ✅ **AI歌词优化**：利用GPT模型优化歌词排版，智能分行
- ✅ **高质量封面**：提取专辑封面并应用高斯模糊背景
- ✅ **渐变遮罩**：优雅的顶部/底部渐变遮罩效果
- ✅ **纯音乐识别**：自动识别纯音乐并生成封面页
- ✅ **批量处理**：支持多文件并发处理
- ✅ **PPT合并**：智能合并多个PPT并同步音频文件

### 🎯 支持格式
- **音频格式**：FLAC、MP3、WAV、M4A
- **输出格式**：PPTX（PowerPoint 2007+）

---

## 🛠️ 工具说明

### 1️⃣ music_to_ppt（音乐转PPT）

**功能**：将音频文件批量转换为PPT演示文稿

**输入**：`music/` 文件夹中的音频文件  
**输出**：`output/` 文件夹中的PPT文件

**主要特点**：
- 自动提取歌词、标题、歌手、封面
- AI智能清洗歌词（去除翻译、宣传信息等）
- 根据歌词长度自动调整字体大小
- 支持纯音乐模式（仅生成封面页）
- 多线程并发处理（默认4线程）

### 2️⃣ hebing（PPT合并工具）

**功能**：合并多个PPT文件并同步对应音频

**输入**：`output/` 文件夹中的PPT + `music/` 文件夹中的音频  
**输出**：`ppt_output/` 文件夹中的合并文件

**主要特点**：
- 交互式选择PPT顺序
- 自动匹配并复制同名音频文件
- 基于PowerPoint COM接口，保留原始格式

---

## 💻 环境要求

### Python版本（仅源码运行）
- Python 3.7 或更高版本

### 系统要求
- **Windows 7/8/10/11**（推荐Windows 10+）
- **Microsoft PowerPoint**（hebing工具必需）

### Python依赖库（仅源码运行）
```bash
pip install pillow mutagen python-pptx openai pywin32
```

> **注意**：如果使用 `.exe` 可执行文件，**无需安装Python和依赖库**。

---

## 🚀 快速开始

### 方式A：使用可执行文件（推荐）

1. **准备音频文件**
   ```
   项目文件夹/
   ├── music_to_ppt.exe
   ├── hebing.exe
   └── music/              # 将音频文件放这里
       ├── 歌曲1.flac
       ├── 歌曲2.mp3
       └── ...
   ```

2. **转换音乐为PPT**
   - 双击运行 `music_to_ppt.exe`
   - 等待处理完成（会在 `output/` 文件夹生成PPT）

3. **合并PPT（可选）**
   - 双击运行 `hebing.exe`
   - 按提示选择要合并的PPT编号
   - 在 `ppt_output/` 查看结果

### 方式B：运行Python源码

```bash
# 1. 克隆或下载项目
git clone <repository-url>
cd music-to-ppt

# 2. 安装依赖
pip install -r requirements.txt

# 3. 放置音频文件到 music/ 文件夹

# 4. 运行转换
python music_to_ppt.py

# 5. （可选）合并PPT
python hebing.py
```

---

## 📖 详细使用

### 步骤1：配置AI功能（可选但推荐）

首次运行 `music_to_ppt` 会自动生成 `ai_config.json`：

```json
{
    "enabled": true,
    "api_key": "sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
    "base_url": "https://api.openai.com/v1",
    "model": "gpt-3.5-turbo",
    "max_retries": 3,
    "max_workers": 4
}
```

**配置说明**：
- `enabled`: 是否启用AI清洗（`true`/`false`）
- `api_key`: 你的OpenAI API密钥（**必填**）
- `base_url`: API地址（可用国内转发）
- `model`: 使用的模型（推荐 `gpt-3.5-turbo`）
- `max_retries`: AI调用失败重试次数
- `max_workers`: 并发线程数

> **提示**：如果未配置API密钥，程序仍会运行，但不会进行歌词智能清洗。

### 步骤2：转换音乐为PPT

```bash
# 方式A：双击 music_to_ppt.exe
# 方式B：运行 python music_to_ppt.py
```

**处理流程**：
1. 自动整理根目录的音频文件到 `music/`
2. 提取元数据（标题、歌手、歌词、封面）
3. AI清洗歌词（如果启用）
4. 生成PPT（封面页 + 歌词页）
5. 输出到 `output/` 文件夹

**输出示例**：
```
[总耗时] 45.32 秒
[成功] 10 首
[失败] 0 首
-----------
[纯音乐] 3 首 (平均 1.2 秒/首)
[带歌词] 7 首 (平均 5.8 秒/首)
```

### 步骤3：合并PPT和音频

```bash
# 方式A：双击 hebing.exe
# 方式B：运行 python hebing.py
```

**交互示例**：
```
========================================
      PPT 合并与音频精准匹配工具
========================================
[1] 歌曲A.pptx
[2] 歌曲B.pptx
[3] 歌曲C.pptx

请选择 PPT 合并顺序 (输入数字序号，用空格分隔)。

请输入: 1 3 2  ← 按需输入顺序，或直接回车按默认顺序
```

**输出**：
- `ppt_output/合并后.pptx`
- 对应的音频文件（自动从 `music/` 复制）

---

## ⚙️ 配置说明

### PPT布局参数（高级）

在 `music_to_ppt.py` 的 `AudioToPPT` 类中可自定义：

```python
# 幻灯片尺寸（16:9）
self.SLIDE_W_INCH = 13.333
self.SLIDE_H_INCH = 7.5

# 封面页专辑封面大小
ALBUM_COVER_SIZE_VAL = 4.8

# 歌词页封面大小
self.LYRIC_COVER_SIZE_VAL = 3.5

# 歌词行间距
self.LINE_SPACING = Inches(1.35)

# 字体样式
self.STYLE_ACTIVE = {'size': 40, 'bold': True, 'color': (255, 255, 255)}
self.STYLE_NORMAL = {'size': 24, 'bold': False, 'color': (160, 160, 160)}
```

### AI清洗规则

AI会按以下规则处理歌词：
1. ✅ 删除作词/作曲等头部元数据
2. ✅ 删除尾部宣传信息
3. ✅ 删除翻译内容（仅保留原文）
4. ✅ 保留原有换行格式
5. ✅ 长句智能分行（超过18字符插入 `^` 符号）
6. ✅ 检测纯音乐（返回 `[PURE_MUSIC]` 标记）

---

## ❓ 常见问题

### Q1: 提示"缺少 openai 库"
**A**: 运行 `pip install openai` 或禁用AI功能（`"enabled": false`）

### Q2: PPT合并时报错"无法启动PowerPoint"
**A**: 确保已安装Microsoft PowerPoint，并且没有其他PPT在编辑中

### Q3: 歌词显示不完整
**A**: 
- 检查音频文件是否包含嵌入歌词
- 验证AI配置是否正确（歌词可能被误判为纯音乐）

### Q4: 处理速度慢
**A**: 
- 调整 `max_workers` 增加并发数（建议不超过CPU核心数）
- 使用 `.exe` 版本（性能略优于Python解释器）

### Q5: 生成的PPT背景模糊
**A**: 这是设计特性（高斯模糊），可在代码中调整：
```python
bg_final = bg_final.filter(ImageFilter.GaussianBlur(radius=60))  # 减小 radius 值
```

### Q6: exe文件被杀毒软件拦截
**A**: 
- 这是误报（PyInstaller打包的exe常见问题）
- 添加到杀毒软件白名单
- 或直接使用Python源码运行

---

## 🔧 技术说明

### 依赖库版本
```
Pillow>=9.0.0
mutagen>=1.45.0
python-pptx>=0.6.21
openai>=1.0.0
pywin32>=300 (仅Windows)
```

### 图层渲染顺序（歌词页）
1. **背景层**：高斯模糊封面
2. **歌词层**：当前句高亮 + 上下文预览
3. **遮罩层**：顶部/底部渐变遮罩（PNG透明）
4. **封面层**：左侧专辑封面 + 歌曲信息

### 性能优化
- 使用 `BytesIO` 内存缓存图片避免磁盘I/O
- `ThreadPoolExecutor` 并发处理多文件
- LANCZOS重采样保证图片质量

---

## 📄 文件结构

```
项目文件夹/
├── music_to_ppt.py         # 主程序（音乐转PPT）
├── music_to_ppt.exe        # 可执行文件
├── hebing.py               # 合并工具
├── hebing.exe              # 可执行文件
├── ai_config.json          # AI配置（自动生成）
├── README.md               # 本文档
├── music/                  # 输入：音频文件
├── output/                 # 输出：生成的PPT
└── ppt_output/             # 输出：合并后的文件
```

---

## 📝 使用建议

1. **首次使用**：先用1-2个音频文件测试，确认效果满意再批量处理
2. **API成本**：使用GPT-3.5-Turbo成本较低，约 $0.002/首歌
3. **歌词质量**：高质量音频文件的嵌入歌词效果更好
4. **备份原文件**：建议备份原始音频文件

---

## 📧 反馈与支持

- **问题反馈**：提交Issue或联系开发者
- **功能建议**：欢迎提交PR

---

## 📜 许可证

本项目遵循 MIT 许可证，可自由使用和修改。

---

## 🎉 更新日志

### v0.0.2 (2025-12-21)
- ✅ 支持FLAC/MP3/WAV/M4A格式
- ✅ AI智能歌词清洗
- ✅ 高质量封面处理
- ✅ PPT合并与音频同步


