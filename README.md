# PolyglotNote-Automator 

PolyglotNote-Automator 是一款专为外语学习者(比如我)设计的自动化西语/英语笔记工具。它可以让你在观看看视频或阅读时，一键完成“查词-翻译-排版-记录”的全过程。写这个小程序主要是为了我自己日常学习使用，起因是在使用谷歌字幕插件Trancy的时候，发现有些西语单词翻译得很奇怪，并且不太方便记录并且学习。我不是很喜欢动笔整理（笔记本经常丢qvq）于是自己捣鼓了一下。

## 功能特性
- **全局捕获**: 选中文字（例如Trancy字幕）按下快捷键即可触发。
- **生词记录**: 调用 Google 词典接口，自动识别并记录西语/英语单词或短语。
- **智能排版**: 自动生成蓝色标题、等线体、15号字体的专业 Word 笔记。
- **本地联动**: 支持一键唤起欧陆软件出品的《西班牙语助手》，一个我常用的本地西语查词软件。使用的时候必须先下载到本地。
- **自动存档**: 每日首次运行自动创建新文档，避免覆盖。

## 快捷键说明
- `S + D + C`: 识别西语单词（查词 + 记录 + 唤起本地助手，可以在软件内部点击星号收藏）
- `S + C`: Google翻译西语模式（仅记录）
- `E + C`: Google翻译英语模式（仅记录）
- `Esc`: 安全退出并保存时间戳

## 安装步骤
1. 克隆仓库: `git clone https://github.com/HarryHe6215/LexiFlow.git`
2. 安装依赖: `pip install -r requirements.txt`
3. 运行程序: `python main.py`

其他的用户需根据实际安装位置修改代码中的软件路径哦！

## 运行效果展示

1.选中字幕

<img width="1894" height="1296" alt="elegir las palabras" src="https://github.com/user-attachments/assets/72a2ccc2-e0ee-4ef1-839d-e8417d3590c6" />

2.查询生词

<img width="1887" height="1308" alt="boscando palabras con asistente" src="https://github.com/user-attachments/assets/e682520e-d148-4694-9a3b-5df155acb470" />

3.笔记生成

<img width="1315" height="975" alt="un ejemplo de la nota" src="https://github.com/user-attachments/assets/72cf4a95-0735-4579-89a9-6453be7ca916" />
