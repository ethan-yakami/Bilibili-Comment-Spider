# Bilibili Comment Spider (B站评论爬虫)

本项目是一个用于抓取 Bilibili 视频评论的 Python 爬虫工具。
支持批量抓取多个视频，支持二级回复（楼中楼）的抓取，并将结果导出为格式美观的 Excel 文件。

## 功能特性
- **两级评论抓取**：不仅抓取一级评论，还能自动递归抓取其下的所有二级回复（楼中楼）。
- **Wbi 签名计算**：内置 Bilibili 最新的 Wbi 签名算法（`w_rid`），防止接口报 `-403` 被拦截。
- **OID 自动解析**：直接输入 `BVID`（如 `BV1xx411c7XD`），程序自动转换为真实的视频 `OID` (aid) 并抓取。
- **批量处理**：支持在一个纯文本文件中配置多个 BV 号，程序自动排队循环抓取。
- **Excel 直观输出**：自带层级标记，一级评论整行蓝色高亮，二级回复内容缩进展示，排版极佳。
- **隐私保护**：Cookie 单独存放并加入 `.gitignore`，防止误传到 GitHub 导致账号被盗。

## 目录结构
```text
Bilibili_Comment_Spider/
├── main.py              # 核心爬虫程序
├── config_example.json  # 配置文件模板，供参考
├── config.json          # 真实的配置文件，存放你本人的 Cookie (被 git ignore)
├── bvid_list.txt        # 待爬取视频清单，每行填一个 BV 号
├── requirements.txt     # Python 依赖清单
├── .gitignore           # Git 忽略配置
└── README.md            # 项目说明文档
```

## 快速安装运行环境
1. 确保你已安装 Python 3.8 或更高版本。
2. 安装所需依赖库：
   ```bash
   pip install -r requirements.txt
   ```

## 配置方法

### 第一步：获取 Cookie
你必须要有自己的 B站 Cookie 才能调用接口获取完整数据。
1. 在浏览器（推荐 Chrome）中登录 [Bilibili](https://www.bilibili.com)。
2. 按 `F12` 打开开发者工具，切换到 **Network (网络)** 标签页。
3. 刷新页面，在左侧列表中随意点击一个请求（如 `nav` 或任意图片）。
4. 在右侧的 **Headers (标头)** -> **Request Headers (请求标头)** 中找到 `Cookie:` 字段。
5. 将 `Cookie:` 后面那一长串字符串全部复制。

### 第二步：配置 config.json
1. 在项目根目录下，复制 `config_example.json` 并重命名为 `config.json`（如果文件已存在可跳过）。
2. 将刚才复制的字符串填入 `cookie` 对应的值中（保留双引号）。
3. 也可以顺便更新 `user_agent` 为你本人的浏览器标识。

*注意：`config.json` 已在 `.gitignore` 中，你的私密信息不会被提交到版本库。*

### 第三步：添加待抓取的视频（BV 号）
打开 `bvid_list.txt`，填入你想要抓取的视频的 BV 号，每行一个。例如：
```text
BV1MxcLzyEVW
BV1xx411c7XD
```
*提示：以 `#` 开头的行会被当作注释跳过。*

## 运行程序
配置完成后，在项目目录下执行：
```bash
python main.py
```

## 输出说明
爬取结束后，程序会在当前目录自动生成一个 `output/` 文件夹。
每个视频都会生成一份独立的 Excel，命名如 `bilibili_comments_BV1MxcLzyEVW.xlsx`。

Excel 表头包含：
- **层级** (`▶ 评论` 或 `└ 回复`)
- **用户ID**
- **用户名**
- **评论内容**
- **点赞数**
- **回复数** (一级评论下的回复总量)
- **发布时间**
- **父评论ID** (若是二级回复，会标识它属于哪条主评论)

## 开源协议
MIT License
