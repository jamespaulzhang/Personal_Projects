```markdown
# Steam 游戏技术支持信息提取器

![Tampermonkey](https://img.shields.io/badge/Tampermonkey-4.0+-black?logo=tampermonkey)
![GitHub license](https://img.shields.io/github/license/jamespaulzhang/steam-support-info-extractor)
![GitHub release](https://img.shields.io/github/v/release/jamespaulzhang/steam-support-info-extractor)

自动提取 Steam 游戏商店页面的技术支持邮箱和客服站点链接，并在页面顶部显示可关闭的信息栏。

## ✨ 功能特性

- 🔍 自动识别当前游戏的 appid
- 📧 提取官方客服邮箱地址（自动添加 mailto: 链接）
- 🌐 获取客服站点链接（支持多语言）
- 🎨 Steam 风格原生界面集成
- ✖️ 一键关闭显示框
- 🚀 即时加载，无页面刷新

## 📥 安装方法

### 基本安装
1. 安装 [Tampermonkey](https://www.tampermonkey.net/) 浏览器扩展
   - [Chrome 版](https://chrome.google.com/webstore/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
   - [Firefox 版](https://addons.mozilla.org/firefox/addon/tampermonkey/)
   - [Edge 版](https://microsoftedge.microsoft.com/addons/detail/tampermonkey/iikmkjmpaadaobahmlepeloendndfphd)

2. [点击一键安装脚本](https://greasyfork.org/zh-CN/scripts/541158-steam-游戏技术支持信息获取)

3. 访问任意 Steam 游戏商店页面（如 `https://store.steampowered.com/app/292030`）即可使用

### 手动安装
1. 复制[脚本代码](src/script.user.js)
2. 在 Tampermonkey 中点击 "新建脚本"
3. 粘贴代码并保存

## 🛠️ 开发指南

### 克隆仓库
```bash
git clone https://github.com/jamespaulzhang/steam-support-info-extractor.git
cd steam-support-info-extractor
```

### 项目结构
```
src/
  └── script.user.js       # 主脚本文件
.editorconfig              # 代码风格配置
.gitignore                 # Git忽略规则
LICENSE                    # MIT许可证
README.md                  # 本文件
```

### 构建与测试
1. 安装依赖（如需）：
   ```bash
   npm install
   ```

2. 开发时建议使用 Tampermonkey 的 "File" 选项直接加载本地脚本文件

**更新日志**：
   ```markdown
   ## 📜 更新日志

   ### v1.0
   - 初始发布版本
   - 支持基础的联系信息提取
   - 添加可关闭的信息栏
   ```

**常见问题**：
   ```markdown
   ## ❓ 常见问题

   Q: 为什么有些游戏找不到联系方式？  
   A: 部分游戏开发商可能未在Steam页面提供明确的客服信息。
   ```

## 🤝 贡献指南

欢迎通过 Issues 或 Pull Requests 贡献代码！  
请确保您的修改：
- 保持代码风格一致
- 包含适当的注释
- 通过基础功能测试

## ⚖️ 开源协议

本项目采用 [MIT License](LICENSE) 开源。
```