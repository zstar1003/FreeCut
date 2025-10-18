# FreeCut Office Add-in 开发和部署指南

## 🎯 项目概述

FreeCut Office Add-in是一个基于Web技术的现代PowerPoint插件，使用HTML、CSS、JavaScript和Office.js API开发。相比传统的VSTO插件，这种方案具有更好的跨平台兼容性和更简单的部署过程。

## ✨ 技术栈

- **前端**: HTML5, CSS3, JavaScript (ES6+)
- **Office API**: Office.js
- **构建工具**: Webpack, Babel
- **开发工具**: Node.js, npm
- **部署**: 本地sideload或Office商店

## 🛠️ 开发环境要求

### 必需环境
- **Node.js** 16.0+ ([下载地址](https://nodejs.org/))
- **PowerPoint** 2016或更新版本
- **现代浏览器** (Chrome, Edge, Firefox)

### 可选工具
- **Visual Studio Code** (推荐编辑器)
- **Office Developer Tools** 扩展

## 📦 项目结构

```
FreeCut/
├── package.json              # 项目配置和依赖
├── webpack.config.js         # 构建配置
├── manifest.xml              # Office插件清单
├── build.bat                 # 构建脚本
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html     # 主界面
│   │   └── taskpane.js       # 主逻辑
│   └── function-file/
│       └── function-file.html # Ribbon按钮处理
├── assets/                   # 图标和资源
└── dist/                     # 构建输出
```

## 🚀 快速开始

### 1. 环境准备

```bash
# 检查Node.js版本
node --version

# 如果未安装，访问 https://nodejs.org/ 下载安装
```

### 2. 安装依赖

```bash
# 在项目根目录运行
npm install
```

### 3. 构建项目

```bash
# 开发构建
npm run build:dev

# 生产构建
npm run build
```

### 4. 启动开发服务器

```bash
# 启动本地HTTPS服务器 (https://localhost:3000)
npm run dev-server
```

### 5. 安装插件到PowerPoint

#### 方法一: 自动sideload (推荐)
```bash
npm run sideload
```

#### 方法二: 手动安装
1. 打开PowerPoint
2. 文件 → 选项 → 信任中心 → 信任中心设置
3. 受信任的加载项目录 → 添加目录
4. 选择项目的dist文件夹
5. 重启PowerPoint
6. 插入 → 我的加载项 → 查看所有 → 选择FreeCut

## 🔧 开发指南

### Office.js API使用

```javascript
// 初始化Office插件
Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        // 插件逻辑
    }
});

// 操作PowerPoint
await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();
    // 处理幻灯片
});
```

### 主要功能模块

#### 1. 幻灯片获取
```javascript
async function getSlides() {
    return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        return slides.items;
    });
}
```

#### 2. 设置管理
```javascript
// 保存设置到localStorage
localStorage.setItem('freeCutSettings', JSON.stringify(settings));

// 加载设置
const settings = JSON.parse(localStorage.getItem('freeCutSettings'));
```

#### 3. 裁剪算法
```javascript
function detectContentBounds(imageData, settings) {
    // 实现自动边界检测
    // 由于Office.js限制，复杂的图像处理需要服务器端支持
}
```

## 📋 功能特性

### ✅ 已实现功能
- 📱 现代化Web界面
- ⚙️ 完整的设置面板
- 🔄 实时设置保存
- 📊 幻灯片信息显示
- 🎛️ Ribbon集成

### 🚧 待完善功能
- 🖼️ 实际图像裁剪 (需要服务器端或Canvas API)
- 📄 真实PDF生成 (需要PDF库集成)
- 👁️ 实时预览 (需要图像处理能力)
- 🔍 高级检测算法

### 💡 实现方案

由于Office.js的安全限制，以下功能需要额外的实现：

#### 图像处理方案
1. **Canvas API**: 在浏览器中处理图像
2. **Web Worker**: 后台处理避免阻塞UI
3. **服务器端**: 发送到服务器处理
4. **WebAssembly**: 运行本地代码

#### PDF生成方案
1. **PDF-lib**: JavaScript PDF生成库
2. **jsPDF**: 客户端PDF创建
3. **服务器API**: 后端PDF生成服务

## 🚀 部署方案

### 本地部署
```bash
# 构建生产版本
npm run build

# 启动生产服务器
npm start
```

### Office商店部署
1. 完善manifest.xml配置
2. 准备商店资源 (图标、截图、描述)
3. 提交到Microsoft AppSource
4. 等待审核通过

### 企业部署
1. 部署到内部Web服务器
2. 配置HTTPS证书
3. 更新manifest.xml中的URL
4. 通过组策略或管理中心分发

## 🔒 安全考虑

### HTTPS要求
- Office Add-in必须使用HTTPS
- 开发时使用自签名证书
- 生产环境需要有效SSL证书

### 权限控制
```xml
<!-- manifest.xml中的权限设置 -->
<Permissions>ReadWriteDocument</Permissions>
```

### 跨域访问
```javascript
// 配置CORS头
headers: {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE',
    'Access-Control-Allow-Headers': 'X-Requested-With, content-type'
}
```

## 🐛 故障排除

### 常见问题

**Q: 插件无法加载**
```
A: 检查HTTPS证书，确保manifest.xml路径正确
   清除Office缓存：%LOCALAPPDATA%\Microsoft\Office\16.0\Wef
```

**Q: JavaScript错误**
```
A: 打开浏览器开发者工具查看控制台
   检查Office.js API版本兼容性
```

**Q: 无法sideload**
```
A: 确保PowerPoint支持开发者模式
   检查信任中心设置
   尝试手动加载manifest.xml
```

**Q: 构建失败**
```
A: 检查Node.js版本 (需要16.0+)
   删除node_modules重新安装
   检查网络连接和npm源
```

### 调试技巧

#### 1. 启用开发者模式
```
1. 文件 → 选项 → 信任中心 → 信任中心设置
2. 加载项 → 要求由受信任的发布者对应用程序加载项进行签名 (取消勾选)
3. 受信任的加载项目录 → 添加项目目录
```

#### 2. 查看日志
```javascript
// 使用console输出调试信息
console.log('Debug info:', data);

// 使用Office的日志系统
Office.context.mailbox.diagnostics.hostName;
```

#### 3. 网络调试
```
F12开发者工具 → Network → 查看API调用
检查manifest.xml是否正确加载
验证HTTPS证书状态
```

## 📈 性能优化

### 构建优化
```javascript
// webpack.config.js
module.exports = {
    optimization: {
        splitChunks: {
            chunks: 'all'
        }
    }
};
```

### 运行时优化
```javascript
// 使用防抖减少API调用
const debounce = (func, wait) => {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
};
```

## 🔮 升级路径

### 向后兼容
- 保持manifest.xml版本兼容
- API版本管理
- 设置数据迁移

### 功能扩展
- 添加新的Office.js API
- 集成第三方服务
- 支持更多Office应用

## 📚 参考资源

### 官方文档
- [Office Add-ins 文档](https://docs.microsoft.com/office/dev/add-ins/)
- [Office.js API 参考](https://docs.microsoft.com/javascript/api/office)
- [PowerPoint Add-ins 指南](https://docs.microsoft.com/office/dev/add-ins/powerpoint/)

### 开发工具
- [Office Add-in 验证器](https://github.com/OfficeDev/office-addin-validator)
- [Yeoman 生成器](https://github.com/OfficeDev/generator-office)
- [Script Lab](https://appsource.microsoft.com/product/office/WA104380862)

### 社区资源
- [Office Dev Center](https://developer.microsoft.com/office)
- [Microsoft 365 开发者计划](https://developer.microsoft.com/microsoft-365/dev-program)
- [Stack Overflow: office-js](https://stackoverflow.com/questions/tagged/office-js)

---

**🎉 现在你已经拥有了一个完整的Office Add-in项目框架！这个方案不需要Visual Studio，可以在任何支持Node.js的环境中开发和构建。**