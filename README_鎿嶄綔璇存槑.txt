PKR现金记账APP 操作说明

一、最快试用（网页版APP）
1. 打开 https://stackblitz.com/
2. Create Project -> React 或 Vite React
3. 上传/复制本项目文件
4. 等自动安装依赖后，点击 Preview 预览
5. 手机打开预览网址，可以添加到桌面使用

二、本地电脑运行
1. 安装 Node.js LTS
2. 解压本项目
3. 在项目目录打开命令行，输入：npm install
4. 输入：npm run dev
5. 浏览器打开命令行显示的网址

三、离线说明
数据保存在本机浏览器 localStorage 中。
项目包含 service worker，可以支持已打开后的离线缓存。
换手机、清理浏览器数据、换浏览器，原数据不会自动同步。

四、Excel模板
模板文件路径：public/templates/payment_approval_template.xlsx
不要改文件名，否则导出审批表会失败。

五、导出
点击“导出当日审批表”，生成：付款审批单_发票清单_日期.xlsx
点击“导出总台账”，生成：PKR现金总台账_日期.xlsx

六、真APK
这个项目是PWA网页APP。要打成真APK，建议先部署到一个网址，然后用PWABuilder生成APK；或者用Capacitor打包。
