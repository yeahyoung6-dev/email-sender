# 邮件发送工具

批量点对点发送个性化邮件的桌面应用。

## 功能特点

- 📧 批量发送个性化邮件
- 📊 支持Excel文件导入（姓名、人员编码、卡号、密码、邮箱地址）
- 📝 邮件模板支持变量替换
- ⏱️ 自动控制发送间隔，防止被邮件服务器拒绝
- 🖥️ 打包成独立exe，无需安装Node.js

## 使用方法

### 方式一：直接运行（需要Node.js）

```bash
npm install
npm start
```

然后在浏览器打开 http://localhost:3000

### 方式二：打包成exe（推荐分发）

在Windows上运行：
```bash
npm install
npm run build:win
```

生成的exe文件在 `release/` 目录。

## GitHub Actions自动打包

1. 将代码推送到GitHub
2. 打标签触发打包：
   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```
3. 在GitHub Actions页面下载打包好的exe文件

## Excel文件格式

| 姓名 | 人员编码 | 卡号 | 密码 | 邮箱地址 |
|------|----------|------|------|----------|
| 张三 | 001 | 123456 | abc123 | zhangsan@example.com |

## 邮件模板变量

在邮件主题和正文中可以使用以下变量：
- `{姓名}`
- `{人员编码}`
- `{卡号}`
- `{密码}`
- `{邮箱地址}`

## SMTP配置示例

| 邮箱 | 服务器 | 端口 |
|------|--------|------|
| QQ邮箱 | smtp.qq.com | 465 |
| 163邮箱 | smtp.163.com | 465 |
| Gmail | smtp.gmail.com | 587 |

**注意**：QQ/163邮箱需要在邮箱设置中开启SMTP服务并获取授权码。

## License

MIT