# 自愿赞助与 PayPal 接入

**OfficeBridge 完全免费，所有功能都不要求赞助。** 赞助仅用于自愿支持维护，不解锁功能，也不影响检查、下载或离线使用。

当前没有配置收款账户或链接，**项目尚不接受实际赞助，应用中的赞助入口隐藏**。`.github/FUNDING.yml` 也只保留注释，没有启用收款地址。

## 以后怎样接入

1. 维护者自行登录自己的 PayPal，取得公开的 PayPal.Me 收款链接，并核对收款主页显示的是自己的账户。是否可以使用 PayPal.Me 取决于账户及所在地区；获取方式见 [PayPal.Me 官方说明](https://www.paypal.com/us/cshelp/article/paypalme-frequently-asked-questions-help432)。
2. 在 [`src/support-config.js`](../src/support-config.js) 中，将 `donationUrl` 的空字符串换成该公开链接。`label` 是页脚入口文字，可以保留默认值：

   ```js
   export const donationUrl = '';
   export const label = '请我喝杯咖啡';
   ```

   上面保持空配置，不包含示例收款人。推荐使用不附加金额、查询参数或片段的 HTTPS PayPal.Me 主页链接。地址格式校验不能证明收款人身份；地址无效或留空时，入口仍然隐藏。
3. 如需同时显示 GitHub 仓库的 Sponsor 按钮，在 [`.github/FUNDING.yml`](../.github/FUNDING.yml) 中添加 `custom` 字段，值填同一条公开链接并用引号包住，提交到仓库默认分支；按需要在仓库 Settings → General → Features 中启用 Sponsorships。参见 [GitHub 官方设置说明](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/displaying-a-sponsor-button-in-your-repository)。
4. 重新构建和打包，再分发更新后的应用：

   ```sh
   npm ci
   npm run release:local
   npm run release:verify
   ```

   `release:local` 会执行测试、网页构建、单 HTML 构建和打包。已经发出去的旧 HTML 不会自动更新。若之前已经公开发布，应更新版本号并创建新版本，不覆盖旧 Release 的文件。发布前查看新文件的页脚，确认链接去向和收款人；不需要进行测试付款。

不想自行改配置时，**以后只需把自己的公开 PayPal 收款链接发给我们**，即可继续完成接入。不要发送密码、验证码、API 令牌、银行卡资料或其他登录凭证。配置收款入口不包括代为创建支付账户、登录账户或发起交易。

入口只是一条由使用者主动点击的外部链接。赞助在 PayPal 页面单独完成，应用不内嵌支付表单或支付 SDK，也不会因未赞助而限制文档功能。停用时清空 `donationUrl`、移除 GitHub 的 `custom` 配置并重新发布即可。
