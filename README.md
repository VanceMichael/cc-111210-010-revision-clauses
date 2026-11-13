# 联展协议修订稿的条款采集

本项目用于授权资料接入的单次查看与接收留存，使用 Node.js 22。来源协议见 `docs/source-contract.md`。本地演练服务只绑定回环地址，数据不含真实个人资料或账号。

## 安装与运行

依赖由 `package-lock.json` 固定。首次准备环境运行 `npm ci --ignore-scripts`，浏览器项目使用 Playwright 1.56.1 对应的 Chromium；浏览器及系统库在镜像准备阶段安装，不在业务命令中下载。

```sh
npm run build
npm test
npm run demo
npm start
```

`npm run demo` 启动本地来源和接入服务，实际获取单次查看结果并写入临时接收记录，显示完整 JSON，随后关闭服务并清理临时目录。通常在十秒内结束。`npm start` 启动常驻服务，默认端口3000，可通过 PORT 与 DATA_DIR 修改监听端口和接收目录。设置 SOURCE_URL 可连接另一进程中已启动的本地来源；不设置时自动创建演练来源。恢复演练可先用 `npm run source` 在另一终端启动来源，记下输出地址，保持该进程运行，仅重启设置了 SOURCE_URL 的接入服务。

## 接口与模块

`GET /health` 查看状态，`GET /inspect` 查看来源（旧查看器，仅列部件与文本，不作为法律正文依据），`GET /collect?view=accept|reject|pending` 按修订视图采集条款（view 必填，三种视图结果不混用），`GET /receive?key=sample[&view=...]` 保存单次接收（带 view 时留存条款采集结果），`GET /receipts/sample` 读取回执。相同key和相同内容可重复接收，不同内容拒绝覆盖——同一 key 不能以另一视图的结果覆盖。

`src/ooxml.mjs` 按 OOXML 命名空间解析 XML、部件关系与路径；`src/clauses.mjs` 实现条款采集与跨版本比对；`src/connector.mjs` 解释单次来源响应并组装采集入口，`fixtures/source.mjs` 提供可重复启动的本地来源，`src/service.mjs` 组织访问入口，`src/store.mjs` 保存查看结果。`test/connector.test.mjs` 验证协议原语和演练源，`test/clauses.test.mjs` 验证条款采集语义，`test/service.test.mjs` 验证服务与留存边界。

## 采集原则

条款身份以 `w14:paraId` 为据（缺失时以部件内路径为据并注明 `identityBasis`），条款编号只来自 `numbering.xml` 的抽象编号与 `startOverride` 等覆盖规则，不从正文数字推断。插入、删除、移动来源与目标按修订标记归属到三种视图；移动按 `w:name`（或样本中的修订 `w:id`）配对，跨段落移动同样成对。表格单元格段落归属条款并标记 `inTable`；页眉页脚不进入条款，列入 `excludedParts`；脚注按 `w:footnoteReference` 与脚注部件对应。每条结果保留部件、段落路径、修订ID与编号依据（`number.basis`），`compareCollections` 据此对齐修订前后版本，区分内容变化与仅重新编号。

移动配对不完整、关系缺失、编号或脚注无法解析、段落标记删除合并、非标准容器归属不明等情况一律写入结果 `review` 数组交人工复核，采集器不自行猜测。

生产接入地址与凭据需由部署配置给出；演练客户端拒绝访问回环以外的地址。文档中的接口不向外提供任意目标代理。异常以明确错误响应返回，接收失败不覆盖原回执。
