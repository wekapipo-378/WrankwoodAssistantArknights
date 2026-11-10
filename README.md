# 蓝卡坞方舟助手 WAA

明日方舟 PC 端的轴写入器、查看器与执行器。攻略者的赛博括约肌。按模板把轴写进 Excel，程序对照全局逻辑帧逐条执行部署、开技能、撤退。

- 可实时在游戏内显示帧号信息浮窗，适配欠费、气压状态
- OCR 识别干员信息，实时写回 Excel
- 用 Excel 公式编辑操作逻辑（如生命值小于 100 时撤退）
- Excel 实时显示当前地图示意图
- 随时可暂停并写入新操作

## 记帧规则

不读内存，不改游戏文件。全局逻辑帧靠费用条反推：

- 冷却初始值为 1（定点编码后存入计数器）。
- 费用条显示长度和冷却进度成正比。
- 从费用条读取到的视觉数值，映射回定点冷却计数器。
- 每一个逻辑帧，冷却计数器固定扣减 69905 / 2^21。
- 程序反向计算：根据费用条读出的剩余冷却值，反推已经走过多少帧。
- 这个帧号是估算值，不是游戏原生的逻辑帧真值。
- 视觉动画和逻辑结算不同步：动画延迟会造成反推帧号和真实逻辑帧存在偏差。
- EPS=167 判定规则不变：回费由底层逻辑计数器低于EPS触发，不以费用条视觉归零为准。

局内负费时回费减半，累加器相位原样保留，不重新起锚。回费速度还可以自动在轴里声明的几个档位之间检测切换。

干员、技能、费用数字靠 OCR 读画面，操作靠模拟触控。本质是看着屏幕替你点鼠标。

UI 缩放、窗口比例、遮挡都会影响，目前只在 16:9、UI 100%、全屏模式下验证过。

## 操作方法

- 提前量按轴设置：默认 1X 提前 4 帧、2X 提前 5 帧，可自己在轴里改。
- 到动作前，程序先暂停在目标帧之前。
- 点最右待部署卡，进入子弹时间。
- 子弹时间内逐帧步进：左半继续、右半暂停，或按 166ms 一帧走。
- 走到目标帧，执行部署、开技能或撤退。
- 执行完再用画面确认结果，成功才继续下一个动作。

## 环境

Windows 10 / 11，Python 3.13，明日方舟 PC 客户端。

硬性要求：16:9、UI 缩放 100%、窗口模式。必须以管理员权限运行，否则触摸注入会被系统静默丢弃。

## 安装

```bat
py -3.13 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

启动 waa_console.bat，会自己请求提权。

自检环境：

```bat
.venv\Scripts\python.exe tools\axis_console.py --selftest
```

## 使用

控制台里 load 或 run 打开一张轴（axes/ 下有示例），进游戏选好关卡进编队，程序自己进图开跑。F5 记当前帧并加一行，F6 暂停或恢复。更多命令看控制台 help。

启动会弹一个悬浮帧窗，显示当前帧。不想要就 showframe auto off，想开回来 showframe auto on。

## 状态

预览版 0.1.0。核心能跑通，没有主窗口。轴格式和控制台命令后面可能改，接口冻结前不会打 1.0.0。

## 打包

```bat
build_exe.bat
build_exe.bat --with-resources
```

产物在 release\WAA-v0.1.0，双击 WAA.exe 即用，目标机不用装 Python。加 --with-resources 把地图、头像、角色表一起打进去（约 330MB），用户不用自己拉数据。

## 测试

tools/ 下有 25 个测试脚本。打包前跑这 15 项回归：

```bat
for %t in (test_console test_inplace_report test_deploy_f6 test_skill_flow test_action_retry test_avatar_flags test_frame_counter_state test_resolution test_rate_action test_rate_auto test_blind_step synth_neg_test test_map_live test_status_composite test_pause_action) do .venv\Scripts\python.exe tools\%t.py
```

## 目录

```
waa/        执行器、帧计数、视觉识别、Excel 面板
tools/      入口和测试脚本
axes/       轴文件
out/        运行输出、随包资源
release/    打包产物
archive/    归档的一次性脚本和旧产物
```

## 计划

按优先级从高到低：

1. 脱离费用条依赖 — 现在帧号靠费用条相位反推，无回费或费用条被挡住就失效，想换成不依赖单一视觉特征的算法。
2. 提高自动操作效率 — 压缩每步之间的等待和重试，能并行准备的动作提前到前面的帧。
3. 完善报错 — 现在的提示偏内部视角，要改成能说清哪一步、为什么、用户该做什么。
4. 自动进图 — 省掉手动选关、点开始。
5. 主窗口 — 控制台之外给一个正式界面，降低上手门槛。
6. 模拟器支持 — 现在只支持 PC 客户端，模拟器的分辨率和渲染差异要单独适配。
7. 敌人路径演算 — 将查看出怪路线功能整合进此工具

## 致谢

格式和思路参考了下面几个项目，代码是另写的：

- [ArknightsGameResource](https://github.com/yuanyan3060/ArknightsGameResource) — 地图、头像、角色表数据
- [prts-plus](https://github.com/jue-ce-zhe/prts-plus) — Excel 轴结构、坐标与朝向约定
- [ArknightsFrameAssistant](https://github.com/CloudTracey/arknights-frame-assistant) — 费用条相位推帧
- [ArknightsCostBarRuler](https://github.com/ZeroAd-06/ArknightsCostBarRuler) — 费用条读数与回费模型
- [MaaAssistantArknights](https://github.com/MaaAssistantArknights/MaaAssistantArknights) — 界面识别与资源组织

## 声明

仅供学习交流，请勿用于商业用途。用自动化跑图可能违反游戏用户协议，风险自负。本工具不读内存、不改文件，也不对封号或数据损失负责，仅用于帧记录、时序分析、操作日志复盘，供学习研究使用。

## 许可证

[MIT](LICENSE)

作者：威猛先生（[@wekapipo-378](https://github.com/wekapipo-378)）
