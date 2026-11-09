# 蓝卡坞方舟助手

WrankwoodAssistantArknights，简称 WAA。

明日方舟 PC 端的全局帧轴执行器：按帧执行部署、技能、撤退，配一个 Excel 实时面板和悬浮帧窗。
帧号由费用条画面推算，不读内存，也不改游戏文件。

当前是早期预览版 0.1.0。功能还不完善，只在开发机（16:9、UI 100%、管理员权限）上验证过，
没有主窗口，轴格式和控制台命令也可能随时改。

## 名字

对外介绍用「蓝卡坞方舟助手」，代码、目录、日志、控制台里用 WAA 或
WrankwoodAssistantArknights。第一次出现写「蓝卡坞方舟助手（WAA）」。

## 开发环境

```
py -3.13 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
waa_console.bat
```

waa_console.bat 会自己提权。必须是管理员，否则触摸注入会被系统丢掉，表现就是点了没反应。

启动后默认会打开悬浮帧窗。不想要的话，临时用 WAA.exe --no-showframe 起一次，
或者在控制台里输入 showframe auto off 永久关掉，showframe auto on 再打开。

想确认环境有没有问题，跑自检：

```
.venv\Scripts\python.exe tools\axis_console.py --selftest
```

它会检查运行环境、第三方依赖、随包资源、out 目录写权限、帧窗程序、文件对话框、OCR 语言包，
七项都过才算可用。自检不碰游戏，也不会抢单实例锁，控制台开着的时候也能跑。

## 打包

```
build_exe.bat
build_exe.bat --with-resources
```

产物在 release\WAA-v0.1.0，双击 WAA.exe 就能用，目标机器不需要装 Python。
目录名跟着 waa/__init__.py 里的版本号走，打包时还会顺手生成同名 zip。
加 --with-resources 会把地图、头像、角色表一起打进去（多 330MB），用户就不用装 Git 去拉数据。

build_exe.bat 只是个壳，真正干活的是 tools\build_release.py。

bat 里不能写中文。cmd 按 OEM 代码页解析 .bat，UTF-8 的中文会让它把行拆错，
报出来的错误是「'xxx' 不是内部或外部命令」。要加提示就交给 Python 打印。

## 测试

改完代码跑这 15 项，全过再打包：

```
for %t in (test_console test_inplace_report test_deploy_f6 test_skill_flow test_action_retry test_avatar_flags test_frame_counter_state test_resolution test_rate_action test_rate_auto test_blind_step synth_neg_test test_map_live test_status_composite test_pause_action) do .venv\Scripts\python.exe tools\%t.py
```

## 目录

```
waa/            执行器、帧计数、视觉识别、Excel 面板
tools/          入口（axis_console / run_axis / frame_window）和 24 个测试
axes/           轴文件
out/            运行输出；ArknightsGameResource 和 gamedata 是地图、头像、角色表
release/        打包产物（WAA-v<版本>/ 和同名 zip）
packaging/      面向用户的说明，以及打包时要用的文件
reference/      开发时参考的项目（MAA、prts-plus、AFC 等），运行时不读
archive/        归档的一次性脚本和旧产物
```

## 计划

按优先级从高到低，顺序大致是这样：

1. 脱离费用条依赖。现在帧号靠费用条相位推算，遇到无回费或者界面挡住费用条的情况就失效，
   想换成不依赖单一视觉特征的算法。
2. 提升自动操作效率。压缩每步之间的等待和重试，能并行准备的动作提前到前面的帧。
3. 完善报错信息。现在的提示偏内部视角，要改成能直接说清哪一步、为什么、用户该做什么。
4. 自动进图。省掉手动选关、点开始的步骤。
5. 主窗口。控制台之外给一个正式界面，降低上手门槛。
6. 模拟器支持。目前只支持 PC 客户端，模拟器的分辨率和渲染差异要单独适配。
7. 敌人路径演算。按出怪路线预判，而不是只执行写死的动作序列。

## 致谢

这个项目的格式和思路参考了下面几个项目，代码是另写的，没有直接复制：

- ArknightsGameResource（yuanyan3060）：地图、头像、角色表的数据来源，程序运行时从该仓库拉取
- prts-plus：Excel 轴的表格结构、地图坐标与朝向约定都沿用它
- ArknightsFrameAssistant（AFC）：全局逻辑帧由费用条相位推算的思路
- ArknightsCostBarRuler（https://github.com/ZeroAd-06/ArknightsCostBarRuler）：费用条读数与回费模型的参考实现
- MaaAssistantArknights：界面识别和资源组织方式的参考

如果哪里漏了署名，告诉我补上。

## 免责

仅供学习和个人使用。用它跑自动化可能违反游戏用户协议，账号风险自己承担。
本项目不读取游戏内存、不修改游戏文件，但也不对任何封号或数据损失负责。
与鹰角网络及上述参考项目的作者均无关联。

## 许可证

MIT。见 LICENSE 文件。

作者：威猛先生（GitHub @wekapipo-378）。


