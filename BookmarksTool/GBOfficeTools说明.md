﻿# GBOfficeTools说明 （仅供自己阅读，不对外发布）
【简书】BookmarksTool更新说明：https://www.jianshu.com/p/0bc0a4c52347

## Todo:

- 考虑实现根据配置文件，生成项目文件目录的功能。
- 考虑实现生成文件夹内所有word文件名称的功能，写入并生成txt文件。



## 更新记录



### 20230123 GBOfficeTools V2.0.0.0
1. 软件改名为GBOfficeTools，版本号从V2.0开始。
2. 优化软件UI。
3. 制作软件安装包，便于发布安装及更新。
4. Aspose.Words.dll更新为V19.11，Word书签工具运行速度比Aspose.Words.dll V18.7快5倍。

### 20230121 BookmarksTool V2.5.1.2
1. 制作软件安装包。
2. 取消项目生成单个文件功能，NuGet卸载了Fody 4.2.1、Costura.Fody 3.3.3。

### 20230116 BookmarksTool V2.5.1.1
1. 增加自动更新功能。
2. 新增Update.xml，用于自动更新。Nugut安装Autoupdater.NET.Official 1.6.4版本。
3. 使用HMS搭建简单HTTP服务器，把Update.xml和更新文件放入服务器。
4. 采用gitee实现软件自动更新，详见AutoUpdate.cs、Update.xml和 https://gitee.com/yaoleistable/GBOfficeTool 用于软件发布，和GBOfficeTools私人程序不一样。
5. Update.xml的更新说明，采用简书发布，对本文字更新即可：https://www.jianshu.com/p/0bc0a4c52347
6. 软件默认强制升级模式，可随时控制软件是否可用。

### 20230116 BookmarksTool V2.5.0
1. 优化UI界面，增加左侧导航栏。

### 20230114 BookmarksTool V2.4.1
1. 优化UI界面，感谢开源项目SunnyUI，https://gitee.com/yhuse/SunnyUI。

### 20230112 BookmarksTool V2.3.5
1. 禁止程序窗口放到和缩小。
2. 解决高分辨率电脑下，缩放失真问题。
3. 从.NET Framework 4.0 升级为.NET Framework 4.5，运行速度提升10%。
4. 目前版本Word批量转PDF功能，经测试，转换33各Word文件用时11.41~12.39秒，之前“WordTOPDF Office Required”版本用时2分08秒，速度提升了10倍。
5. 帮助文件中增加“软件将于2025年到期，，联系作者免费获取新版本。

### 20230110 BookmarksTool V2.3.4
1. 解决文件路径中存在~$.doc文件时，出现异常的问题，解决办法为先删除掉~$.doc临时（异常退出产生）的文件。
2. 修复word文件被打开时，软件无法转换PDF抛出异常的问题。

### 20230110 BookmarksTool V2.3.3
1. 修改参评范围图片存放路径，为word文件所在路径。
2. 软件打开后，txtbox不显示软件帮助信息。

### 20230108 BookmarksTool V2.3.2
1. 实现软件到指定日期不可用的功能。
2. 设置软件过期时间为2025年，Form1.cs第34行。
3. 软件打开后，txtbox显示软件帮助信息。
4. 运行软件最小化。

### 20230108 BookmarksTool V2.3
1. 书签替换还是用普通循环，并行循环容易卡死。
2. 调整UI大小，更美观些。

### 20230107 BookmarksTool V2.2
1. 优化文本框清楚内容的效果。
2. Word转PDF实现了普通循环转换和并行循环,并行循环计算时间节省约50%以上。

### 20230106 BookmarksTool V2.1
1. 全新升级，重新构建，可选择对于excel文件和word文件。
2. 把所有常用功能，集合到LeiTools类库中。
3. 增加Word批量转换PDF功能。
4. 增加文件夹选择功能，可以选择执行文件夹下面的Word报告生成和PDF转换。

### 20210128 BookmarksToolPro V1.5
1. 取消了加密功能，所有电脑均可运行。

### 20200405 BookmarksTool V1.4：
1. 增加了加密功能，与硬件绑定，防止软件外传。

### 20191025 BookmarksTool V1.0：
1. 设计了图像界面。
2. Word中插入英文书签，可对word模板进行批量替换。
3. 采用Costura.Fody.4.1.0、Fody.6.0.2，是最终程序变为一个单独文件。

### 20191023 BookmarksTool V0.1 No GUI:
1. 无图形界面，命令行运行。