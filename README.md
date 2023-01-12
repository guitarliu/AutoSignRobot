# AutoSignRobot
Auto Create Word/Excel Documents with Personal Signs

本程序为一键签到助手，可将选中人员电子签名（人员姓名、人员部门、人员电话）图片直接写入Word或Excel文档中；本程序主要包含三个模块：党务工作签到、公司继续教育签到、部门内部培训签到，同时各个模块均支持增加或删除人员信息功能（不包括增删人员电子签功能）。

## 功能
- [x] 一键生成党务工作签到表（Word）
- [x] 一键生成继续教育签到表（Word）
- [x] 一键生成内部培训签到表（Excel） 

## 文件替换

程序内提供的人员名单文档PersonnelList.db可用记事本进行编辑，人员名单格式相关信息如下表：
|项目|内容|备注|
|:-:|:-:|:-:|
|人员名单格式|“姓名\t部门\t专业\t电话\tZZ面貌”||
|人员名单文件路径|安装路径/AutoSignRobot/AutoSignRobot/DataRources/PersionnelList.db||
|人员电子签名格式|“XX姓名.svg”、“XX部门”、“XX专业”|XX为人员姓名，SVG格式图片分辨率在缩放过程中保持不变|
|人员电子签名路径|安装路径/AutoSignRobot/AutoSignRobot/DataRources/SignImages|**SVG图片**，**建议去背景**，大小分别为**姓名签350*200像素**、**部门签500*200像素**、**电话签500*200像素**|

## 声明

本程序为辅助签到工具，任何个人或团体不得将本工具用于从事任务商业、非法或侵权活动；由此对他人造成的权利侵害由用户自行承担；