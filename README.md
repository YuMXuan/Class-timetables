# Class-timetables
This program only works for SWJTU-LEEDS joint school.  
这个程序仅适用于西南交通大学利兹学院课程表。  

Based on Python, the project is for converting the curriculum in Excel format to the .ICS format that can be recognized by calendar apps on computers, mobile phones, and so on in order to automatically remind the students of the time, place and related content of the class by the calendar software.  

基于Python的，有关于将Excel格式的课程表转化为电脑和手机以及其他设备上的日历APP能识别的.ICS格式，以达到能自动化地由日历软件提醒学生上课的时间、地点以及相关内容的项目。  

'Pre' is the preprocessor of the imported Excel file.  
'Main' is the main program.  
'conf_classInfo.json' is the export file of the preprocessed file.  
'conf_classTime.json' is the time of the class schedule.  
'*.xlsx' is the imported class timetable file.  
'*.ics' is the exported calendar file.  

建议使用日历软件搭配服务器使用，可以做到服务端更新，客户端自动同步更新。


订阅地址可以尝试：
-
1. https://drymx.com/cal/MEY4N.ics #最后文件名为专业（ME、EE、CS、CE）+年级（Y1、Y2、Y3、Y4）+是否提前30分钟提醒（Y/N）+.ics。
2. https://drymx.cn/cal/MEY4N.ics #国内服务器

没怎么更新维护代码，感兴趣的可以优化一下，本人并不是CS专业的。

!2019年后不保证及时更新。且并无2019至2020大一版本。


#Version 2019.09.27 (2.1) 国庆节版本
-
CE：
	Y2-20190924
	Y3-20190924
	Y4-20190912
CS：
	Y2-20190925
	Y3-20190825
	Y4-20190826
EE：
	Y2-20190923
	Y3-20190925
	Y4-20190925
ME：
	Y2-20190925
	Y3-20190925
	Y4-20190925

#Just do it. 
