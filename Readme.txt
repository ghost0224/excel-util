使用说明
将原始数据放在指定目录下，如：C:/excel-data/sample.xlsx
运行cmd，切换到excel-util.jar所在的目录，通过如下方法运行程序，第一段路径为原始文件路径，第二段路径为生成的文件路径：
java -jar excel-util-1.0.0-RELEASE.jar C:/excel-data/sample.xlsx C:/excel-data/打卡记录.xlsx

源码路径：https://github.com/ghost0224/excel-util.git

注：
目前不支持原始数据跨月情况，可以调打卡记录时按月拉取，或者手工调整后使用。