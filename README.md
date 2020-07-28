# py-dingding-live-statistics
> 钉钉直播数据的汇总统计(Deprecated)


### 使用方法
1. 确保使用的钉钉版本为4.7  [http://www.xitongzhijia.net/soft/34129.html]
2. 在钉钉班级群的直播回放中，下载[直播数据]，确保文件格式是csv
3. 将需要汇总统计的所有直播csv文件[如整个星期的课]放入指定文件夹[参数 live_data_path]
4. 设定班级学生名单excel文件[参数 class_info_path]
5. 设定班级名称[参数 class_info_name]
6. 设定输出汇总excel文件路径[参数 output_path]
7. 设定指定时长(秒)以下标记为红色样式[参数 red_style_duration_second]
8. 设定指定时长(秒)以下标记为黄色样式[参数 yellow_style_duration_second]



### 演示
<div align=center><img src="https://github.com/bjc5233/py-dingding-live-statistics/raw/master/resources/demo.png"/></div>




### 注意点
1. python安装模块(pandas xlrd xlwt)
2. 姓名等数据已经修改，非真实信息
