#!/usr/bin/python
#! -*- coding: utf8 -*-
# 说明
#   钉钉直播数据的汇总统计，学生在课程中的听课时长
# 参考
#   https://zhuanlan.zhihu.com/p/111196174
# 操作流程
#   1. 确保使用的钉钉版本为4.7  [http://www.xitongzhijia.net/soft/34129.html]
#   2. 在钉钉班级群的直播回放中，下载[直播数据]，确保文件格式是csv
#   3. 将需要汇总统计的所有直播csv文件[如整个星期的课]放入指定文件夹[参数 live_data_path]
#   4. 设定班级学生名单excel文件[参数 class_info_path]
#   5. 设定班级名称[参数 class_info_name]
#   6. 设定输出汇总excel文件路径[参数 output_path]
#   7. 设定指定时长(秒)以下标记为红色样式[参数 red_style_duration_second]
#   8. 设定指定时长(秒)以下标记为黄色样式[参数 yellow_style_duration_second]
# external
#   date       2020-03-13 16:14:19
#   face       ●﹏●
#   weather    Shanghai Cloudy 12℃
import pandas
import os
import re
import xlrd
import xlwt

# ================ 基本配置 ================
live_data_path = 'C:\\path\\python\\dingding\\resources\\liveData'
class_info_path = 'C:\\path\\python\\dingding\\resources\\五5名单新.xlsx'
class_info_name = "五5班"
school_name = "ZZ小学"
output_path = 'C:\\path\\python\\dingding\\直播数据汇总统计.xls'
red_style_duration_second = 10           #默认10秒钟(未观看时长计为0)
yellow_style_duration_second = 1800      #默认30分钟
# ================ 基本配置 ================



# ================ 公共类\方法 ================
def time_to_second(t):
    if ":" in t:
        h, m, s = t.strip().split(":")
        return int(h) * 3600 + int(m) * 60 + int(s)
    else:
        return 0    
def time_to_str(t):
    time_str = ""
    if ":" in t:
        h, m, s = t.strip().split(":")
        if (h != "00" and h != "0"):
            time_str += h + "小时"
        if (m != "00"):
            time_str += m + "分"
        if (s != "00"):
            time_str += s + "秒"
    return time_str
week_name_dict = { 0: '周一', 1: '周二', 2: '周三', 3: '周四', 4: '周五', 5: '周六', 6: '周日'}
def format_to_dateweek(datetime):
    week_name = week_name_dict[datetime.dayofweek]  #Monday=0, Sunday=6.
    return datetime.strftime('%m.%d') + week_name
def format_time(time_str):
    if (time_str == "无"):
        return time_str
    else:
        datetime = pandas.to_datetime(time_str, format='%H:%M:%S')
        return datetime.strftime('%H:%M:%S')
        
        
def build_cell_style():
    # https://blog.csdn.net/weixin_44065501/article/details/88899257
    borders = xlwt.Borders()    # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    alignment = xlwt.Alignment()
    alignment.horz = 0x02   # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.vert = 0x01   # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    style = xlwt.XFStyle()
    style.borders = borders
    style.alignment = alignment
    return style
def add_red_cell_style(normal_style):
    font = xlwt.Font()
    font.colour_index = 2
    font.bold = True
    style = xlwt.XFStyle()
    style.borders = normal_style.borders
    style.alignment = normal_style.alignment
    style.font = font
    return style
def add_yellow_cell_style(normal_style):
    font = xlwt.Font()
    font.colour_index = 52  # 字体颜色索引值, 推荐51 52     https://blog.csdn.net/weixin_44065501/article/details/88874643
    font.bold = True
    style = xlwt.XFStyle()
    style.font = font
    style.borders = normal_style.borders
    style.alignment = normal_style.alignment
    return style
def add_header_cell_style(normal_style):
    font = xlwt.Font()
    font.height = 18 * 20   # 11为字号，20为衡量单位
    font.bold = True
    font.name = "华文行楷"
    style = xlwt.XFStyle()
    style.font = font
    style.borders = normal_style.borders
    style.alignment = normal_style.alignment
    return style
def build_duration_style(duration_str=None):
    if (duration_str is None):
        duration_second = 0
    else:
        duration_second = time_to_second(duration_str)
    if (duration_second < red_style_duration_second):
        duration_style = red_style
    elif (duration_second < yellow_style_duration_second):
        duration_style = yellow_style
    else:
        duration_style = normal_style
    return duration_str, duration_style
    
class Live_Day:
    def __init__(self, date, live_data):
        self.date = date
        self.live_data = live_data
    def __lt__(self, other): # override <操作符
        if self.date < other.date:
            return True
        return False
        

class Live_Data:
    re_teacher = re.compile(r'[(].*[)]', re.S)
    re_isolate_student = re.compile(r'[(].*[)]|妈妈|爸爸|妈妈|哥哥|姐姐|爷爷|奶奶|外公|外婆|阿姨|家长', re.S)
    def __init__(self, file_path):
        self.file_path = file_path
        self.basic_info = pandas.read_csv(file_path, sep='\\t', engine='python', skiprows=1, nrows=1, encoding='utf-16')
        self.data_frame = pandas.read_csv(file_path, sep='\\t', engine='python', skiprows=5, encoding='utf-16')
        # base_info
        file_name = os.path.basename(file_path)
        self.live_name = os.path.splitext(file_name)[0]
        self.live_time = pandas.to_datetime(self.basic_info.iat[0, 0], format='%Y-%m-%d %H:%M:%S')
        self.live_time_str = self.live_time.strftime('%Y-%m-%d %H:%M:%S')
        self.live_class = self.basic_info.iat[0, 1]
        self.live_teacher = re.sub(Live_Data.re_teacher, "", self.data_frame.iat[0, 0])
        self.live_length = time_to_str(self.basic_info.iat[0, 2])
        # duration
        self.students = self.data_frame.loc[(self.data_frame['部门']).str.isspace() == False]
        student_duration_dict = {}
        for one_record in self.students.to_numpy():
            #同一个学生的家长可以有多个账号(爸爸妈妈)，去除家长信息后，同一个学生会有多个观看记录，保留时间最大记录
            names = self.match_student_name(self.isolate_student_name(one_record[2]))
            duration = format_time(one_record[7])
            for name in names:
                live_student_data_old = student_duration_dict.get(name)
                if (live_student_data_old):
                    if (time_to_second(duration) > time_to_second(live_student_data_old.duration)):
                        student_duration_dict[name] = Live_Student_Data(name, duration)
                else:
                    student_duration_dict[name] = Live_Student_Data(name, duration)
        self.student_duration_dict = student_duration_dict
    def isolate_student_name(self, name_str):
        # 去除括号内部的类别、去除家长称谓(阿姨爸爸妈妈哥哥姐姐爷爷)、有同一家庭多学生情况\只保留指定班级的那个学生
        return re.sub(Live_Data.re_isolate_student, "", name_str)
    def match_student_name(self, name_str):
        if ("/" in name_str):
            need_check_names = name_str.strip().split("/")
        else:
            need_check_names = [name_str]
        valid_names = []
        for name in need_check_names:   # 与<班级学生名单>姓名做比较, 去除不在其中的记录
            if (name in class_student_names):
                valid_names.append(name)
        return valid_names
    def __lt__(self, other): # override <操作符
        if self.live_time < other.live_time:
            return True
        return False
        
class Live_Student_Data:    # 这里定义成类为了更好扩展
    def __init__(self, name, duration):
        self.name = name
        self.duration = duration
    def __str__(self):
        return self.name + "\t" + self.duration
# ================ 公共类\方法 ================



# ================ 基础数据 ================
os.system("title 直播数据汇总统计")
class_info_students = pandas.read_excel(class_info_path, sheet_name=0, skiprows=2, encoding='utf-16')
class_student_names = class_info_students["姓名"].to_list()


csv_files = []
live_day_dict = {}      #{"03.10周二", [Live_Data]}
for csv in os.listdir(live_data_path):
    if (os.path.splitext(csv)[1] == '.csv'):
        csv_files.append(os.path.join(live_data_path, csv))
        
for csv in csv_files:
    ld = Live_Data(csv)    
    date_str = format_to_dateweek(ld.live_time)
    live_day_one_day = live_day_dict.get(date_str)
    if (live_day_one_day):
        live_day_one_day.append(ld)
    else:
        live_day_dict[date_str] = [ld]
    
for day, live_one_day in live_day_dict.items():
    live_one_day.sort()
# ================ 基础数据 ================




# ================ 构建统计excel ================
workbook = xlwt.Workbook(encoding = 'utf-8')
sheet = workbook.add_sheet('Sheet1')
normal_style = build_cell_style()
red_style = add_red_cell_style(normal_style)
yellow_style = add_yellow_cell_style(normal_style)
header_style = add_header_cell_style(normal_style)

# 学生姓名列
sheet.write_merge(1, 2, 0, 0, "学生姓名", normal_style)    #行row1 row2 列col1 col2
sheet.col(0).width = 15 * 256   #256是基本单位
student_row_index = 3
for name in class_student_names:
    sheet.write(student_row_index, 0, name, normal_style)
    student_row_index += 1
# 全课总时长
sheet.write(student_row_index, 0, "全课总时长", normal_style)


#课程节次 学习时长
live_day_col_index = 1
for day, live_one_day in live_day_dict.items():
    live_day_col_index2 = live_day_col_index + len(live_one_day) - 1
    sheet.write_merge(1, 1, live_day_col_index, live_day_col_index2, day, normal_style)   #行row1 row2 列col1 col2
    
    for index, live in enumerate(live_one_day):
        live_index_name = "第" + str(index+1) + "节"
        #live_index_name = live.live_name
        live_col_index = live_day_col_index + index
        print(day + "\t" + live.live_name + "\t\t\t" + live.live_time_str + "\t\t\t" + live.live_length)
        sheet.col(live_col_index).width = 15 * 256   #256是基本单位
        sheet.write(2, live_col_index, live_index_name, normal_style)
        sheet.write(student_row_index, live_col_index, live.live_length, normal_style)
        # 学生学习时长
        student_row_index = 3
        student_duration_dict = live.student_duration_dict
        for name in class_student_names:
            student_duration = student_duration_dict.get(name)
            if (student_duration):
                duration_str, duration_style = build_duration_style(student_duration.duration)
                sheet.write(student_row_index, live_col_index, duration_str, duration_style)
                student_row_index += 1        
    live_day_col_index = live_day_col_index2 + 1
    
    
sheet.write_merge(0, 0, 0, live_day_col_index-1, school_name + "   " + class_info_name + "   线上教学学生参与情况汇总表", header_style)
workbook.save(output_path)
print("\t\t\t")
print("表格生成完毕:" + output_path)
os.system('pause')
exit(0)
# ================ 构建统计excel ================