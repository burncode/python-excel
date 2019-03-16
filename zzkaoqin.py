#!/usr/bin/python
#author:burncode
#date:2019.3.16
import xlrd
import xlwt
import logging
import time

class Kqtools():

    # 设置日志
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                        filename="kq.log")
    logger = logging.getLogger(__name__)

    def __init__(self,input_filename,output_filename,total_users=None,Agroup =None,Agroupname="A组",Bgroup = None,Bgroupname="B组"):
        """

        :param input_filename:  待处理excel文件
        :param output_filename: 输出结果报存的excel文件
        :param total_users:     总员工数列表
        :param Agroup:          特定的A组列表
        :param Agroupname:      A组名称 默认为A组
        :param Bgroup:          B组
        :param Bgroupname:      B组名称 默认为B组
        """
        self.input_filename = input_filename
        self.output_filename =output_filename
        self.Agroup = Agroup
        self.Agroupname = Agroupname
        self.Bgroup = Bgroup
        self.Bgroupname = Bgroupname
        self.total_users =total_users

    def times_to_seconds(self,t1):
        # 时间转秒
        h1, m1, s1 = t1.hour, t1.minute, t1.second
        t1_secs = s1 + 60 * (m1 + 60 * h1)
        return (t1_secs)

    def secend_to_time(self,secend):

        # 秒转时间
        m, s = divmod(secend, 60)
        h, m = divmod(m, 60)
        stime = "%d:%02d" % (h, m)
        return stime

    def float_to_secends(self,excel_time):
        # float转化时间 单位秒
        x = excel_time  # a float
        my_time = int(x * 24 * 3600)  # convert to number of seconds
        # my_time = time(x // 3600, (x % 3600) // 60, x % 60)  # hours, minutes, seconds
        return my_time

    def dict_Avg(self,Dict):

        # 字典平均值
        lens = len(Dict)  # 取字典中键值对的个数
        sums = sum(Dict.values())  # 取字典中键对应值的总和
        avgs = sums / lens
        return avgs

    def fix_value(self,old_dict):
        new_dict = {}
        for k, v in old_dict.items():
            if isinstance(v, dict):
                v = self.fix_value(v)
            if v == u"打卡异常":
                new_dict[k] = float(v.replace("打卡异常", "0"))
            else:
                new_dict[k] = v
        return new_dict

    def excel_style(self,blod =False,bc = False):
        # --------------------样式设置---------------------
        style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
        al = xlwt.Alignment()
        al.horz = 0x02  # 设置水平居中
        al.vert = 0x01  # 设置垂直居中
        font = xlwt.Font()  # 为样式创建字体
        font.name = 'Times New Roman'
        font.bold = blod  # 黑体
        borders = xlwt.Borders()  # Create Borders
        borders.left = xlwt.Borders.THIN
        style.borders = borders
        if bc==True:
            pattern = xlwt.Pattern()  # Create the Pattern
            pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
            pattern.pattern_fore_colour = 22
            style.pattern = pattern
        style.font = font
        style.alignment = al
        return  style

    def handle_execl(self):

        wb = xlrd.open_workbook(self.input_filename)
        colx = []
        new_dict_list=[]

        for i in wb.sheets():

            colx.append(
                dict(zip(i.col_values(0, start_rowx=1, end_rowx=None), i.col_values(2, start_rowx=1, end_rowx=None))))
        for i in range(len(colx)):
            new_dict_list.append(self.fix_value(colx[i]))
        dict2 = {}
        total_result = {}
        for i in new_dict_list:
            for key, value in i.items():
                if key not in dict2:
                    dict2[key] = []
                dict2[key].append(value)

        logging.info("--时间段: %s  员工打卡记录--", wb.sheet_names())
        logging.info("员工姓名        打卡次数            总工作时长 ")
        for k, v in dict2.items():
            logging.info("{}               {}                       {}".format(k, len(v), self.secend_to_time(self.float_to_secends(sum(v)))))
            total_result[k] = self.float_to_secends(sum(v) / len(v)) # result 为员工总时长字典

        # bug修正 一个组里面有多个工作时长一样大的 或者一样小的员工
        # 部门平均时间计算 【这里我有点疑惑 几个sheet里面某个人都没有打卡的话，即某个人这几天都没有打卡,总人数就会缺1，所以组人员的列表我使用了自定义来源，而不是从列表总人中取，并做了判断】
        total_user_in_excel = total_result.keys()
        no_recode_users = total_user_in_excel ^ self.total_users
        if  no_recode_users == set():
            self.Bgroup = (set(self.total_users) - set(self.Agroup))
        else:
            print("没有打卡记录员工: %s" % list(no_recode_users))
            self.Agroup = set(self.Agroup) - no_recode_users
            self.Bgroup = (set(self.total_users) - no_recode_users-set(self.Agroup))
        Asubdict = dict([(key, total_result[key]) for key in self.Agroup])
        Bsubdict = dict([(key, total_result[key]) for key in self.Bgroup])
        Aavgtime = self.secend_to_time(self.dict_Avg(Asubdict))
        Bavgtime = self.secend_to_time(self.dict_Avg(Bsubdict))
        Amaximum_time = max(Asubdict.values())
        Aminimum_time = min(Asubdict.values())
        Amaximum_user = [k for k, v in Asubdict.items() if v == Amaximum_time]
        Aminimum_user = [k for k, v in Asubdict.items() if v == Aminimum_time]
        Bmaximum_time = max(Bsubdict.values())
        Bminimum_time = min(Bsubdict.values())
        Bmaximum_user = [k for k, v in Bsubdict.items() if v == Bmaximum_time]
        Bminimum_user = [k for k, v in Bsubdict.items() if v == Bminimum_time]
        logging.info('--------------------------------------------')
        logging.info("A组最大值:%8s    姓名: %-8s", self.secend_to_time(Amaximum_time), Amaximum_user)
        logging.info("A组最小值:%8s    姓名: %-8s",self.secend_to_time(Aminimum_time),Aminimum_user)
        logging.info("B组最大值:%8s    姓名: %-8s", self.secend_to_time(Bmaximum_time), Bmaximum_user)
        logging.info("B组最小值:%8s    姓名: %-8s", self.secend_to_time(Bminimum_time), Bminimum_user)
        logging.info('--------------------------------------------')
        #写表
        outfile = xlwt.Workbook()
        xlsheet = outfile.add_sheet('%s 时间段 员工工时统计结果 '% "-".join(wb.sheet_names()),cell_overwrite_ok=True)

        table_header = ["团队", "人数", "平均工时", "团队MAX", "团队Min"]
        Acontent = [self.Agroupname, len(self.Agroup), Aavgtime, '{}（ {}小时 ）'.format(",".join(Amaximum_user), self.secend_to_time(Amaximum_time)),
                    '{}（ {}小时 ）'.format(",".join(Aminimum_user), self.secend_to_time(Aminimum_time))]
        Bcontent = [self.Bgroupname, len(self.Bgroup), Bavgtime, '{}（ {}小时 ）'.format(",".join(Bmaximum_user), self.secend_to_time(Bmaximum_time)),
                    '{}（ {}小时 ）'.format(",".join(Bminimum_user), self.secend_to_time(Bminimum_time))]
        headerlen = len(table_header)

        for i in range(headerlen):
            xlsheet.col(i).width = 0x0d00 + i * 180
            xlsheet.write(0, i, table_header[i],self.excel_style(blod=True,bc=True))
            xlsheet.write(1, i, Acontent[i],self.excel_style())
            xlsheet.write(2, i, Bcontent[i],self.excel_style())
        print("开始写入excel文件,记得关闭查看过的生成文件...")
        outfile.save(self.output_filename)

def run_job(data_from_excel =True):
    """

    :param data_from_excel:  员工是否来源于excel表格导入，如果为false，则需要自定义总员工列表
    """
    Agroup = ["王刚", "李刚", "陈四", "陈五", "陈六"]
    if data_from_excel:
        wbu = xlrd.open_workbook('./data/user.xlsx')
        wbu_sheets = wbu.sheets()[0]
        Total = wbu_sheets.col_values(0, start_rowx=1, end_rowx=None)
    else:
        Total = ['陈一', '崔二', '邓三', '陈四', '陈五', '陈六', '戴七', '巩一', '郝十三', '蒋十八', '李刚', '林火火', '刘门', '戎正', '尚方', '申通', '王铁',
                 '王刚', '许证', '许通', '叶成', '张丁', '章兴', '朱买', '范开', '祝福', "maohan", "fefef"]
    now = time.strftime("%Y-%m-%d", time.localtime(time.time()))
    kqtools = Kqtools(input_filename="./data/kaoqin.xlsx", output_filename="./output/{}-result.xlsx".format(now), total_users=Total,
                      Agroup=Agroup)
    kqtools.handle_execl()

if __name__ == '__main__':
    #其中data下面的kaoqin为考勤结果，user为员工列表
    run_job()


