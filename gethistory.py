from xlfuns import xlApp
import numpy as np
from datetime import datetime
from functools import reduce


# TODO 数据用 numpy 存储，须要修改地方， check_data
# 几个 get index 最后取数据地方
# TODO 期号、开奖号、试机号等都统一用 string 形式

# class LotteryDataFilter()
# 写成类的形式，然后3d p3 ssq 分别继承并改写
class D3DataFilter():
    start_year = 2002 # 首次发行的年份
    current_year = datetime.now().year # 当前年份
    interval = 360 # 每年发行期数，根据彩票类型而不同，可以设置
    caizhong = ""

    def __init__(self,data=[],type="3d"):
        self.data = data
        self.caizhong = type.lower()
        self.__check_data(data)
        # 数据还要转换为 int64 类型，可以使用 numpy
        pass

    def __check_data(self,data):
        # TODO 检查数据合法性。想法是根据数据长度，做个简单的检查，3d 的数据应该
        # 应该等于 7 : 期号，开奖，试机号。p3 的暂时未定，到时候再修改，也可以用
        # 于 ssq ,到时候在修改吧。检测成功返回 True 否则 False
        if data:
            fa = np.array(self.data,dtype=np.int64)
            self.data = fa.tolist()
            return True
        else:
            print('未能正确获取到数据，请检查！')
            return False

    def __check_issue(self,issue):
        # 对输入的期号进行检查，要么是 2024xxx 或者是小于 365
        # use try except
        try:
            int_issue = int(issue)
        except Exception as e:
            print(e)
            print(f"你输入的期号有错： {issue}")
            return False

        # 根据彩种不同，这里函数也是不同的
        match self.caizhong:
            case "3d" | "p3":
                # 期号是0-366 之间的数字
                if int_issue > 0 and int_issue < 370:
                    return True
                # 7 位数字形式，暂定21 - 23 世纪
                elif int_issue > 2000000 and int_issue < 2200001:
                    return True
                else:
                    print(f"你输入的期号有错，in DataFilter.__check_issue: {int_issue}")
                    return False
            case _ :
                return False

    def __chehck_jianghao(self,data:str):
        # 字符串转为数字在 000 - 999 之间
        try:
            tmp = int(data)
            if tmp >= 0 and tmp <= 999:
                return True
            else:
                return False
        except Exception as e:
            print(e)
            return False

    def __get_data_index_by_full_issue(self,issue:str):
        """根据一个全期号，比如 2024001 查找期号在数据中的 index 并返回，出错或
        没查到，返回 None"""
        if len(issue) != 7:
            print(f"你输入的期号不是 2024xxx 这样7 位长度： {issue}")
            # 不能返回负数，因为python list can use negative number as index
            return None
        # 类似 excel 中 row index 方法来处理，不过这次数据直接在内存中
        # if not self.__check_data(self.data):
        #     print("你输出的数据有错，请检查")
        #     return None
        int_issue = int(issue)
        year = int(int_issue/1000)
        index = len(self.data) -1
        issue_in_year = int_issue % 1000

        base_issue = self.data[index][0]
        base_year = int(base_issue / 1000)
        issue_without_year = base_issue % 1000
        while  base_year != year :
            index = index -  (issue_without_year-1) + int(self.interval/2) - \
                (base_year - year) * self.interval
            if index > -1:
                base_issue = self.data[index][0]
                base_year = int(base_issue / 1000)
                issue_without_year = base_issue % 1000
            else:
                return None

        # 进一步检查，数据是否正确。
        if base_year == year and index > -1:
            index = index - issue_without_year + issue_in_year

            if index >= len(self.data) or index < 0:
                return None
            if self.data[index][0] == int_issue:
                # return int(index)
                return index
            else:
                return None
        else:
            return None

    def __compare_with_jianghao(self,data:list,jianghao:int)->bool:
        if len(data) != 3 or not self.__chehck_jianghao(str(jianghao)):
            return False
        else:
            if reduce(lambda x,y:x*10+y,data) == jianghao:
                return True
            else:
                return False

    def __get_data_index_by_kjh(self,kjh):
        # try str(kjh) 先转换一下，后面就直接使用
        # 通过开奖号来查找
        result = []
        try:
            str_kjh = str(kjh)
            int_kjh = int(kjh)
        except Exception as e:
            return result
        for index in range(0,len(self.data)):
            if self.__compare_with_jianghao(self.data[index][1:4], int_kjh):
                result.append(index)
        return result

    def __get_data_index_by_sjh(self,sjh):
        # try str(kjh) 先转换一下，后面就直接使用
        # 通过开奖号来查找
        result = []
        try:
            str_sjh = str(sjh)
            int_sjh = int(sjh)
        except Exception as e:
            return result
        for index in range(0,len(self.data)):
            if self.__compare_with_jianghao(self.data[index][4:7], int_sjh):
                result.append(index)
        return result
    # TODO 解决参数的类型不统一问题，特别是 issue 和 jianghao 要统一，不管外面传
    # 什么，内部只用一种类型，方便使用

    def get_data_by_issue(self,issue):
        # 主要用来获取同期数据，特别是历史同期
        indexes = self.__get_data_index_by_issue(issue)
        return self.__get_data_by_index(indexes,count=1)

    def get_last_n_sjh_kjh(self,n:int=11):
        """返回最后 n 期的试机号和开奖号数据
        prarameter:
            n : int 返回的期数，默认是 11
        return:
            [[试机号,...],[开奖号,...]]
            列表中有两个列表，第一个是试机号的数据，第二个是开奖号的数据
        """
        # rlt =[[],[]]
        rlt = []
        for row in self.data[-n:]:
            rlt.append((self.__convert_list_to_str(row[4:7]), self.__convert_list_to_str(row[1:4])))
        return rlt


    def get_last_n_data(self,n=11)->list:
        """ 获取最近几期的开奖数据，默认11 行"""
        return self.data[-n:]

    def __convert_list_to_str(self,li:list)->str:
        return  reduce(lambda x,y:str(x)+str(y),li)

    def get_last_nth_kjh(self,n:int)->str:
        """获取数据中倒数第 n 个开奖号"""
        if n > len(self.data):
            print(f"你输入的 index 超过了 list 最大长度，请检查： {n}")
            return ""
        else:
            return self.__convert_list_to_str(self.data[-n][1:4])

    def get_last_nth_sjh(self,n:int)->str:
        """获取数据中倒数第n 个试机号"""
        if n > len(self.data):
            print(f"你输入的 index 超过了 list 最大长度，请检查： {n}")
            return ""
        else:
            return self.__convert_list_to_str(self.data[-n][4:7])

    def get_kjh_gensui(self):
        """开奖号跟随，默认取开奖号的后两期
        返回值： [[xxx,...],[xxx,xxx,...]]
        """
        kjh = self.__convert_list_to_str( self.data[-1][1:4] )

        count = 3
        indexes = self.__get_data_index_by_kjh(kjh)
        finds = self.get_data_by_kjh(kjh,count,"down")
        # 再次提取，只提取出开奖号
        # 就算有n 期也能处理
        rlt=[]
        for i in range(1,count):
            rlt.append([])

        for find in finds[:-1]:
            # for row in find[1:]:
            for i in range(0,count-1):
                rlt[i].append(self.__convert_list_to_str(find[i+1][1:4]))
            #     tmp.append(self.__convert_list_to_str(row[1:4]))
            # rlt.append(tmp)
        return rlt
        # if len(rlt) > 1:
        #     return rlt[:-1]
        # else:
        #     return []

    def get_kjh_previous_gensui(self)->list:
        """倒数第二期的开奖号跟随，取开奖号的后1期
        返回值： [xxx,xxx,...]
        """
        kjh = self.__convert_list_to_str( self.data[-2][1:4] )

        indexes = self.__get_data_index_by_kjh(kjh)
        finds = self.get_data_by_kjh(kjh,2,"down")
        # 再次提取，只提取出开奖号
        rlt = []
        for find in finds:
            for row in find[1:]:
                rlt.append(self.__convert_list_to_str(row[1:4]))
        return rlt

    def get_kjh_previous_gensui2(self)->list:
        """倒数第二期的开奖号跟随，取开奖号的第二期数据
        返回值： [xxx,xxx,...]
        """
        kjh = self.__convert_list_to_str( self.data[-2][1:4] )

        indexes = self.__get_data_index_by_kjh(kjh)
        finds = self.get_data_by_kjh(kjh,3,"down")
        # 再次提取，只提取出开奖号
        rlt = []
        for find in finds:
            rlt.append(self.__convert_list_to_str(find[2][1:4]))
            # for row in find[2:]:
            #     rlt.append(self.__convert_list_to_str(row[1:4]))
        return rlt[:-1]

    def get_sjh_gensui(self,sjh:str)->list:
        """今天最新试机号跟随，默认取包含当前试机号往后 count 期的开奖号
        prarame:sjh string 最新的试机号数据，字符串
        返回值： [xxx,xxx,...]
        """
        indexes = self.__get_data_index_by_sjh(sjh)
        finds = self.get_data_by_sjh(sjh,1,"down")
        # 再次提取，只提取出开奖号
        rlt = []
        for find in finds:
            for row in find:
                rlt.append(self.__convert_list_to_str(row[1:4]))
        return rlt

    def get_sjh_previous_gensui(self)->list:
        """昨天的试机号跟随，默认取包含当前试机号往后 count 期的开奖号
        返回值： [xxx,xxx,...]
        """
        sjh = self.__convert_list_to_str( self.data[-1][4:7] )
        indexes = self.__get_data_index_by_sjh(sjh)
        finds = self.get_data_by_sjh(sjh,1,"down")
        # 再次提取，只提取出开奖号
        rlt = []
        for find in finds:
            for row in find:
                rlt.append(self.__convert_list_to_str(row[1:4]))
        return rlt

    def get_lishichuhao(self,issue:str)->list:
        """历史出号：获取具有想同期号后三位的开奖数据
        prarm : issue string 7 或3 位数字组成的期号，
        return : [xxx,xxx,...] 开奖号组成的 list
        """
        if len(issue) == 7:
            tmp= issue[-3:]
        elif len(issue) == 3:
            tmp=issue
        else:
            return[]
        data = self.get_data_by_issue(tmp)
        rlt =[]
        for row in data:
            rlt.append(self.__convert_list_to_str(row[0][1:4]))
        return rlt


    def get_previous_lishichuhao(self)->list:
        # 上一期的历史出号
        issue = str(self.data[-1][0])
        return self.get_lishichuhao(issue)

    def get_data_by_kjh(self,kjh,count=3,direction="down"):
        indexes = self.__get_data_index_by_kjh(kjh)
        return self.__get_data_by_index(indexes,count,direction)

    def get_data_by_sjh(self,sjh,count=2,direction="down"):
        indexes = self.__get_data_index_by_sjh(sjh)
        return self.__get_data_by_index(indexes,count,direction)

    def __get_data_by_index(self,indexes,count=1,direction="center"):

        # direction 有 "center", "up"和 "down"
        # 主要用于历史同期数据取号
        # 甚至可以取出多期数据，返回到当某一年的出号情况。
        # 根据 count 的值和 direction 进行划分，生成一个范围，然后直接取值并返回
        # 返回的数据类型是 [[]] 组成的列表，直接切片生成。
        result = []
        # 这是 center 情况处理，如果遇到不能均分时候，上面号要进量比下面多一个
        if count < len(self.data) :
            for index in indexes:
                if count > 0:
                    match direction.lower():
                        case "center":
                            if count % 2 != 0:
                                index_begin= index - int((count-1)/2)
                            else:
                                index_begin= index - int(count/2) -1
                        case "up":
                                index_begin= index - count + 1
                        case "down" :
                                index_begin= index

                    index_end = index_begin + count
                    # 检测 index_begin 和 index_end 不超出数据范围。
                    if index_begin < 0:
                        index_gegin = 0
                        index_end = index_begin + count
                    if index_end > len(self.data):
                        index_end = len(self.data)
                        index_begin = index_end - count

                    result.append(self.data[index_begin:index_end])

                # else: # count <= 0
                #     return []

        else:
            result.apped(self.data)
        return result

    def __get_data_index_by_issue(self,issue)->list:
        """传入一个期号，如果是7 位 或者小于等于 3 位的数字，那么就从数据中查找
        期号在数据中的位置，并以列表 list 方式返回。"""
        if not self.__check_issue(issue):
            print(f"你输入的期号有错： {issue}")
            return []
        str_issue = str(issue)
        find_result=[]

        if len(str_issue) == 7:
            index = self.__get_data_index_by_full_issue(str_issue)
            if index:
                find_result.append(index)
        else:
            if len(str_issue) < 3:
                str_issue = str_issue.rjust(3,"0")
            for year in range(self.start_year,self.current_year+1):
                index = self.__get_data_index_by_full_issue(str(year)+str_issue)
                if index:
                    find_result.append(index)
        return find_result

    def get_last_nth_issue(self,n:int):
        return self.data[-n][0] if n<= len(self.data) and n>0 else 0

    def excel_cell_expand(self,cell:str,row:int,column:int)->str:
        """将一个单元格地址扩展为一个一个 row 行，column 列的 range address
        prarameter:
            cell : string 一个字符串表示的 excel 单元格地址
            row : int 行数
            column : int 列数
        return : 扩展后的 range 范围，例如 "A2:D5"
        """

        # TODO 方向控制，主要是跟 excle 中 xlup xldown xlright xlleft 一致
        row_start, column_start = self.get_row_column_num_from_cell_addr(cell)
        row_end = row_start + row -1
        column_end = column_start + column -1

        str_col_start = self.xlcolumn_num_to_label(column_start)
        str_col_end = self.xlcolumn_num_to_label(column_end)
        return f"{str_col_start}{row_start}:{str_col_end}{row_end}"

    def get_row_column_num_from_cell_addr(self,cell:str):
        """将一个合法的 excel cell address 转为行列号返回"""
        col=""
        row=0
        for ele in cell:
            tmp = ele.upper()
            if tmp >= 'A' and tmp <='Z':
                col = col + tmp
            else:
                break
        return int(cell[len(col):]),self.xlcolumn_label_to_num(col)

    def xlcolumn_num_to_label(self,number, major_ver=14):
        """excel 列号转标签

        Parameters:
            :number:(int) 列号
            :major_ver:(int) 主版本号，版本不同，最大列数是不同的
        Return:(str) 合法列号时候返回列标签，非法列号返回空字符串""
        """
        #  excel 2007--12  2010--14 2013--15 2016--16
        max_column = self.xlcolumn_get_max_column(major_ver)

        if number >= 1 and number <= max_column:
            ordA = ord('A')
            rlt = []
            while number > 0:
                remainder = number % 26
                number = (int)(number/26)
                if remainder == 0:
                    number -= 1
                    remainder = 26
                rlt.append(chr(ordA+remainder-1))

            rlt.reverse()  # reverse是针对当前列表，不会返回列表
            return "".join(rlt)

        else:
            return ""


    def xlcolumn_label_to_num(self,label, major_ver=14):
        """excel列标签转为列号

        Parameters:
            :label:(string) 要转换的列标签
        :returns:(int) 合法列标签返回相应列号；非法时候返回0

        """
        base = 1
        rlt = 0
        ordA = ord('A')
        for i in range(len(label)-1, -1, -1):
            char = label[i].upper()
            if char >= "A" and char <= "Z":
                rlt += base*(ord(char)-ordA+1)
                base *= 26
            else:
                return 0
        #  判断下是否超出列标范围
        max_column = self.xlcolumn_get_max_column(major_ver)
        return rlt if rlt <= max_column else 0


    def xlcolumn_get_max_column(self,major_ver):
        # TODO:  <15-06-20, yourname> 需要对major_ver合法性进行判断
        return 16384 if major_ver >= 12 else 256
