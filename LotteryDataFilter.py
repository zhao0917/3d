
# 类设计时候，先要考虑这个东西怎么用，谁要用
# 我这里设计目的，就是用类存储一组结构化的数据，然后对数据进行过滤处理。
# 写成类的形式，然后3d p3 ssq 分别继承并改写

class LotteryDataFilter():
    start_year = 2002 # 首次发行的年份
    current_year = datetime.now().year # 当前年份
    interval = 360 # 每年发行期数，根据彩票类型而不同，可以设置
    caizhong = ""

    def __init__(self,data=[],type="3d"):
        self.data = data
        self.caizhong = type.lower()
        # 数据还要转换为 int64 类型，可以使用 numpy
        pass

    def __check_and_convert_data(self,data):
        # TODO 检查数据合法性。想法是根据数据长度，做个简单的检查，3d 的数据应该
        # 应该等于 7 : 期号，开奖，试机号。p3 的暂时未定，到时候再修改，也可以用
        # 于 ssq ,到时候在修改吧。检测成功返回 True 否则 False
        pass

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

    def __check_jianghao(self,data:str):
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
        if not self.__check_data(self.data):
            print("你输出的数据有错，请检查")
            return None
        int_issue = int(issue)
        year = int(int_issue/1000)
        index = len(self.data) -1
        issue_in_year = int_issue % 1000

        base_issue = self.data[index][0]
        base_year = int(base_issue / 1000)
        issue_without_year = base_year % 1000
        while  base_year != year :
            index = index -  (issue_without_year-1) + int(self.interval/2) - \
                (base_year - year) * self.interval
            if index > -1:
                base_issue = self.data[index][0]
                base_year = int(base_issue / 1000)
                issue_without_year = base_year % 1000
            else:
                return None

        # 进一步检查，数据是否正确。
        if base_year == year and index > -1:
            index = index - issue_without_year + issue_in_year
            if self.data[index][0] == int_issue:
                return index
            else:
                return None
        else:
            return None

    def __compare_with_jianghao(self,data:list,jianghao:int)->bool:
        if len(data) != 3 or not self.__check_jianghao(str(jianghao)):
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
                                index_begin= index - (count-1)/2
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
                        index_end = self.data
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
