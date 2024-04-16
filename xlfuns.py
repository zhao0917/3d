# 这个函数封装了一些使用xlwings访问excel的一些接口函数，
#  方便我们使用

import xlwings as xw


class xlApp(object):
    """保存一个xlwings app对象用来操作excel和错误处理"""

    def __init__(self, pid=0, visible=True, book=""):
        """
        :pid:int 默认值0.当存在excel进程时候，且pid！=0，会尝试连接
            进程号为pid的excel进程，如果失败或pid为0，则连接到xw.apps.keys()中
            第一个进程号所在的进程。  如果不存在已经
            打开的excel进程则新打开一个excel进程。
        :visible:boolean 默认True 控制excel app的可见性
        :book:str 默认是""，一个excel book的绝对路径
        """
        self.pid = pid
        self.new_app = False  # 是否新建了excel app，如果新建了要注意关闭
        self.visible = visible
        self.book_path = book
        self.books = []  # 字典保存是否更好，然后也可以编号，类似excel 用 item(i) 来访问
        self.active_book = None

        self.__connect_to_app()

    def __connect_to_app(self):
        #  连接到一个app，如果没有就创建一个
        if xw.apps.count == 0:
            self.app = xw.apps.add()
            self.pid = self.app.pid
            self.new_app = True
        else:
            #  当存在excel进程时候，试图取得pid指定的那个进程，
            #  如果此进程不存在，或pid=0，取得xw.apps.keys()[0]对应
            #  的那个进程
            if not self.pid == 0:
                try:
                    self.app = xw.apps[self.pid]
                except Exception:
                    self.pid = xw.apps.keys()[0]
                    self.app = xw.apps[self.pid]
            else:
                self.pid = xw.apps.keys()[0]
                self.app = xw.apps[self.pid]

        self.app.visible = self.visible
        self.open(self.book_path)

    def open(self, path=""):
        if not path == "":
            try:
                self.active_book = self.app.books.open(path)
                self.books.append(self.active_book)
            except Exception as e:
                print(e)

    @property
    def active_wb(self):
        return self.active_book

    def activate_book(self, book, steal_focus=False):
        """ 激活一个book
        Parameters:
        :book:要激活的工作簿的名或一个整数代表的index，表示类中存储的已打开的
            book中的index
        :steal_focus (bool, default False) – If True, make frontmost
            window and hand over focus from Python to Excel.
        """
        if isinstance(book, str):
            for wb in self.books:
                if wb.name == book:
                    wb.activate(steal_focus)
                    self.active_book = wb
        elif isinstance(book, int):
            if book < len(self.books):
                self.books[book].activate(steal_focus)
                self.active_book = wb

    def close_book(self, name="", save=True, save_as=""):
        # 还要把关闭的 wb 从 books 移除
        for book in self.books:
            if book.name == name:
                if save:
                    if save_as == "":
                        book.save()
                        self.books.remove(book)
                    else:
                        try:
                            book.save(save_as)
                            self.books.remove(book)
                        except Exception:
                            pass
                book.close()
                break

    def close_all_wb(self):
        # 关闭所有的wb
        pass

    def appclose(self, save_all=True):
        # TODO:  <14-06-20, yourname>
        # 新建的app会打开默认的工作簿，这个不保存
        for book in self.books:
            if save_all:
                if '.' in book.name:
                    book.save()
            book.close()
        if self.new_app:
            self.app.quit()
