import os
import xls_process


class SAPChecker:
    def __init__(self):
        print('\n'
              'SAP 主数据申请表检查工具\n'
              '填写模板时请勿删除原模板内容\n'
              'https://github.com/PrinnyWan/SAP_checker\n'
              )
        os.system("pause")
        self.root = './'
        self.file_names = os.listdir(self.root)

    def run(self):
        file_counter, excel_counter = 0, 0
        for x in self.file_names:
            if '.' in x:
                file_counter += 1
            if '.xls' in x:
                excel_counter += 1
                processer = xls_process.Xlps(self.root + x)
                if processer.check():
                    print(x + ' 通过检查')
                else:
                    print(x + ' 请修改后重新检查\n')
        os.system("pause")
        return


if __name__ == '__main__':
    SAPChecker().run()
