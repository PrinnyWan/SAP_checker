import os
import xls_process


class SAPChecker:
    def __init__(self):
        print('\n'
              'Test. Write the document here'
              '\n')
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
                processer = xls_process.xlps(self.root + x)
                if processer.check():
                    print(x + ' 通过审核')
                else:
                    print(x + ' 请修改后重新审核\n')
        os.system("pause")
        return


if __name__ == '__main__':
    SAPChecker().run()
