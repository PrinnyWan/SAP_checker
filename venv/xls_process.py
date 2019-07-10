import re

import xlrd


class xlps():
    def __init__(self, path):
        self.workbook = xlrd.open_workbook(path)
        self.severe_problems = []
        self.small_problems = []
        self.status = True
        # 是否通过了检查

    def check(self):
        if '供应商数据申请' in self.workbook._sheet_names:
            self._check_supplier_sheet(self.workbook.sheet_by_name('供应商数据申请'))
        elif '客户主数据收集模板' in self.workbook._sheet_names:
            self._check_customer_sheet(self.workbook.sheet_by_name('客户主数据收集模板'))
        else:
            return False
        return self.status

    def _check_supplier_sheet(self, sheet):
        for x in range(6, sheet.nrows):
            if sheet.row_values(x)[1] == '' and sheet.row_values(x)[3] == '' and sheet.row_values(x)[4] == '':
                continue
            print('\n公司名称：', sheet.row_values(x)[3])
            if sheet.row_values(x)[1] != '1000供应商-第三方':
                self.severe_problems.append('供应商、分包商帐户组非第三方：暂不适用检查非第三方账户组/或在excel下拉菜单中选择第三方正确格式')
            if sheet.row_values(x)[3] == '':
                self.severe_problems.append('供应商、分包商名称不能为空')
            elif len(str(sheet.row_values(x)[3])) > 35:
                self.severe_problems.append('供应商、分包商名称不能超过35个字符')
            if sheet.row_values(x)[4] == '':
                self.severe_problems.append('供应商、分包商类别不能为空')
            elif len(set(str(sheet.row_values(x)[4]).split(' ')[0].split('；')) | set(
                    ['E设备供应商', 'M材料供应商', 'S服务供应商', 'D设计分包商', 'C施工分包商', 'O运营分包商', 'L劳务分包商', 'OT其他'])) != 8:
                self.severe_problems.append('供应商、分包商类别填写有误，请按照要求填写/检查符号；/删除多余空格')
            if sheet.row_values(x)[5] == '':
                self.severe_problems.append('国家不能为空')
            elif len(str(sheet.row_values(x)[5]).split(' ')) != 2:
                self.severe_problems.append('国家格式错误，请按照要求填写/从下拉框选取')
            if sheet.row_values(x)[6] == '':
                self.small_problems.append('国外地址，不填写 省/直辖市,请确认')
            elif len(str(sheet.row_values(x)[6]).split(' ')) != 2:
                self.severe_problems.append('省/直辖市格式错误，请按照要求 国外地址不填写/从下拉框选取')
            if sheet.row_values(x)[7] == '':
                self.severe_problems.append('城市不能为空')
            if sheet.row_values(x)[8] == '':
                self.severe_problems.append('注册地址不能为空')
            elif len(str(sheet.row_values(x)[8])) > 35:
                self.small_problems.append('注册地址超出35字符限制，超出部分会被省略')
            if sheet.row_values(x)[9] == '':
                self.severe_problems.append('邮政编码不能为空')
            if sheet.row_values(x)[10] != '':
                self.severe_problems.append('第三方贸易伙伴不可填写')
            if sheet.row_values(x)[11] not in {'0000 法人', '0001 总经理', '0002 采购', '0003 销售', '0004 机构', '0005 管理',
                                               '0006 生产', '0007 质量保证', '0008 秘书', '0009 财务部', '0010 法律部'}:
                self.severe_problems.append('联系人部门填写错误，请按照要求 从下拉框选取')
            if sheet.row_values(x)[12] == '':
                self.severe_problems.append('联系人姓名不可为空')
            if sheet.row_values(x)[13] == '':
                self.severe_problems.append('联系人电话不可为空')
            if str(sheet.row_values(x)[14]) != '2202040100.0':
                self.severe_problems.append('第三方供应商、分包商统驭科目为2202040100，请修改')
            if sheet.row_values(x)[15] == '':
                self.small_problems.append('外国公司可不填写企业统一社会代码，请确认')
            elif re.compile("[\u4e00-\u9fa5]+").search(str(sheet.row_values(x)[15])):
                self.severe_problems.append('企业统一社会代码不可含中文，请检查/无法填写社会统一代码请留空并在附件注明')
            if sheet.row_values(x)[16] == '':
                self.severe_problems.append('法人代表不可为空')
            elif len(str(sheet.row_values(x)[16])) > 20:
                self.small_problems.append('法人代表超出20字符限制，超出部分会被省略')
            if sheet.row_values(x)[17] == '':
                self.severe_problems.append('公司地址不可为空')
            elif len(str(sheet.row_values(x)[17])) > 110:
                self.small_problems.append('公司地址超出110字符限制，超出部分会被省略')
            if sheet.row_values(x)[18] == '':
                self.severe_problems.append('经营范围不可为空')
            if sheet.row_values(x)[19] == '':
                self.severe_problems.append('交易币种不可为空')
            elif len(str(sheet.row_values(x)[19]).split(' ')[0]) != 3 or re.compile("[\u4e00-\u9fa5]+").search(
                    str(sheet.row_values(x)[19]).split(' ')[0]):
                self.severe_problems.append('交易币种填写错误，请按照要求填/从下拉框选取')
            if sheet.row_values(x)[20] == '':
                self.small_problems.append('特殊情况可不填写注册资本金，请确认')
            elif re.compile("[\u4e00-\u9fa5]+").search(str(sheet.row_values(x)[20])):
                self.severe_problems.append('注册资本金应为纯数字，请检查')
            elif int(sheet.row_values(x)[20]) > 100000:
                self.small_problems.append('注册资本金数额较大，注意该数字以万元为单位，请检查')
            if sheet.row_values(x)[21] == '':
                self.severe_problems.append('注册资本金币种不可为空')
            elif len(str(sheet.row_values(x)[21]).split(' ')[0]) != 3 or re.compile("[\u4e00-\u9fa5]+").search(
                    str(sheet.row_values(x)[21]).split(' ')[0]):
                self.severe_problems.append('注册资本金币种填写错误，请按照要求填写/从下拉框选取')
            self._out_put()
        return self.status

    def _check_customer_sheet(self, sheet):
        for x in range(6, sheet.nrows):
            if sheet.row_values(x)[1] == '' and sheet.row_values(x)[3] == '' and sheet.row_values(x)[4] == '':
                continue
            print('\n客户名称：', sheet.row_values(x)[3])
            if sheet.row_values(x)[1] != '1000客户-第三方':
                self.severe_problems.append('客户帐户组非第三方：暂不适用检查非第三方账户组/或在excel下拉菜单中选择第三方正确格式')
            if sheet.row_values(x)[3] == '':
                self.severe_problems.append('客户名称不能为空')
            elif len(str(sheet.row_values(x)[3])) > 35:
                self.severe_problems.append('客户名称不能超过35个字符')
            if sheet.row_values(x)[4] == '':
                self.severe_problems.append('搜索项不能为空')
            elif len(sheet.row_values(x)[4]) > 10:
                self.small_problems.append('搜索项长度大于10个字符，多余部分无法保存')
            if sheet.row_values(x)[5] == '':
                self.severe_problems.append('街道不能为空')
            elif len(sheet.row_values(x)[5]) > 35:
                self.small_problems.append('街道长度大于35个字符，多余部分无法保存')
            if sheet.row_values(x)[6] == '':
                self.severe_problems.append('公司地址不能为空')
            elif len(sheet.row_values(x)[6]) > 110:
                self.small_problems.append('公司地址长度大于110个字符，多余部分无法保存')
            if sheet.row_values(x)[7] == '':
                self.severe_problems.append('城市不能为空')
            if sheet.row_values(x)[8] == '':
                self.severe_problems.append('国家不能为空')
            elif len(str(sheet.row_values(x)[8]).split(' ')[0]) != 2 or re.compile("[\u4e00-\u9fa5]+").search(
                    str(sheet.row_values(x)[8]).split(' ')[0]):
                self.severe_problems.append('国家填写格式错误，请按照要求填写/从下拉框选取')
            if sheet.row_values(x)[9] == '':
                self.severe_problems.append('邮政编码不能为空')
            if sheet.row_values(x)[10] == '':
                self.severe_problems.append('联系电话不能为空')
            if sheet.row_values(x)[11] == '':
                self.severe_problems.append('联系人不能为空')
            if str(sheet.row_values(x)[12]) != '1122040100.0':
                self.severe_problems.append('第三方客户统驭科目为1122040100，请修改')
            if sheet.row_values(x)[13] == '':
                self.small_problems.append('外国公司可不填写企业统一社会代码，请确认并附件注明')
            elif re.compile("[\u4e00-\u9fa5]+").search(str(sheet.row_values(x)[13])):
                self.severe_problems.append('企业统一社会代码不可含中文，请检查/无法填写社会统一代码请留空并在附件注明')
            self._out_put()
        return self.status

    def _out_put(self):
        strings = ['严重问题:', '其他问题:']
        problems_list = [self.severe_problems, self.small_problems]
        if self.severe_problems:
            self.status = False
        for i in range(2):
            string = strings[i]

            print(string)
            if not problems_list[i]:
                print('无')
            for x in problems_list[i]:
                print(problems_list[i].pop(0))
        return
