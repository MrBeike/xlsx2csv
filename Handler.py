import os,sys
import pandas as pd
import configparser
import PySimpleGUI as sg

class Handler:
    def __init__(self):
        sg.theme('LightBrown8')
        self.code = False
        self.configReader()
        
    def appPath(self, relativepath):
        """Returns the base application path."""
        if hasattr(sys, 'frozen'):
            basePath = os.path.dirname(sys.executable)
            # Handles PyInstaller
        else:
            basePath = os.path.dirname(__file__)
        print(os.path.join(basePath, relativepath))
        return os.path.join(basePath, relativepath)

    def createINI(self):
        # 避免打包—复制的复杂操作，直接将新建一个ini文件再写入内容。
        if not self.code:
            self.code = sg.popup_get_text('请输入单位统一信用代码', font=("微软雅黑", 12),title="单位统一信用代码")
        content = f'''[organization]
        ！# code即报送单位统一信用代码
        ！code = {self.code}
        ！[danwei]
        ！# 单位贷款报表命名规范(一般不变)
        ！csv_file_code = CLDKXX
        ！zip_file_code = DWDKXX
        ！[nonghu]
        ！# 农户贷款报表命名规范(一般不变)
        ！csv_file_code = NHZJ
        ！zip_file_code = NHZJ
        '''
        # 为了缩进好看点，字符串有几段多了空格，此处处理下
        content = content.split('！')
        content = [x.rstrip(' ') for x in content]
        with open(self.appPath('config.ini'), 'w+', encoding='utf-8') as file:
          file.writelines(content)
        sg.popup('ini文件创建成功', font=("微软雅黑", 12), title='提示')
        # 若重新创建ini文件，则需要重新读取ini文件
        self.configReader()
    
    def configReader(self):
        try:
            config = configparser.ConfigParser()
            config.read(self.appPath('config.ini'), encoding="utf-8")
            self.organ_code = config.get('organization', 'code')
            self.danwei_csv_code = config.get('danwei', 'csv_file_code')
            self.danwei_zip_code = config.get('danwei', 'zip_file_code')
            self.nonghu_csv_code = config.get('nonghu', 'csv_file_code')
            self.nonghu_zip_code = config.get('nonghu', 'zip_file_code')
            return True
        except configparser.NoSectionError:
            sg.popup('ini文件不存在或配置文件格式被破坏。将重新生成ini文件',font=("微软雅黑", 12),title='提示')
            self.code = sg.popup_get_text('请输入单位统一信用代码', font=("微软雅黑", 12),title="单位统一信用代码")
            self.createINI()
            return False

    def readFile(self,filePath,type):
        '''
        利用pandas读取工作表数据
        return：sheet_datas : 工作簿中所有工作表对象 list
        '''
        # 1.读取文件,只取第一张sheet
        data = pd.read_excel(filePath, sheet_name=0,header=0,na_values=[0],
                                    keep_default_na=False)
        if type == 'danwei':
            data = data[:-7]
        return data

    def writeFile(self,data,type,compression=False):
        '''
        :params  data,dataframe,the data read form excel file
        :params  type,str,danwei or nonghu
        :params  compression,bool,output file compression option
        '''
        if type =='nonghu':
            csv_filename = f'{self.organ_code}_{self.nonghu_csv_code}_{self.date}.csv'
            zip_filename = f'{self.organ_code}_{self.nonghu_zip_code}_{self.date}.zip'
        else:
            csv_filename = f'{self.organ_code}_{self.danwei_csv_code}_{self.date}.csv'
            zip_filename = f'{self.organ_code}_{self.danwei_zip_code}_{self.date}.zip'
        if compression:
            compression_opts = dict(method='zip',archive_name=csv_filename)
            data.to_csv(zip_filename,sep='|',header=False,index=False,compression=compression_opts)
        data.to_csv(csv_filename,sep='|',header=False,index=False,encoding='utf-8')
        return

    def gui(self):
        '''
        简单GUI
        :return: filepath  读取到的文件地址+文件名
        '''
        layout = [
            [sg.Text('ini配置文件调整：更改统一社会信用代码。', size=(40, 1),font=("微软雅黑", 12)), sg.Button(
                button_text='生成ini文件', key='createINI', font=("微软雅黑", 10))],
            [sg.Radio('农户贷款借款人信息表',"type",key ='nonghu',default=True,font=("微软雅黑", 12),size=(18,1)),sg.Radio(
                '单位贷款基础数据表',"type",key='danwei',font=("微软雅黑", 12),size=(15,1)),sg.Checkbox('生成Zip压缩包',key='compression',font=("微软雅黑", 12))],
            [sg.Text('报表数据日期(YYYYMMDD)',size=(20, 1),font=("微软雅黑", 12)),sg.InputText('', key="date",font=("微软雅黑",10))],
            [sg.Text('请选择文件所在路径', size=(15, 1), font=("微软雅黑", 12), auto_size_text=False, justification='left'),
             sg.InputText('表格路径', font=("微软雅黑", 12)), sg.FileBrowse(button_text='浏览', font=("微软雅黑", 10))],
            [sg.Submit(button_text='  提 交 ', font=("微软雅黑", 10), auto_size_button=True, pad=[5, 5]), sg.Cancel(
                button_text='  退 出  ', key='Cancel', font=("微软雅黑", 10), auto_size_button=True, pad=[5, 5])]
        ]
        window = sg.Window(
            'CSV格式化存储工具', default_element_size=(40, 3)).Layout(layout)
        # TODO这里面的逻辑需要优化
        while True:
            button, values = window.Read()
            if button == 'createINI':
                self.createINI()
            elif button in (None, 'Cancel'):
                break
                return False
            else:
                file_path = values['浏览']
                book_type = 'nonghu' if values['nonghu'] else 'danwei'
                compression = values['compression']
                self.date = values['date']
                type = file_path.split('.')[-1]
                if type in ('xlsx', 'xls'):
                    data = self.readFile(file_path, book_type)
                    self.writeFile(data, book_type, compression)
                    sg.popup('生成成功，文件位于程序同级路径下',
                             font=("微软雅黑", 12), title='提示')
                else:
                    sg.popup('所选文件非Excel工作簿类型文件，请重试',
                             font=("微软雅黑", 12), title='提示')


# 主程序入口
if __name__ == '__main__':
    H = Handler()
    H.gui()