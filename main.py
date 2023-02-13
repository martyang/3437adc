import os
import xlutils3.copy
import time
import serial
import pyvisa as visa
import xlwt
import xlrd
from pyvisa import VisaIOError
from serial import SerialException


def writeFile(sheet1, sheet2, volt_test, data_list, row, col):
    """
    将数据volt，data写入到sheet表指定row，col
    """
    max_data = max(data_list)
    min_data = min(data_list)
    avg_data = int(sum(data_list) / len(data_list))
    sheet1.write(row, 0, volt_test)
    sheet1.write(row, col, min_data)
    sheet1.write(row, col+1, max_data)
    sheet1.write(row, col+2, avg_data)
    sheet2.write(row, 0, volt_test)
    sheet2.write(row, col, hex(min_data))
    sheet2.write(row, col+1, hex(max_data))
    sheet2.write(row, col+2, hex(avg_data))


def powerSetVolt(power, volt):
    sysTem = power.query('*IDN?')
    if 'DH1766A' in sysTem:
        power.write('INST CH1\n')
        power.write('VOLT %f\n' % volt)
    else:
        power.write('VOLT %f,(@1)\n' % volt)


def powerSetCurrent(power, current):
    sysTem = power.query('*IDN?')
    if 'DH1766A' in sysTem:
        power.write('CURR %f\n' % current)
    else:
        power.write('CURR %f,(@1)\n' % current)


def powerON(power):
    sysTem = power.query('*IDN?')
    if 'DH1766A' in sysTem:
        power.write('OUTP ON\n')
    else:
        power.write('OUTP ON,(@1)\n')


class AdcTest:

    def __init__(self):
        self.__work_path = os.getcwd()
        config_file = open(os.path.join(self.__work_path, 'config.txt'), 'rb')
        content = config_file.read().decode('utf-8')
        self.address = content.split('\n')[0].strip().split(' ')[1]
        self.port = content.split('\n')[1].strip().split(' ')[1]
        self.baud = content.split('\n')[2].strip().split(' ')[1]
        # self.cmd = content.split('\n')[3].strip().split(' ')[1]
        self.cmd = '01 e0 fc 0b 04 01'

    def start_test(self):
        try:
            ser_com = serial.Serial(self.port, self.baud, timeout=5)
            rm = visa.ResourceManager()
            power = rm.open_resource(self.address, open_timeout=1000)
        except SerialException:
            print('串口无法打开！')
            time.sleep(2)
        except VisaIOError:
            print('电源无法打开！')
            time.sleep(2)
        else:
            broad = input('请输入测试板编号:')
            ser_com.write(bytes.fromhex(self.cmd))
            adc_set = 'SETP_ADC'

            volt_test = 0
            volt_end = 4.1
            row_write = 2
            col_write = 1
            if not os.path.exists(os.path.join(self.__work_path, '3437 ADC测试.xls')):
                print('新建xls')
                workbook = xlwt.Workbook()
                sheet1 = workbook.add_sheet('十进制')
                sheet2 = workbook.add_sheet('十六进制')
                sheet1.write(0, 2, broad)
                sheet1.write(1, 0, '电压')
                sheet1.write(1, 1, 'Min')
                sheet1.write(1, 2, 'Max')
                sheet1.write(1, 3, 'Avg')
                sheet2.write(0, 2, broad)
                sheet2.write(1, 0, '电压')
                sheet2.write(1, 1, 'Min')
                sheet2.write(1, 2, 'Max')
                sheet2.write(1, 3, 'Avg')
            else:
                xlsfile = xlrd.open_workbook_xls(os.path.join(self.__work_path, '3437 ADC测试.xls'), formatting_info=True)
                print('读取xls')
                col_write = xlsfile.sheets()[0].ncols   # xlrd打开的文件只能读取不能写入
                workbook = xlutils3.copy.copy(xlsfile)   # 写入需要copy一个文件副本，先写入到副本然后再保存到原文件
                sheet1 = workbook.get_sheet(0)
                sheet2 = workbook.get_sheet(1)
                sheet1.write(0, col_write+1, broad)
                sheet1.write(1, col_write, 'Min')
                sheet1.write(1, col_write+1, 'Max')
                sheet1.write(1, col_write+2, 'Avg')
                sheet2.write(0, col_write+1, broad)
                sheet2.write(1, col_write, 'Min')
                sheet2.write(1, col_write+1, 'Max')
                sheet2.write(1, col_write+2, 'Avg')

            powerSetVolt(power, volt_test)
            powerSetCurrent(power, 0.5)
            powerON(power)
            ser_com.reset_input_buffer()
            time.sleep(0.5)
            while volt_test < volt_end:
                play_count = 0
                data_list = []
                while play_count < 10:
                    if ser_com.inWaiting():
                        data1 = str(ser_com.readline().decode("utf-8"))
                        print(data1)
                        if adc_set in data1:
                            adc_data = data1.split('=')[1]
                            dec_data = adc_data.split(',')[0]
                            data_list.append(int(dec_data))
                            print(dec_data)
                            play_count += 1
                writeFile(sheet1, sheet2, volt_test, data_list, row_write, col_write)
                workbook.save('3437 ADC测试.xls')
                row_write += 1
                volt_test += 0.1
                powerSetVolt(power, volt_test)
                print('测试电压%f' % volt_test)
                time.sleep(1)
            # ser_com.colse()
            # power.colse()
            print('completed test')


if __name__ == "__main__":
    cycle = AdcTest()
    cycle.start_test()
