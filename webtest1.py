import xlwt
import time


class inforsec:

    def __init__(self):
        self.vs_list = []
        self.vs_current_connection = {}
        self.vs_total_connection = {}
        self.vs_total_bytes_in = {}
        self.vs_total_bytes_out = {}
        self.real_server_list = []
        self.real_server_current_connection = {}
        self.real_server_total_connection = {}
        self.real_server_total_bytes_in = {}
        self.real_server_total_bytes_out = {}
        self.real_server_average_response_time = {}


    def get_realservice_status(self):
        with open("E:\python\dmzssl.txt") as log:
            f = log.readlines()
        for index in range(len(f)):
            if f[index].find("Real service") == -1:
                continue
            else:
                self.real_server_current_connection[f[index].rstrip().replace("Real service ", "").replace("UP ACTIVE", "")] = f[index+3].split(":")[1].strip()
                self.real_server_total_bytes_in[f[index].rstrip().replace("Real service ", "").replace("UP ACTIVE", "")] = f[index+6].split(":")[1].strip()
                self.real_server_total_bytes_out[f[index].rstrip().replace("Real service ", "").replace("UP ACTIVE", "")] = f[index+7].split(":")[1].strip()
                self.real_server_list.append(f[index].rstrip().replace("Real service ", "").replace("UP ACTIVE", ""))
                self.real_server_average_response_time[f[index].rstrip().replace("Real service ", "").replace("UP ACTIVE", "")] = f[index+12].split(":")[1].strip().replace(" ms", "")

    def get_virtualservice_status(self):
        with open("E:\python\dmzssl.txt") as log:
            f = log.readlines()
        for index in range(len(f)):
            if str(f[index]).find("tcp virtual service"):
                continue
            else:
                self.vs_current_connection[f[index].split('"')[1]] = f[index+2].split(":")[1].strip()
                self.vs_total_connection[f[index].split('"')[1]] = f[index+3].split(":")[1].strip()
                self.vs_total_bytes_in[f[index].split('"')[1]] = f[index + 4].split(":")[1].strip()
                self.vs_total_bytes_out[f[index].split('"')[1]] = f[index + 5].split(":")[1].strip()
                self.vs_list.append(f[index].split('"')[1])
        for index in range(len(f)):
            if str(f[index]).find("http virtual service"):
                continue
            else:
                self.vs_current_connection[f[index].split('"')[1]] = f[index+2].split(":")[1].strip()
                self.vs_total_connection[f[index].split('"')[1]] = f[index+3].split(":")[1].strip()
                self.vs_total_bytes_in[f[index].split('"')[1]] = f[index + 4].split(":")[1].strip()
                self.vs_total_bytes_out[f[index].split('"')[1]] = f[index + 5].split(":")[1].strip()
                self.vs_list.append(f[index].split('"')[1])
        for index in range(len(f)):
            if str(f[index]).find("https virtual service"):
                continue
            else:
                self.vs_current_connection[f[index].split('"')[1]] = f[index+2].split(":")[1].strip()
                self.vs_total_connection[f[index].split('"')[1]] = f[index+3].split(":")[1].strip()
                self.vs_total_bytes_in[f[index].split('"')[1]] = f[index + 4].split(":")[1].strip()
                self.vs_total_bytes_out[f[index].split('"')[1]] = f[index + 5].split(":")[1].strip()
                self.vs_list.append(f[index].split('"')[1])
        for index in range(len(f)):
            if str(f[index]).find("tcps virtual service"):
                continue
            else:
                self.vs_current_connection[f[index].split('"')[1]] = f[index+2].split(":")[1].strip()
                self.vs_total_connection[f[index].split('"')[1]] = f[index+3].split(":")[1].strip()
                self.vs_total_bytes_in[f[index].split('"')[1]] = f[index + 4].split(":")[1].strip()
                self.vs_total_bytes_out[f[index].split('"')[1]] = f[index + 5].split(":")[1].strip()
                self.vs_list.append(f[index].split('"')[1])

    def real_server_sorted(self):
        self.real_server_current_connection_sorted = sorted(self.real_server_current_connection.items(),key = lambda item:int(item[1]),reverse=True)
        self.real_server_average_response_time_sorted = sorted(self.real_server_average_response_time.items(),key = lambda item:float(item[1]),reverse=True)
        self.real_server_total_bytes_in_sorted = sorted(self.real_server_total_bytes_in.items(),key = lambda item:int(item[1]),reverse=True)
        self.real_server_total_bytes_out_sorted = sorted(self.real_server_total_bytes_out.items(),key = lambda item:int(item[1]),reverse=True)
        return

    def vs_sorted(self):
        self.vs_current_connection_sorted = sorted(self.vs_current_connection.items(),key=lambda item:int(item[1]),reverse=True)
        self.vs_total_connection_sorted = sorted(self.vs_total_connection.items(),key=lambda  item:int(item[1]),reverse=True)
        self.vs_total_bytes_in_sorted = sorted(self.vs_total_bytes_in.items(),key=lambda  item:int(item[1]),reverse=True)
        self.vs_total_bytes_out_sorted = sorted(self.vs_total_bytes_out.items(),key=lambda  item:int(item[1]),reverse=True)
        print(len(self.vs_list))
        return

    def write_excel(self):
        book = xlwt.Workbook(encoding='utf-8', style_compression=0)
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.bold = True
        style.font = font
        datetime = time.strftime('%Y-%m-%d',time.localtime(time.time()))
        print(datetime)
        sheet1 = book.add_sheet('真实服务当前连接数', cell_overwrite_ok=True)
        sheet2 = book.add_sheet('真实服务平均响应时间', cell_overwrite_ok=True)
        sheet3 = book.add_sheet('真实服务总入向流量', cell_overwrite_ok=True)
        sheet4 = book.add_sheet('真实服务总出向流量', cell_overwrite_ok=True)
        sheet5 = book.add_sheet('虚服务当前连接', cell_overwrite_ok=True)
        sheet6 = book.add_sheet('虚服务总连接', cell_overwrite_ok=True)
        sheet7 = book.add_sheet('虚服务总入向流量', cell_overwrite_ok=True)
        sheet8 = book.add_sheet('虚服务总出向流量', cell_overwrite_ok=True)
        for index in range(len(self.real_server_list)):
            sheet1.write(0, 0, "真实服务名称", style)
            sheet1.write(0, 1, "当前连接数", style)
            sheet1.write(index+1, 0, str(self.real_server_current_connection_sorted[index][0]))
            sheet1.write(index+1, 1, self.real_server_current_connection_sorted[index][1])
        for index in range(len(self.real_server_list)):
            sheet2.write(0, 0, "真实服务名称", style)
            sheet2.write(0, 1, "当前平均响应时间(s)", style)
            sheet2.write(index+1, 0, str(self.real_server_average_response_time_sorted[index][0]))
            sheet2.write(index+1, 1, self.real_server_average_response_time_sorted[index][1])
        for index in range(len(self.real_server_list)):
            sheet3.write(0, 0, "真实服务名称", style)
            sheet3.write(0, 1, "总入向流量(Byte)", style)
            sheet3.write(index+1, 0, str(self.real_server_total_bytes_in_sorted[index][0]))
            sheet3.write(index+1, 1, self.real_server_total_bytes_in_sorted[index][1])
        for index in range(len(self.real_server_list)):
            sheet4.write(0, 0, "真实服务名称", style)
            sheet4.write(0, 1, "总出向流量(Byte)", style)
            sheet4.write(index+1, 0, str(self.real_server_total_bytes_out_sorted[index][0]))
            sheet4.write(index+1, 1, self.real_server_total_bytes_out_sorted[index][1])
        for index in range(len(self.vs_list)):
            sheet5.write(0, 0, "虚拟服务名称", style)
            sheet5.write(0, 1, "当前连接数", style)
            sheet5.write(index+1, 0, str(self.vs_current_connection_sorted[index][0]))
            sheet5.write(index+1, 1, self.vs_current_connection_sorted[index][1])
        for index in range(len(self.vs_list)):
            sheet6.write(0, 0, "虚拟服务名称", style)
            sheet6.write(0, 1, "总连接数", style)
            sheet6.write(index+1, 0, str(self.vs_total_connection_sorted[index][0]))
            sheet6.write(index+1, 1, self.vs_total_connection_sorted[index][1])
        for index in range(len(self.vs_list)):
            sheet7.write(0, 0, "虚拟服务名称", style)
            sheet7.write(0, 1, "总入向流量(Byte)", style)
            sheet7.write(index+1, 0, str(self.vs_total_bytes_in_sorted[index][0]))
            sheet7.write(index+1, 1, self.vs_total_bytes_in_sorted[index][1])
        for index in range(len(self.vs_list)):
            sheet8.write(0, 0, "虚拟服务名称", style)
            sheet8.write(0, 1, "总出向流量(Byte)", style)
            sheet8.write(index+1, 0, str(self.vs_total_bytes_out_sorted[index][0]))
            sheet8.write(index+1, 1, self.vs_total_bytes_out_sorted[index][1])
        book.save(str(datetime)+".xls")
        return

    def main(self):
        #i.modity_config()
        i.get_realservice_status()
        i.get_virtualservice_status()
        i.real_server_sorted()
        i.vs_sorted()
        i.write_excel()
        return


i = inforsec()
i.main()
