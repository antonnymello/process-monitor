from asyncio.windows_events import NULL
import os
from time import sleep
import psutil
from openpyxl import load_workbook
from datetime import datetime

process_to_monitor = str(input("Enter the process name: "))

if len(process_to_monitor) == 0:
    raise Exception("Process name are required")

process_list = []

for proc in psutil.process_iter():
    try:
        process_name = proc.name()
        process_pid = proc.pid

        if process_name == process_to_monitor:
            process_list.append(proc)
            print('Process found!:', process_name, ' ::: ', process_pid)
            pass

    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        pass

if(len(process_list) == 0):
    raise Exception("Process not found")


def get_memory_usage_by_pid(pid):
    process = psutil.Process(pid)
    memory_usage = process.memory_info().rss
    return float(memory_usage) / (1024 * 1024)


def append_memory_usage_to_excel(pid, name, memory_usage, now):
    wb = load_workbook('report.xlsx')
    ws = wb.active
    ws.append([pid, name, memory_usage, now])
    wb.save('report.xlsx')


while True:
    for process in process_list:
        process_pid = process.pid
        now = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        memory_usage_in_mb = get_memory_usage_by_pid(process.pid)

        append_memory_usage_to_excel(
            process_pid, process_to_monitor, memory_usage_in_mb, now)

        print('Memory usage of process:', process_to_monitor, 'of PID:',
              process_pid, 'is:', memory_usage_in_mb, 'MB')

    sleep(60)

# print(get_memory_usage_by_pid(process_list.pid))
