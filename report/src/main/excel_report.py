#%%
import sys

sys.path.insert(0, r'C:\Users\Babatunde\Desktop\Digital_Skola_File\Homework\project_1\project_1\report\src\operators')
sys.path.insert(0, r'C:\Users\Babatunde\Desktop\Digital_Skola_File\Homework\project_1\project_1\report\src\utils\discord')


from xlsx_report_plugin import ExcelReportPlugin
from discord_webhook import send_to_discord

import os
import json
# import json

base_path = os.sep.join(os.getcwd().split(os.sep)[:-3])
print(f'base path: {base_path}')

input_data = base_path + '/input_data/supermarket_sales.xlsx'
output_data = base_path + '/output_data/daily_report_3.xlsx'

configs = open(base_path + '/configs/webhook.json')
webhook_url = json.load(configs)['webhook_url']

automate = ExcelReportPlugin(
    input_data=input_data,
    output_data=output_data
)

if __name__ == "__main__":
    automate.main()
    send_to_discord(webhook_url, output_data)

# %%
