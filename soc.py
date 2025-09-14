from DrissionPage import ChromiumPage,ChromiumOptions
import re
import json
import random
from urllib.parse import urljoin
import requests
from openpyxl import Workbook
import tkinter as tk
import time
from tkinter import messagebox
import matplotlib.pyplot as plt


class SingleDataInput:
    def __init__(self, root):
        self.root = root
        self.root.title("单个数据输入")
        self.root.geometry("300x150")
        self.data = None
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="请输入数据:", font=("SimHei", 10)).pack(pady=10)
        self.data_entry = tk.Entry(self.root, width=30, font=("SimHei", 10))
        self.data_entry.pack(pady=5)

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="保存", command=self.save_data, width=8).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="查看", command=self.show_data, width=8).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="清空", command=self.clear_data, width=8).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="退出", command=self.root.quit, width=8).pack(side=tk.LEFT, padx=5)

    def save_data(self):
        input_value = self.data_entry.get().strip()

        if not input_value:
            messagebox.showwarning("提示", "输入不能为空!")
            return

        self.data = input_value
        messagebox.showinfo("成功", f"数据已保存:\n{input_value}")

    def show_data(self):
        if self.data is None:
            messagebox.showinfo("提示", "暂无保存的数据")
        else:
            messagebox.showinfo("已保存数据", self.data)

    def clear_data(self):
        self.data_entry.delete(0, tk.END)
        self.data = None
        messagebox.showinfo("提示", "已清空输入")
def remove_illegal_characters(value):
    if isinstance(value, str):
        return re.sub(r'[\x00-\x1F]+', '', value)
    return value

cp=ChromiumPage()
wb = Workbook()
ws = wb.active
'''
cp.get('https://login.taobao.com/member/login.jhtml')
cp.ele('css:#fm_login_id').input('18921231168')
cp.ele('css:#fm_login_password').input('wzf060623')
cp.ele('css:.fm_button.fm_submit.password_login.button_low_light').click()
'''
base_url = "https://detail.tmall.com/item.htm"
param_prefix = "?ali_trackid=2%3Amm_5539777028_3063500466_115725900203%3A1757479001265_554188484_0&bxsign=tbkBld-mKOgvg-EX-HRgQMVxB5EaUJ_FL_bU6HR7VZpFvXnwI6wAzvP0344a0nelX9L9P9cVzNO_yoMXOsPaxfpyHwgqjMj6mImVrF0xTbjaMuAo3QuvycwfA_eWm1Shvx4ymTVKvrCzRMJZKDIQhOwxWIORF457PMIZWht_k2b6jYr6JJh1Y2ab--iptszOx52&id="
param_suffix = "&rootPageId=20150318020018244&scm=20140767.59990_11_81_18_11_468_1757477773103.7%7Citem%7C0.0&spm=a2e1u.27659560.d166185234306.10&u_channel=bybtqdyh&umpChannel=bybtqdyh&union_lens=lensId%3AOPT%401710136889%400b87bd25_0d45_18e2c1a0fce_934a%40026RBOCx6kDH2LDr9rdvBMoD%40eyJmbG9vcklkIjo4NzgwMSwiic3JjRmxvb3JJZCI6IjM4OTU3In0ie%3Bscm%3A1007.15348.109552.0_87801_2a7fa964-7fbf-4e65-b248-1d91aef5d888%3Brecoveryid%3A557196790_0%401757477770063%3Bprepvid%3A201_33.44.149.39_4562884_1757477770379&skuId=5138049605923"

print('输入商品id')
root = tk.Tk()
root.option_add("*Font", "SimHei 10")
app = SingleDataInput(root)
root.mainloop()
a=app.data
full_params = param_prefix + a + param_suffix

detail_url = urljoin(base_url, full_params)
cp.get(detail_url)
try:
    target_element = cp.ele('css:span[class^="text-"]', timeout=10)
    value = target_element.text
    print(f"提取到的价格：{value}")

except Exception as e:
    print(f"提取失败：{e}")
try:
    img_ele = cp.ele('css:	img[src$="item_pic.jpg_.webp"]')
    src_url = img_ele.attr("src")
    print(f"提取到的 src 网址：{src_url}")
except Exception as e:
    print(f"提取失败：{e}")
url = src_url
r = requests.get(url, stream=True)
with open('aaa_soc.webp', 'wb') as f:
    for chunk in r.iter_content(chunk_size=8192):
        if chunk:
            f.write(chunk)
cp.listen.start('h5/mtop.taobao.rate.detaillist.get/6.0')
cp.ele('css:div[class^="ShowButton--"]').click()
time.sleep(random.randint(1, 2))
aaa=[]
try:
    inquery = cp.listen.wait(count=10,timeout=10,fit_count=False)
    print(inquery)
    if  inquery==[] :
        print(f"未收到评论数据的响应")
    for item in inquery:
        enquery = item.response.body
        detail_pattern = re.compile(r'mtopjsonppcdetail\d+\((.*?)\)', re.DOTALL)
        detail_data = json.loads(re.findall(detail_pattern, enquery)[0])
        aaa.append(detail_data)
except (Exception, json.decoder.JSONDecodeError) as e:
    print(json.loads(re.findall(detail_pattern, enquery)))
    print(f"处理评论数据时出错: {e}")
good_list=[]
middle_list=[]
bad_list=[]
try:
    for ii in aaa:
        comment_list = ii['data']['rateList']
        for item in comment_list:
            detail_comment = remove_illegal_characters(item['feedback'])
            score=remove_illegal_characters(item["extraInfoMap"]["userGrade"])
            score=int(score)
            if score ==5:
                good_list.append(detail_comment)
            elif 1<score and score<5:
                middle_list.append(detail_comment)
            else:
                bad_list.append(detail_comment)
except KeyError as e:
    print(f"解析评论列表时出错: {e}")
except Exception as e:
    print(f"未知错误: {e}")
sizes = [len(good_list),len(middle_list),len(bad_list)]
labels = ['好评','中评','差评']
valid_data = [(label, size) for label, size in zip(labels, sizes) if size > 0]
if not valid_data:
    print("所有列表的长度均为0，无法生成饼图")
plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
fig, ax = plt.subplots(figsize=(8, 6))
wedges, texts, autotexts = ax.pie(
    sizes,
    labels=labels,
    autopct='%1.1f%%',
    startangle=90,
    wedgeprops=dict(width=0.3)
)
plt.setp(texts, size=10)
plt.setp(autotexts, size=8, color="black", weight="bold")
ax.set_title('各列表长度占比分布')
ax.axis('equal')
plt.tight_layout()
plt.show()
all_list=[good_list,middle_list,bad_list]
for item in all_list:
    ws.append([remove_illegal_characters(cell) for cell in item])
try:
    wb.save(r'淘宝评论soc.xlsx')
    print("数据成功保存到 '淘宝评论soc.xlsx'")
except Exception as e:
    print(f"保存 Excel 文件时出错: {e}")


