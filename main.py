import sys

import pandas as pd
import openpyxl as op
import xlsxwriter as xl
import os
import re
import argparse
from datetime import datetime
import logging

#加载配置文件
from config import VALID_RULES
# 解决Windows控制台输出中文乱码�������

#文件路径
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
IN_PATH = os.path.normpath(os.path.join(BASE_PATH, "./in"))
LOG_PATH = os.path.normpath(os.path.join(BASE_PATH, "./LOG"))

# 输出路径处理，如果影刀没传output，就用脚本同目录out文件
OUT_PATH = os.path.join(BASE_PATH, "out")
os.makedirs(OUT_PATH, exist_ok=True)

# 自动建文件夹
for p in [IN_PATH, OUT_PATH, LOG_PATH]:
    os.makedirs(p, exist_ok=True)

# ---------------- 接收命令行参数 ----------------
parser = argparse.ArgumentParser(description="订单数据校验脚本")
parser.add_argument("--input", required=True, help="待校验csv文件的完整绝对路径")
parser.add_argument("--output", help="输出异常结果csv完整路径(可选)")
args = parser.parse_args()

# 拿到影刀传过来的输入文件路径
input_csv_path = args.input
out_csv = args.output
if not out_csv:
 out_dir = os.path.dirname(out_csv)
# 判断是否传了输出参数


# 日志配置
logging.basicConfig(
    filename=os.path.join(LOG_PATH, f"run_{datetime.now().strftime('%Y%m%d')}.log"),
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8"
)
logger = logging.getLogger()

#手机号正则表达式
PHONE_REGEX = r'^1[3-9]\d{9}$'#^表示字符串的开始

#手机号校验
def check_phonenumber(phonenumber):
    if pd.isna(phonenumber):
        return False
    else:
        return bool(re.match(PHONE_REGEX,str(phonenumber).strip()))
#金额校验
def check_amount(amount):
    if pd.isna(amount) or amount < 0:
        return False
    elif amount >= 0:
        return True


#核心逻辑函数
def main():
    # 加载文件
    try:
        file_path = os.path.join(input_csv_path)
        if not os.path.exists(file_path):
            logger.error(f"文件不存在：{file_path}")
            sys.exit(1)  # 失败返回1，影刀识别出错

        else:
            #读取数据
            df = pd.read_csv(file_path, encoding="utf-8")
            print(df.iloc[0])
            logger.info(f"文件读取成功：{file_path}, 数据行数：{len(df)}")
            logger.info(df)

            #数据处理
            #处理空数据
            df = df.dropna(subset=['order_id','total_amount','discount_amount','shipping_fee','tax_amount'])
            logger.info(f"空数据处理完成，处理后数据行数：{len(df)}")

            #循环校验每一行的数据是否合法
            error_list = []
            for index,row in df.iterrows():
                error_res = []
                if VALID_RULES["phone_check"]:
                    if not check_phonenumber(row['phone']):
                        error_res.append("手机号校验失败")

                if VALID_RULES["jine_check"]:
                    if (not check_amount(row['total_amount'])
                            or not check_amount(row['discount_amount'])
                            or not check_amount(row['shipping_fee'])
                            or not check_amount(row['tax_amount'])):
                        error_res.append("金额校验失败")
                if error_res:
                    # 把当前一整行全部转为字典
                    row_dict = row.to_dict()
                    # 额外增加异常信息字段
                    row_dict["error_res"] = ";".join(error_res)
                    error_list.append(row_dict)

            #输出错误数据
            df_error = pd.DataFrame(error_list)

            # 输出汇总表、异常表
            summary_file = os.path.join(out_csv, f"summary_result.xlsx")
            error_file = os.path.join(out_csv, f"error_list_result.xlsx")
            # 汇总表（原始全量）
            with pd.ExcelWriter(summary_file, engine="xlsxwriter") as writer:
                df.to_excel(writer, sheet_name="订单汇总", index=False)
                # 异常sheet标红
                if len(df_error) > 0:
                    df_error.to_excel(writer, sheet_name="异常清单", index=False)
                    #获取异常清单这个sheet对象
                    ws = writer.sheets["异常清单"]
                    #设置异常清单这个sheet的E列（手机号列）的格式为红色
                    red_format = writer.book.add_format({"font_color": "red"})
                    ws.set_column("B:B", 20, red_format)
                    ws.set_column("P:P", 20, red_format)

                    logger.info(f"处理完成，总数据：{len(df)}，异常数量：{len(df_error)}")
                    print(f"完成！异常数量：{len(df_error)}，报表在out文件夹")
                    return sys.exit(0) # 正常结束

    except Exception as e:
        logger.error(f"程序报错，报错原因{str(e)}")



if __name__ == '__main__':
    main()



