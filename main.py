# -i http://pypi.douban.com/simple/ --trusted-host pypi.douban.com
import time
from os import path
import pandas as pd
from tkinter import *
from tkinter.filedialog import askopenfilename

# select file
root = Tk()  # 创建一个Tkinter.Tk()实例
root.withdraw()

default_dir = __file__
file_path = askopenfilename(title=u'选择文件', initialdir=(path.abspath(path.dirname(default_dir))))

# read file
df_raw = pd.read_excel(file_path)

# check format
valid_format = {"类型", "名称", "业务单号", "支付流水号", "关联单号", "交易来源地", "账务主体", '账户', "收入(元)", "支出(元)", "余额(元)", "支付方式", "交易对手",
                "渠道", "下单时间", "入账时间", "操作人", "附加信息", "备注"}
if set(df_raw.columns.values) != valid_format:
    print("文件格式不正确:")
    print(df_raw.columns.values)
    time.sleep(5)
    sys.exit()

df_result = pd.DataFrame(
    {'下单时间': [], '入账时间': [], '业务单号': [], '商品名称': [], '订单金额': [], '手续费': [], '退款金额': [], '支付流水号': []})

df_income = df_raw[df_raw['类型'] == '订单入账'].reset_index(drop=False)
df_refund = df_raw[df_raw['类型'] == '退款'].reset_index(drop=False)
df_transaction_fee = df_raw[df_raw['类型'] == '交易手续费'].reset_index(drop=False)


def get_fee(order_id):
    fee = -1
    x = df_transaction_fee[df_transaction_fee['业务单号'] == order_id]['支出(元)']
    if len(x.values) == 1:
        fee = x.values[0]
    return fee


def get_refund_and_payment_id(order_id):
    refund = -1
    payment_id = '-'
    x = df_refund[df_refund['业务单号'] == order_id]['支出(元)']
    if len(x.values) == 1:
        refund = x.values[0]
        payment_id = df_refund[df_refund['业务单号'] == order_id]['支付流水号'].values[0]
    return [refund, payment_id]


def row_transfer(row):
    transaction_time = row['下单时间']
    accounting_time = row['入账时间']
    order_id = row['业务单号']
    item_name = row['名称']
    order_amount = row['收入(元)']

    fee = get_fee(order_id)

    x = get_refund_and_payment_id(order_id)
    refund = x[0]
    refund_payment_id = x[1]

    new_row = [transaction_time, accounting_time, order_id, item_name, order_amount, fee, refund, refund_payment_id]
    df_result.loc[row['index']] = new_row


def save_result(result):
    print('--- 正在保存文件... ---')

    # reset index and encode, then output result
    result.reset_index(drop=False)
    result.to_csv('对账单.csv', encoding='utf_8_sig')


def save_missing_refunds(result):
    x = result[result['退款金额'] > -1]['业务单号'].values
    y = df_refund['业务单号']
    missing_records = [v for v in y if v not in x]
    print('missing_refunds', missing_records)
    pd.DataFrame(missing_records).to_csv('未找到匹配订单的退款.csv')


def save_missing_fees(result):
    print("--- 检查数据完整性... ---")
    x = result[result['手续费'] > -1]['业务单号'].values
    y = df_transaction_fee['业务单号']
    missing_records = [v for v in y if v not in x]
    print('missing_fees', missing_records)
    pd.DataFrame(missing_records).to_csv('未找到匹配订单的手续费.csv')


def initiate():
    df_income.apply(row_transfer, axis=1)
    save_result(df_result)
    save_missing_refunds(df_result)
    print(df_result.loc[0])


if __name__ == '__main__':
    print('--- 正在处理文件... ---')
    print(default_dir)
    initiate()
    print('--- 完成 ---')
