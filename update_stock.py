import akshare as ak
import pandas as pd
import os
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from datetime import datetime
import warnings

warnings.filterwarnings("ignore")


# ==============================
# 交易日判断
# ==============================

def is_today_trading_day():
    """判断今天是否为A股交易日"""

    try:

        today = datetime.now().strftime("%Y%m%d")

        df = ak.tool_trade_date_hist_sina()

        df["trade_date"] = pd.to_datetime(df["trade_date"])
        df["trade_date"] = df["trade_date"].dt.strftime("%Y%m%d")

        return today in df["trade_date"].values

    except Exception as e:

        print(f"交易日判断失败：{e}")

        return True


# ==============================
# 配置
# ==============================

CONFIG = {

    "stocks": [

        {"name": "湖南白银", "code": "002716", "start_price": 6.92},
        {"name": "蓝色光标", "code": "300058", "start_price": 11.52},
        {"name": "利欧股份", "code": "002131", "start_price": 5.64},

    ],

    "excel_file": "stock_data.xlsx",

    "email": {

        "sender": os.getenv("EMAIL_SENDER"),
        "password": os.getenv("EMAIL_PASSWORD"),
        "receivers": os.getenv("EMAIL_RECEIVER", "").split(","),

        "smtp_server": "smtp.qq.com",
        "smtp_port": 465,

    }

}


# ==============================
# 获取股票数据（带重试）
# ==============================

def get_stock_data(code, retry=5):

    for i in range(retry):

        try:

            df = ak.stock_zh_a_hist(

                symbol=code,
                period="daily",
                start_date="20200101",
                adjust=""

            )

            close_price = df.iloc[-1]["收盘"]

            max_price = df["最高"].max()

            return close_price, max_price

        except Exception as e:

            print(f"{code} 获取失败，第{i+1}次尝试")

            if i < retry - 1:

                time.sleep(5)

            else:

                raise e


# ==============================
# 指标计算
# ==============================

def calculate_metrics(start_price, close_price, max_price):

    rise = (close_price / start_price - 1) * 100

    drawdown_max = (close_price / max_price - 1) * 100

    drawdown_start = (close_price / start_price - 1) * 100

    return rise, drawdown_max, drawdown_start


# ==============================
# 更新Excel
# ==============================

def update_excel():

    rows = []

    for idx, stock in enumerate(CONFIG["stocks"], start=1):

        print(f"获取 {stock['name']} 数据")

        try:

            close_price, max_price = get_stock_data(stock["code"])

        except Exception as e:

            print(f"{stock['name']} 数据获取失败：{e}")

            continue

        rise, drawdown_max, drawdown_start = calculate_metrics(

            stock["start_price"],
            close_price,
            max_price

        )

        rows.append({

            "序号": idx,
            "股票名称": stock["name"],
            "股票代码": stock["code"],
            "起始日价": stock["start_price"],
            "起始日起最高价（自动更新）": round(max_price, 2),
            "每日收盘价（自动更新）": round(close_price, 2),
            "涨幅": f"{rise:.2f}%",
            "最高价回撤": f"{drawdown_max:.2f}%",
            "起始价回撤": f"{drawdown_start:.2f}%"

        })

    df = pd.DataFrame(rows)

    df.to_excel(CONFIG["excel_file"], index=False)

    print("Excel生成成功")

    return CONFIG["excel_file"]


# ==============================
# 发送邮件
# ==============================

def send_email_with_attachment(file_path):

    sender = CONFIG["email"]["sender"]
    password = CONFIG["email"]["password"]
    receivers = CONFIG["email"]["receivers"]

    if not sender or not password or receivers == [""]:

        print("邮箱配置缺失，跳过邮件发送")

        return

    msg = MIMEMultipart()

    msg["Subject"] = f"股票数据更新 {datetime.now().strftime('%Y-%m-%d')}"
    msg["From"] = sender
    msg["To"] = ",".join(receivers)

    msg.attach(MIMEText("今日股票数据已更新，请查看附件。"))

    with open(file_path, "rb") as f:

        part = MIMEApplication(f.read())

    part.add_header(

        "Content-Disposition",
        "attachment",
        filename=os.path.basename(file_path)

    )

    msg.attach(part)

    with smtplib.SMTP_SSL(

            CONFIG["email"]["smtp_server"],
            CONFIG["email"]["smtp_port"]

    ) as server:

        server.login(sender, password)

        server.sendmail(sender, receivers, msg.as_string())

    print("邮件发送成功")


# ==============================
# 主函数
# ==============================

def main():

    print("开始更新股票数据...")

    if not is_today_trading_day():

        print("今天不是A股交易日，跳过更新")

        return

    excel_file = update_excel()

    send_email_with_attachment(excel_file)

    print("任务完成")


# ==============================

if __name__ == "__main__":

    main()
