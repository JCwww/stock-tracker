import akshare as ak
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# 配置信息（建议通过环境变量设置，避免硬编码）
CONFIG = {
    "stocks": [
        {"name": "湖南白银", "code": "002716", "start_price": 6.92},
        {"name": "蓝色光标", "code": "300058", "start_price": 11.52},
        {"name": "利欧股份", "code": "002131", "start_price": 5.64}
        # 后期增加股票直接在这里添加即可
    ],
    "excel_file": "stock_data.xlsx",
    "email": {
        "sender": os.getenv("EMAIL_SENDER"),  # 发件人邮箱（从环境变量获取）
        "password": os.getenv("EMAIL_PASSWORD"),  # 邮箱授权码
        "receiver": os.getenv("EMAIL_RECEIVER"),  # 收件人邮箱
        "smtp_server": "smtp.qq.com",  # QQ邮箱示例，其他邮箱需修改
        "smtp_port": 465
    }
}

def get_stock_data(stock_code):
    """获取股票最新数据"""
    try:
        # 获取股票历史行情（前复权）
        stock_zh_a_hist_df = ak.stock_zh_a_hist(
            symbol=stock_code,
            period="daily",
            start_date="20230101",
            end_date=datetime.now().strftime("%Y%m%d"),
            adjust="qfq"
        )
        
        if stock_zh_a_hist_df.empty:
            return None
        
        # 获取最新收盘价
        latest_close = stock_zh_a_hist_df.iloc[-1]["收盘"]
        # 获取从起始日至今的最高价
        highest_price = stock_zh_a_hist_df["收盘"].max()
        
        return {
            "latest_close": round(latest_close, 2),
            "highest_price": round(highest_price, 2),
            "update_date": datetime.now().strftime("%Y-%m-%d")
        }
    except Exception as e:
        print(f"获取股票{stock_code}数据失败: {e}")
        return None

def calculate_metrics(start_price, latest_close, highest_price):
    """计算涨幅、回撤等指标"""
    # 涨幅 = (最新收盘价 - 起始价) / 起始价 * 100%
    increase_rate = round((latest_close - start_price) / start_price * 100, 2)
    # 最高价回撤 = (最新收盘价 - 最高价) / 最高价 * 100%
    high_retracement = round((latest_close - highest_price) / highest_price * 100, 2)
    # 起始价回撤 = (最新收盘价 - 起始价) / 起始价 * 100%
    start_retracement = round((latest_close - start_price) / start_price * 100, 2)
    
    return {
        "increase_rate": f"{increase_rate}%",
        "high_retracement": f"{high_retracement}%",
        "start_retracement": f"{start_retracement}%"
    }

def update_excel():
    """更新Excel表格"""
    # 检查文件是否存在，不存在则创建
    if not os.path.exists(CONFIG["excel_file"]):
        # 创建初始DataFrame
        df = pd.DataFrame({
            "序号": [],
            "股票名称": [],
            "股票代码": [],
            "起始日价": [],
            "起始日起最高价（自动更新）": [],
            "每日收盘价（自动更新）": [],
            "涨幅": [],
            "最高价回撤": [],
            "起始价回撤": []
        })
    else:
        df = pd.read_excel(CONFIG["excel_file"])
    
    # 遍历股票列表更新数据
    new_data = []
    for idx, stock in enumerate(CONFIG["stocks"], 1):
        stock_data = get_stock_data(stock["code"])
        if not stock_data:
            continue
        
        # 计算指标
        metrics = calculate_metrics(
            stock["start_price"],
            stock_data["latest_close"],
            stock_data["highest_price"]
        )
        
        # 构建行数据
        row = {
            "序号": idx,
            "股票名称": stock["name"],
            "股票代码": stock["code"],
            "起始日价": stock["start_price"],
            "起始日起最高价（自动更新）": stock_data["highest_price"],
            "每日收盘价（自动更新）": stock_data["latest_close"],
            "涨幅": metrics["increase_rate"],
            "最高价回撤": metrics["high_retracement"],
            "起始价回撤": metrics["start_retracement"]
        }
        new_data.append(row)
    
    # 更新DataFrame
    new_df = pd.DataFrame(new_data)
    # 保存到Excel
    new_df.to_excel(CONFIG["excel_file"], index=False)
    print(f"Excel表格更新完成，更新时间：{datetime.now()}")
    return CONFIG["excel_file"]

def send_email_with_attachment(file_path):
    """发送带附件的邮件"""
    if not os.path.exists(file_path):
        print("附件文件不存在，发送失败")
        return False
    
    # 构建邮件
    msg = MIMEMultipart()
    msg["From"] = CONFIG["email"]["sender"]
    msg["To"] = CONFIG["email"]["receiver"]
    msg["Subject"] = f"股票数据每日更新 - {datetime.now().strftime('%Y-%m-%d')}"
    
    # 邮件正文
    body = f"""
    您好！
    附件是{datetime.now().strftime('%Y-%m-%d')}最新的股票数据表格，请注意查收。
    
    自动发送，无需回复。
    """
    msg.attach(MIMEText(body, "plain", "utf-8"))
    
    # 添加附件
    with open(file_path, "rb") as f:
        part = MIMEApplication(f.read())
        part.add_header("Content-Disposition", "attachment", filename=os.path.basename(file_path))
        msg.attach(part)
    
    # 发送邮件
    try:
        server = smtplib.SMTP_SSL(CONFIG["email"]["smtp_server"], CONFIG["email"]["smtp_port"])
        server.login(CONFIG["email"]["sender"], CONFIG["email"]["password"])
        server.sendmail(CONFIG["email"]["sender"], CONFIG["email"]["receiver"], msg.as_string())
        server.quit()
        print("邮件发送成功")
        return True
    except Exception as e:
        print(f"邮件发送失败: {e}")
        return False

def main():
    """主函数"""
    print("开始更新股票数据...")
    # 更新Excel
    excel_file = update_excel()
    # 发送邮件
    send_email_with_attachment(excel_file)
    print("任务完成！")

if __name__ == "__main__":
    main()