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

# 新增：交易日检查函数
def is_today_trading_day() -> bool:
    """检查今天（北京时间）是否为A股交易日"""
    try:
        today = datetime.now().strftime("%Y%m%d")
        df = ak.tool_trade_date_hist_sina()
        # 统一转换为字符串格式 YYYYMMDD
        if pd.api.types.is_datetime64_any_dtype(df['trade_date']):
            df['trade_date_str'] = df['trade_date'].dt.strftime("%Y%m%d")
        else:
            df['trade_date_str'] = df['trade_date'].str.replace('-', '')
        return today in df['trade_date_str'].values
    except Exception as e:
        print(f"交易日检查失败: {e}，默认视为交易日")
        return True

# 配置信息
CONFIG = {
    "stocks": [
        {"name": "湖南白银", "code": "002716", "start_price": 6.92},
        {"name": "蓝色光标", "code": "300058", "start_price": 11.52},
        {"name": "利欧股份", "code": "002131", "start_price": 5.64}
    ],
    "excel_file": "stock_data.xlsx",
    "email": {
        "sender": os.getenv("EMAIL_SENDER"),
        "password": os.getenv("EMAIL_PASSWORD"),
        "receivers": os.getenv("EMAIL_RECEIVER", "").split(","),
        "smtp_server": "smtp.qq.com",
        "smtp_port": 465
    }
}

# ... （中间的 get_stock_data, calculate_metrics, update_excel, send_email_with_attachment 函数保持不变）

def main():
    """主函数"""
    print("开始更新股票数据...")
    
    # 新增：先检查交易日
    if not is_today_trading_day():
        print("今天不是A股交易日，跳过数据更新。")
        return
    
    # 更新Excel
    excel_file = update_excel()
    # 发送邮件
    send_email_with_attachment(excel_file)
    print("任务完成！")

if __name__ == "__main__":
    main()
