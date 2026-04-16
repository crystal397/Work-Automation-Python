from apscheduler.schedulers.blocking import BlockingScheduler
from collector import run_bulk_collection
from datetime import datetime, timedelta


def daily_update():
    """전날(하루치) 데이터만 수집하여 DB 갱신"""
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")
    print(f"[{datetime.now()}] 일배치 실행: {yesterday}")
    run_bulk_collection(start_date=yesterday, end_date=yesterday)


scheduler = BlockingScheduler()
scheduler.add_job(daily_update, "cron", hour=6, minute=0)  # 매일 오전 6시
scheduler.start()