import schedule
import time
import threading

class TaskScheduler:
    """
    定时任务调度器
    使用 schedule 库实现每日指定时刻自动执行任务
    定时时间：12:35 和 15:35
    """

    def __init__(self, execute_callback, logger):
        self.execute_callback = execute_callback
        self.logger = logger
        self.is_running = False
        self.scheduler_thread = None
        self.stop_event = threading.Event()
        self.schedule_times = ["12:35", "15:35"]

    def setup_schedule(self, times=None):
        """清空旧的定时任务，重新设置"""
        if times is not None:
            self.schedule_times = times
        schedule.clear()
        for t in self.schedule_times:
            schedule.every().day.at(t).do(self._scheduled_execute)
            self.logger.info(f"[Scheduler] Registered: daily at {t}")

    def _scheduled_execute(self):
        """定时触发时的回调函数"""
        self.logger.info("[Scheduler] === Scheduled task triggered ===")
        try:
            self.execute_callback()
        except Exception as e:
            self.logger.error(f"[Scheduler] Task exception: {str(e)}")

    def start(self, times=None):
        """启动调度器"""
        if self.is_running:
            return False
        self.setup_schedule(times)
        self.is_running = True
        self.stop_event.clear()
        self.scheduler_thread = threading.Thread(target=self._run_scheduler, daemon=True)
        self.scheduler_thread.start()
        self.logger.info("[Scheduler] Scheduler started")
        return True

    def _run_scheduler(self):
        """调度器主循环"""
        while self.is_running and not self.stop_event.is_set():
            schedule.run_pending()
            time.sleep(1)

    def stop(self):
        """停止调度器"""
        if not self.is_running:
            return False
        self.is_running = False
        self.stop_event.set()
        schedule.clear()
        if self.scheduler_thread:
            self.scheduler_thread.join(timeout=3)
        self.logger.info("[Scheduler] Scheduler stopped")
        return True
