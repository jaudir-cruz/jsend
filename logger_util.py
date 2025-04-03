# logger_util.py
import logging
import queue
from config import SESSION_ID

class LoggerUtil:
    def __init__(self):
        self.log_queue = queue.Queue()
        self.logger = self.setup_logging()

    def setup_logging(self):
        log_filename = f"log_execucao_{SESSION_ID}.txt"
        logging.basicConfig(
            filename=log_filename,
            level=logging.DEBUG,
            format="%(asctime)s - %(levelname)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        return logging.getLogger()

    def log_message(self, message, level=logging.INFO):
        self.logger.log(level, message)
        self.log_queue.put(message)

    def get_log_queue(self):
        return self.log_queue

    def get_log_filename(self):
        return f"log_execucao_{SESSION_ID}.txt"
