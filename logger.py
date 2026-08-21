# logger.py
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from datetime import datetime
import sys
import os

# Global logger instance
_logger = None

def get_logger(name="TelephoneScrapping", log_dir="Logs", log_level=logging.INFO):
    """
    Get or create a logger instance
    
    Args:
        name (str): Logger name
        log_dir (str): Directory to store log files
        log_level (int): Logging level
    
    Returns:
        logging.Logger: Configured logger instance
    """
    global _logger
    
    if _logger is not None:
        return _logger
    
    # Create log directory with date
    date_str = datetime.now().strftime("%Y-%m-%d")
    log_path = Path(log_dir) / date_str
    log_path.mkdir(parents=True, exist_ok=True)
    
    # Create log file with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_path / f"{name}_{timestamp}.log"
    
    # Create logger
    logger = logging.getLogger(name)
    logger.setLevel(log_level)
    logger.propagate = False
    
    # Clear any existing handlers
    if logger.handlers:
        logger.handlers.clear()
    
    # Formatter
    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(filename)s:%(lineno)d | %(funcName)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    
    # File handler with rotation
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=10 * 1024 * 1024,  # 10MB
        backupCount=10,
        encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    _logger = logger
    
    logger.info("=" * 70)
    logger.info(f"Logger initialized: {log_file}")
    logger.info("=" * 70)
    
    return logger

# Convenience functions
def debug(message):
    get_logger().debug(message)

def info(message):
    get_logger().info(message)

def warning(message):
    get_logger().warning(message)

def error(message):
    get_logger().error(message)

def critical(message):
    get_logger().critical(message)

def exception(message):
    get_logger().exception(message)

# Example usage
if __name__ == "__main__":
    logger = get_logger()
    logger.info("Test log message")
    logger.debug("Debug message")
    logger.warning("Warning message")