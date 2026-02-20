import logging
import os
import sys
from logging.handlers import RotatingFileHandler

# Create Log Directory if it doesn't exist
LOG_DIR = 'logs'
os.makedirs(LOG_DIR, exist_ok=True)

# Create Logger
logger = logging.getLogger('AMS_Orders_Logger')
logger.setLevel(logging.DEBUG)

# Rotating File Handler (prevents unlimited growth)
log_file = os.path.join(LOG_DIR, 'ams_orders.log')
file_handler = RotatingFileHandler(
    log_file,
    maxBytes=10*1024*1024,  # 10 MB per file
    backupCount=5,  # Keep 5 backup files (ams_orders.log.1, .2, .3, .4, .5)
    encoding='utf-8'
)
file_handler.setLevel(logging.DEBUG)

# Prefer the real stderr, fall back to the original stderr or devnull
stderr_stream = getattr(sys, 'stderr', None) or getattr(sys, '__stderr__', None)
if stderr_stream is None:
    # Open devnull so handler always has a writeable stream in GUI/exe builds
    stderr_stream = open(os.devnull, 'w', encoding='utf-8')

console_handler = logging.StreamHandler(stderr_stream)
console_handler.setLevel(logging.INFO)

# Formatter
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Handlers + Logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)