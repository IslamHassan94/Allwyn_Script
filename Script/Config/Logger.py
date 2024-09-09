import os
import sys

sys.path.append("../..")
import logging
import logging.config

LOGGING_DIR = '../../Files/Logs/'
# Add this block to ensure the directory exists
if not os.path.exists(LOGGING_DIR):
    os.makedirs(LOGGING_DIR)
logfilename = os.path.join(LOGGING_DIR, 'application.log')


def init_Logger():
    logging_config = {
        'version': 1,
        'disable_existing_loggers': False,
        'formatters': {
            'standard': {
                'format': '%(asctime)s [%(levelname)s] %(name)s: %(message)s'
            },
        },
        'handlers': {
            'default_handler': {
                'class': 'logging.handlers.RotatingFileHandler',
                'level': 'DEBUG',
                'mode': 'w',
                'formatter': 'standard',
                'filename': logfilename,
                'encoding': 'utf8',
                'maxBytes': 5000000,
                'backupCount': 3
            },
        },
        'loggers': {
            '': {
                'handlers': ['default_handler'],
                'level': 'DEBUG',
                'propagate': False
            }
        }
    }
    logging.config.dictConfig(logging_config)
    stdout_handler = logging.StreamHandler(sys.stdout)
    logging.getLogger().addHandler(stdout_handler)


if __name__ == '__main__':
    init_Logger()
