import logging
import os
import time

from datetime import datetime, timezone

from aw_client import ActivityWatchClient
from aw_core.log import setup_logging
from aw_core.models import Event

from .config import parse_args
from .windows import get_outlook_activity, get_active_process_name


logger = logging.getLogger(__name__)
log_level = os.environ.get("LOG_LEVEL")
if log_level:
    logger.setLevel(logging.__getattribute__(log_level.upper()))


def main():
    # initialization phase
    args = parse_args()
    poll_interval = args.poll_time
    testing = args.testing

    # set up logging
    setup_logging(
        name="aw-watcher-outlook",
        testing=args.testing,
        verbose=args.verbose,
        log_stderr=True,
        log_file=True,
    )

    # set up client
    client = ActivityWatchClient("aw-watcher-outlook", testing=testing)

    # Create bucket if missing
    BUCKET_NAME = f"{client.client_name}_{client.client_hostname}"
    client.create_bucket(
        BUCKET_NAME,
        event_type="outlookitem",
    )

    # main watcher loop
    logger.info("aw-watcher-outlook started")
    client.wait_for_start()

    with client:
        current_state = {}  # use purely as helper for logging
        while True:
            try:
                # get data
                data = {}
                active_process = get_active_process_name()
                if active_process.lower() == "outlook.exe":
                    data = get_outlook_activity()

                # purely logging
                logger.debug("Data:", data)
                if current_state != data:
                    logger.info(f"Changed state from {current_state} to {data}")
                    current_state = data

                # send to server
                if data != {}:
                    event = Event(
                        timestamp=datetime.now(timezone.utc),
                        data=data
                    )
                    client.heartbeat(
                        BUCKET_NAME,
                        event,
                        pulsetime=poll_interval * 2,
                        queued=True,
                    )

                time.sleep(poll_interval)
            except KeyboardInterrupt:
                logger.info("aw-watcher-outlook stopped by keyboard interrupt")
                break
