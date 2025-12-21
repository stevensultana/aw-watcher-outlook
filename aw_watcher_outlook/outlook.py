from datetime import datetime, timezone
import time

from aw_client import ActivityWatchClient
from aw_core.models import Event

from .windows import get_outlook_activity, get_active_process_name

DEBUG = True

def main(poll_interval: float, testing: bool):
    # initialization phase
    # set up client
    client = ActivityWatchClient("aw-watcher-outlook", testing=testing)
    client.wait_for_start()

    # Create bucket if missing
    BUCKET_NAME = f"{client.client_name}_{client.client_hostname}"
    client.create_bucket(
        BUCKET_NAME,
        event_type="outlookitem",
    )

    # main watcher loop
    with client:
        while True:
            try:
                outlook_active = get_active_process_name().lower() == "outlook.exe"
                if outlook_active:
                    data = get_outlook_activity()
                else:
                    data = None
                if DEBUG: print("Data:", data)

                if data is not None:
                    event = Event(
                        timestamp=datetime.now(timezone.utc),
                        data=data
                    )
                    client.heartbeat(BUCKET_NAME, event, pulsetime=poll_interval * 2)

                time.sleep(poll_interval)
            except KeyboardInterrupt:
                print("aw-watcher-outlook stopped by keyboard interrupt")
                break
