import os

from app import app, read_serial_data, start_mqtt_client
import threading

# Start hardware integrations once per WSGI worker only when explicitly enabled.
if os.environ.get("ENABLE_HARDWARE", "0") == "1":
    threading.Thread(target=read_serial_data, daemon=True).start()
    start_mqtt_client()

