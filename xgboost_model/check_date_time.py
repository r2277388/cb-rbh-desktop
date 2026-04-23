from datetime import datetime
import pytz

# Define your local timezone
local_timezone = pytz.timezone("America/Los_Angeles")  # Replace with your timezone

# Get the current time in UTC and convert to local time
current_time_utc = datetime.now(pytz.utc)
current_time_local = current_time_utc.astimezone(local_timezone)

print("Current UTC time:", current_time_utc)
print("Current local time:", current_time_local)
