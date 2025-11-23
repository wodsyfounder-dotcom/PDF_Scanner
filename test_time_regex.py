import re

TIME_REGEX = re.compile(
    r"""
    \b(?:
        # Clock time format: HH:MM:SS or HH:MM with optional AM/PM
        (?:[01]?\d|2[0-3]):[0-5]\d(?::[0-5]\d)?\s*(?:[AP]M|[ap]m)?
        |
        # Duration with units (supports decimals like 2.5 minutes)
        [-+]?(?:\d+(?:\.\d+)?)\s*(?:
            # Nanoseconds
            ns|nsec|nanosec|nanosecond|nanoseconds
            |
            # Microseconds
            us|usec|microsec|microsecond|microseconds|µs|μs
            |
            # Milliseconds
            ms|msec|millisec|millisecond|milliseconds
            |
            # Seconds
            s|sec|secs|second|seconds
            |
            # Minutes
            m|min|mins|minute|minutes
            |
            # Hours
            h|hr|hrs|hour|hours
            |
            # Days
            d|day|days
            |
            # Weeks
            w|wk|wks|week|weeks
        )\b
    )
    """,
    re.VERBOSE | re.IGNORECASE
)

print("Testing enhanced TIME_REGEX:\n")

test_values = [
    # Clock times
    "14:30", "2:30 PM", "23:59:59", "12:00 am",
    # Seconds - with and without spaces
    "30s", "30 s", "45sec", "45 sec", "60seconds", "60 seconds", "2.5secs",
    # Minutes - attached and spaced
    "5m", "5 m", "10min", "10 min", "15minutes", "2.5 mins",
    # Hours - attached and spaced
    "2h", "2 h", "3hr", "3 hr", "4hours", "1.5 hrs", "1.5hrs",
    # Days - attached and spaced
    "7d", "7 d", "14days", "14 days",
    # Weeks - attached and spaced
    "2w", "2 w", "3weeks", "1wk", "1 wk",
    # Milliseconds - attached and spaced
    "500ms", "500 ms", "100milliseconds", "50msec", "50 msec",
    # Microseconds - attached and spaced
    "100us", "100 us", "250microseconds", "500µs",
    # Nanoseconds - attached and spaced
    "1000ns", "1000 ns", "500nanoseconds",
    # Mixed text
    "Wait 30seconds before proceeding",
    "Soak time: 2.5hours",
    "Duration 45min",
    "Test for 100ms then 2.5hrs",
]

for val in test_values:
    match = TIME_REGEX.search(val)
    if match:
        print(f"✓ '{val:45}' -> matched: '{match.group(0)}'")
    else:
        print(f"✗ '{val:45}' -> NO MATCH")
