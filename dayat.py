import datetime
import sys

update_daytime = datetime.datetime.strptime(sys.argv[1], '%Y-%m-%d')
# uday 15:30 ~ uday+1 8:00
min_update_time = update_daytime + datetime.timedelta(hours=15,minutes=30)
max_update_time = update_daytime + datetime.timedelta(days=1, hours=8)

now = datetime.datetime.now()

# send 3 to %errorlevel%
if now<min_update_time or now>max_update_time: sys.exit(3)