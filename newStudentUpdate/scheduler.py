import schedule
import time
import powerschoolscript
 
def run(script):
    script = powerschoolscript.main()
# Execute this task every 5 seconds
schedule.every(1).minutes.do(run, "Running script")
 
# Runs code in an infinite loop
while True:
    schedule.run_pending()
    time.sleep(1)