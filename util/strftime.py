import time

def time_transfer(milliseconds):
    Elapsedtime_transfer = time.strftime('%H:%M:%S', time.gmtime(milliseconds / 1000))
    return Elapsedtime_transfer