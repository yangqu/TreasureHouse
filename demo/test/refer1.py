import threading
import time

def add(x,y):
 print(x+y)
t = threading.Timer(10,add,args=(4,5))
t.start()
time.sleep(12)
t.cancel()
print("===end===")