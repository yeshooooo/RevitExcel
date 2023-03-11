
import threading
import time
def thread1():
    time.sleep(10)
    print("进程1执行")

if __name__ == '__main__':
    print("程序开始执行")
    t = threading.Thread(target=thread1)
    t.start()

    print("主线程结束")