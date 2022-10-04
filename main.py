from step.step import Step
import os

if __name__ == "__main__":
    print("體溫=t，InsightIQ=i，都要=a，貼上圖檔=p")
    want = input("你今天想要做什麼?：")
    step = Step()
    if want.lower() == "t":
        step.temp()

    elif want.lower() == "i":
        step.insight()

    elif want.lower() == "a":
        step.temp()
        step.insight()

    elif want.lower() == "p":
        step.pic()

    else:
        print("字都打不好? 睡了 晚安")
        os._exit(0)
