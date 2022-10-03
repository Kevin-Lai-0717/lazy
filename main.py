import datetime
import os
import threading
from step.tools.temperature import Temperature
from step.tools.insight import Insight

tocc_data = {"acc": "zxc578623",
             "pas": "asd578623", }

Insight_data = {"acc": "administrator",
                "pas": "P@ssw0rd", }

# 下周目標:底下code移去step + 自動貼照片至word
if __name__ == "__main__":
    print("體溫=t，InsightIQ=i，都要=a，貼上圖檔=p")
    step = input("你今天想要做什麼?：")
    weekday = datetime.date.today().weekday()
    # -----選擇要體溫-----
    if step.lower() == "t":

        i = 0
        # -----如果是周一 自動填TOCC-----
        if weekday != 0:
            Temp = Temperature(tocc_data.get("acc"), tocc_data.get("pas"), mask_off=True)
            Temp.open_chrome()
            Temp.connect()
            Temp.my_temperature()

        else:
            # -----詢問是否曾脫口罩-----
            while True:
                print("上周末是否曾外出脫下口罩?")
                mask_off = input("y or n：")
                print("================================================================")
                if mask_off.lower() not in {"y", "n"}:
                    i += 1
                    print("別亂輸入 不然把你的頭給擰下來")
                    if i < 4:
                        continue
                    print("不玩了 罷工 自己去手動填")
                    os._exit(0)
                else:
                    break
            Temp = Temperature(tocc_data.get("acc"), tocc_data.get("pas"), mask_off)
            Temp.open_chrome()
            Temp.connect()
            Temp.my_tocc()
            Temp.my_temperature()
    elif step.lower() == "i":
        insight = Insight(Insight_data.get("acc"), Insight_data.get("pas"))
        insight.open_chrome()
        insight.connect()
        insight.get_data()
        insight.edit_text_to_docx()

    elif step.lower() == "p":
        if weekday ==4:
            print("今天不是星期五 不要貼上圖檔")
            os._exit(0)
        else:
            insight = Insight(Insight_data.get("acc"), Insight_data.get("pas"))
            insight.add_pictures()

    elif step.lower() == "a":
        i = 0
        # -----如果是周一 自動填TOCC-----
        if weekday != 0:
            Temp = Temperature(tocc_data.get("acc"), tocc_data.get("pas"), mask_off=True)
            Temp.open_chrome()
            Temp.connect()
            Temp.my_temperature()
            insight = Insight(Insight_data.get("acc"), Insight_data.get("pas"))
            insight.open_chrome()
            insight.connect()
            insight.get_data()
            insight.edit_text_to_docx()
        else:
            # -----詢問是否曾脫口罩-----
            while True:
                print("上周末是否曾外出脫下口罩?")
                mask_off = input("y or n：")
                print("================================================================")
                if mask_off.lower() not in {"y", "n"}:
                    i += 1
                    print("別亂輸入 不然把你的頭給擰下來")
                    if i < 4:
                        continue
                    print("不玩了 罷工 自己去手動填")
                    os._exit(0)
                else:
                    break
            Temp = Temperature(tocc_data.get("acc"), tocc_data.get("pas"), mask_off)
            Temp.open_chrome()
            Temp.connect()
            Temp.my_tocc()
            Temp.my_temperature()
            insight = Insight(Insight_data.get("acc"), Insight_data.get("pas"))
            insight.open_chrome()
            insight.connect()
            insight.get_data()
            insight.edit_text_to_docx()

    else:
        print("字都打不好? 睡了 晚安")
        os._exit(0)
