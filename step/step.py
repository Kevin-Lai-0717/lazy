import datetime
import os
from step.tools.temperature import Temperature
from step.tools.insight import Insight

tocc_data = {"acc": "zxc578623",
             "pas": "asd578623", }

Insight_data = {"acc": "administrator",
                "pas": "P@ssw0rd", }


class Step:

    def temp(self):
        weekday = datetime.date.today().weekday()
        i = 0

        # -----不是周一 就不用寫tocc-----
        if weekday != 0:
            temp = Temperature(tocc_data.get("acc"), tocc_data.get("pas"), mask_off=True)
            temp.open_chrome()
            temp.connect()
            temp.my_temperature()

        else:
            # -----詢問是否曾脫口罩-----
            while True:
                print("上周末是否曾外出脫下口罩?")
                mask_off = input("y or n：")
                mask_off = mask_off.lower()
                print("================================================================")
                if mask_off not in {"y", "n"}:
                    i += 1
                    print("別亂輸入 不然把你的頭給擰下來")
                    if i < 4:
                        continue
                    print("不玩了 罷工 自己去手動填")
                    os._exit(0)
                else:
                    break
            temp = Temperature(tocc_data.get("acc"), tocc_data.get("pas"), mask_off)
            temp.open_chrome()
            temp.connect()
            temp.my_tocc()
            temp.my_temperature()

    def insight(self):
        insight = Insight(Insight_data.get("acc"), Insight_data.get("pas"))
        insight.open_chrome()
        insight.connect()
        insight.get_data()
        insight.edit_text_to_docx()

    def pic(self):
        weekday = datetime.date.today().weekday()
        if weekday != 4:
            print("今天不是星期五 不要貼上圖檔")
            os._exit(0)
        else:
            insight = Insight(Insight_data.get("acc"), Insight_data.get("pas"))
            insight.add_pictures()
