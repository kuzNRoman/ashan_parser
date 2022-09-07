import pandas as pd
from numpy import add
from requests_html import HTMLSession
import json
import win32com.client as win32


class HelpfullFunctions:
    def getRightDay(today):
        """
        Классифицируем сегодняшний день
        """
        tday = None
        if today == 1:
            tday = "пн"
        if today == 2:
            tday = "вт"
        if today == 3:
            tday = "ср"
        if today == 4:
            tday = "чт"
        if today == 5:
            tday = "пт"
        return tday

    def getSortData():
        """
        Функция сортирует и отбирает нужные нам столбцы из большого Excel файла (особенности внутри компании были)
        """
        # открываем большой файл, где хранится отчетность и хранятся другие столбцы, также необходимые в целом компании, но для парсера они нужны не все
        excel_data_df = pd.read_excel("ashan.xlsx", sheet_name="sheet1")
        data_sort = pd.concat(
            [
                excel_data_df["Shop"],
                excel_data_df["Комментарии"],
                excel_data_df["поставка"],
                excel_data_df["Day"],
                excel_data_df["e-mail"],
                excel_data_df["City"],
                excel_data_df["Tovar1"],
                excel_data_df["Tovar2"],
                excel_data_df["Tovar3"],
                excel_data_df["Квант"],
            ],
            axis=1,
        )
        return data_sort

    def parseUrls(data_sort, tday):
        """
        Функция парсит данные за сегодняшний день, проходя по ссылкам по каждому товару и вносит их в DataFrame.
        В конце получаем отсортированный df, который потом запишем в excel
        """
        session = HTMLSession()
        # создаем пустые DataFrame для дальнейшего заполнения
        df = pd.DataFrame(columns=["Shop", "Tovar1"])
        df2 = pd.DataFrame(columns=["Shop", "Tovar2"])
        df3 = pd.DataFrame(columns=["Shop", "Tovar3"])

        data_today = data_sort.loc[data_sort.Day == tday]

        for shop, ssilka in zip(data_today["Shop"], data_today["File_T"]):
            try:
                r = session.get(ssilka)
                my_json = r.content.decode("utf8")
                data = json.loads(my_json)
                df.loc[len(df)] = [shop, data["stock"]["qty"]]
            except:
                data = "НД"
                df.loc[len(df)] = [shop, data]
                continue

        for shop2, ssilka2 in zip(data_today["Shop"], data_today["File_P"]):
            try:
                r2 = session.get(ssilka2)
                my_json2 = r2.content.decode("utf8")
                data2 = json.loads(my_json2)
                df2.loc[len(df2)] = [shop, data2["stock"]["qty"]]
            except:
                data2 = "НД"
                df2.loc[len(df2)] = [shop2, data2]
                continue

        for shop3, ssilka3 in zip(data_today["Shop"], data_today["Farsh"]):
            try:
                r3 = session.get(ssilka3)
                my_json3 = r3.content.decode("utf8")
                data3 = json.loads(my_json3)
                df3.loc[len(df3)] = [shop, data3["stock"]["qty"]]
            except:
                data3 = "НД"
                df3.loc[len(df3)] = [shop3, data3]
                continue

        df_summary = pd.concat([df, df2["FilePiksha"], df3["FarshTreska"]], axis=1)

        data_today_sort = pd.concat(
            [
                data_today["Day"],
                data_today["поставка"],
                data_today["e-mail"],
                data_today["City"],
                data_today["Комментарии"],
                data_today["Квант"],
            ],
            axis=1,
        )
        data_to_excel = pd.concat(
            [df_summary.reset_index(), data_today_sort.reset_index()], axis=1
        )
        return data_to_excel

    def emailSender(data_to_excel):
        """
        Функция автоматически генерирует сообщение, с учетом кол-ва остатков у каждого магазина и дня, и рассылает сообщения.
        Также автоматически высылается сообщение-отчет руководителям отдела.
        """
        days_dict = {
            "пн": "понедельникам",
            "вт": "вторникам",
            "ср": "средам",
            "чт": "четвергам",
            "пт": "пятницам",
        }
        for (
            address,
            day,
            postavka,
            amount_tovar1,
            amount_tovar2,
            amount_tovar3,
            kvant,
        ) in zip(
            data_to_excel["e-mail"],
            data_to_excel["Day"],
            data_to_excel["поставка"],
            data_to_excel["FileTreska"],
            data_to_excel["FilePiksha"],
            data_to_excel["FarshTreska"],
            data_to_excel["SteykTreska"],
            data_to_excel["Квант"],
        ):
            try:
                day = days_dict[day]
                postavka = days_dict[postavka]

                if amount_tovar1 < 5:
                    amount_tovar1_text = (
                        "Просим обратить внимание, что у Вас заканчивается товар1!  "
                    )
                else:
                    amount_tovar1_text = ""

                if amount_tovar2 < 5:
                    amount_tovar2_text = (
                        "Просим обратить внимание, что у Вас заканчивается товар2!  "
                    )
                else:
                    amount_tovar2_text = ""

                if amount_tovar3 < 5:
                    amount_tovar3_text = (
                        "Просим обратить внимание, что у Вас заканчивается товар3!  "
                    )
                else:
                    amount_tovar3_text = ""

                kvant2 = str(kvant)
                text1 = (
                    "Уважаемые партнеры, добрый день!\nПоставщик №77777777\nСегмент №777\nМинимальный объем поставки по Вам: "
                    + kvant2
                    + " Рублей.\n График доставки: прием заявок по "
                    + day
                    + ", поставка по "
                    + postavka
                    + ".\nГотовы к поставкам следующих позиций:\nПОЗИЦИЯ1\nПОЗИЦИЯ2\n"
                    + amount_tovar1_text
                    + amount_tovar2_text
                    + "\n"
                    + amount_tovar3_text
                    + "\n Благодарю за внимание!"
                )

                outlook = win32.Dispatch("outlook.application")
                mail = outlook.CreateItem(0)
                mail.To = address
                mail.Subject = "Тема сообщения"
                mail.Body = text1
                mail.Send()
                outlook.Quit()

            except:
                continue

        # Отправка внутреннего отчета начальству
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.To = "email1@email.info; email2@email2.info"
        mail.Subject = "Отчет"
        mail.Body = (
            "Добрый день!\nВ приложении отсортированный набор магазинов с датой заявки на сегодня"
            + " и информацией по остаткам на сегодняшний день.\n Также были автоматически разосланы сообщения по всем магазинам на сегодняшний день (с учетом кванта, дат заявок и отгрузок).\n Сообщение сгенерировано автоматически, отвечать не нужно."
        )
        attachment = "path_to_file"
        mail.Attachments.Add(attachment)
        mail.Send()
        outlook.Quit()
