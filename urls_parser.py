from requests_html import HTMLSession
import json
import pandas as pd

session = HTMLSession()

urls_tovar1 = {
    # example
    "Мытищи": "https://www.auchan.ru/v1/catalog/product-detail?code=tovar-soup-mor-450g&merchantId=1",
    "Коммунарка": "https://www.auchan.ru/v1/catalog/product-detail?code=tovar-soup-mor-450g&merchantId=2",
    "Марфино": "https://www.auchan.ru/v1/catalog/product-detail?code=tovar-soup-mor-450g&merchantId=3",
}

urls_tovar2 = {
    # example
    "Мытищи": "https://www.auchan.ru/v1/catalog/product-detail?code=tovar-soup-mor-450g&merchantId=1",
    "Коммунарка": "https://www.auchan.ru/v1/catalog/product-detail?code=tovar-soup-mor-450g&merchantId=2",
    "Марфино": "https://www.auchan.ru/v1/catalog/product-detail?code=tovar-soup-mor-450g&merchantId=3",
}


urls_tovar3 = {
    # example
    "Мытищи": "https://www.auchan.ru/v1/catalog/product-detail?code=tovar-soup-mor-450g&merchantId=1",
    "Коммунарка": "https://www.auchan.ru/v1/catalog/product-detail?code=tovar-soup-mor-450g&merchantId=2",
    "Марфино": "https://www.auchan.ru/v1/catalog/product-detail?code=tovar-soup-mor-450g&merchantId=3",
}

# парсим ссылки
def parsingUrls(urls_data_tovar1, urls_data_tovar2, urls_data_tovar3):
    session = HTMLSession()
    df = pd.DataFrame({"Магазин": "Наименование", "Товар1": "Штук"}, index=[0])
    df2 = pd.DataFrame({"Магазин": "Наименование", "Товар2": "Штук"}, index=[0])
    df3 = pd.DataFrame({"Магазин": "Наименование", "Товар3": "Штук"}, index=[0])
    for shop, ssilka in zip(urls_data_tovar1[0], urls_data_tovar1[1]):
        try:
            r = session.get(ssilka)
            my_json = r.content.decode("utf8")
            data = json.loads(my_json)
            df.loc[len(df)] = [shop, data["stock"]["qty"]]
        except:
            data = "НД"
            df.loc[len(df)] = [shop, data]
            continue

    for shop2, ssilka2 in zip(urls_data_tovar2[0], urls_data_tovar2[1]):
        try:
            r2 = session.get(ssilka2)
            my_json2 = r2.content.decode("utf8")
            data2 = json.loads(my_json2)
            df2.loc[len(df2)] = [shop2, data2["stock"]["qty"]]
        except:
            data2 = "НД"
            df2.loc[len(df2)] = [shop2, data2]
            continue

    for shop3, ssilka3 in zip(urls_data_tovar3[0], urls_data_tovar3[1]):
        try:
            r3 = session.get(ssilka3)
            my_json3 = r3.content.decode("utf8")
            data3 = json.loads(my_json3)
            df3.loc[len(df3)] = [shop3, data3["stock"]["qty"]]
        except:
            data3 = "НД"
            df3.loc[len(df3)] = [shop3, data3]
            continue

    df_summary = pd.concat([df, df2["Товар2"], df3["Товар3"]], axis=1)
    return df_summary


# вызов парсера с выгрузкой в xlsx данных
data_tovar1 = pd.DataFrame(urls_tovar1.items())
data_tovar2 = pd.DataFrame(urls_tovar2.items())
data_tovar3 = pd.DataFrame(urls_tovar3.items())
final_df = parsingUrls(data_tovar1, data_tovar2, data_tovar3)
writer = pd.ExcelWriter("output.xlsx")
final_df.to_excel(writer)
writer.save()
