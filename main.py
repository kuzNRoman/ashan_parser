#!/usr/bin/python
# -*- coding: utf-8 -*-
import pandas as pd
import datetime
from helpfull_functions import HelpfullFunctions


def dailyJob():
    hf = HelpfullFunctions()
    today = datetime.datetime.today().isoweekday()
    tday = hf.getRightDay(today)
    data_sort = hf.getSortData()
    data_to_excel = hf.parseUrls(data_sort, tday)

    writer = pd.ExcelWriter("output_today.xlsx")
    data_to_excel.to_excel(writer)
    writer.save()

    send_messages = hf.emailSender(data_to_excel)
    return "done"


# dailyJob()
