import eikon as ek

def EikonDownloader():
    print("Start Eikon")
    ek.set_app_key('52100a58c95f42ffbc0b80d245888cc811196d79')
    ek.get_news_headlines('R:LHAG.DE', date_from='2019-03-06T09:00:00', date_to='2019-03-06T18:00:00')