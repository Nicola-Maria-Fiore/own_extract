import eikon as ek

def EikonDownloader():
    print("Start Eikon")
    ek.set_app_key('52100a58c95f42ffbc0b80d245888cc811196d79')
    data = ek.get_data("ENEI.MI", ["TR.HoldingsDate"], parameters={"SDate":"2000-12-31", "CH":"Fd", "RH":"IN"})
    print(data)