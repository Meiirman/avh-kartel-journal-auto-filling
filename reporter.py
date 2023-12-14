import requests


def get_default_data():
    myDict = {
        i.split(" ")[0] : i.split(" ")[1] 
        for i in open("default.txt" , 'r' , encoding='utf-8').read().split("\n")
    }
    return myDict['LINK']


def get_curent_time():
    import datetime

    current_time = datetime.datetime.now()
    formatted_time = current_time.strftime("%d.%m.%Y %H:%M:%S")
    return formatted_time

def send_report(text="Текста нет", process="Процесс без названия", responsible="Неизвестный ответственный"):
    URL = get_default_data()
    
    response = requests.post(f"{URL}?time={get_curent_time()}&process={process}&responsible={responsible}&text={text}")
    return response
