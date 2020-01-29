import Functions_for_Milk_parser

def Milk_parser():
    #Основная программа
    Perekrestok = Functions_for_Milk_parser.Perekrestok_milk_parser()

    #Блок_1: Создаем директорию и файл Эксель Milk_parser с рабочим листом формата год-месяц-день и возвращаем имя рабочего листа. 
    sheet_name = Perekrestok.open_excel_and_put_information_to_file()
    #Блок_2: Получаем все ссылки на молоко
    all_links = Perekrestok.get_all_milk_links()
    links_q_ty = len(all_links)
    all_milk_data = []
    for i in range(0, links_q_ty):
        #Парсим сайт перекрестка по каждой ссылке и получаем необходимую информацию в формате множества {} 
        milk_data = Perekrestok.get_information_from_link(all_links[i])
        all_milk_data.append(milk_data)
    #print(all_milk_data)
    #Блок_3: Записываем полученную информацию     
    Perekrestok.load_information_to_excel_file(sheet_name, all_milk_data) 


Milk_parser()



