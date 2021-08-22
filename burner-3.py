
import os
import math
import xlsxwriter
import json
import datetime

from collections import namedtuple
from gas_control_section import Gas_control_section
from classics import Burner,SSV,Silencer,Star_delta,VSD,Booster_station,Burners
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter.ttk import Combobox


def burner_choice (capacity,resistance):
    
    for item in Burners[fuel]:
        minload=float(item.minload)
        maxload=float(item.maxload)
        if capacity>minload and capacity<maxload:
            point_pressure=float(item.point_pressure(capacity))
            if resistance<point_pressure:
                print(item)
                return item


def check_dimensions():
    diameter=float(entry_diameter.get())
    length=float(entry_length.get())
    required_diameter=lambda x: 0.0244*x**3-1.9659*x**2+89.4752*x+538.3134
    required_length=lambda x: 2399.8229*x**0.3333
    required_diameter=required_diameter(capacity)
    required_length=required_length(capacity)

    if diameter<required_diameter:
        shortage=round(100*(required_diameter-diameter)/required_diameter)
        text_check_dimensions=f"Required diameter: {round(required_diameter,-1)} mm --> shortage: {shortage} %"
    else:
        text_check_dimensions=f"Required diameter: {round(required_diameter,-1)} mm --> OK"
    Label(frame_chooser,text=text_check_dimensions).grid(row=4,column=2,columnspan=2,padx=1,pady=1,sticky="w")

    if length<required_length:
        shortage=round(100*(required_length-length)/required_length)
        text_check_dimensions=f"Required length: {round(required_length,-1)} mm --> shortage: {shortage} %"
    else:
        text_check_dimensions=f"Required length: {round(required_length,-1)} mm --> OK"
    Label(frame_chooser,text=text_check_dimensions).grid(row=5,column=2,columnspan=2,padx=1,pady=1,sticky="w")



def inside_marks(string):
    """Возвращает текст из строки внутри кавычек «»,
    если кавычек нет, саму строку"""
    try:
        open_mark = string.index("«")
        close_mark = string.index("»")
        return string[open_mark+1:close_mark]
    except ValueError as er:
        return string



def quotation_maker():

    current_date = datetime.date.today().strftime("%d.%m.%Y")


    with open("prices.json","r",encoding="utf-8") as pr:
        prices=json.load(pr)

    try:
        status_booster_station_=status_booster_station.get()
    except NameError:
        status_booster_station_="delete"

    try:
        status_flow_meter_=status_flow_meter.get()
    except NameError:
        status_flow_meter_="delete"

    try:
        status_fuel_fitting_station_=status_fuel_fitting_station.get()
    except NameError:
        status_fuel_fitting_station_="delete"



    equipment_dict={"burner":"main", "fuel_fitting_station":status_fuel_fitting_station_, "SSV":status_SSV.get(),
                    "gas_control_section":status_gas_control_section.get(), "booster_station":status_booster_station_, "flow_meter":status_flow_meter_,
                    "flame_monitoring":status_flame_monitoring.get(), "seavis":status_seavis.get(), "star_delta":status_star_delta.get(),
                    "silencer":status_silencer.get(), "monoblock_vsd":status_monoblock_vsd.get(), "seavis_vsd":status_seavis_vsd.get(), 
                    "vsd":status_vsd.get(), "load_feedback":status_load_feedback.get(), "BUS_interface":status_bus_interface.get(),
                    "O2_trim":status_o2_trim.get(), "O2_system":status_o2_system.get()}


    if float(chosen_burner.motor_rating.replace(",","."))>=55:
        del equipment_dict["star_delta"]

    main_set={"burner","fuel_fitting_station","SSV","silencer","booster_station","flow_meter","gas_control_section"}
    control_set={"seavis","star_delta","monoblock_vsd","seavis_vsd","vsd",
                "load_feedback","BUS_interface","O2_trim","O2_system","flame_monitoring"}

    main_list=[]
    control_list=[]
    option_list=[]

    for equipment,status in equipment_dict.items():
        if status=="main" and equipment in main_set:
            main_list.append(equipment)
        if status=="main" and equipment in control_set:
            control_list.append(equipment)
        if status=="option":
            option_list.append(equipment)


    discount=entry_discount.get()
    number_of_burners=int(combo_number_of_burners.get())

    with open("directory.txt","r") as f:
        folder=f.read()
    if not folder:
        folder=os.path.expanduser("~")
    company = entry_company.get()
    company = inside_marks(company)
    mypath=os.path.join(folder,entry_number.get()+" "+str(chosen_burner)+" "+entry_boiler_name.get().strip()+" "+entry_boiler_capacity.get().strip()+" МВт "+ company.strip() +".xlsx")
    print(mypath)
    wb=xlsxwriter.Workbook(mypath)
    ws=wb.add_worksheet()

    bold_Arial10_format = wb.add_format({"bold": True,"font_name":"Arial","font_size":10})
    stand_Arial10_format = wb.add_format({"font_name":"Arial","font_size":10})
    bold_Arial11_format = wb.add_format({"bold": True,"font_name":"Arial","font_size":11})
    stand_Arial11_format = wb.add_format({"font_name":"Arial","font_size":11})

    shapka_format = wb.add_format({"align":"center","valign":"top","bold": True,"font_name":"Arial","font_size":10,"text_wrap": True,"border": 1})
    shapka_yellow_format = wb.add_format({"align":"center","valign":"top","bold": True,"font_name":"Arial","font_size":10,"text_wrap": True,"border": 1,"bg_color":"#FFFF99"})
    header_format = wb.add_format({"align":"center","valign":"center","bold": True,"font_name":"Arial","font_size":11,"text_wrap": True})
    header_italic_format = wb.add_format({"align":"center","valign":"center","bold": True, "italic": True, "font_name": "Arial","font_size":11,"text_wrap": True})
    easy_eq_format = wb.add_format({"valign":"top","font_name":"Arial","font_size":10,"text_wrap": True,"border": 1})

    b_format = wb.add_format({"align":"left", "valign": "top", "border": 1, "text_wrap": True,"font_name":"Arial","font_size":10})
    c_format = wb.add_format({"align":"right","valign":"top","font_name":"Arial","font_size":11,"num_format": "#,##0.00","bg_color":"#FFFF99","border": 1})
    d_e_f_format = wb.add_format({"align":"center","valign":"top","font_name":"Arial","font_size":11,"bg_color":"#FFFF99","border": 1})
    g_h_i_k_format = wb.add_format({"align":"right","valign":"top","font_name":"Arial","font_size":11,"num_format": "#,##0.00","border": 1})
    j_format = wb.add_format({"align":"center","valign":"top","font_name":"Arial","font_size":11,"border": 1})
    result_format = wb.add_format({"align":"right","valign":"top","font_name":"Arial","font_size":11,"num_format": "#,##0.00","border": 1,"bold": True})
    
    #Делаю ширину колонок
    ws.set_column("A:A", 4)
    ws.set_column("B:B", 75)
    ws.set_column("C:C", 12,None,{'level': 1})
    ws.set_column("D:D", 8,None,{'level': 1})
    ws.set_column("E:E", 8,None,{'level': 1})
    ws.set_column("F:F", 8,None,{'level': 1})
    ws.set_column("G:G", 12)
    ws.set_column("H:H", 10)
    ws.set_column("I:I", 12)
    ws.set_column("J:J", 4)
    ws.set_column("K:K", 12)


    ws.merge_range("A1:B1",entry_company.get(),bold_Arial11_format)
    

    ws.write("A3","СПЕЦИФИКАЦИЯ:"+"                  "+"№ "+entry_number.get(),bold_Arial11_format)
    ws.write_rich_string("A4",bold_Arial11_format,"Приложение к контракту:"+"      ",stand_Arial11_format,"№ "+entry_number.get() + " от " + current_date + " г.")
    ws.write_rich_string("A6",bold_Arial11_format,"Условия оплаты:"+"                    ",stand_Arial11_format,"предоплата 70%, 30% перед отгрузкой с завода в Германии")
    ws.write("A8","Условия поставки:"+"                 "+"DDP Москва, включая 20% НДС",bold_Arial11_format)
    ws.write_rich_string("A9",bold_Arial11_format,"Срок поставки:"+"                       ",stand_Arial11_format, entry_delivery_time.get() + " недель после согласования всех технических условий")
    ws.write("A10","                                                   и поступления предоплаты на счёт продавца",stand_Arial11_format)

    #Делаю шапку
    ws.write("A13","Поз.",shapka_format)
    ws.write("B13","Наименование товара",shapka_format)
    ws.write("C13","SPC",shapka_yellow_format)
    ws.write("D13","%",shapka_yellow_format)
    ws.write("E13","1,00",shapka_yellow_format)
    ws.write("F13","1,00",shapka_yellow_format)
    ws.write("G13","Цена за единицу в Евро, без НДС",shapka_format)
    ws.write("H13","НДС по ставке 20%, в Евро",shapka_format)
    ws.write("I13","Цена  в Евро за ед., вкл. НДС",shapka_format)
    ws.write("J13","Кол-во",shapka_format)
    ws.write("K13","Сумма в Евро, вкл. НДС",shapka_format)

    if fuel=="NG":
        main_header_text_1="Газовая"
    elif fuel=="NG/LFO":
        main_header_text_1="Комбинированная"
    elif fuel=="LFO":
        main_header_text_1="Жидкотопливная"
    main_header_text=main_header_text_1 + " горелка SAACKE тип " + chosen_burner.type_ + " " + chosen_burner.size + " с принадлежностями"
    ws.merge_range("A14:K14",main_header_text,header_format)

    

    def burner_field(b_cell,c_cell):
        if combo_type_of_boiler.get()=="steam":
            insert_1="парового "
            insert_4="паропроизводительностью"
        elif combo_type_of_boiler.get()=="hot water":
            insert_1="водогрейного "
            insert_4="тепловой мощностью"
        elif combo_type_of_boiler.get()=="TOH":
            insert_1="термомасляного котла "
            insert_4="тепловой мощностью"


        if combo_type_of_furnace.get()=="watertube":
            insert_2="водотрубного котла "
            insert_6=f"Размеры топки котла: ширина {entry_furnace_width.get()} мм, высота {entry_furnace_height.get()} мм, глубина {entry_furnace_length.get()} мм.\n"
        elif combo_type_of_furnace.get()=="flametube":
            insert_2="жаротрубного "
            insert_6=f"Размеры топки котла: длина {entry_length.get()} мм, диаметр {entry_diameter.get()} мм.\n"
        elif combo_type_of_furnace.get()=="TOH":
            insert_2=""

       
        if combo_type_of_furnace.get()=="flametube" and combo_type_of_ftb_furnace.get()=="3-pass":
            insert_3="трехходового котла"
        elif combo_type_of_furnace.get()=="flametube" and combo_type_of_ftb_furnace.get()=="2-pass":
            insert_3="двухходового котла"
        elif combo_type_of_furnace.get()=="flametube" and combo_type_of_ftb_furnace.get()=="inverse":
            insert_3="котла с инверсионной топкой"
        else:
            insert_3=""
       

        if fuel=="NG":
            insert_5="природный газ по ГОСТ 5542-2014."
            insert_8=("Тепловая мощность горелки: минимальная – " + str(chosen_burner.capacity_range[0][0]).replace(".",",") + " МВт, требуемая – " + str(capacity).replace(".",",") +
                    " МВт, максимальная – " + str(chosen_burner.capacity_range[0][1]).replace(".",",") + " МВт.\n")
            burner_field_0="Газовая блочная промышленная горелка "
            burner_field_9=(f"Регулирование мощности – модулируемое.\n"
                    f"Диапазон рабочего регулирования – {chosen_burner.turndown_ratio[0]}.\n" 
                    f"Регулирование соотношения топливо/воздух – электронное.\n")
            
        elif fuel=="NG/LFO":
            insert_5="природный газ по ГОСТ 5542-2014, дизельное топливо по ГОСТ 305-2013."
            insert_8=("Тепловая мощность горелки: минимальная при работе на газе – " + str(chosen_burner.capacity_range[0][0]).replace(".",",") + 
                    " МВт, на жидком топливе - " + str(chosen_burner.capacity_range[1][0]).replace(".",",") +" МВт, "
                    f"требуемая – " + str(capacity).replace(".",",") + " МВт, максимальная – " + str(chosen_burner.capacity_range[0][1]).replace(".",",") +" МВт.\n")
            burner_field_0="Комбинированная блочная промышленная горелка "
            burner_field_9=(f"Регулирование мощности – модулируемое.\n"+
                    f"Диапазон рабочего регулирования при работе на газе – {chosen_burner.turndown_ratio[0]}, " 
                    f"на дизельном топливе – {chosen_burner.turndown_ratio[1]}.\n" +
                    f"Регулирование соотношения топливо/воздух – электронное.\n")
        elif fuel=="LFO":
            insert_5="дизельное топливо по ГОСТ 305-2013."
            insert_8=("Тепловая мощность горелки: минимальная – " + str(chosen_burner.capacity_range[1][0]) + " МВт, требуемая – " + str(capacity).replace(".",",") +
                    " МВт, максимальная – " + str(chosen_burner.capacity_range[1][1]) +" МВт.\n")
            burner_field_0="Жидкотопливная блочная промышленная горелка "
            burner_field_9=(f"Регулирование мощности – модулируемое.\n"
                    f"Диапазон рабочего регулирования – {chosen_burner.turndown_ratio[1]}.\n" 
                    f"Регулирование соотношения топливо/воздух – электронное.\n")


        if combo_type_of_boiler.get()=="hot water":
            insert_7="МВт"
        elif combo_type_of_boiler.get()=="steam":
            insert_7="т/ч"

        global remark_counter
        remark_counter=1
        
        burner_field_1="SAACKE"
        burner_field_2=" тип "
        burner_field_3=f"{chosen_burner.type_} {chosen_burner.size} "
        burner_field_4="для " + insert_1 + insert_2 + insert_3 + " "
        burner_field_5=entry_boiler_name.get()
        burner_field_6=" " + insert_4 + " "
        burner_field_7=entry_boiler_capacity.get() + " " + insert_7 + ".\n"
        burner_field_8=(insert_6 +
                    "Сопротивление газоходов котла - " + str(resistance).replace(".",",") +" мбар.\n" 
                    f"Топливо – " + insert_5 + "\n" + insert_8)
        burner_field_10=(f"Примечание {remark_counter}: Стандартные условия эксплуатации: высота над уровнем моря - 250 м макс.," +
                        " установка в котельной. Температура воздуха горения - мин. 5 °С, макс. 55 °С. " +
                        "Исполнение горелки - не взрывозащищенное. Напряжение силовое - 400В,50Гц, напряжение цепей управления - 230В,50Гц.")
        
        remark_counter+=1

        ws.write_rich_string(b_cell,
                        stand_Arial10_format,burner_field_0,
                        bold_Arial10_format,burner_field_1,
                        stand_Arial10_format,burner_field_2,
                        bold_Arial10_format,burner_field_3,
                        stand_Arial10_format,burner_field_4,
                        bold_Arial10_format,burner_field_5,
                        stand_Arial10_format,burner_field_6,
                        bold_Arial10_format,burner_field_7,
                        stand_Arial10_format,burner_field_8,
                        bold_Arial10_format,burner_field_9,
                        bold_Arial10_format,burner_field_10,
                        b_format)

        if chosen_burner in Burners["NG/LFO"][12:]:
            if capacity <= 18.6:
                add_price = prices["fuel_fitting_station"]["18,6"]
            elif capacity > 18.6:
                add_price = prices["fuel_fitting_station"]["26,1"]
            burner_price = prices["Burner Mono"][str(chosen_burner)] + add_price
        else:
            burner_price = prices["Burner Mono"][str(chosen_burner)]
        ws.write_formula(c_cell,burner_price,c_format)


    def SSV_field(b_cell,c_cell):
        chosen_SSV=SSV(capacity)
        SSV_field_1="Газовый защитный участок со встроенным стабилизатором давления"
        SSV_field_2=(f" в блочном исполнении для расхода газа {chosen_SSV.flow()} нм³/час макс.\n"
                    f"Входное давление газа – {chosen_SSV.inlet_pressure()} мбар, 500 мбар макс.\n" 
                    f"Состав: 2 отсечных клапана с электромагнитным приводом {chosen_SSV.diameter()}, реле давления газа макс., " 
                    f"устройство контроля герметичности газовых клапанов на базе реле давления газа мин., кнопка аварийного выключения, компенсатор тепловых расширений {chosen_SSV.diameter()}.")
        
        ws.write_rich_string(b_cell,
                            bold_Arial10_format,SSV_field_1,
                            stand_Arial10_format,SSV_field_2,
                            b_format)
        ws.write_formula(c_cell,prices["SSV"][chosen_SSV.diameter()],c_format)
        ws.set_row(int(b_cell[1:])-1,96.75)

    def gas_control_section_field(b_cell,c_cell):
        chosen_gas_control_section = Gas_control_section(capacity)
        gas_control_section_field_1 = "Газовый регулирующий участок"
        gas_control_section_field_2 = (f" высокого давления {chosen_gas_control_section.size()[0]}-{chosen_gas_control_section.size()[1]}/{chosen_gas_control_section.size()[2]}-{chosen_gas_control_section.size()[3]} "
                                    f"для расхода газа {chosen_gas_control_section.max_flow()} нм³/час., входное давление газа – 3 бар, {chosen_gas_control_section.max_pressure()} бар макс., "
                                    "температура газа -10...50 °C, температура окружающей среды -10...60 °C. \n"
                                    f"Состав: ручная запорная заслонка DN {chosen_gas_control_section.size()[0]}, фильтр DN {chosen_gas_control_section.size()[0]}, "
                                    f"регулятор давления газа DN {chosen_gas_control_section.size()[2]} со встроенным ПЗК, "
                                    "ПСК, термометр, шаровой кран/вентиляция, два манометра с отсечными кранами.")
        ws.write_rich_string(b_cell,
                            bold_Arial10_format,gas_control_section_field_1,
                            stand_Arial10_format,gas_control_section_field_2,
                            b_format)
        ws.write_formula(c_cell,prices["gas_control_section"][chosen_gas_control_section.price_str()],c_format)

    def silencer_field(b_cell,c_cell):
        chosen_silencer=Silencer(chosen_burner)
        ws.write(b_cell,"Встроенный шумоглушитель.",easy_eq_format)
        ws.write_formula(c_cell,prices["silencer"][str(chosen_burner.size)],c_format)


    def flame_monitoring_field(b_cell,c_cell):
        ws.write(b_cell,"Датчик факела FLS09, реле факела FLUS06, включены в стоимость горелки.",easy_eq_format)
        ws.write_formula(c_cell,prices["flame_monitoring"],c_format)
    

    def seavis_field(b_cell,c_cell):
        global remark_counter
        seavis_field_1="Автоматика "
        seavis_field_2="Se@vis Compact"
        seavis_field_3=(" с функциями топочного автомата, электронно-связанного регулятора соотношения топливо/воздух, " +
                        "вывода текстовых и графических сообщений о режимах работы и неисправностях горелки, с современным цветным дисплеем для визуализации.\n" +
                        "Шкаф с автоматикой смонтирован на горелке.\n")
        seavis_field_4=f"Примечание {remark_counter}: Допустимая температура окружающей среды – 5 ÷ 50 °С.\n"
        remark_counter+=1
        seavis_field_5=f"Примечание {remark_counter}: Со стороны заказчика обеспечивается: наличие выключателя для подвода напряжения  (230 В, 50/60 Гц, ~ 1000 ВА), аварийный выключатель."
        remark_counter+=1
        
        ws.write_rich_string(b_cell,
                            stand_Arial10_format,seavis_field_1,
                            bold_Arial10_format,seavis_field_2,
                            stand_Arial10_format,seavis_field_3,
                            bold_Arial10_format,seavis_field_4,
                            bold_Arial10_format,seavis_field_5,
                            b_format)

        ws.write_formula(c_cell,prices["seavis_integrated"][chosen_burner.type_],c_format)


    def star_delta_field(b_cell,c_cell):
        global remark_counter
        chosen_star_delta=Star_delta(chosen_burner.motor_rating)
        star_delta_field_1=(f"Аппаратура для защиты и управления электродвигателя дутьевого вентилятора {chosen_star_delta.motor_rating} кВт, " +
                            "смонтированная на горелке (подключение звезда-треугольник).\n")
        star_delta_field_2=f"Примечание {remark_counter}: Со стороны заказчика обеспечивается: подвод напряжения и защита предохранителями (400 В, 50 Гц).\n"
        remark_counter+=1
        star_delta_field_3=f"Примечание {remark_counter}: При применении частотного регулирования двигателя дутьевого вентилятора данная аппаратура не требуется."
        remark_counter+=1
        ws.write_rich_string(b_cell,
                            stand_Arial10_format,star_delta_field_1,
                            bold_Arial10_format,star_delta_field_2,
                            bold_Arial10_format,star_delta_field_3,
                            b_format)
        ws.write_formula(c_cell,prices["star_delta"][chosen_star_delta.motor_rating],c_format)
        ws.set_row(int(b_cell[1:])-1,87.75)


    def monoblock_vsd_field(b_cell,c_cell):
        ws.write(b_cell,"Исполнение двигателя дутьевого вентилятора для частотного регулирования.",b_format)
        ws.write_formula(c_cell,prices["monoblock_vsd"],c_format)


    def seavis_vsd_field(b_cell,c_cell):
        monoblock_vsd_field_1="Исполнение автоматики "
        monoblock_vsd_field_2="Se@vis Compact"
        monoblock_vsd_field_3=" для частотного регулирования двигателя дутьевого вентилятора."
        ws.write_rich_string(b_cell,
                            stand_Arial10_format,monoblock_vsd_field_1,
                            bold_Arial10_format,monoblock_vsd_field_2,
                            stand_Arial10_format,monoblock_vsd_field_3,
                            b_format)
        ws.write_formula(c_cell,prices["seavis_integrated_vsd"][chosen_burner.type_],c_format)
        ws.set_row(int(b_cell[1:])-1,33)

    def vsd_field(b_cell,c_cell):
        global remark_counter
        chosen_vsd=VSD(chosen_burner.motor_rating)
        vsd_field_1=(f"Частотный преобразователь электродвигателя дутьевого вентилятора {chosen_burner.motor_rating} кВт," +
                     " IP55 (настенное исполнение).\n")
        vsd_field_2=(f"Примечание {remark_counter}: При применении частотного регулирования оборотов двигателя дутьевого вентилятора"+
                    "установка аппаратуры для защиты и управления двигателя вентилятора не требуется.")
        remark_counter+=1
        ws.write_rich_string(b_cell,
                            stand_Arial10_format,vsd_field_1,
                            bold_Arial10_format,vsd_field_2,
                            b_format)
        ws.write_formula(c_cell,"=(" + prices["vsd"][chosen_vsd.motor_rating] + "*1.1)/(1.2*1.12*0.45)",c_format)
        ws.set_row(int(b_cell[1:])-1,75)
    
    def load_feedback_field(b_cell,c_cell):
        ws.write(b_cell,"Модуль информационный 4-20 мА об изменении нагрузки горелки.",b_format)
        ws.write_formula(c_cell,prices["load_feedback"],c_format)


    def bus_interface_field(b_cell,c_cell):
        bus_interface_field_1="Преобразователь интерфейса RS232 в RS485 ProfiBUS DP для автоматики "
        bus_interface_field_2="Se@vis Compact"
        bus_interface_field_3=" (другие протоколы по дополнительному запросу)."
        ws.write_rich_string(b_cell,
                            stand_Arial10_format,bus_interface_field_1,
                            bold_Arial10_format,bus_interface_field_2,
                            stand_Arial10_format,bus_interface_field_3,
                            b_format)
        ws.write_formula(c_cell,prices["BUS_interface"],c_format)
        ws.set_row(int(b_cell[1:])-1,33)
    
    
    def o2_trim_field(b_cell,c_cell):
        ws.write(b_cell,"Модуль для подачи 4-20 мА для корректировки по О2.",b_format)
        ws.write_formula(c_cell,prices["O2_trim"],c_format)


    def o2_system_field(b_cell,c_cell):
        o2_system_field=("Система измерения O2 в дымовых газах LS2/LT3.\n" +
                        "Для снижения тепловых потерь с уходящими газами при работе на природном газе и дизельном топливе.\n" +
                        "Состоит из следующих основных компонентов:\n" + 
                        "Лямбда-зонд с пробоотборным устройством\n" +
                        " - Макс. температура дымовых газов 300 °C\n" +
                        " - Глубина погружения в газоход макс. 300 мм\n" +
                        " - Соединительный кабель 5 м, с разъемом\n" +
                        "Передатчик в отдельном корпусе.")
        ws.write(b_cell,o2_system_field,b_format)
        ws.write_formula(c_cell,prices["O2_system"],c_format)
        ws.set_row(int(b_cell[1:])-1,126)

    def booster_station_field(b_cell,c_cell):
        global remark_counter,chosen_booster_station
        chosen_booster_station=Booster_station(capacity)
        booster_station_field_1=(f"Насосный агрегат выносной на поддоне MOG {chosen_booster_station.size()}, {chosen_booster_station.motor_rating()} "
                                "кВт, 380 В, 50 Гц высокого давления 30 бар, включая: "
                                "перепускной клапан, предохранительный клапан, фильтр, манометр, вакуумметр, реле давления мин., запорный кран, газовоздухоотделитель.\n")
        booster_station_field_2=f"Примечание {remark_counter}: Со стороны заказчика обеспечивается: подвод напряжения и защита предохранителями (400 В, 50 Гц)."
        remark_counter+=1
        ws.write_rich_string(b_cell,
                            stand_Arial10_format,booster_station_field_1,
                            bold_Arial10_format,booster_station_field_2,
                            b_format)
        ws.write_formula(c_cell,prices["booster_station"][chosen_booster_station.size()],c_format)

    def flow_meter_field(b_cell,c_cell):
        flow_meter_field=f"Жидкотопливный расходомер VZO {chosen_booster_station.flow_meter_size()}."
        ws.write(b_cell,flow_meter_field,b_format)
        ws.write_formula(c_cell,prices["flow_meter"][chosen_booster_station.flow_meter_size()],c_format)

    def fuel_fitting_station_field(b_cell,c_cell):
        fuel_fitting_station_field=("Выносная жидкотопливная арматура на поддоне.\n"
                                    "Включая: механически связанные шаровые краны в подающей и обратной линиях, "
                                    "грязеуловитель, реле давления топлива мин. и макс., "
                                    "2 предохранительных клапана с пневмоприводами, перепускной клапан, "
                                    "управляющие электромагнитные клапаны, манометры.\n"
                                    "Включена в стоимость горелки.")
        ws.write(b_cell,fuel_fitting_station_field,b_format)
        ws.write_formula(c_cell,"=0",c_format)
        ws.set_row(int(b_cell[1:])-1,84)

    def equipment_stroke(row,func,discount):
        row=str(row)
        a_cell="A"+row
        b_cell="B"+row
        c_cell="C"+row
        d_cell="D"+row
        e_cell="E"+row
        f_cell="F"+row
        g_cell="G"+row
        h_cell="H"+row
        i_cell="I"+row
        j_cell="J"+row
        k_cell="K"+row

        ws.write(a_cell,str(current_number)+".",j_format)
        func(b_cell,c_cell)
        ws.write(d_cell,discount,d_e_f_format)
        ws.write(e_cell,1,d_e_f_format)
        ws.write(f_cell,1.12,d_e_f_format)
        g_formula="=ROUND(("+c_cell+"-("+c_cell+"/100)*"+d_cell+")*"+e_cell+"*"+f_cell+",2)"

        ws.write_formula(g_cell,g_formula,g_h_i_k_format)
        h_formula="=ROUND("+g_cell+"*0.2,2)"
        ws.write_formula(h_cell,h_formula,g_h_i_k_format)
        i_formula="="+g_cell+"+"+h_cell
        ws.write_formula(i_cell,i_formula,g_h_i_k_format)
        ws.write(j_cell,1,j_format)
        k_formula="=("+g_cell+"+"+h_cell+")*"+j_cell
        ws.write_formula(k_cell,k_formula,g_h_i_k_format)
    

    function_list={"burner":burner_field, "SSV":SSV_field, "gas_control_section":gas_control_section_field, "booster_station":booster_station_field,
                "flow_meter":flow_meter_field, "silencer":silencer_field, "flame_monitoring":flame_monitoring_field, "seavis":seavis_field, "star_delta":star_delta_field,
                "monoblock_vsd":monoblock_vsd_field, "seavis_vsd":seavis_vsd_field, "vsd":vsd_field, "load_feedback":load_feedback_field,
                "BUS_interface":bus_interface_field, "O2_trim":o2_trim_field, "O2_system":o2_system_field, "fuel_fitting_station":fuel_fitting_station_field}

    current_string=15
    current_number=1
    result_rows=[]

    for equipment in main_list:
        result_rows.append(current_string)
        equipment_stroke(current_string,function_list[equipment],discount)
        current_string+=1
        current_number+=1

    merged_range="A" + str(current_string) + ":" + "K" + str(current_string)
    ws.merge_range(merged_range,"")
    ws.write("A"+str(current_string),"Автоматика  управления  и безопасности горелки",header_format)

    if control_list:
        current_string+=1
        for equipment in control_list:
            result_rows.append(current_string)
            equipment_stroke(current_string,function_list[equipment],discount)
            current_string+=1
            current_number+=1

    current_string+=1
    summation_string=current_string
    ws.write("A"+str(current_string),"Стоимость товара в Евро (без дополнительного оборудования):",bold_Arial11_format)
    result_price_1_burner="=K" + str(result_rows[0])
    for row in result_rows[1:]:
        result_price_1_burner += "+K" + str(row)
    ws.write_formula("K"+str(current_string),result_price_1_burner,result_format)

    current_string+=1

    ws.write("B"+str(current_string),"НДС, 20%",bold_Arial11_format)
    nds_1_burner="=H" + str(result_rows[0])
    for row in result_rows[1:]:
        nds_1_burner += "+H" + str(row)
    ws.write_formula("K"+str(current_string),nds_1_burner,result_format)

    current_string+=1

    ws.write("B"+str(current_string),"Транспортные расходы:",bold_Arial11_format)
    ws.write("K"+str(current_string),0,result_format)

    current_string+=2
    ws.write("A"+str(current_string),"ИТОГО сумма к оплате в Евро за 1 горелку с принадлежностями",bold_Arial11_format)
    current_string+=1
    ws.write("A"+str(current_string),"(без дополнительного оборудования) на условиях поставки DDP Москва,",bold_Arial11_format)
    current_string+=1
    ws.write("A"+str(current_string),f"включая 20% НДС, нетто (со скидкой для {entry_company.get()}):",bold_Arial11_format)
    ws.write_formula("K"+str(current_string),"=K" + str(summation_string) + "+K" + str(summation_string+2),result_format)
    result_string_1_burner=current_string

    if number_of_burners>1:

        current_string+=2
        if 1<number_of_burners<5:
            itogo_string=f"ИТОГО сумма к оплате в Евро за {number_of_burners} горелки с принадлежностями"
        elif 4<number_of_burners<11:
            itogo_string=f"ИТОГО сумма к оплате в Евро за {number_of_burners} горелок с принадлежностями"
        ws.write("A"+str(current_string),itogo_string,bold_Arial11_format)
        current_string+=1
        ws.write("A"+str(current_string),"(без дополнительного оборудования) на условиях поставки DDP Москва,",bold_Arial11_format)
        current_string+=1
        ws.write("A"+str(current_string),f"включая 20% НДС, нетто (со скидкой для {entry_company.get()}):",bold_Arial11_format)
        ws.write_formula("K"+str(current_string),"=K" + str(result_string_1_burner) + "*" + str(number_of_burners),result_format)
    
    
    current_string+=2

    if option_list:
        merged_range="A" + str(current_string) + ":" + "K" + str(current_string)
        ws.merge_range(merged_range,"Дополнительное оборудование - по желанию заказчика",header_format)

        current_string+=1

        for equipment in option_list:
            equipment_stroke(current_string,function_list[equipment],discount)
            current_string+=1
            current_number+=1

        current_string+=1

    merged_range="A" + str(current_string) + ":" + "K" + str(current_string)
    ws.merge_range(merged_range,"Дополнительные услуги, предоставляемые бесплатно:",header_format)

    current_string+=1

    service_strings=["Техническое сопровождение проекта.",
                    "Разрешительные документы: Сертификат ТР ТС, Паспорт на горелку",
                    "Техническая документация на монтаж и эксплуатацию по стандарту ЗААКЕ на русском языке."]
    for row in range(3):
        ws.write("A" + str(current_string),str(current_number)+".",j_format)
        merged_range="B" + str(current_string) + ":" + "K" + str(current_string)
        ws.merge_range(merged_range,service_strings[row],easy_eq_format)
        current_string+=1
        current_number+=1

    merged_range="A" + str(current_string) + ":" + "K" + str(current_string)
    ws.merge_range(merged_range,"По дополнительному запросу:",header_format)

    ws.set_row(current_string, 30)
    current_string+=1
    
    
    ws.write("A" + str(current_string),str(current_number)+".",j_format)
    merged_range="B" + str(current_string) + ":" + "K" + str(current_string)
    ws.merge_range(merged_range,("Склад запасных частей, сервисный центр в Москве, услуги по шеф -монтажу и шеф-наладке," +
                                "обучение персонала особенностям эксплуатации и технического обслуживания силами штатных высококвалифицированных сервисных инженеров компании SAACKE."),easy_eq_format)

    current_string+=2
    merged_range="A" + str(current_string) + ":" + "K" + str(current_string)
    ws.merge_range(merged_range,"Стоимость принадлежностей действительна при покупке вместе с горелкой",header_format)

    current_string+=1
    merged_range="A" + str(current_string) + ":" + "K" + str(current_string)
    ws.merge_range(merged_range,"Окончательная стоимость - после согласования техусловий",header_format)
       
    current_string+=1
    merged_range="A" + str(current_string) + ":" + "K" + str(current_string)
    ws.merge_range(merged_range,"Гарантия на оборудование - 12 мес. с момента пуска, но не более 18 мес. с момента поставки",header_format)

    current_string+=2
    merged_range="A" + str(current_string) + ":" + "K" + str(current_string)
    ws.merge_range(merged_range,"Для продления срока действия предложения, внесения изменений в техническую и",header_italic_format)

    current_string+=1
    merged_range="A" + str(current_string) + ":" + "K" + str(current_string)
    ws.merge_range(merged_range,"коммерческую части просим Вас обращаться к специалистам нашего представительства",header_italic_format)

    ws.print_area(0,0,current_string,10)
    ws.set_print_scale(70)
    ws.set_margins(bottom=0.75)
    ws.set_margins(right=0.64)

    with open("footer.txt","r",encoding="utf-8") as f:
        footer=f.read()

    ws.set_header('&C&G', {'image_center': 'log_saacke.jpg'})
    ws.set_footer(f'&L&09&"Arial,Regular"Ответственный: Лозицкий Дмитрий\nПодготовил: {footer}\nТел.: +7 (495) 788-80-88\n&D' +
                '&R&09&"Arial,Regular"Стр. &P из &N')
                

    wb.close()
    os.startfile(mypath)

def main():
    try:
        global capacity,resistance,chosen_burner,fuel,mono_duo
        capacity=float(entry_capacity.get())
    except ValueError:
        messagebox.showwarning("Error","Capacity field should contain a number")
    try:
        resistance=float(entry_resistance.get())
    except ValueError:
        messagebox.showwarning("Error","Resistance field should contain a number")
    
    fuel=combo_fuel.get()
    mono_duo=combo_mono_duo.get()
    
    chosen_burner=burner_choice(capacity,resistance)
    label_chosen_burner.config(text=chosen_burner)
    
    assert 2<capacity<28, "Неверный ввод"

    if chosen_burner:
        label_enter_diameter.grid(row=4,column=0,sticky=W)
        entry_diameter.grid(row=4,column=1)
        label_enter_lenght.grid(row=5,column=0,sticky=W)
        entry_length.grid(row=5,column=1)
        button_check_dimensions.grid(row=6,column=1)
        button_make_quotation.grid(row=4,column=0)


def new_project():
    
    def common_data():
        frame_common_data.grid(row=1,column=0,padx=10, pady=10)
        label_company.grid(row=0,column=0,sticky=W)    
        entry_company.grid(row=0,column=1)   
        label_number.grid(row=1,column=0,sticky=W)
        entry_number.grid(row=1,column=1)
        label_delivery_time.grid(row=2,column=0,sticky=W)
        entry_delivery_time.grid(row=2,column=1)
        Label(frame_common_data,text="Discount, %").grid(row=3,column=0,sticky=W)
        entry_discount.grid(row=3,column=1)
        
        Label(frame_common_data,text="Number of burners").grid(row=4,column=0,sticky=W)
        combo_number_of_burners.grid(row=4,column=1)
        combo_number_of_burners.set(1)
        

    fuel=combo_fuel.get()
    mono_duo=combo_mono_duo.get()

    def complectation():
        
        if frame_complectation.winfo_children():
            for widget in frame_complectation.winfo_children():
                print(widget)
                widget.grid_forget()
        frame_complectation.grid(row=1,column=1,padx=10, pady=10)
        
        
        Label(frame_complectation,text="Main").grid(row=0,column=1)
        Label(frame_complectation,text="Optional").grid(row=0,column=2)
        Label(frame_complectation,text="Delete").grid(row=0,column=3)

        equipment_list_NG = [equipment_SSV_radiobutton, equipment_flame_monitoring_radiobutton, equipment_seavis_radiobutton,
                            equipment_star_delta_radiobutton, equipment_gas_control_section, equipment_silencer_radiobutton, 
                            equipment_monoblock_vsd_radiobutton, equipment_seavis_vsd_radiobutton, equipment_vsd_radiobutton, 
                            equipment_load_feedback_radiobutton, equipment_bus_interface_radiobutton, 
                            equipment_o2_trim_radiobutton, equipment_o2_system_radiobutton]

        equipment_list_NG_LFO = [equipment_SSV_radiobutton, equipment_booster_station_radiobutton, equipment_flow_meter_radiobutton,
                            equipment_flame_monitoring_radiobutton, equipment_seavis_radiobutton,
                            equipment_star_delta_radiobutton, equipment_gas_control_section, equipment_silencer_radiobutton, 
                            equipment_monoblock_vsd_radiobutton, equipment_seavis_vsd_radiobutton,
                            equipment_vsd_radiobutton, equipment_load_feedback_radiobutton, equipment_bus_interface_radiobutton, 
                            equipment_o2_trim_radiobutton, equipment_o2_system_radiobutton]

        if fuel=="NG/LFO" and chosen_burner in Burners["NG/LFO"][12:]:
            equipment_list_NG_LFO.insert(1,equipment_fuel_fitting_station_radiobutton)

        if fuel=="NG":
            for row_, func in enumerate(equipment_list_NG,start=1):
                func(row_)
            
        if fuel=="NG/LFO":
            for row_, func in enumerate(equipment_list_NG_LFO,start=1):
                func(row_)

        
    common_data()
    complectation()
    button_ready.grid(row=3,column=1,sticky=N)


window=Tk()
window.title("TEMINOX CHOOOSER")
window.geometry("1000x700")

left_frame=Frame(window)
left_frame.grid(row=0,column=0,sticky=N)
right_frame=Frame(window)
right_frame.grid(row=0,column=1,sticky=N)

frame_chooser=LabelFrame(left_frame,text="Teminox chooser",container=False,height=900,width=900,bd=2)



def equipment_SSV_radiobutton(row_):
    Label(frame_complectation,text="SSV").grid(row=row_,column=0,sticky=W)
    global status_SSV
    status_SSV=StringVar()
    Radiobutton(frame_complectation,variable=status_SSV,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_SSV,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_SSV,value="delete").grid(row=row_,column=3)
    status_SSV.set("main")

def equipment_gas_control_section(row_):
    Label(frame_complectation,text="Gas_control_section").grid(row=row_,column=0,sticky=W)
    global status_gas_control_section
    status_gas_control_section=StringVar()
    Radiobutton(frame_complectation,variable=status_gas_control_section,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_gas_control_section,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_gas_control_section,value="delete").grid(row=row_,column=3)
    status_gas_control_section.set("option")

def equipment_flame_monitoring_radiobutton(row_):
    Label(frame_complectation,text="Flame_monitoring").grid(row=row_,column=0,sticky=W)
    global status_flame_monitoring
    status_flame_monitoring=StringVar()
    Radiobutton(frame_complectation,variable=status_flame_monitoring,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_flame_monitoring,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_flame_monitoring,value="delete").grid(row=row_,column=3)
    status_flame_monitoring.set("main")

def equipment_seavis_radiobutton(row_):
    Label(frame_complectation,text="se@vis").grid(row=row_,column=0,sticky=W)
    global status_seavis
    status_seavis=StringVar()
    Radiobutton(frame_complectation,variable=status_seavis,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_seavis,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_seavis,value="delete").grid(row=row_,column=3)
    status_seavis.set("main")

def equipment_star_delta_radiobutton(row_):
    Label(frame_complectation,text="star-delta").grid(row=row_,column=0,sticky=W)
    global status_star_delta
    status_star_delta=StringVar()
    Radiobutton(frame_complectation,variable=status_star_delta,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_star_delta,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_star_delta,value="delete").grid(row=row_,column=3)
    status_star_delta.set("main")

def equipment_silencer_radiobutton(row_):
    Label(frame_complectation,text="silencer").grid(row=row_,column=0,sticky=W)
    global status_silencer
    status_silencer=StringVar()
    Radiobutton(frame_complectation,variable=status_silencer,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_silencer,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_silencer,value="delete").grid(row=row_,column=3)
    status_silencer.set("option")

def equipment_monoblock_vsd_radiobutton(row_):
    Label(frame_complectation,text="Monoblock VSD").grid(row=row_,column=0,sticky=W)
    global status_monoblock_vsd
    status_monoblock_vsd=StringVar()
    Radiobutton(frame_complectation,variable=status_monoblock_vsd,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_monoblock_vsd,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_monoblock_vsd,value="delete").grid(row=row_,column=3)
    status_monoblock_vsd.set("option")

def equipment_seavis_vsd_radiobutton(row_):
    Label(frame_complectation,text="Se@vis VSD").grid(row=row_,column=0,sticky=W)
    global status_seavis_vsd
    status_seavis_vsd=StringVar()
    Radiobutton(frame_complectation,variable=status_seavis_vsd,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_seavis_vsd,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_seavis_vsd,value="delete").grid(row=row_,column=3)
    status_seavis_vsd.set("option")

def equipment_vsd_radiobutton(row_):
    Label(frame_complectation,text="VSD").grid(row=row_,column=0,sticky=W)
    global status_vsd
    status_vsd=StringVar()
    Radiobutton(frame_complectation,variable=status_vsd,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_vsd,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_vsd,value="delete").grid(row=row_,column=3)
    status_vsd.set("option")

def equipment_load_feedback_radiobutton(row_):
    Label(frame_complectation,text="load_feedback").grid(row=row_,column=0,sticky=W)
    global status_load_feedback
    status_load_feedback=StringVar()
    Radiobutton(frame_complectation,variable=status_load_feedback,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_load_feedback,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_load_feedback,value="delete").grid(row=row_,column=3)
    status_load_feedback.set("option")

def equipment_bus_interface_radiobutton(row_):
    Label(frame_complectation,text="BUS_interface").grid(row=row_,column=0,sticky=W)
    global status_bus_interface
    status_bus_interface=StringVar()
    Radiobutton(frame_complectation,variable=status_bus_interface,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_bus_interface,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_bus_interface,value="delete").grid(row=row_,column=3)
    status_bus_interface.set("option")

def equipment_o2_trim_radiobutton(row_):
    Label(frame_complectation,text="O2_trim").grid(row=row_,column=0,sticky=W)
    global status_o2_trim
    status_o2_trim=StringVar()
    Radiobutton(frame_complectation,variable=status_o2_trim,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_o2_trim,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_o2_trim,value="delete").grid(row=row_,column=3)
    status_o2_trim.set("option")

def equipment_o2_system_radiobutton(row_):
    Label(frame_complectation,text="O2_system").grid(row=row_,column=0,sticky=W)
    global status_o2_system
    status_o2_system=StringVar()
    Radiobutton(frame_complectation,variable=status_o2_system,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_o2_system,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_o2_system,value="delete").grid(row=row_,column=3)
    status_o2_system.set("option")

def equipment_booster_station_radiobutton(row_):
    Label(frame_complectation,text="Booster station").grid(row=row_,column=0,sticky=W)
    global status_booster_station
    status_booster_station=StringVar()
    Radiobutton(frame_complectation,variable=status_booster_station,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_booster_station,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_booster_station,value="delete").grid(row=row_,column=3)
    status_booster_station.set("main")

def equipment_flow_meter_radiobutton(row_):
    Label(frame_complectation,text="Flow meter").grid(row=row_,column=0,sticky=W)
    global status_flow_meter
    status_flow_meter=StringVar()
    Radiobutton(frame_complectation,variable=status_flow_meter,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_flow_meter,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_flow_meter,value="delete").grid(row=row_,column=3)
    status_flow_meter.set("main")

def equipment_fuel_fitting_station_radiobutton(row_):
    Label(frame_complectation,text="Fuel fitting station").grid(row=row_,column=0,sticky=W)
    global status_fuel_fitting_station
    status_fuel_fitting_station=StringVar()
    Radiobutton(frame_complectation,variable=status_fuel_fitting_station,value="main").grid(row=row_,column=1)
    Radiobutton(frame_complectation,variable=status_fuel_fitting_station,value="option").grid(row=row_,column=2)
    Radiobutton(frame_complectation,variable=status_fuel_fitting_station,value="delete").grid(row=row_,column=3)
    status_fuel_fitting_station.set("main")

def calculate():
    efficiency=float(entry_boiler_efficiency.get())
    if combo_type_of_boiler.get() in ("hot water","TOH"):
        boiler_capacity=float(entry_boiler_capacity.get())
        entry_capacity.delete(0,END)
        entry_capacity.insert(END,str(round(100*boiler_capacity/efficiency,2)))
    if combo_type_of_boiler.get()=="steam":
        steam_rate=float(entry_boiler_capacity.get())
        entry_capacity.delete(0,END)
        entry_capacity.insert(END,str(round(100*steam_rate*0.65/efficiency,2)))



frame_chooser.grid(row=0,column=0,padx=10,pady=10,ipadx=10,ipady=5,sticky="ne")
combo_fuel=Combobox(frame_chooser,values=("NG","NG/LFO","LFO"),width=10)
combo_fuel.set("NG")
combo_fuel.grid(row=0,column=3,padx=5,pady=1)
combo_mono_duo=Combobox(frame_chooser,values=("Mono","Duo"),width=10)
combo_mono_duo.set("Mono")
combo_mono_duo.grid(row=1,column=3,padx=5,pady=1)
Label(frame_chooser,text="Enter capacity of burner, MW:").grid(row=0,column=0,padx=5,pady=1,sticky=W)
entry_capacity=Entry(frame_chooser)
entry_capacity.grid(row=0,column=1,padx=5,pady=1)
Label(frame_chooser,text="Enter resistance, mbar:").grid(row=1,column=0,padx=5,pady=1,sticky=W)
entry_resistance=Entry(frame_chooser)
entry_resistance.grid(row=1,column=1,padx=5,pady=1)

Button(frame_chooser,text="choose",command=main).grid(row=2,column=1,padx=5,pady=1)
label_chosen_burner=Label(frame_chooser,text="Burner")
label_enter_diameter=Label(frame_chooser,text="Enter diameter, mm:")
label_enter_lenght=Label(frame_chooser,text="Enter length, mm:")
entry_diameter=Entry(frame_chooser)
entry_length=Entry(frame_chooser)
button_check_dimensions=Button(frame_chooser,text="check dimensions?",command=check_dimensions)
button_make_quotation=Button(left_frame,text="make quotation",command=new_project)
Button(frame_chooser,text="Calculate",command=calculate).grid(row=0,column=2,padx=5,pady=1,sticky="w")



label_chosen_burner.grid(row=3,column=1)
frame_boiler_data=LabelFrame(right_frame,text="boiler data",height=200,width=500,bd=2,padx=5,pady=1)
entry_boiler_name=Entry(frame_boiler_data)
entry_boiler_capacity=Entry(frame_boiler_data)
entry_boiler_efficiency=Entry(frame_boiler_data)
combo_type_of_boiler=Combobox(frame_boiler_data,values=("hot water","steam","TOH"),width=10)
combo_type_of_furnace=Combobox(frame_boiler_data,values=("watertube","flametube","TOH"),width=10)
combo_type_of_ftb_furnace=Combobox(frame_boiler_data,values=("3-pass","2-pass","inverse"),width=10)
entry_furnace_length=Entry(frame_boiler_data,width=6)
entry_furnace_width=Entry(frame_boiler_data,width=6)
entry_furnace_height=Entry(frame_boiler_data,width=6)


def boiler_data():

    
    
    
    combo_type_of_boiler.set("hot water")
    frame_boiler_data.grid(row=0,column=1, padx=10, pady=10, ipadx=10, ipady=5,sticky="ne")
    Label(frame_boiler_data,text="Enter boiler type:").grid(row=0,column=0,padx=5,pady=1,sticky=W)
    combo_type_of_boiler.grid(row=0,column=1,padx=5,pady=1,sticky=W)
    lb1=Label(frame_boiler_data,text="Enter boiler's capacity, MW:")
    lb1.grid(row=1,column=0,padx=5,pady=1,sticky=W)
    entry_boiler_capacity.grid(row=1,column=1,padx=5,pady=1)
    Label(frame_boiler_data,text="Enter efficiency, %:").grid(row=2,column=0,padx=5,pady=1,sticky=W)
    Label(frame_boiler_data,text="Enter boiler name:").grid(row=3,column=0,padx=5,pady=1,sticky=W)
    entry_boiler_efficiency.grid(row=2,column=1,padx=5,pady=1,sticky=W)
    entry_boiler_name.grid(row=3,column=1,padx=5,pady=1)
    Label(frame_boiler_data,text="Enter type of furnace:").grid(row=4,column=0,padx=5,pady=1,sticky=W)
    combo_type_of_furnace.set("flametube")
    combo_type_of_furnace.grid(row=4,column=1,padx=5,pady=1,sticky=W)
    combo_type_of_ftb_furnace.grid(row=5,column=1,padx=5,pady=1,sticky=W)
    combo_type_of_ftb_furnace.set("3-pass")
    
    def change_type_of_boiler(ivent):
        if combo_type_of_boiler.get() in ("hot water","TOH"):
            lb1.config(text="Enter boiler's capacity, MW:")
            return
        if combo_type_of_boiler.get()=="steam":
            lb1.config(text="Enter steam rate, t/h:")

    def change_type_of_furnace(ivent):
        for widget in frame_boiler_data.grid_slaves():
            if widget.grid_info()["row"]>=5:
                widget.grid_remove()
        if combo_type_of_furnace.get()=="flametube":
            combo_type_of_ftb_furnace.grid(row=5,column=1,sticky=W)
        elif combo_type_of_furnace.get()=="watertube":
            Label(frame_boiler_data,text="L, mm:").grid(row=5,column=0,padx=5,pady=1,sticky=E)
            entry_furnace_length.grid(row=5,column=1,padx=5,pady=1,sticky=W)
            Label(frame_boiler_data,text="W, mm:").grid(row=6,column=0,padx=5,pady=1,sticky=E)
            entry_furnace_width.grid(row=6,column=1,padx=5,pady=1,sticky=W)
            Label(frame_boiler_data,text="H, mm:").grid(row=7,column=0,padx=5,pady=1,sticky=E)
            entry_furnace_height.grid(row=7,column=1,padx=5,pady=1,sticky=W)


    combo_type_of_boiler.bind("<<ComboboxSelected>>",change_type_of_boiler)
    combo_type_of_furnace.bind("<<ComboboxSelected>>",change_type_of_furnace)

    
boiler_data()



frame_common_data=LabelFrame(left_frame,text="common data",height=200,width=500,bd=2,padx=5,pady=5)
label_company=Label(frame_common_data,text="Company")
entry_company=Entry(frame_common_data)
entry_company.insert(END,"ООО «»")
label_number=Label(frame_common_data,text="Number of offer")
entry_number=Entry(frame_common_data)
label_delivery_time=Label(frame_common_data,text="Delivery time, weeks")
entry_delivery_time=Entry(frame_common_data)
entry_delivery_time.insert(END,"17-19")
entry_discount=Entry(frame_common_data)
entry_discount.insert(END,"30")
combo_number_of_burners=Combobox(frame_common_data,values=(1,2,3,4,5,6,7,8,9,10),width=3)


button_ready=Button(right_frame,text="Ready!",command=quotation_maker,width=20,height=5,bd=3,font=22)
frame_complectation=LabelFrame(right_frame,text="complectation",height=200,width=500,bd=2,padx=5,pady=5)

def choose_path():
    file_footer=filedialog.askdirectory()
    with open("directory.txt","w") as f:
        f.write(file_footer)

def set_footer():
    footer=entry_footer.get()
    with open("footer.txt","w",encoding="utf-8") as f:
        f.write(footer)
    footer_window.destroy()

def choose_footer():
    global footer_window
    footer_window=Toplevel(window)
    footer_window.title("Enter your first and last name")
    footer_window.geometry("300x100")
    global entry_footer
    entry_footer=Entry(footer_window,width=25)
    with open("footer.txt","r",encoding="utf-8") as f:
        footer=f.read()
    entry_footer.insert(END,footer)
    entry_footer.grid(row=0,column=0,padx=3,pady=3)
    button_set_footer=Button(footer_window,text="save",command=set_footer)
    button_set_footer.grid(row=0,column=1,padx=3,pady=3)


menu_bar=Menu(window)
window.config(menu=menu_bar)
settings_menu=Menu(menu_bar)
menu_bar.add_cascade(label="Settings",menu=settings_menu)
settings_menu.add_command(label="Save directory",command=choose_path)
settings_menu.add_command(label="Footer",command=choose_footer)



window.mainloop()


















