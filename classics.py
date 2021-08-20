import math

class Burner:
    def __init__(self, type_, size, turndown_ratio, motor_rating, capacity_range, minload, maxload, statpressure):
        self.type_=type_
        self.size=size
        self.turndown_ratio=turndown_ratio
        self.motor_rating=motor_rating
        self.capacity_range=capacity_range
        self.minload=minload
        self.maxload=maxload
        self.statpressure=statpressure
    def __str__(self):
        return self.type_+" "+self.size
    def point_pressure(self,capacity):
        for i in range (len(self.statpressure)-1):
            if capacity>=self.statpressure[i][0] and capacity<=self.statpressure[i+1][0]:
                point=self.statpressure[i+1][1]+(self.statpressure[i][1]-self.statpressure[i+1][1])*(
                    self.statpressure[i+1][0]-capacity)/(self.statpressure[i+1][0]-self.statpressure[i][0])
                break
        return point

   


G_50_7=Burner("G", "50-7,5", ("1:7",None), "7,5", [(0.5,5.0),None], 2, 3.5, [(2,31.5),(2.5,31.5),(3.55,20)])
G_50_11=Burner("G", "50-11", ("1:7",None), "11", [(0.5,5.0),None], 2, 5, [(2,31.5),(3.55,31.5),(5,13.5)])
G_50_15=Burner("G", "50-15", ("1:7",None), "15", [(0.5,5.0),None], 2, 5, [(2,40),(3.55,40),(5,22.5)])
G_70_15=Burner("G", "70-15", ("1:7",None), "15", [(0.7,7.0),None], 4, 6.3, [(4,26),(5,26),(6.3,13.5)])
G_70_18=Burner("G", "70-18", ("1:7",None), "18,5", [(0.7,7.0),None], 4, 7, [(4,30.5),(5,30.5),(7,12.5)])
G_70_22=Burner("G", "70-22", ("1:7",None), "22", [(0.7,7.0),None], 4, 7, [(4,39),(5,39),(7,20.5)])
G_100_18=Burner("G", "100-18", ("1:7",None), "18,5", [(1.0,10.0),None], 6, 8.9, [(6,30),(7.1,30),(8.9,12)])
G_100_22=Burner("G", "100-22", ("1:7",None), "22", [(1.0,10.0),None], 6, 10, [(6,37.5),(7.1,37.5),(10,11)])
G_100_30=Burner("G", "100-30", ("1:7",None), "30", [(1.0,10.0),None], 6, 10, [(6,51.5),(7.1,51.5),(10,27.5)])
G_140_30=Burner("G", "140-30", ("1:7,5",None), "30", [(1.4,14.1),None], 8, 12.6, [(8,30),(10,30),(12.6,14)])
G_140_37=Burner("G", "140-37", ("1:7,5",None), "37", [(1.4,14.1),None], 8, 14.1, [(8,47.5),(10,47.5),(14.1,16)])
G_140_45=Burner("G", "140-45", ("1:7,5",None), "45", [(1.4,14.1),None], 8, 14.1, [(8,54),(10,54),(14.1,25)])
G_200_37=Burner("G", "200-37", ("1:8",None), "37", [(1.8,20.0),None], 12, 15.9, [(12,37.5),(14.2,37.5),(15.9,26)])
G_200_45=Burner("G", "200-45", ("1:8",None), "45", [(1.8,20.0),None], 12, 17.8, [(12,37.5),(14.2,37.5),(17.8,26)])
G_200_55=Burner("G", "200-55", ("1:8",None), "55", [(1.8,20.0),None], 12, 20, [(12,58),(14.2,58),(20,22.5)])
G_200_75=Burner("G", "200-75", ("1:8",None), "75", [(1.8,20.0),None], 12, 20, [(12,72),(14.2,72),(20,44)])
G_280_75=Burner("G", "280-75", ("1:8",None), "75", [(2.5,28.2),None], 17, 28.2, [(17,49),(20,49),(28.2,12)])

GL_50_7=Burner("GL", "50-7,5", ("1:7","1:3"), "7,5", [(0.5,5.0),(1.4,5.0)], 2, 3.5, [(2,28.5),(2.5,28.5),(3.55,17)])
GL_50_11=Burner("GL", "50-11", ("1:7","1:3"), "11", [(0.5,5.0),(1.4,5.0)], 2, 5, [(2,28.5),(3.55,28.5),(5,8)])
GL_50_15=Burner("GL", "50-15", ("1:7","1:3"), "15", [(0.5,5.0),(1.4,5.0)], 2, 5, [(2,37.5),(3.55,37.5),(5,17)])
GL_70_15=Burner("GL", "70-15", ("1:7","1:3"), "15", [(0.7,7.0),(1.8,7.0)], 4, 6.3, [(4,23),(5,23),(6.3,9.5)])
GL_70_18=Burner("GL", "70-18", ("1:7","1:3"), "18,5", [(0.7,7.0),(1.8,7.0)], 4, 7, [(4,27.5),(5,27.5),(7,7.5)])
GL_70_22=Burner("GL", "70-22", ("1:7","1:3"), "22", [(0.7,7.0),(1.8,7.0)], 4, 7, [(4,36),(5,36),(7,16)])
GL_100_18=Burner("GL", "100-18", ("1:7","1:3"), "18,5", [(1.0,10.0),(2.5,10.0)], 6, 8.9, [(6,27),(7.1,27),(8.9,8)])
GL_100_22=Burner("GL", "100-22", ("1:7","1:3"), "22", [(1.0,10.0),(2.5,10.0)], 6, 10, [(6,34),(7.1,34),(10,6.5)])
GL_100_30=Burner("GL", "100-30", ("1:7","1:3"), "30", [(1.0,10.0),(2.5,10.0)], 6, 10, [(6,48.5),(7.1,48.5),(10,22)])
GL_140_30=Burner("GL", "140-30", ("1:7,5","1:3"), "30", [(1.4,14.1),(3.3,14.1)], 8, 12.6, [(8,27.5),(10,27.5),(12.6,10)])
GL_140_37=Burner("GL", "140-37", ("1:7,5","1:3"), "37", [(1.4,14.1),(3.3,14.1)], 8, 14.1, [(8,44.5),(10,44.5),(14.1,11)])
GL_140_45=Burner("GL", "140-45", ("1:7,5","1:3"), "45", [(1.4,14.1),(3.3,14.1)], 8, 14.1, [(8,51.5),(10,51.5),(14.1,20.5)])
GL_200_37=Burner("GL", "200-37", ("1:8","1:3"), "37", [(1.8,20.0),(4.7,20.0)], 12, 15.9, [(12,34),(14.2,34),(15.9,22)])
GL_200_45=Burner("GL", "200-45", ("1:8","1:3"), "45", [(1.8,20.0),(4.7,20.0)], 12, 17.8, [(12,34),(14.2,34),(17.8,22)])
GL_200_55=Burner("GL", "200-55", ("1:8","1:3"), "55", [(1.8,20.0),(4.7,20.0)], 12, 20, [(12,55),(14.2,55),(20,17.5)])
GL_200_75=Burner("GL", "200-75", ("1:8","1:3"), "75", [(1.8,20.0),(4.7,20.0)], 12, 20, [(12,69),(14.2,69),(20,38.5)])
GL_280_75=Burner("GL", "280-75", ("1:8","1:3"), "75", [(2.5,28.2),(6.8,25.2)], 17, 28.2, [(17,46),(20,46),(28.2,7)])




Burners={"NG":[G_50_7,G_50_11,G_50_15,G_70_15,G_70_18,G_70_22,G_100_18,G_100_22,G_100_30,
                G_140_30,G_140_37,G_140_45,G_200_37,G_200_45,G_200_55,G_200_75,G_280_75],
        "NG/LFO":[GL_50_7,GL_50_11,GL_50_15,GL_70_15,GL_70_18,GL_70_22,GL_100_18,GL_100_22,
                GL_100_30,GL_140_30,GL_140_37,GL_140_45,GL_200_37,GL_200_45,GL_200_55,GL_200_75,GL_280_75]}


class SSV:
    def __init__(self, capacity, LHV=34.0):
        self.capacity=capacity
        self.LHV=LHV

    def flow(self):
        return round(3600*self.capacity/self.LHV,-2)
    def diameter(self):
        if 300<=self.flow()<=500:
            return "DN 65"
        if 500<self.flow()<=800:
            return "DN 80"
        if 600<self.flow()<=1200:
            return "DN 100"
        if 1200<self.flow()<=2000:
            return "DN 125"
        if 2000<self.flow()<=2900:
            return "DN 150"
    def inlet_pressure(self):
        if self.flow()<=1500:
            return 300
        if self.flow()<=2900:
            return 330
    def description(self):
        description=(f"Газовый защитный участок со встроенным стабилизатором давления в блочном исполнении для расхода газа {self.flow()} нм³/час макс.\n"
                    f"Входное давление газа – {self.inlet_pressure()} мбар, 500 мбар макс.\n" 
                    f"Состав: 2 отсечных клапана с электромагнитным приводом {self.diameter()}, реле давления газа макс., " 
                    f"устройство контроля герметичности газовых клапанов на базе реле давления газа мин., кнопка аварийного выключения, компенсатор тепловых расширений {self.diameter()}.")
        return description


class Booster_station:
    def __init__(self,capacity):
        self.capacity=capacity
    def max_capacity(self):
        if self.capacity<=4.8:
            return 4.8
        if self.capacity<=6.2:
            return 6.2
        if self.capacity<=9.5:
            return 9.5
        if self.capacity<=12.3:
            return 12.3
        if self.capacity<=16.8:
            return 16.8
        if self.capacity<=20.2:
            return 20.2
        if self.capacity<=26.9:
            return 26.9
    def size(self):
        if self.capacity<=4.8:
            return "19065"
        if self.capacity<=6.2:
            return "1949"
        if self.capacity<=9.5:
            return "1950"
        if self.capacity<=12.3:
            return "1951"
        if self.capacity<=16.8:
            return "1951-1"
        if self.capacity<=20.2:
            return "1952"
        if self.capacity<=26.9:
            return "1953"
    def motor_rating(self):
        if self.size() == "19065":
            return "1,5"
        if self.size() == "1949":
            return "2,2"
        if self.size() == "1950":
            return "3"
        if self.size() in ["1951","1951-1"]:
            return "4"
        if self.size() == "1952":
            return "5,5"
        if self.size() == "1953":
            return "7,5"
    def flow_meter_size(self):
        if self.size() in ["19065","1949","1950","1951"]:
            return "DN 20"
        if self.size() in ["1951-1","1952"]:
            return "DN 25"
        if self.size() == "1953":
            return "DN 40"




class Silencer:
    def __init__(self,burner):
        self.burner=burner


class Star_delta:
    def __init__(self,motor_rating):
        self.motor_rating=motor_rating



class VSD:
    def __init__(self,motor_rating):
        self.motor_rating=motor_rating