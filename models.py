

class LoginInfo():
    def __init__(self, login_site:str, id:str, password:str, no_of_tabs:int):
        self.login_site = login_site
        self.id = id
        self.password = password

class ReservationInfo():
    def __init__(self, wanted_date: int, wanted_time1: int, wanted_time2: int, wanted_time3: int):
        self.wanted_date = wanted_date
        self.wanted_time1 = wanted_time1
        self.wanted_time2 = wanted_time2
        self.wanted_time3 = wanted_time3