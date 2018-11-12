class Unit:
    def __init__(self, name, soacrow, slidesrow, slidesrowend):
        self.name = name
        self.soacrow = soacrow
        self.slidesrow = slidesrow
        self.slidesrowend = slidesrowend

        self.lastweek = 0.0
        self.monthtotal = 0.0
        self.wip = 0.0
        self.yearlytotal = 0.0
        self.quartertotal = 0.0
        self.monthgoal = 0.0
        self.quartergoal = 0.0
        self.yearlygoal = 0.0

        self.maintenance = 0
        self.service = 0
        self.conv = 0
        self.sold = 0

        self.mon = 0.0
        self.tue = 0.0
        self.wed = 0.0
        self.thu = 0.0
        self.fri = 0.0
        self.sat = 0.0
        self.sun = 0.0

        self.week1 = 0.0
        self.week2 = 0.0
        self.week3 = 0.0
        self.week4 = 0.0
        self.week5 = 0.0

        self.month1 = 0.0
        self.month2 = 0.0
        self.month3 = 0.0