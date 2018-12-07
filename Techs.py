class Tech:
    def __init__(self, name, row, businessunit):
        self.name = name
        self.row = row
        self.businessunit = businessunit
        self.lastweek = 0.0
        self.monthtotal = 0.0
        self.wip = 0.0
        self.maintenance = 0
        self.service = 0
        self.convcalls = 0
        self.soldcalls = 0
        self.servconv = 0
        self.servsold = 0
        self.maintconv = 0
        self.maintsold = 0
        self.iaqconv = 0
        self.iaqcls = 0
        self.sdconv = 0
        self.sdcls = 0