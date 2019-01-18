class CSR:
    def __init__(self, firstname, lastname):
        self.firstname = firstname
        self.lastname = lastname
        self.name = firstname + " " +  lastname
        self.initials = firstname[0] + lastname[0]
        self.jobs = 0
        self.opps = 0
        self.sold = 0
        self.revenue = 0.0
        self.row = ""