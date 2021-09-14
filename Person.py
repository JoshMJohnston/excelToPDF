class Person:
    def __init__(self, name, dob, id):
        self.name = name
        self.dob = dob
        self.id = id
    
    def setName(self, name):
        self.name = name

    def setDob(self, dob):
        self.dob = dob

    def setId(self, id):
        self.id = id

    def getName(self):
        return self.name

    def getDob(self):
        return self.dob

    def getId(self):
        return self.id

    def toString(self):
        print(f"{self.name}\t{self.dob}\t{self.id}")