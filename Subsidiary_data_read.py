import xlrd
file_location = "G:/FAME/Subsidiaries/2.xlsx"

workbook = xlrd.open_workbook(file_location)

sheet = workbook.sheet_by_index(1)

class Vertex:
    def __init__(self,BvDnumber):
        self.ID = BvDnumber
        self.neighbors = list()
        self.network = 0

    def add_neighbor(self,v):
        if v not in self.neighbors:
            self.neighbors.append(v)
            self.neighbors.sort()
            

#list of companies
companies = []

#list of adjacencies
adjacencies = []

# read first row for keys  
keys = sheet.row_values(0)


for r in range(1, sheet.nrows):
    
    mother_company = sheet.cell(r,13)

    #add the mother company to the company list if new
    if mother_company.value == '':
        mother_company.value = previous_ID
    else:
        previous_ID = mother_company.value
        
    if mother_company.value not in companies:
        companies.append(mother_company.value)
        adjacencies.append(([],mother_company.value))
    
    #add the subsidiary to the mother company list of adjacencies 
    subsidiary = sheet.cell(r,15)
    
    if subsidiary.value not in adjacencies[companies.index(mother_company.value)][0]:
        adjacencies[companies.index(mother_company.value)][0].append(subsidiary.value)

#create the vertices in the Networks
vertices = []
for c in companies:
    v = Vertex(c)
    vertices.append((v,c))

print(vertices[:,0])
       
       
       
