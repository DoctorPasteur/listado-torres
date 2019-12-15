import xlwt

#Estas funciones limitan las entradas para que sean valores válidos.
def introducirEspesoresEnPulgadas():
    opcionesEspesoresEnPulgadas=["1/8", "3/16", "5/16", "1/4", "3/8", "1/2", "5/8", "1" ]
    Entrada = input("[+]Introduzca el espesor en fracción de pulgadas:    ")
    if not Entrada in opcionesEspesoresEnPulgadas:
        return introducirEspesoresEnPulgadas()
    return Entrada

def introducirAlaEnPulgadas():
    opcionesAlaEnPulgadas=["6", "5", "4", "3", "2-1/2", "2", "1-1/2", "1-1/4", "1", "3/4" ]
    Entrada = input("[+]Introduzca el ala en fracción de pulgadas:    ")
    try:
        if not Entrada in opcionesAlaEnPulgadas:
            return introducirAlaEnPulgadas()
        return Entrada
    except ValueError:
        print("introduzca un número válido")
        return introducirAlaEnPulgadas()

def introducirValorEntero(solicitud):
    try:
        valorEntero = int(input(solicitud))
        return valorEntero
    except ValueError:
        print("introduzca un número válido")
        return introducirValorEntero(solicitud)




def reemplazarFracciones(espesorPulg):
    espesorPulg = espesorPulg.replace('2-1/2"', "63.5")
    espesorPulg = espesorPulg.replace('1-1/2"', "38.1")
    espesorPulg = espesorPulg.replace('1-1/4"', "31.75")
    espesorPulg = espesorPulg.replace('1/8"', "3.175")
    espesorPulg = espesorPulg.replace('3/16"', "4.76")
    espesorPulg = espesorPulg.replace('5/16"', "7.94")
    espesorPulg = espesorPulg.replace('1/4"', "6.35")
    espesorPulg = espesorPulg.replace('3/8"', "9.52")
    espesorPulg = espesorPulg.replace('1/2"', "12.7")
    espesorPulg = espesorPulg.replace('5/8"', "15.87")
    espesorPulg = espesorPulg.replace('3/4"', "19.05")
    espesorPulg = espesorPulg.replace('7/8"', "22.22")
    espesorPulg = espesorPulg.replace('1"', "25.4")
    espesorPulg = espesorPulg.replace('2"', "50.8")
    espesorPulg = espesorPulg.replace('3"', "76.2")
    espesorPulg = espesorPulg.replace('4"', "101.8")
    espesorPulg = espesorPulg.replace('5"', "127")
    espesorPulg = espesorPulg.replace('6"', "152.4")

    #print(espesorPulg)
    return espesorPulg





class elementoAngular:
        def __init__(self, designacion, alaPulg, espesorPulg, cantidad, longitudMm):
            self.designacion = designacion
            self.alaPulg = str(alaPulg) + '"'
            self.espesorPulg = espesorPulg + '"'
            self.descripcion = 'L ' + self.alaPulg + " x " + self.alaPulg + ' x ' + espesorPulg + '"'

            self.alamm = reemplazarFracciones(self.alaPulg)
            self.espesormm = reemplazarFracciones(self.espesorPulg)
            
            self.longitud = int(longitudMm)
            self.cantidad = int(cantidad)

        def calcularPesoUnit(self):
            self.pesoUnit = ((((2 * float(self.alamm) - float(self.espesormm)) * float(self.espesormm))*self.longitud)/1000000000)*7850
            self.pesoUnit = round(self.pesoUnit, 2)

        def calcularPesoAcum(self):
            self.pesoAcum = self.pesoUnit*self.cantidad
            self.pesoAcum = round(self.pesoAcum, 2)


class elementoCartela:
        def __init__(self, designacion, ancho, largo, espesorPulg, cantidad):
            self.designacion = designacion
            self.ancho = int(ancho)
            self.largo = int(largo)
            self.espesorPulg = espesorPulg + '"'
            self.descripcion = 'PL ' + self.espesorPulg + " x " + str(self.ancho) + ' x ' + str(self.largo) + ' mm'

            self.espesormm = reemplazarFracciones(self.espesorPulg)
            
            self.cantidad = int(cantidad)

        def calcularPesoUnit(self):
            self.pesoUnit = ((float(self.ancho) * float(self.largo) * float(self.espesormm))/1000000000)*7850
            self.pesoUnit = round(self.pesoUnit, 2)

        def calcularPesoAcum(self):
            self.pesoAcum = self.pesoUnit * self.cantidad
            self.pesoAcum = round(self.pesoAcum, 2)

            



print("--PROGRAMA PARA LISTAR ELEMENTOS DE TORRE--")

ch1 = 1
ch2 = 5
listado = {}
indListado = 1

while ch1 != 4:
    print("\n+-+-+-+-+-+-+ INICIO +-+-+-+-+-+-+\n")
    print("1. AGREGAR ELEMENTO\n2. VER LISTA\n3. EXPORTAR LISTA\n4. Salir")
    ch1 = int(input("Elija una opción\t"))

    if ch1 == 1 :
        while ch2 != 4:
            print("\n+-+-+-+-+-+-+ AGREGAR ITEM+-+-+-+-+-+-+\n")
            print("1. ELEMENTO ANGULAR\n2. ELEMENTO CARTELA\n3. OTRO ELEMENTO\n4. VOLVER AL INICIO")
            ch2 = int(input("Elija una opción\t"))

            if ch2 == 1 :
                print("\n+-+-+-+-+-+-+ AGREGAR ELEMENTO ANGULAR+-+-+-+-+-+-+\n")
                designacion1 = str(input("[+]Introduzca la designacion del elemento:    "))
                alaPulgadas = introducirAlaEnPulgadas()
                espesorPulgadas = introducirEspesoresEnPulgadas()

                solicitudLongitudElemento = "[+]Introduzca la longitud del elemento en mm:    "
                longitudElemento = introducirValorEntero(solicitudLongitudElemento)
                
                solicitudCantidadElementos = "[+]Introduzca la cantidad de unidades:    "
                cantidadDeUnidades = introducirValorEntero(solicitudCantidadElementos)
                
                anguloTemp = elementoAngular(designacion1, alaPulgadas, espesorPulgadas, cantidadDeUnidades, longitudElemento)
                anguloTemp.calcularPesoUnit()
                anguloTemp.calcularPesoAcum()
                listado[indListado] = {"Designacion": anguloTemp.designacion, "Descripcion": anguloTemp.descripcion, "Longitud": anguloTemp.longitud, "Cantidad": anguloTemp.cantidad, "PesoUnitario": anguloTemp.pesoUnit, "PesoAcum": anguloTemp.pesoAcum}
                indListado += 1

            if ch2 == 2 :
                print("\n+-+-+-+-+-+-+ AGREGAR ELEMENTO CARTELA+-+-+-+-+-+-+\n")
                designacion1 = str(input("[+]Introduzca la designacion del elemento:    "))
                espesorPulg1 = introducirEspesoresEnPulgadas()

                solicitudAncho = "[+]Introduzca el ancho de la cartela en milimetros:    "
                ancho1 = introducirValorEntero(solicitudAncho)

                solicitudLargo = "[+]Introduzca la longitud de la cartela en mm:    "
                largo1 = introducirValorEntero(solicitudLargo)

                solicitudCantidadElementos = "[+]Introduzca la cantidad de unidades:    "
                cantidadDeUnidades = introducirValorEntero(solicitudCantidadElementos)

                cartelaTemp = elementoCartela(designacion1, ancho1, largo1, espesorPulg1, cantidadDeUnidades)
                cartelaTemp.calcularPesoUnit()
                cartelaTemp.calcularPesoAcum()
                listado[indListado] = {"Designacion": cartelaTemp.designacion, "Descripcion": cartelaTemp.descripcion, "Longitud": cartelaTemp.largo, "Cantidad": cartelaTemp.cantidad, "PesoUnitario": cartelaTemp.pesoUnit, "PesoAcum": cartelaTemp.pesoAcum}
                indListado += 1


#Esta es la parte del código que imprime el listado
    if ch1 == 2:
        print("\n+-+-+-+-+-+-+LISTADO DE ELEMENTOS+-+-+-+-+-+-+\n")
        for id, info in listado.items():
            #print("\nID:", id)
            print(f"Designación: {listado[id]['Designacion']}")
            print(f"Descripción: {listado[id]['Descripcion']}")
            print(f"Longitud: {listado[id]['Longitud']}")
            print(f"Cantidad: {listado[id]['Cantidad']}")
            print(" ")
            #for key in info:
                #print(key + ':', info[key])
    if ch1 == 3:
        #LAS LINEAS DE ABAJO DETERMINAN EL NOMBRE DE LA HOJA Y DEL ARCHIVO
        nomArchivo = str(input("[+] Introduzca el nombre del archivo:    "))
        nomArchivo = nomArchivo + '.xls'
        nomHoja = str(input("[+] Introduzca el nombre de la hoja:    "))
        wb = xlwt.Workbook()
        ws = wb.add_sheet(nomHoja)
        ws.write(0, 0, 'DESIGNACION')
        ws.write(0, 1, 'DESCRIPCION')
        ws.write(0, 2, 'LONGITUD')
        ws.write(0, 3, 'CANTIDAD')
        ws.write(0, 4, 'PESO UNITARIO')
        ws.write(0, 5, 'PESO ACUMULADO')

        for id, info in listado.items():
            designacionSal = listado[id]['Designacion']
            descripcionSal = listado[id]['Descripcion']
            longitudSal = listado[id]['Longitud']
            cantidadSal = listado[id]['Cantidad']
            pesounitSal = listado[id]['PesoUnitario']
            pesoacumSal = listado[id]['PesoAcum']
            ws.write(id, 0, designacionSal)
            ws.write(id, 1, descripcionSal)
            ws.write(id, 2, longitudSal)
            ws.write(id, 3, cantidadSal)
            ws.write(id, 4, pesounitSal)
            ws.write(id, 5, pesoacumSal)

        wb.save(nomArchivo)
            #print("\nID:", id)
            #print(listado[id]['Designacion'])
        



# class Student:
#     name = 'abc'
#     roll_no = 1

#     def print_info (self):
#         print(Student.name,"\t",Student.roll_no)

#     def input_info (self):
#         Student.roll_no = int(input("Enter Student roll number  "))
#         Student.name = input("Enter Student name  ")

# ch=1
# students = []

# while ch!=3:
#     print("1.Create new Student\n2. Display students\n3.Exit")
#     ch=int(input("Enter choice\t"))`enter code here`
#     if ch==1 :
#         tmp=Student()
#         tmp.input_info()
#         students.append(tmp)
#     elif ch==2:
#         print("---Students Details---")
#         print("Name\tRoll No")
#     for i in students:
#         i.print_info()