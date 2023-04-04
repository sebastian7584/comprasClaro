from selenium.common.exceptions import NoSuchElementException, NoSuchFrameException, NoSuchWindowException
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from subprocess import CREATE_NO_WINDOW
from selenium.webdriver.support.ui import Select
import chromedriver_autoinstaller
from selenium import webdriver
from subprocess import Popen
from turtle import title
from tkinter import *
from tkinter import ttk
import time
from datetime import datetime
import pandas as pd
import subprocess
import sys
import threading


class Web_Controller:

    def __init__(self):
        self.chromedriver()
        self.openChrome()

    def chromedriver(self):
        chromedriver_autoinstaller.install()
    
    def validate(funcion):
        def execute(self,*args, **kwargs):
            proof = True
            contador = 1
            while proof:
                try:
                    data = funcion(self,*args, **kwargs)
                    proof= False
                    return data
                except:
                    if contador < 60:
                        print(f'intento numero {contador}')
                        time.sleep(1)
                        contador +=1
                    else:
                        raise('Excedio el numero de intentos')
        return execute
    
    def openChrome(self):
        service = ChromeService('chromedriver')
        options =  webdriver.ChromeOptions()
        self.browser = webdriver.Chrome(chrome_options= options)
    

    @validate
    def selectPage(self,link):
        self.browser.get(link)
    
    @validate
    def insert(self, byStr, text, by='xpath'):
        if by == "xpath": find = self.browser.find_element_by_xpath(byStr)
        elif by == "id": find = self.browser.find_element_by_id(byStr)
        elif by == "name": find = self.browser.find_element_by_name(byStr)
        else: find =None
        if find is not None:
            find.send_keys(text)
    
    @validate
    def click(self, byStr, by='xpath'):
        if by == "xpath": find = self.browser.find_element_by_xpath(byStr)
        elif by == "id": find = self.browser.find_element_by_id(byStr)
        elif by == "name": find = self.browser.find_element_by_name(byStr)
        else: find =None
        if find is not None:
            find.click()
    
    @validate
    def read(self, byStr, by='xpath'):
        if by == "xpath": find = self.browser.find_element_by_xpath(byStr)
        elif by == "id": find = self.browser.find_element_by_id(byStr)
        elif by == "name": find = self.browser.find_element_by_name(byStr)
        if find is not None:
            return find.text
        else: return "none"

    def readNoValidate(self, byStr, by='xpath'):
        if by == "xpath": find = self.browser.find_element_by_xpath(byStr)
        elif by == "id": find = self.browser.find_element_by_id(byStr)
        elif by == "name": find = self.browser.find_element_by_name(byStr)
        if find is not None:
            return find.text
        else: return "none"

    @validate
    def wait(self, byStr, condition ,by='xpath'):
        if by == "xpath": find = self.browser.find_element_by_xpath(byStr)
        elif by == "id": find = self.browser.find_element_by_id(byStr)
        elif by == "name": find = self.browser.find_element_by_name(byStr)
        if find is not None:
            if find.text is not None:
                if condition in find.text:
                    raise('error')
    
class ExcelConvert:

    def export(result):
        df = pd.DataFrame(result[1:], columns=result[0])
        df.to_excel('archivo_excel.xlsx', index=False)

class Navegacion:

    def __init__(self, user, password, init, end, root, label):
        self.link= 'https://190.144.217.66/Front_PortalComercial/controlseguridad/login-dos.asp'
        self.link2='https://190.144.217.66/Front_PortalComercial/Portal/consultas/con_factura.asp'
        self.link3='https://190.144.217.66/Front_PortalComercial/Portal/consultas/con_seriales_factura.asp'
        self.user= user
        self.password= password
        self.fechaIni = init
        self.fechaFin = end
        self.root = root
        self.label = label
        self.init()
        self.getFacturas()
        self.getSeriales()
        self.close()

    def init(self):
        self.navegador = Web_Controller()
        self.navegador.selectPage(self.link)
        self.navegador.click('details-button','id')
        self.navegador.click('proceed-link','id')
        self.navegador.insert('/html/body/section/form/input[1]', self.user)
        self.navegador.insert('password', self.password, 'id')
        self.navegador.insert('SelServicio', 'Pedidos en Línea', 'id')
        self.navegador.click('/html/body/section/form/button')
        self.navegador.insert('sel_regionlogin', 'Occidente', 'id')
        self.navegador.insert('SelOrgCanalSector', 'Kit Prepago', 'id')
        self.navegador.click('Button1', 'id')
        self.navegador.selectPage(self.link2)
        self.navegador.insert('/html/body/form/table/tbody/tr/td/table/tbody/tr[2]/td[1]/select', 'Factura')
        self.navegador.insert('FecIni', self.fechaIni, 'id')
        self.navegador.insert('FecFin', self.fechaFin, 'id')
        self.navegador.click('/html/body/form/table/tbody/tr/td/table/tbody/tr[3]/td[1]/table/tbody/tr/td[4]/input')
        data=self.navegador.wait('/html/body/form/table/tbody/tr/td/table/tbody/tr[7]/td/table/tbody/tr/td/div/table/tbody/tr/td', 'Informe los campos del filtro para hacer la seleccion')

    def getFacturas(self):
        self.facturas = []
        table = True
        group = 1
        conteo = 1
        while table:
            dataTable = True
            try:
                factura =     self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[7]/td/table/tbody/tr/td/div/table[{group}]/tbody/tr[2]/td[1]')
                fecha =       self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[7]/td/table/tbody/tr/td/div/table[{group}]/tbody/tr[2]/td[2]')
                vencimiento = self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[7]/td/table/tbody/tr/td/div/table[{group}]/tbody/tr[2]/td[6]')
                total =       self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[7]/td/table/tbody/tr/td/div/table[{group}]/tbody/tr[2]/td[7]')
                fila = 2
                renglones=[]
                while dataTable:
                    try:
                        codMaterial= self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[7]/td/table/tbody/tr/td/div/table[{group+1}]/tbody/tr[{fila}]/td[1]')
                        producto=    self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[7]/td/table/tbody/tr/td/div/table[{group+1}]/tbody/tr[{fila}]/td[2]')
                        cantidad=    self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[7]/td/table/tbody/tr/td/div/table[{group+1}]/tbody/tr[{fila}]/td[3]')
                        unidad=      self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[7]/td/table/tbody/tr/td/div/table[{group+1}]/tbody/tr[{fila}]/td[4]')
                        valor=       self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[7]/td/table/tbody/tr/td/div/table[{group+1}]/tbody/tr[{fila}]/td[5]')
                        posicion=    self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[7]/td/table/tbody/tr/td/div/table[{group+1}]/tbody/tr[{fila}]/td[6]')
                        condicion=   self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[7]/td/table/tbody/tr/td/div/table[{group+1}]/tbody/tr[{fila}]/td[7]')
                        desCondicion=self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[7]/td/table/tbody/tr/td/div/table[{group+1}]/tbody/tr[{fila}]/td[8]')
                        subTotal=    self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[7]/td/table/tbody/tr/td/div/table[{group+1}]/tbody/tr[{fila}]/td[9]')
                        renglon= {'codMaterial':codMaterial, 'producto':producto, 'cantidad':cantidad, 'unidad':unidad, 'valor':valor, 'posicion':posicion, 'condicion':condicion, 'desCondicion':desCondicion, 'subTotal':subTotal}
                        renglones.append(renglon)
                        fila +=1
                    except:
                        dataTable = False
                item = {'factura':factura, 'fecha':fecha, 'vencimiento':vencimiento, 'total':total, 'renglones':renglones}
                self.facturas.append(item)
                group +=3
                self.label.config(text=conteo)
                self.root.update()
                conteo +=1
            
            except:
                table = False
    
    def getSeriales(self):
        self.navegador.selectPage(self.link3)

        self.result = [['serial','costoSinIva','codigo','producto','fecha','factura','tipo','vencimiento','iva','totalConIva']]
        for i in self.facturas:
            self.navegador.insert('text_Factura', i['factura'], 'id')
            self.navegador.click('/html/body/form/table/tbody/tr/td/table/tbody/tr[2]/td[2]/input')
            documentData = True
            fila = 2
            
            while documentData:
                try:
                    serial =   self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[4]/td/table/tbody/tr/td/div/table/tbody/tr[{fila}]/td[3]')
                    producto = self.navegador.readNoValidate(f'/html/body/form/table/tbody/tr/td/table/tbody/tr[4]/td/table/tbody/tr/td/div/table/tbody/tr[{fila}]/td[2]')
                    ivaTaza=0
                    descuento = 0
                    costo = 0
                    tipo= ''
                    for j in i['renglones']:
                        valor = j["valor"].replace(".","").replace(",",".")
                        cantidad = j['cantidad']
                        if producto == j['producto']:
                            codigo = j['codMaterial']
                            if j["desCondicion"] == 'IVA repercutido':
                                if valor == "0.00":
                                    ivaTaza = 0
                                elif valor == "19000.00":
                                    ivaTaza= 0.19
                                else: print(f'error con iva de {valor}')
                            elif j["desCondicion"] == 'Dcto Comercial':
                                descuento = int(valor.replace(".00",""))
                            elif j["desCondicion"] == 'Precio SIMCARD':
                                costo = int(valor.replace(".00",""))
                                tipo= 'SIM'
                            elif j["desCondicion"] == 'Prec sin IVA sin SIM':
                                costo = int(valor.replace(".00","")) 
                                tipo= 'KIT'
                            else:
                                print(f'error  {j["desCondicion"]} {valor}')
                    costoSinIva = costo + descuento
                    fecha = i['fecha']
                    factura = i['factura']
                    vencimiento = datetime.strptime(i['vencimiento'], '%d/%m/%Y').date() - datetime.strptime(fecha, '%d/%m/%Y').date()
                    vencimiento = vencimiento.days
                    iva= costoSinIva * ivaTaza
                    totalConIva= costoSinIva + iva
                    renglon = [str(serial).lstrip('0'),costoSinIva,str(codigo).lstrip('0'),producto,fecha,factura,tipo,vencimiento,iva,totalConIva]
                    self.result.append(renglon)
                    fila +=1    

                except Exception as e:
                    print(e)
                    documentData = False
    
    def close(self):
        self.navegador.browser.close()
        # self.navegador.quit()

class App:
    def __init__(self):
        self.root = Tk()
        self.root.title("Validacion Factura")
        self.colorTeam = '#E11419'
        self.root.geometry('300x175')
        self.root.config(bg='#fff')

        self.espacio1 = Label(self.root, text=" ", bg='white')
        self.espacio1.grid(row=0, column=0)

        self.label_usuario = Label(self.root, text="Usuario:", bg='white')
        self.label_usuario.grid(row=1, column=0)

        self.entry_usuario = Entry(self.root)
        self.entry_usuario.grid(row=1, column=1)

        self.label_clave = Label(self.root, text="Clave:", bg='white')
        self.label_clave.grid(row=2, column=0)

        self.entry_clave = Entry(self.root, show="*")
        self.entry_clave.grid(row=2, column=1)

        self.label_fecha_inicial = Label(self.root, text="Fecha inicial:", bg='white')
        self.label_fecha_inicial.grid(row=3, column=0)

        self.entry_fecha_inicial = Entry(self.root)
        self.entry_fecha_inicial.grid(row=3, column=1)

        self.label_fecha_final = Label(self.root, text="Fecha final:", bg='white')
        self.label_fecha_final.grid(row=4, column=0)

        self.entry_fecha_final = Entry(self.root)
        self.entry_fecha_final.grid(row=4, column=1)

        self.espacio2 = Label(self.root, text=" ", bg='white')
        self.espacio2.grid(row=5, column=0)

        self.button_imprimir = Button(self.root, text="Generar Informe", command=self.start_thread, bg=self.colorTeam, fg='white')
        self.button_imprimir.place(relx=0.3, rely= 0.67, relwidth = 0.33, relheight= 0.30)

    def start_thread(self):
        thread = threading.Thread(target=self.ejecuccion)
        thread.start()
       
    
    def ejecuccion(self):
        user = self.entry_usuario.get()
        password = self.entry_clave.get()
        fechaIni = self.entry_fecha_inicial.get()
        fechaFin = self.entry_fecha_final.get()
        navegacion = Navegacion(user,password,fechaIni,fechaFin, self.root, self.espacio1)
        ExcelConvert.export(navegacion.result)
        openFile()

class openFile:

    def __init__(self) -> None:
        p = Popen("openExcel.bat")
        stdout, stderr = p.communicate()


# user='cjgomez'
# password='123456'
# fechaIni = '27/03/2023'
# fechaFin = '27/03/2023'

root = App()
root.root.mainloop()



