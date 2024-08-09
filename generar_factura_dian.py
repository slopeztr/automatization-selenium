from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time, csv, traceback, sys, re
from datetime import datetime


WEB_DRIVER_PATH = '/usr/local/chromedriver'

class AutomaticDianBilling( object ):

    def __init__( self, csvfile= '' ):
        self.driver = webdriver.Chrome(
          executable_path= WEB_DRIVER_PATH,
        )
        self.oprint( 'Webdriver inicializado' )

        self.ciudad2departamento = self.import_ciudades_departamentos()
        self.oprint( 'Ciudades cargadas' )

        self.login()
        self.oprint( 'Sesión abierta' )

        self.data = self.import_csv( csvfile )
        self.oprint( f'Datos CSV importados rows={len(self.data)}' )

        self.data_validator()
        self.oprint( 'Datos importados válidos' )

        self.generate()
        self.oprint( 'Datos generados a la DIAN' )

        self.driver.quit()

    def generate( self ):
        for i, record in enumerate( self.data ) :
            saved = False
            max_attempt = 2
            attempt_count = 0

            while not saved :
                try :
                    if attempt_count >= max_attempt : 
                        raise Exception( "Numero de intentos superados" )

                    self.form( record )
                    self.eprint( f"Registro facturado pedido={record[ 'Número de pedido' ]}" )

                    saved = True
                except TimeoutException :
                    attempt_count += 1
                    time.sleep( 5 )
                except Exception :
                    self.eprint( f"Generando factura row={i+2} pedido={record[ 'Número de pedido' ]}" )
                    self.eprint( traceback.format_exc() )
                    break

    def clear( self, element_name ) :
        self.driver.find_element_by_name( 
            element_name
        ).clear()

    def set_value( self, element_name, value ) :
        self.clear( element_name )

        self.driver.find_element_by_name(
            element_name
        ).send_keys( value )


    def select_index( self, element_name, index ) :
        Select( 
            self.driver.find_element_by_name(
                element_name
            )
        ).select_by_index( index )

    def select_value( self, element_name, value ) :
        Select( 
            self.driver.find_element_by_name(
                element_name
            )
        ).select_by_value( value )


    def select_product( self, product ) :
        # Detalles de Productos
        self.driver.find_element_by_css_selector(
            'a.ioFacturo-button'
        ).click()

        self.driver.find_element_by_css_selector(
            'div#ProductsTable_filter.dataTables_filter label input[type="search"]'
        ).send_keys( product )

        self.driver.find_element_by_link_text( product ).click()

        WebDriverWait( self.driver, 10 ).until( 
            EC.invisibility_of_element_located(
                (    
                    By.CSS_SELECTOR, 
                    'div#ModalProductsTablemodal.fade.ioFacturo-Modal.in'
                )
            )
        )

        return True


    def form( self, record ):
        self.driver.get( 
            'https://facturaciongratuita.dian.gov.co/IoFacturo/Documents/FacturaElectronica' 
        )

        # Datos del documento
        driver = self.driver

        fecha = datetime.strptime(
            record[ 'Fecha del pedido' ], 
            '%Y-%m-%d'
        ).strftime('%d/%m/%Y')

        self.set_value( 'Documento.Encabezado.IdDoc.FechaEmis', fecha )
        self.set_value( 'Documento.Encabezado.IdDoc.FechaVenc', fecha )

        self.select_index( 'Documento.Encabezado.IdDoc.MedioPago', 2 )
        self.select_index( 'Documento.Encabezado.IdDoc.Serie', 1 )
        self.select_index( 'Documento.Encabezado.IdDoc.TipoNegociacion', 1 )
        self.select_index( 'Documento.Encabezado.IdDoc.Plazo', 1 )

        driver.execute_script( 
            'arguments[0].click();', 
            driver.find_element_by_name(
                'Documento.Encabezado.IdDoc.IVANotIncluded'
            )
        )

        # Datos del emisor
        self.select_index( 'Documento.Encabezado.Emisor.RegimenContable', 2 )
        self.set_value( 'Documento.Encabezado.Emisor.DomFiscal.Ciudad', 'Itagui' )
        self.set_value( 'Documento.Encabezado.Emisor.DomFiscal.Calle', 'Calle 52 #49-23' )

        self.select_index( 'Documento.Encabezado.Emisor.DomFiscal.Departamento', 2 )
        driver.implicitly_wait( 10 )

        self.select_index( 'Documento.Encabezado.Emisor.DomFiscal.Municipio', 57 )


        # Datos del Receptor
        tipo_dni = self.word_normalizer( record[ 'Tipo DNI' ] )

        self.select_index( 
            'Documento.Encabezado.Receptor.DocRecep.TipoDocRecep', 
            ( 2 if tipo_dni == 'CEDULA' else 5 ) 
        )
        self.set_value( 'Documento.Encabezado.Receptor.DocRecep.NroDocRecep', record[ 'DNI' ] )

        self.select_index( 
            'Documento.Encabezado.Receptor.TipoContribuyenteR',
            ( 1 if tipo_dni == 'CEDULA' else 0 ) 
        )

        if tipo_dni == 'NIT':
            self.select_index( 'Documento.Encabezado.Receptor.RegimenContableR', 1 )

        ciudad = self.city_normalizer( record[ 'Ciudad (facturación)' ] )

        self.set_value( 'Documento.Encabezado.Receptor.NombreRecep.PrimerNombre', record[ 'Nombre (facturación)' ] )
        self.set_value( 'Documento.Encabezado.Receptor.NmbRecep', record[ 'Nombre (facturación)' ] )
        self.set_value( 'Documento.Encabezado.Receptor.DomFiscalRcp.Ciudad', ciudad )
        self.select_index( 'Documento.Encabezado.Receptor.DomFiscalRcp.Pais', 45 )

    
        self.select_value(
            'Documento.Encabezado.Receptor.DomFiscalRcp.Departamento',
            self.ciudad2departamento[ ciudad ]
        )

        driver.implicitly_wait( 10 )
        self.select_value(
            'Documento.Encabezado.Receptor.DomFiscalRcp.Municipio',
            ciudad
        )

        self.set_value( 'Documento.Encabezado.Receptor.ContactoReceptor[0].eMail', record[ 'Correo electrónico (facturación)' ] )
        self.set_value( 'Documento.Encabezado.Receptor.ContactoReceptor[0].Telefono', record[ 'Teléfono (facturación)' ] )
        self.select_product( record[ 'Product Id' ] )

        driver.execute_script( "arguments[0].value = ''", 
            self.driver.find_element_by_name(
                'Documento.Detalle[0].QtyItem'
            )
        )
        self.set_value( 'Documento.Detalle[0].QtyItem', record[ 'Artículo #' ] )

        driver.execute_script( "arguments[0].value = ''", 
            self.driver.find_element_by_name(
                'Documento.Detalle[0].PrcBrutoItem'
            )
        )

        self.set_value( 'Documento.Detalle[0].PrcBrutoItem', record[ 'Importe total del pedido' ] )

        WebDriverWait( driver, 10 ).until( 
        EC.element_to_be_clickable(
            (
            By.NAME, 
            'Documento.Detalle[0].MontoTotalItem'
            )
        )
        ).click()
        

        if tipo_dni == 'NIT':
            self.set_value( 'Personalizados.DocPersonalizado.campoString[0].Value', 
                """
                Servicio educativo excluido de IVA, según Art 476 del E. Trib.
                - Servicios de educación prestados por establecimientos reconocidos por el MEN.
                Res Nº 11813 de mar 24 de 2017 Itaguí 
                - ETDH geard. No somos autorretenedores. Ret 2.5% 
                """ 
            )

        driver.execute_script( 
            'arguments[0].click();', 
            driver.find_element_by_css_selector(
                'div.form-group.col-sm-6.col-md-6.col-lg-6 button.btn.btn-primary.ioFacturo-btn'
            )
        )

        # WebDriverWait( driver, 30 ).until( 
        #     EC.element_to_be_clickable(
        #         (
        #             By.CSS_SELECTOR, 
        #             'button#btn-save.btn.btn-medium.btn-success.ioFacturo-btn'
        #         )
        #     )
        # ).click()


        # WebDriverWait( driver, 60*4 ).until(
        #     EC.url_to_be( 
        #         'https://facturaciongratuita.dian.gov.co/Document/Sent' 
        #     ) 
        # )


    def login( self ):
        try :
            driver = self.driver
            driver.get( 
              'https://facturaciongratuita.dian.gov.co/Account/Login' 
            )

            driver.find_element_by_name(
              'UserName'
            ).send_keys( 'johndoe@gmail.com' )

            driver.find_element_by_name(
              'Password'
            ).send_keys( 'KQ49HYZG' )

            driver.find_element_by_id(
              'login-btn'
            ).click()

            assert( driver.current_url == 'https://facturaciongratuita.dian.gov.co/' )
        except Exception as e :
            self.eprint( f'Iniciando sesión error={str(e)}') 

            sys.exit()


    def import_ciudades_departamentos( self ) :
        ciudades = {}

        try:
            data = csv.reader(
              open( 
                'departamentos.csv',
                newline=''
              ),
              delimiter=','
            )

            for row in data :
                ciudades[ row[0] ] = row[1]  
                    
        except Exception as e:
            self.eprint( f'Importando ciudades error={str(e)}' )

            sys.exit()
        
        return ciudades


    def city_normalizer( self, city ):
        city = self.word_normalizer( city )
        city = re.sub( r'^BOGOTA$', 'BOGOTA, D.C.', city )
        city = re.sub( r'^SANTIAGO DE CALI$', 'CALI', city )
        city = re.sub( r'^(DESCONOCIDO)$', 'ITAGUI', city )

        if not city in self.ciudad2departamento :
            city = 'ITAGUI'

        return city

    def word_normalizer( self, word ):
        word = word.upper().translate( 
            str.maketrans( 
                { 'Á':'A', 'É':'E', 'Í':'I', 'Ó':'O', 'Ú': 'U' }
            )
        )

        return word
    
    def import_csv( self, csvfile ):
        try:
            data = csv.DictReader(
              open( 
                csvfile,
                newline='' 
              ),
              delimiter=','
            )
          
        except Exception as e :
            self.eprint( f'Importando CSV error={str(e)}' )

            sys.exit()
        
        return list( data )


    def data_validator( self ):
        columns_format = {
            'Número de pedido': r'^\d+$',
            'Estado del pedido': r'^Completado$',
            'Fecha del pedido': r'^\d\d\d\d-\d\d-\d\d$',
            'Nombre (facturación)': r'^.+$',
            'Tipo DNI': r'^(Cédula|NIT)$',
            'DNI': r'^\d+$',
            'Ciudad (facturación)': r'^.+$',
            'Correo electrónico (facturación)': r'^.+$',
            'Teléfono (facturación)': r'^\d+$',
            'Importe total del pedido': r'^\d+$',
            'Product Id': r'^\d+$'
        }

        driver = self.driver
        driver.get( 
            'https://facturaciongratuita.dian.gov.co/IoFacturo/Product' 
        )

        #hack para cambiar el valor de la opcion
        driver.execute_script(
            "arguments[0].setAttribute('value', '9999')",
            driver.find_element_by_css_selector(
                '#ProductsTable_length select > option:nth-child(2)'
            )
        )

        self.select_index( 'ProductsTable_length', 1 )
        time.sleep(5)

        td = WebDriverWait( driver, 10 ).until(
            EC.presence_of_all_elements_located( 
                ( 
                    By.CSS_SELECTOR,
                    '#ProductsTable tbody > tr > td:nth-child(2)'
                )
            ) 
        )

        products_id = [ element.text for element in td ]
        for i, record in enumerate( self.data ) :
            
            if record[ 'Product Id' ] not in products_id :
                self.eprint( f"No se encontró el producto ID={record[ 'Product Id' ]}" )

                sys.exit()


            for col,value in record.items() :
                if col in columns_format :
                    if re.match( columns_format[ col ], value ) is None :
                        self.eprint( f"Error de formato en col={col} row={i+2}" )

                        sys.exit()

    def oprint( self, msg ):
        print( datetime.now().strftime( '%Y-%m-%d %H:%M:%S' ), msg )

    def eprint( self, msg ):
        print( datetime.now().strftime( '%Y-%m-%d %H:%M:%S' ), msg, file=sys.stderr )



if __name__ == "__main__":
    AutomaticDianBilling( 'facturas.csv' )