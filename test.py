from selenium import webdriver
from selenium.webdriver.support.select import Select
import time

WEB_DRIVER_PATH = '/usr/local/chromedriver'

def extract_columnaid_product():
  driver = webdriver.Chrome(
    executable_path= WEB_DRIVER_PATH,
  )

  login( driver )  

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

  Select( 
    driver.find_element_by_name(
      'ProductsTable_length'
    )
  ).select_by_index( 1 )

  time.sleep(15)

  tds = driver.find_elements_by_css_selector(
    '#ProductsTable tbody > tr > td:nth-child(2)'
  )

  for td in tds: 
    print(td.text)  

  

  time.sleep(20000)


def login( driver ):
    try :
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
        sys.exit( f'======= No início sesión =======' )



extract_columnaid_product()