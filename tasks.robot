*** Settings ***
Documentation       Template robot main suite.
Library    RPA.Browser.Selenium
Library    RPA.Excel.Application
Library    RPA.Windows
Library    Collections
Library    DateTime
Library    Prices

*** Variables ***
@{data}
${equal}    =
*** Tasks ***
Obtener datos de ofertas Nike
    ${datos}=    Create List
    Extraer Datos Nike    ${datos}
    Cargar Productos a Excel    ${datos}


*** Keywords ***
Extraer Datos Nike
    [Arguments]    ${datos}
    
    Open Available Browser    https://www.nike.com/gb/w/mens-football-shoes-1gdj0znik1zy7ok    
    ${prev_width}    ${prev_height}    Get Element Size    css:body
    ${actual_height}=    Set Variable    ${0}
    ${selector}=    Set Variable    css:span.related_categories__title
    ${count}=    Set Variable    ${0}    
    WHILE    ${prev_height} != ${actual_height}
        ${prev_height}=    Set Variable    ${actual_height}
        Scroll Element Into View    ${selector}
        Sleep    7
        ${prev_width}    ${actual_height}    Get Element Size    css:body 
        ${count}=    Set Variable    ${count + 1}
        IF    ${count} == ${10}
            ${prev_height}=    Set Variable    ${actual_height}
        END      
    END
    Log To Console    ${count}
    @{productos}=    Get WebElements    css:div.product-card
    FOR    ${producto}    IN    @{productos}
        Log    ${producto}
        ${titulo}=    Get WebElement    ${producto.find_element(by='class name',value="product-card__title")}
        ${imagen}=    Get WebElement    ${producto.find_element(by='tag name',value='img')} 
        ${precio}=    Get WebElement    ${producto.find_element(by='class name',value='product-price')}  
        ${url_detalle}=    Get WebElement    ${producto.find_element(by='class name',value='product-card__link-overlay')}
        @{ofertas}=   Create List    ${titulo.text}    ${imagen.get_attribute('src')}    ${precio.text}    ${url_detalle.get_attribute('href')}
        Append To List    ${datos}    ${ofertas}
        Log To Console    ${precio.text}
        Log To Console    \n\n
    END
    Log To Console    ${datos}
    Close Browser


Cargar Productos a Excel
    [Arguments]    ${data}
    RPA.Excel.Application.Open Application    visible=${True}
    Add New Workbook
    Set Active Worksheet    sheetname=Sheet1
    ${date}=    Get Current Date    result_format=datetime
    Write To Cells    Sheet1    1    2    Datos Extraidos Tienda Nike
    ${format}=    Set Variable    [$-es-BO]dddd, d "de" mmmm "de" yyyy
    Write To Cells    worksheet=Sheet1    row=1    column=5    number_format=${format}      value=${date.date()}
    Write To Cells    worksheet=Sheet1    row=1    column=6    number_format=hh:mm    value=${date}  
    Write To Cells    worksheet=Sheet1    row=3    column=2    value=TITULO
    Write To Cells    worksheet=Sheet1    row=3    column=3    value=PRECIO
    Write To Cells    worksheet=Sheet1    row=3    column=4    value=DIVISA
    Write To Cells    worksheet=Sheet1    row=3    column=5    value=IMAGEN
    ${data_length}=    Get Length    ${data}
    Sleep    3
    FOR    ${i}    IN RANGE    4    ${data_length}
        ${price}    ${curr}    Get Price And Currency    ${data}[${i - 4}][2]  
        Write To Cells    worksheet=Sheet1    row=${i}    column=2    formula==HYPERLINK("${data}[${i - 4}][3]","${data}[${i - 4}][0]")
        Write To Cells    worksheet=Sheet1    row=${i}    column=3    value=${price}    number_format=0,00
        Write To Cells    worksheet=Sheet1    row=${i}    column=4    value=${curr}
        Write To Cells    worksheet=Sheet1    row=${i}    column=5    formula==IMAGE("${data}[${i - 4}][1]","i",3,40,100)
        Log To Console    ${price}
    END
    Sleep    5
    Save Excel As    filename=Datos_nike_${date.day}_${date.month}_${date.year}    autofit=True
    Sleep    6