*** Settings ***
Documentation     Proyecto que busca automatizar consultas hacia Civil Pjud.
...               Se necesitan las librerias instaladas para poder hacer funcionar el script.
Library           SeleniumLibrary
Library           ExcelLibrary
Library           clipboard
Library           String
Library           ImageHorizonLibrary
Library           DateTime

*** Variables ***
${Url}            https://civil.pjud.cl/CIVILPORWEB    # Direccion de la pagina a realizar las consultas
${PathExcel}      resultado/Nombres.xls    #Ubicacion de archivo Excel a consultar.
${NombreHojaExcel}    nombres    #Nombre de la hoja excel que se consulta.
${Contador}       1    #Contador que recorrera el total de valores de archivo excel.
${NombreCopiar}    ${EMPTY}    #Nombre que se extrae de excel
${ApellidoPaternoCopiar}    ${EMPTY}
${ApellidoMaternoCopiar}    ${EMPTY}
${RutCopiar}      ${EMPTY}
${ContadorCasos}    ${EMPTY}
${ContadorDeCasosInternos}    1
${SiTienenCaso}    ${EMPTY}

*** Test Cases ***
TestFinal
    Open Excel    ${PathExcel}
    ${Count1}    Get Row Count    ${NombreHojaExcel}    #Total de filas
    @{Count1}    Get column values    ${NombreHojaExcel}    1    #Valores de la columna 1
    FOR    ${Var1}    IN    @{Count1}    #Recorre    cada fila de archivo excel
        BuscadorDeCasos
        log    ${Contador}
    #Contador
        Sleep    5s
        ValidarTotalCasos
        ContadorCasosInternosReset
        Close Browser
        AumentadorDeNumeroPorCaso
        Log    ${Var1}
    END

*** Keywords ***
BuscadorDeCasos
    [Documentation]    Rescata variables desde Excel.
    Open Excel    ${PathExcel}
    Open Browser    ${Url}    chrome    #Apertura de explorador
    Sleep    10s    \    #Espera de 10 segundos
    Select Frame    name=body
    Click Element    //td[contains(@id,'tdCuatro')]
    log    ${Contador}
    ${NombreCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    B${Contador}
    clipboard.Copy    ${NombreCopiar}
    ${NombreCopiar}    Set Suite Variable    ${NombreCopiar}
    Log    ${NombreCopiar}
    Click Element    //input[contains(@name,'NOM_Consulta')]
    Press Keys    none    CTRL+V
    ${ApellidoPaternoCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    D${Contador}
    clipboard.Copy    ${ApellidoPaternoCopiar}
    ${ApellidoPaternoCopiar}    Set Suite Variable    ${ApellidoPaternoCopiar}
    Log    ${ApellidoPaternoCopiar}
    Click Element    //input[contains(@name,'APE_Paterno')]
    Press Keys    none    CTRL+V
    ${ApellidoMaternoCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    E${Contador}
    clipboard.Copy    ${ApellidoMaternoCopiar}
    ${ApellidoMaternoCopiar}    Set Suite Variable    ${ApellidoMaternoCopiar}
    Log    ${ApellidoMaternoCopiar}
    Click Element    //input[contains(@name,'APE_Materno')]
    Press Keys    none    CTRL+V
    ${RutCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    A${Contador}
    clipboard.Copy    ${RutCopiar}
    ${RutCopiar}    Set Suite Variable    ${RutCopiar}
    Log    ${RutCopiar}
    Sleep    10s
    Press Keys    \    ENTER
    Sleep    10s
    Wait Until Element Is Visible    //th[contains(@id,'Tit1')]
    Sleep    10s

AumentadorDeNumeroPorCaso
    [Documentation]    Contador del total de personas de los cuales se consideraran para las consultas.
    ${temp}    Evaluate    ${Contador} + 1
    Set Test Variable    ${Contador}    ${temp}

ValidarTotalCasos
    [Documentation]    Validar el numero total de casos por rut.
    log    ${Contador}
    Sleep    20s
    Wait Until Element Is Visible    //th[contains(@id,'Tit1')]
    Sleep    5s
    ${Span}=    Get WebElements    //th[contains(@id,'Tit1')]
    Sleep    5s
    Wait Until Element Is Visible    //th[contains(@id,'Tit1')]
    Log Many    ${Span}=
    Sleep    5s
    Wait Until Element Is Visible    //th[contains(@id,'Tit1')]
    Sleep    5s
    ${MyText}=    Get Text    ${Span[0]}
    Wait Until Element Is Visible    //th[contains(@id,'Tit1')]
    Sleep    5s
    Log    ${MyText}
    Wait Until Element Is Visible    //th[contains(@id,'Tit1')]
    Sleep    5s
    ${MyTex2t}=    Remove String Using Regexp    ${MyText}    /^[a-zA-Z\s]*$/;
    Log    ${MyTex2t}
    ${string}=    String.Fetch From Left    ${MyText}    )
    ${string}=    String.Replace String    ${string}    ${Space}    ${EMPTY}
    Log    ${string}
    ${MyText}=    Remove String    ${string}    Causas    [    ]    :    Cantidad
    Log    ${MyText}
    Convert To Number    ${MyText}
    Log    Ahora soy un numero
    Log    ${MyText}
    Set Test Variable    ${ContadorCasos}    ${MyText}
    Log    ${ContadorCasos}
    Run Keyword If    ${MyText}>0    Repeat Keyword    ${ContadorCasos} times    AperturaDeCasos
    ...    ELSE    Close Browser
    Set Test Variable    ${ContadorCasos}    0
    Log    ${ContadorCasos}

AperturaDeCasos
    Log    ${ContadorCasos}
    Log    ${RutCopiar}
    Log    ${ContadorCasos}
    FOR    ${Var2}    IN    ${ContadorCasos}
        Sleep    15s
        Click Element    (//td[contains(@class,'textoC')])[${ContadorDeCasosInternos}]
        Sleep    21s
        Click Element    (//td[contains(.,'Litigantes')])[1]
        Sleep    20s
        ValidarRutExcelHaciaPjud
        Go Back
        Sleep    15s
        Select Frame    name=body
        Sleep    10s
        ContadorCasosInternos
    END

ContadorCasosInternos
    ${temp2}    Evaluate    ${ContadorDeCasosInternos}+2
    Set Test Variable    ${ContadorDeCasosInternos}    ${temp2}

ContadorCasosInternosReset
    ${temp2}    Evaluate    1
    Set Test Variable    ${ContadorDeCasosInternos}    ${temp2}

GuardadorEnExcel
    Open Excel    resultado/Prototipo.xls
    Sleep    5s
    log    ${Contador}
    log    ${RutCopiar}
    log    ${NombreCopiar}
    log    ${ApellidoPaternoCopiar}
    log    ${ApellidoMaternoCopiar}
    Put String To Cell    resultado    0    ${Contador}    ${RutCopiar}
    Put String To Cell    resultado    1    ${Contador}    ${NombreCopiar}
    Put String To Cell    resultado    2    ${Contador}    ${ApellidoPaternoCopiar}
    Put String To Cell    resultado    3    ${Contador}    ${ApellidoMaternoCopiar}
    ${SiTienenCaso}    Get WebElements    (//td[contains(@height,'11')])[1]
    ${NumeroCaso}=    Get Text    ${SiTienenCaso[0]}
    Put String To Cell    resultado    4    ${Contador}    ${NumeroCaso}
    ${timestamp} =    Get Current Date    result_format=%Y-%m-%d-%H-%M
    ${filename} =    Set Variable    resultado-${timestamp}.xls
    Save Excel    resultado/${filename}

ValidarRutExcelHaciaPjud
    log    ${ContadorCasos}
    log    ${NombreCopiar}
    log    ${RutCopiar}
    log    ${ApellidoMaternoCopiar}
    log    ${ApellidoPaternoCopiar}
    Sleep    7s
    Sleep    2s
    ${Span}=    SeleniumLibrary.Get WebElements    //td[@class='texto'][contains(.,'${RutCopiar}')]
    Log    ${Span}
    log    ${RutCopiar}
    ${test}=    Get Element Count    //td[@class='texto'][contains(.,'${RutCopiar}')]
    log    ${test}
    Sleep    5s
    Run Keyword If    ${test}>0    GuardadorEnExcel
    ...    ELSE    log    "No existe Registro Valido"
