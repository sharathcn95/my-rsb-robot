*** Settings ***
Documentation     Insert the sales data for the week and export it as a PDF.
Library           RPA.Browser.Selenium
Library           RPA.HTTP
Library           RPA.Excel.Files
Library           RPA.PDF

*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open the intranet website
    Download the Excel file
    Log in
    Fill the form using the data from the Excel file
    Collect the results
    Export Table as PDF
    [Teardown]    LogOut and Close the Browser

*** Keywords ***
Open the intranet website
    Open Chrome Browser    https://robotsparebinindustries.com/

Log in
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Submit Form
    Assert logged in

Assert logged in
    Wait Until Page Contains Element    id:sales-form
    Location Should Be    https://robotsparebinindustries.com/#/

Fill and Submit the Form for One Person
    [Arguments]    ${sales_rep}
    Input Text    firstname    ${sales_rep}[First Name]
    Input Text    lastname    ${sales_rep}[Last Name]
    Input Text    salesresult    ${sales_rep}[Sales]
    Select From List By Value    salestarget    ${sales_rep}[Sales Target]
    Submit Form

Download the Excel file
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=True

Fill the form using the data from the Excel file
    Open Workbook    SalesData.xlsx
    ${sales_reps}=    Read Worksheet As Table    header=True
    Close Workbook
    FOR    ${sales_rep}    IN    @{sales_reps}
        Fill and Submit the Form for One Person    ${sales_rep}
    END

Collect the results
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png

Export table as PDF
    Wait Until Element Is Visible    id:sales-results
    ${Table_HTML}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${Table_HTML}    ${OUTPUT_DIR}${/}sales_summary.pdf

LogOut and Close the Browser
    Click Button    logout
    Close Browser
