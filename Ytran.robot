*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF
Library    RPA.Browser.Selenium    auto_close=${False}
Library    RPA.Robocorp.Vault
Library    RPA.Robocorp.WorkItems
Library    Collections
Library    BuiltIn
Library    RPA.Excel.Files
Library    DateTime
Library    RPA.Tables
Library    RPA.MSGraph
Library    RPA.Email.ImapSmtp

*** Variables ***
${URL}    https://fusebox-portal-vn-dev.azurewebsites.net/login
${URL_EXLOG}    https://fusebox-portal-vn-dev.azurewebsites.net/Exlog
${Excel_File_Path}    C:/Users/VBPO/Desktop/RPA/output   


*** Test Cases ***
Open 
    Open Available Browser    ${URL}

Login
    ${Data}    Get Secret    swaglabs  
    Input Text    UserName    ${Data}[username]    
    Input Password    Password    ${Data}[password]    
    Click Button    //*[@id="frm_Login"]/input[2]
Login_exlog
    Go To    ${URL_EXLOG}
Get Data From Website   
    ${Get_Data_Table}=    Get WebElements    //table[@id="ErrorLog"]  
    ${list_column}=    Create List     Host    Code    Type    Error    User    Date    Time
    ${table_error}=    Create Table    columns=${list_column}     
   FOR    ${n}    IN RANGE   2    12
            ${host}=    Get Text    //*[@id="ErrorLog"]/tbody/tr[${n}]/td[1]
            ${code}=    Get Text    //*[@id="ErrorLog"]/tbody/tr[${n}]/td[2]
            ${type}=    Get Text    //*[@id="ErrorLog"]/tbody/tr[${n}]/td[3]
            ${error}=   Get Text    //*[@id="ErrorLog"]/tbody/tr[${n}]/td[4]
            ${user}=    Get Text    //*[@id="ErrorLog"]/tbody/tr[${n}]/td[5]
            ${date}=    Get Text    //*[@id="ErrorLog"]/tbody/tr[${n}]/td[6]
            ${time}=    Get Text    //*[@id="ErrorLog"]/tbody/tr[${n}]/td[7]
            Log To Console     \nHost: ${host} ,Code: ${code} ,Type: ${type} ,Error: ${error} ,User: ${user} ,Date: ${date} ,Time: ${time}\n          
            ${data_row_table}=    Create List    ${host}    ${code}    ${type}    ${error}    ${user}    ${date}    ${time}
            Add Table Row    ${table_error}    ${data_row_table}
    END
    Log To Console    ${table_error}
    ${date_time_now}    Get Current Date    result_format=%m-%d-%Y
    Create Excel Files    
    ...    ${Excel_File_Path}    
    ...    ${date_time_now}    
    ...    demo02   
    ...    ${table_error}
    
    Send Mail    y.tran@mainbridgehp.com    nhuy511998@gmail.com    Demo03    CC    C:/Users/VBPO/Desktop/RPA/output/${date_time_now}.xlsx 
      
*** Keywords ***
Create Excel Files
    [Arguments]      ${Excel_File_Path}    ${name_file}     ${sheet_name}    ${table_error}
    
    ${date_time_now}    Get Current Date    result_format=%m-%d-%Y
    Log To Console    ${date_time_now}
    Create Workbook    ${Excel_File_Path}${/}${name_file}.xlsx  
    Save Workbook
    Open Workbook    ${Excel_File_Path}${/}${date_time_now}.xlsx   
    Create Worksheet    name=${sheet_name}    content=${table_error}    header=${True}
    Save Workbook
    Close Workbook
    Log To Console    Create thành công
Send Mail

    [Arguments]    ${data_mail}    ${recipients}    ${subject}    ${body}    ${attachments}
    ${data_mail}    Get Secret    emailCredentials
    Authorize    ${data_mail}[username]   ${data_mail}[password]  smtp_server=smtp.gmail.com  smtp_port=587
    Send Message    sender=${data_mail}[username]    recipients=${recipients}    
    ...    subject=${subject}    body=${body}    attachments=${attachments}       