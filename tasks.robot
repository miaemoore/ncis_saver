*** Settings ***
Documentation       All NCIS files will be converted to PDF and saved to computer from New Account emails.
...                 Under Variables:
...                 Input name of Outlook account/email.
...                 Input name of folder where New Account emails, which have NCIS files attached, are store in Outlook.
...                 Input the folder on computer where the NCIS file should end up (i.e. MDI Working Folder).
...                 Input a folder where the emails that have been processed can be saved (i.e. archive folder).

Library             RPA.Outlook.Application    autoexit=False
Library             RPA.PDF
Library             OperatingSystem
Library             RPA.FileSystem
Library             String
Library             RPA.Excel.Application    autoexit=False
Library             RPA.Excel.Files


*** Variables ***
#Outlook email account
${account_name}=        MiMoore@pathgroup.com
#Email folder where NCIS emails are stored
${folder_name}=         New Accounts
#Path of folder where NCIS files should go
${output_folder}=       I:${/}MDI WORKING FOLDER${/}New Accounts${/}2023${/}
#NCIS forms that have been saved to the MDI folder go to this email folder
${archive_folder}=      New Accounts Archived


*** Tasks ***
Minimal task
    Open Outlook
    Find NCIS emails
    Convert Excel to PDF
    [Teardown]    Close and delete extra files


*** Keywords ***
Open Outlook
    RPA.Outlook.Application.Open Application

Find NCIS emails
    ${emails}=    Get Emails
...    account_name=${account_name}
...    folder_name=${folder_name}
#...    email_filter=[unread]=${True}
...    save_attachments=True
...    attachment_folder=${OUTPUT_DIR}${/}Attachments
...    sort=True
...    sort_key=Received
...    sort_descending=True
    FOR    ${email}    IN    @{emails}
        Run Keyword And Ignore Error    Move Emails
...    account_name=${account_name}
...    source_folder=${folder_name}
...    target_folder=${archive_folder}
    END

Convert Excel to PDF
    ${ncisfiles}=    OperatingSystem.List Files In Directory
    ...    ${OUTPUT_DIR}${/}Attachments
    ...    pattern=*.xlsx
    FOR    ${ncisfile}    IN    @{ncisfiles}
        ${title}=    Get Substring    ${ncisfile}    0    -5
        ${exists}=    Does File Exist    ${output_folder}${title}.pdf
        IF    ${exists}            CONTINUE
        Export As Pdf
        ...    ${OUTPUT_DIR}${/}Attachments${/}${title}.pdf
        ...    ${OUTPUT_DIR}${/}Attachments${/}${ncisfile}
        Extract Pages From Pdf
        ...    ${OUTPUT_DIR}${/}Attachments${/}${title}.pdf
        ...    ${output_folder}${title}.pdf
        ...    1
        RPA.Excel.Application.Close Document
    END

Close and delete extra files
    Close All Pdfs
    RPA.Filesystem.Empty Directory    ${OUTPUT_DIR}${/}Attachments${/}
