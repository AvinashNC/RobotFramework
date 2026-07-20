*** Settings ***
Library    SeleniumLibrary

*** Test Cases ***
Open Google
    Open Browser    https://www.google.com    chrome
    Maximize Browser Window
    Title Should Be    Google
    Close Browser