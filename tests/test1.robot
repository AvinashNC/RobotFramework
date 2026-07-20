*** Settings ***
Library    SeleniumLibrary

*** Variables ***
${URL}    http://localhost:8080

*** Test Cases ***
Verify Door Lock
    Open Browser    ${URL}    chrome
    Click Button    id=lockDoors

    Wait Until Page Contains    Doors Locked

    Element Should Contain    id=status    Doors Locked

    Close Browser