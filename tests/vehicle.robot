*** Settings ***
Library    SeleniumLibrary

*** Test Cases ***
Verify Vehicle Search
    Open Browser    https://example.com    chrome
    Title Should Be    Example Domain
    Close Browser