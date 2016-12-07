# Sync-AD-With-Local-Computer-Descriptions
Searches AD for computers, queries each computer for their system description, updates AD with local descriptions

    .Synopsis
    This script will facilitate synchronizing computer descriptions in Active Directory with the system descriptions stored on each individual computer which by design, are not used by AD.

    .DESCRIPTION
    Enter credentials used to search AD, and same or different credentials to connect to each resulting computer.
    Searches AD for computer objects in AD matching the search term input, connects to each of the resulting computers to pull their system descriptions and compares them in a displayed table.
    
    Options
    [1] Update AD with ALL non-empty local system descriptions matching the entered AD computer name search term
        This will update in one batch all AD search results with their corresponding local system descriptions using the AD credentials supplied earlier. 

    [2] Manually approve each AD update one at a time 
        This will prompt you for approval to update AD with the local system description of each AD computer search result.
