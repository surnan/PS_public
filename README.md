# PowerShell Scripts
## Exchange 

* show_group_members.ps1
  - Prompt for: **_group-name_**
  - Prompt for: **_output_device_**
  - Output:
    * group members
    * output to: command prompt or .csv

* Get_EntraUsersByCreatedDate.ps1
  - Prompt for: **start date**
  - Prompt for: **end date**
  - Output:
    * CONSOLE:
        * DisplayName
        * UserPrincipalName
        * CreatedDateTime
    * .CSV file:
        * name = "EntraUsersByCreatedDate" + <timestamp>
        * Columns:
            * DisplayName
            * UserPrincipalName
            * CreatedDateTime
            * Job Title
            
* Get_EntraUser_Detailse.ps1
  - Prompt for: **email address**
  - Output to console and .csv file
    * If email address matches user as UPN or alias:
        * Display Name
        * UPN
        * Job Title
        * Account Creation Timestamp
        * Email aliases
        * In Table Format:
            * Group Name
            * Group Type
            * Group Email
  
* Get_OffboardingExceptions.ps1
    - No prompt for input
    - Output:
    * CONSOLE & .CSV file:
        * DisplayName
        * UserPrincipalName
        * AccountEnabled
        * Email Address Count
          * if count = 1 then display name instead of number
        * Group Count
          * if count = 1 then display name instead of number