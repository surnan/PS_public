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
        