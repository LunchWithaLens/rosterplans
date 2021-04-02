# rosterplans
Some example PowerShell calls to Graph API to create roster plans

See https://docs.microsoft.com/en-us/office365/planner/disable-roster-containers for the admin portion, then PlannerAdmin.ps1 makes use of the libraries.

RosterPlansInteractiveLogin.ps1 uses MSAL to connect and then shows how to create a roster, create a plan in the roster container and also add members.  I also show the new code to create a plan in a Groups container.
