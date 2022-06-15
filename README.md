# Azure-AD-user-Last-Sign-in

This script will check the last sign activity for the Azure AD users, Last Sign in activity will include interactive and non-interactive sign in , to get the accurate result when this user did a last sign in. as normal powershell command will target only the interactive sign-in not the non-interactive logs.

In common scenarios, the user has a valid token so he is continuing login using non-interactive login and from interactive login, he could be show as no sign in for the last month.

This Script is based on Azure APP with APIs with Application Permission , below are the API permission granted:

AuditLog.Read.All
Directory.Read.All
User.Read.All

Details user and application sign-in activity for a tenant (directory). You must have an Azure AD Premium P1 or P2 license to download sign-in logs using the Microsoft Graph API

lastSignInDateTime :Azure AD maintains interactive sign-ins going back to April 2020
lastNonInteractiveSignInDateTime : Azure AD maintains non-interactive sign-ins going back to May 2020
