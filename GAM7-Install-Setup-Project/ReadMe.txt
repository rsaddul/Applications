GAM7 Project Pre-requisites:

Enable Service "Google Cloud Platform" for the Admin OU 
(https://admin.google.com/ac/settings/serviceonoff?iid=174&aid=923893674115)

Set "All users in this group or org unit are 18 or older" for the Admin OU" (https://admin.google.com/ac/managedsettings/453172576738/age_based_access)

(The below must be done at Root Domain Level)
https://console.cloud.google.com/iam-admin/iam?orgonly=true&organizationId=227786515067&supportedpurview=organizationId,folder,project
Set Admin Account as "Organization Policy Administrator" (orgpolicy.policyAdmin)

https://console.cloud.google.com/iam-admin/orgpolicies/list?organizationId=227786515067&orgonly=true&supportedpurview=organizationId,folder,project
Search for "iam.managed.disableServiceAccountKeyUpload" and Disable "upload/iam.managed.disableServiceAccountKeyUpload" and "Upload/iam.disableServiceAccountKeyUpload" and "Turn these OFF"