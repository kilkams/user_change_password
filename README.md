# user_change_password for MSSQL Job

0. You need to add the linked server ADSI to the root of the forest 
1. Replace LDAP://subdomain.domain.corp/OU=Users,OU=OUUsers,DC=subdomain,DC=domain,DC=CORP to your domain and OU name.
2. Add your exceptions to this list @user_id IN ('jirasupport'), separated by commas
3. Replace src="https://domain.com/logo.png" to your corporate logo URL
4. Replace @profile_name = 'user_not'; to your profile name in MSSQL
5. Replace domain in https://owa.domain.com/owa/auth/logon.aspx?replaceCurrent=1&url=https%3A%2F%2Fowa.domain.com%2Fowa%2F%23path%3D%2Foptions%2Fmyaccount to your domain in all places 
