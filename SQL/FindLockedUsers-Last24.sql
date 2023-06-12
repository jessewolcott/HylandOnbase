SELECT a.[username] as "Username"
	  , "Login Disabled" = 
			Case 
				When disablelogin like '1'
					Then 'Account Disabled'
					else 'Account Enabled'
					end
      ,a.[badlogincount] as "Bad Login Count"
      ,a.[lastpwchange] as "Last PW Change"
      ,a.[realname] as "Real Name"
      ,a.[lockouttime] as "Lockout Time"
      ,a.[lockoutreason] as "Lockout Reason"


  FROM [OnBase].[hsi].[useraccount] a

  where a.disablelogin like '1' 
		and	 a.[lockouttime] >= DATEADD(HH, -24, GETDATE()) 