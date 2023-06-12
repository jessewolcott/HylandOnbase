SELECT TOP (1000) [usernum]
      ,[username]
      ,[disablelogin]
      ,[networkid]
      ,[lastlogon]
     
  FROM [OnBase].[hsi].[useraccount]
  where ([disablelogin] = '0') AND ([lastlogon] < '2023-01-01 00:00:00.000')
  ORDER by lastlogon