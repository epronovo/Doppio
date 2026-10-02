SELECT [Active__c],
       [IsDeleted],
       [bgi_ERP_Reference_Count__c],
       --[Industry],
       [Id],
       [Brand__c],
       [Customer_Service_Location2__c],
       [User_Company__c],
       [bgi_Brand_s_Preferred__c],
       [AccountSource],
       [bgi_golden_id__c],
       [bgi_Dupe_Golden_ID__c],
       [bgi_Potential_Duplicate_Match__c],
       [Aurora_Unique_Id__c],
       [AccountNumber],
       [bgi_EXT_ID__c],
       [bgi_DUNS_Number__c],
       [DW_Id__c],
       [Name],
       [bgi_International_Account_Name__c],
       [BillingStreet],
       [BillingCity],
       [BillingStateCode],
       [BillingState],
       [BillingCountryCode],
       [BillingCountry],
       [BillingPostalCode],
       [Phone],
       [ShippingStreet],
       [ShippingCity],
       [ShippingStateCode],
       [ShippingState],
       [ShippingCountryCode],
       [ShippingCountry],
       [ShippingPostalCode],
       [CurrencyIsoCode],
       [bgi_RPS_Status__c],
       [bgi_RPS_Message__c],
       [ParentId],
       [Customer_Group__c],
       [Continent__c],
       [bgi_Key_Account_Manager__c],
       [Data_Quality_Score__c],
       [Data_Quality_Description__c],
       [SystemModstamp],
       --[LastViewedDate],
       [LastReferencedDate],
       [LastActivityDate],
       [LastModifiedDate]
FROM [DW_BGI_OpStage].[dbo].[Account] a
WHERE Brand__c LIKE '%IGS%'
      OR Brand__c LIKE '%KALLER%'
      OR Brand__c LIKE '%ASRaymond%'
      OR Brand__c LIKE '%Hyson%'
      OR Customer_Service_Location2__c LIKE 'FMC %'
      OR Customer_Service_Location2__c LIKE 'MCS %'
      OR a.User_Company__c LIKE 'FMC %'
ORDER BY [LastReferencedDate] DESC, a.LastActivityDate DESC

