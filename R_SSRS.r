######################################################################################
######################################################################################
###                                                                                ###
###                                     TITLE:                                     ###
###                                     R CODE                                     ###
###                        AUTHOR: 政 ZHUZHENG(IVERSON) ZHOU                       ###
###                                DATE: 2020-06-05                                ###
###                                   Initital VERSION                             ###
###  TOPIC:   Pull SSRS Reports and scheduling metadata Pragmatically in R         ###
###                 From MS SSRS Reporting Server                                  ###
###                                                                                ###
######################################################################################
######################################################################################

library(RODBC)
library(dplyr)

IZconnection <- odbcConnect('driver={SQL Server};
                              server=*****************,1**00;
                              database=ReportServer;
                              trusted_connection=true')

metadata <- sqlQuery(IZconnection,"
SELECT CAST(a.ScheduleID AS SYSNAME) AS JobName
,CONVERT(XML,CONVERT(VARBINARY(MAX),e.Content)) AS reportXML
,a.SubscriptionID
,e.NAME
,e.Path
,d.Description
,d.LastStatus
,d.EventType
,d.LastRunTime
,ItemID	
,Path,ParentID,Type
,Property
,CreationDate
,e.ModifiedDate
,Parameter
,UserName
--,b.date_created
--,b.date_modified
, DataSetName
,DataSourceName
,CommandType
,CommandText
FROM [IZDBAPRDSSRS01].ReportServer.dbo.ReportSchedule AS a
--INNER JOIN msdb.dbo.sysjobs AS b ON CAST(a.ScheduleID AS SYSNAME) = b.NAME
INNER JOIN [IZDBAPRDSSRS01].ReportServer.dbo.ReportSchedule AS c ON CAST(a.ScheduleID AS SYSNAME) = CAST(c.ScheduleID AS SYSNAME)
INNER JOIN [IZDBAPRDSSRS01].ReportServer.dbo.Subscriptions AS d ON c.SubscriptionID = d.SubscriptionID
INNER JOIN (
        		SELECT a.*,
            DataSetName = x.value('@Name', 'NVARCHAR(256)'),
            DataSourceName = x.value('(*:Query/*:DataSourceName)[1]', 'NVARCHAR(260)'),
            CommandType = x.value('(*:Query/*:CommandType)[1]', 'NVARCHAR(15)'),
            CommandText = x.value('(*:Query/*:CommandText)[1]', 'NVARCHAR(MAX)')
			
    FROM (
			Select *,CONVERT(XML,CONVERT(VARBINARY(MAX),Content)) ContentXml from [IZDBAPRDSSRS01].[ReportServer].[dbo].[Catalog] 
	  	--CONVERT(varchar(max), CONVERT(varbinary(max), content)) Fields
		WHERE Content is not null
		AND Type = 2
		)a
    CROSS APPLY ContentXml.nodes('/*:Report/*:DataSets/*:DataSet')  r (x)
		) AS e ON d.Report_OID = e.ItemID
INNER JOIN [IZDBAPRDSSRS01].[ReportServer].[dbo].[Users] f ON e.CreatedByID=f.UserID

  "
)

######################################################################################
### Meta Data Wraggling : Programaticaly extract names,dataset used and queries embedded from reports Wwhere its SQL queries contain specific variable(s) of interest
######################################################################################
# regular expression lookup columns contain specific columns
metadata2<-metadata[grep("Freegifts", metadata$reportXML), ]

# de-dup reports
metadata3<-distinct(metadata2,NAME, .keep_all= TRUE)
