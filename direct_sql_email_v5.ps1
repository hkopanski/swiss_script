$ServerInstance = 'xxxx'
$Database = 'xxxx'
$variable1 = 'xxxx'
$variable2 = 'xxxx'
$Query =  "SET NOCOUNT ON

DECLARE @StartTime as varchar(60)
DECLARE @EndTime as varchar(60)

SET @EndTime = GETDATE()
SET @StartTime = DATEADD(day, -1, @EndTime)

DECLARE @AlarmRaise table 

( 
    EventTime nvarchar(60), 
    ID nvarchar(50), 
    AlarmState nvarchar(20),
    TagID NVARCHAR(100) 
    ) 

DECLARE @Chamber_List table
(
	Chamber_TagName nvarchar (30)
)

insert into @Chamber_List values 

('Chamber_1'), ('Chamber_2')


INSERT @AlarmRaise 
    SELECT EventTime, Alarm_ID, Alarm_State, Source_ProcessVariable 
    FROM Events 
        WHERE EventTime > @StartTime and EventTime < @EndTime 
        and Alarm_State='UNACK_ALM' 
        and (Source_Object in  (select * from @Chamber_List ))
        
DECLARE @AlarmAck table
( 
    EventTime nvarchar(60), 
    ID nvarchar(50), 
    UnAckDuration nvarchar(20),
    TagID NVARCHAR(100),
	Comment nvarchar(4000)
) 

INSERT @AlarmAck 
    SELECT EventTime, Alarm_ID, Alarm_UnAckDurationMs, Source_ProcessVariable, Comment
    FROM Events 
        WHERE EventTime > @StartTime and EventTime < @EndTime 
        and Alarm_Acknowledged=1 
        and  (Source_Object in  (select * from @Chamber_List ))

        
DECLARE @AlarmClear table
( 
    EventTime nvarchar(60), 
    ID nvarchar(50),
    AlarmDuration nvarchar(20),
    TagID NVARCHAR(100) 
    ) 
    
INSERT @AlarmClear 
    SELECT EventTime, Alarm_ID, Alarm_DurationMs, Source_ProcessVariable 
    FROM Events 
        WHERE EventTime > @StartTime and EventTime < @EndTime
     /* and Alarm_DurationMs >= 3600000 */
        and Type='Alarm.Clear' 
        and  (Source_Object in  (select * from @Chamber_List ))
       
--======================--
SELECT s.TagID as TagName
                --'Alarm Life - '+ s.ID, 
                -- CASE 
                -- WHEN a.EventTime > c.EventTime THEN 'Cleared Before Ack' 
                -- WHEN a.EventTime < c.EventTime THEN 'Acked Before Clear' 
                -- ELSE '-' END as Comment 
                ,RIGHT('00' + CONVERT(NVARCHAR(2), DATEPART(DAY, s.EventTime)), 2) + SUBSTRING(UPPER(DATENAME(month, s.EventTime)), 0, 4) + SUBSTRING(UPPER(DATENAME(year, s.EventTime)), 0, 5) + ' ' + RIGHT('00' + CONVERT(varchar, SUBSTRING(UPPER(DATENAME(hour, s.EventTime)), 0, 4)), 2) + RIGHT('00' + CONVERT(NVARCHAR(2), SUBSTRING(UPPER(DATENAME(MINUTE, s.EventTime)), 0, 4)), 2) as AlarmRaised 
                ,ISNULL(RIGHT('00' + CONVERT(NVARCHAR(2), DATEPART(DAY, a.EventTime)), 2) + SUBSTRING(UPPER(DATENAME(month, a.EventTime)), 0, 4) + SUBSTRING(UPPER(DATENAME(year, a.EventTime)), 0, 5) + ' ' + RIGHT('00' + CONVERT(varchar, SUBSTRING(UPPER(DATENAME(hour, a.EventTime)), 0, 4)), 2) + RIGHT('00' + CONVERT(NVARCHAR(2), SUBSTRING(UPPER(DATENAME(MINUTE, a.EventTime)), 0, 4)), 2), 'N/A') as AlarmAcked 
                ,RIGHT('00' + CONVERT(NVARCHAR(2), DATEPART(DAY, c.EventTime)), 2) + SUBSTRING(UPPER(DATENAME(month, c.EventTime)), 0, 4) + SUBSTRING(UPPER(DATENAME(year, c.EventTime)), 0, 5) + ' ' + RIGHT('00' + CONVERT(varchar, SUBSTRING(UPPER(DATENAME(hour, c.EventTime)), 0, 4)), 2) + RIGHT('00' + CONVERT(NVARCHAR(2), SUBSTRING(UPPER(DATENAME(MINUTE, c.EventTime)), 0, 4)), 2) as AlarmRTN 
                --,a.UnAckDuration as UnAckDuration 
                --,a.UnAckDuration as UnAckDuration 
                ,CONVERT(VARCHAR,DATEADD(ms,CAST(c.AlarmDuration AS int),0), 108) as AlarmDuration_hr
                --,CAST(c.AlarmDuration AS float) / 3600000 as AlarmDuration_ms 
								,ISNULL(a.Comment, 'N/A') as AcknowledgeComment
								,@StartTime as QueryStart
								,@EndTime as QueryEnd
				
FROM @AlarmRaise s 
LEFT JOIN @AlarmClear c ON c.ID = s.ID
LEFT JOIN @AlarmAck a ON a.ID = c.ID AND a.EventTime <> c.EventTime
ORDER BY TagName, AlarmRaised asc

SET NOCOUNT OFF"


$data_table_sql = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Username $variable1 -Password $variable2 -Query $Query

$head = @"

 <style>
	body {
	background: #fafafa;
	color: #444;
	font: 100%/30px 'Helvetica Neue', helvetica, arial, sans-serif;
	text-shadow: 0 1px 0 #fff;
}

strong {
	font-weight: bold; 
}

capFormat {
	background: linear-gradient(#777, #444);
    border-left: 1px solid #555;
    border-right: 1px solid #777;
    border-top: 1px solid #555;
    border-bottom: 1px solid #333;
    box-shadow: inset 0 1px 0 #999;
    color: #fff;
    font: 30px 'Helvetica Neue', helvetica, arial, sans-serif;
    font-weight: bold;
    padding: 10px 15px;
    position: relative;
    text-shadow: 0 1px 0 #000; 
}

table {
	background: #f5f5f5;
	border-collapse: separate;
	box-shadow: inset 0 1px 0 #fff;
	font-size: 12px;
	line-height: 24px;
	margin: 30px auto;
	text-align: left;
	width: 800px;
}	

th {
	background: linear-gradient(#777, #444);
	border-left: 1px solid #555;
	border-right: 1px solid #777;
	border-top: 1px solid #555;
	border-bottom: 1px solid #333;
	box-shadow: inset 0 1px 0 #999;
	color: #fff;
  font-weight: bold;
	padding: 10px 15px;
	position: relative;
	text-shadow: 0 1px 0 #000;	
}

th:after {
	background: linear-gradient(rgba(255,255,255,0), rgba(255,255,255,.08));
	content: '';
	display: block;
	height: 25%;
	left: 0;
	margin: 1px 0 0 0;
	position: absolute;
	top: 25%;
	width: 100%;
}

th:first-child {
	border-left: 1px solid #777;	
	box-shadow: inset 1px 1px 0 #999;
}

th:last-child {
	box-shadow: inset -1px 1px 0 #999;
}

td {
	border-right: 1px solid #fff;
	border-left: 1px solid #e8e8e8;
	border-top: 1px solid #fff;
	border-bottom: 1px solid #e8e8e8;
	padding: 10px 15px;
	position: relative;
	transition: all 300ms;
}

td:first-child {
	box-shadow: inset 1px 0 0 #fff;
}	

td:last-child {
	border-right: 1px solid #e8e8e8;
	box-shadow: inset -1px 0 0 #fff;
}	

td:first-child {
  font-weight: bold
}

tr:nth-child(odd) td {
	background: #f1f1f1;	
}

tr:last-of-type td {
	box-shadow: inset 0 -1px 0 #fff; 
}

tr:last-of-type td:first-child {
	box-shadow: inset 1px -1px 0 #fff;
}	

tr:last-of-type td:last-child {
	box-shadow: inset -1px -1px 0 #fff;
}	

tbody:hover td {
	color: transparent;
	text-shadow: 0 0 3px #aaa;
}

tbody:hover tr:hover td {
	color: #444;
	text-shadow: 0 1px 0 #fff;
} 
 </style>
"@


if($data_table_sql -ne $null){

$data_table_sql | Select TagName, AlarmRaised, AlarmAcked, AlarmRTN, AlarmDuration_hr, AcknowledgeComment | ConvertTo-Html -Head $head | Out-File C:\Scripts\chambers_query\B35_Daily_QC_Report.html

$ftable = $data_table_sql | format-table

$ftable > C:\Scripts\chambers_query\chamber_table.txt

$Attachments = $data_table_sql | format-table
$Attachment = 'C:\Scripts\chambers_query\B35_Daily_QC_Report.html'
    
    $sendMailReport = @{
        From = 'xxxx'
        To = 'person@email.com'
        Subject = 'xxxxx'
        Body = "xxxxxx"
        Attachments = $Attachment
        Priority = 'High'
        DeliveryNotificationOption = 'OnSuccess', 'OnFailure'
        SmtpServer = 'xxxxx'
    }
    
Send-MailMessage @sendMailReport

Remove-Item C:\Scripts\chambers_query\*.*

}else{ 

$data_table_sql | Select TagName, AlarmRaised, AlarmAcked, AlarmRTN, AlarmDuration_hr, AcknowledgeComment | ConvertTo-Html -Head $head | Out-File C:\Scripts\chambers_query\B35_Daily_QC_Report.html

$ftable = $data_table_sql | format-table

$ftable > C:\Scripts\chambers_query\chamber_table.txt

$Attachments = $data_table_sql | format-table
$Attachment = 'C:\Scripts\chambers_query\B35_Daily_QC_Report.html'
    
    $sendMailReport = @{
        From = 'xxxxx'
        To = 'xxxxxx'
        Subject = 'Daily Chamber Report'
        Body = "No Alarms found in the last 24 hours"
        Priority = 'High'
        DeliveryNotificationOption = 'OnSuccess', 'OnFailure'
        SmtpServer = 'xxxxx'
    }
    
Send-MailMessage @sendMailReport

}
