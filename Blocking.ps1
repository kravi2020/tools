#set-executionpolicy unrestricted

#Author : Chetan Vishwakarma
#Date : 11Oct2017

$Computers = Get-Content "D:\HealthCheck_Scripts\NewScript\Servers\BackupInstances.txt"
$OutputFile = "D:\HealthCheck_Scripts\NewScript\report\BlockingReport.htm" 

$HTML = '<style type="text/css"> 
    #Header{font-family:"Trebuchet MS", Arial, Helvetica, sans-serif;width:100%;border-collapse:collapse;} 
    #Header td, #Header th {font-size:14px;border:1px solid #98bf21;padding:3px 7px 2px 7px;} 
    #Header th {font-size:14px;text-align:left;padding-top:5px;padding-bottom:4px;background-color:#A7C942;color:#fff;} 
    #Header tr.alt td {color:#000;background-color:#EAF2D3;} 
    </Style>' 
    

  $HTML += "<HTML><BODY><Table border=1 cellpadding=0 cellspacing=0 width=100% id=Header> 
        <TR> 
            <TH><B>Server Name</B></TH> 
            <TH><B>BlockingSessionId</B></TD> 
            <TH><B>DB Name</B></TH> 
            <TH><B>Start_Time</B></TH>
            <TH><B>HoursTaken</B></TH>
            <TH><B>MinuteTaken</B></TH>
            <TH><B>Command</B></TH>

        </TR>"


foreach($Servers in $Computers)
    {


        $Query =   'select @@servername as Servername, blocking_session_id BlockingSessionId ,db_name(dbid) as DBName, Start_Time,DATEDIFF(Minute,start_time,getdate())/60 as HoursTaken,  DATEDIFF(Minute,start_time,getdate()) as MinuteTaken,Command   from sys.dm_exec_requests cross apply sys.dm_exec_sql_text (sql_handle) WHERE blocking_session_id <> 0'
       
   

        $Result = Invoke-Sqlcmd ($Query) -ServerInstance $Servers
        $Result
            foreach ($Item in $Result)
            {
                    $ServerNames = $Item.Servername
                    $BlockingSessionId = $Item.BlockingSessionId
                    $DbName = $Item.DBName
                    $Start_Time = $Item.Start_Time
                    $HoursTaken = $Item.HoursTaken
                    $MinuteTaken = $Item.MinuteTaken
                    $Command = $Item.Command

                     $HTML += "<TR> 
                    <TD>$($ServerNames)</TD> 
                    <TD>$($BlockingSessionId)</TD> 
                    <TD>$($DbName)</TD>   
                    <TD>$($Start_Time)</TD>     
                    <TD>$($HoursTaken)</TD>   
                    <TD>$($MinuteTaken)</TD>   
                    <TD>$($Command)</TD>     
                </TR>" 

                    
            }
           
           
    }

    $HTML += "</Table></BODY></HTML><BR><BR>" 

   $HTML | Out-File $OutputFile

    $Fragments+=$HTML 

    #write the result to a file 
    ConvertTo-Html -head $head -body $fragments `
     | Out-File $OutputFile



   
     

     
     