##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uK1
##Nc3NCtDXThU=
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiW5
##OsHQCZGeTiiZ4tI=
##OcrLFtDXTiW5
##LM/BD5WYTiiZ4tI=
##McvWDJ+OTiiZ4tI=
##OMvOC56PFnzN8u+Vs1Q=
##M9jHFoeYB2Hc8u+Vs1Q=
##PdrWFpmIG2HcofKIo2QX
##OMfRFJyLFzWE8uK1
##KsfMAp/KUzWJ0g==
##OsfOAYaPHGbQvbyVvnQX
##LNzNAIWJGmPcoKHc7Do3uAuO
##LNzNAIWJGnvYv7eVvnQX
##M9zLA5mED3nfu77Q7TV64AuzAgg=
##NcDWAYKED3nfu77Q7TV64AuzAgg=
##OMvRB4KDHmHQvbyVvnQX
##P8HPFJGEFzWE8tI=
##KNzDAJWHD2fS8u+Vgw==
##P8HSHYKDCX3N8u+Vgw==
##LNzLEpGeC3fMu77Ro2k3hQ==
##L97HB5mLAnfMu77Ro2k3hQ==
##P8HPCZWEGmaZ7/K1
##L8/UAdDXTlaDjofG5iZk2Wr8SH0lUuGUuqOqwY+o7NbfsyzfXbIVR1BYgCzuKUq0VbwXTfB1
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
$shift=(Get-Date -Format g).ToString()
$sub="Shift Handover`t"+$shift

$password = ConvertTo-SecureString 'Azure@89' -AsPlainText -Force
$cr = New-Object System.Management.Automation.PSCredential ('mallikarjunar@workspot.com', $password)


function ShiftInfo {
    param (
        $Id,
        [validateset('Task', 'Incident', 'Problem','Document','Jira')]
        [string]$Activity,
        [string]$Info,
        [String][ValidateSet('Yes','No')]$XtraRows
    )
    [pscustomobject]@{
        Job = $Activity
        Number = $Id
        Comments = $Info
        Rows = $Xtrarows
        }
        
}


[pscustomobject]$total=@()
do{

    $result = Invoke-Expression (Show-Command ShiftInfo -PassThru)
    $total += $result
}while($result.Rows -like 'Yes')


# Create a DataTable
$table = New-Object system.Data.DataTable "TestTable"
$col1 = New-Object system.Data.DataColumn Job,([string])
$col2 = New-Object system.Data.DataColumn Id,([string])
$col3 = New-Object system.Data.DataColumn Comments,([string])
$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)

# Add content to the DataTable

for($i=0;$i -lt $total.Count;$i++)
{
    $row = $table.NewRow()
    $row.Job = $total[$i].Job
    $row.Id = $total[$i].Number
    $row.Comments=$total[$i].Comments
    $table.Rows.Add($row)
}

# Create an HTML version of the DataTable
$html = "<table id='tabid'><tr><th>Job</th><th>Id</th><th>Comments</th></tr>"
foreach ($row in $table.Rows)
{ 
    $html += "<tr><td>" + $row[0] + "</td>"+"<td>" + $row[1] + "</td>"+"<td>" + $row[2] + "</td></tr>"
}
$html += "</table>"

# Send the email


$body2="
<style>
#tabid {
  font-family: 'Trebuchet MS', Arial, Helvetica, sans-serif;
  border-collapse: collapse;
  width: 100%;
}

#tabid td, #Job th {
  border: 1px solid #ddd;
  padding: 8px;
}

#tabid tr:nth-child(even){background-color: #f2f2f2;}

#tabid tr:hover {background-color: #ddd;}

#tabid th {
  padding-top: 12px;
  padding-bottom: 12px;
  text-align: left;
  background-color: #4CAF50;
  color: white;
}
</style>

"


$body = "<!DOCTYPE html> <html>Team,<br />Please fine the Shift handover details  below:<br /><br /></br>" +$body2+ $html+"</body></html>"


Send-MailMessage -From "mallikarjunar@workspot.com" -To "Support-India@workspot.com" -Bcc "ashisha@workspot.com" -Body $body -Subject $sub -SmtpServer smtp.office365.com -UseSsl -Credential $cr -BodyAsHtml
