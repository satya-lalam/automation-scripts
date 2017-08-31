function Get-ODBC-Data{
   param([string]$query=$(throw 'query is required.'))
   $conn = New-Object System.Data.Odbc.OdbcConnection
   $conn.ConnectionString = "DSN=<datasource name>;"
   $conn.open()
   $cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
   $ds = New-Object system.Data.DataSet
   (New-Object system.Data.odbc.odbcDataAdapter($cmd)).fill($ds) | out-null
   $conn.close()
   $ds.Tables[0]
}

function Set-ODBC-Data{
  param([string]$query=$(throw 'query is required.'))
  $conn = New-Object System.Data.Odbc.OdbcConnection
  $conn.ConnectionString= "DSN=<datasource name>;"
  $cmd = new-object System.Data.Odbc.OdbcCommand($query,$conn)
  $conn.open()
  $cmd.ExecuteNonQuery()
  $conn.close()
}

$query = "<query>"
$result = Get-ODBC-Data -query $query
set-odbc-data -query $query
Write-Output $result > <location to store result>
