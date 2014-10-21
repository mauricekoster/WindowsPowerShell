$odpAssemblyName = "Oracle.DataAccess, Version=2.112.3.0, Culture=neutral, PublicKeyToken=89b483f429c47342"
[System.Reflection.Assembly]::Load($odpAssemblyName)

# $asm = [appdomain]::currentdomain.getassemblies() | where-object {$_.FullName -eq $odpAssemblyName}
# $asm.GetTypes() | Where-Object {$_.IsPublic} | Sort-Object {$_.FullName } | ft FullName, BaseType | Out-String

function Select-TNS
{
    $enu = New-Object Oracle.DataAccess.Client.OracleDataSourceEnumerator
    Write-Output $enu.GetDataSources()
}


function Connect-Oracle([string] $connectionString = $(throw "connectionString is required"))
{
  $conn= New-Object Oracle.DataAccess.Client.OracleConnection($connectionString)
  $conn.Open()
  Write-Output $conn
}

function Get-ConnectionString($user, $pass, $hostName, $port, $sid)
{
  $dataSource = ("(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={0})  (PORT={1}))(CONNECT_DATA=(SERVICE_NAME={2})))" -f $hostName, $port,  $sid)
  Write-Output ("Data Source={0};User Id={1};Password={2};Connection Timeout=10" -f $dataSource, $user, $pass)
}

function Get-ConfigConnectionString(
    [string] $filename = $(throw "filename is required"),
    [string] $name = $(throw "connection string name is required")
    )
{
    $config = [xml](Get-Content $filename)
    $item = $config.configuration.connectionStrings.add | where {$_.name -eq $name}
    if (!$item) { 
        throw "Failed to find a connection string with name '{0}'" -f $name
    }
    return $item.connectionString
}

function Get-DataTable
{
  Param(
    [Parameter(Mandatory=$true)]
    [Oracle.DataAccess.Client.OracleConnection]$conn,
    [Parameter(Mandatory=$true)]
    [string]$sql
  )
  $cmd = New-Object Oracle.DataAccess.Client.OracleCommand($sql,$conn)
  $da = New-Object Oracle.DataAccess.Client.OracleDataAdapter($cmd)
  $dt = New-Object System.Data.DataTable
  [void]$da.Fill($dt)
  return ,$dt
}

function Get-ConnectionString-COM
{
  Get-ConnectionString COM_OWNER COM_DEV_01 localhost 1521 XE
}
