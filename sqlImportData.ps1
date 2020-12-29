$sqlConnOut = New-Object System.Data.SqlClient.SqlConnection
$sqlConnOut.ConnectionString = "Server=localhost;Integrated Security=true;Initial Catalog=AdventureWorksDW2016"
$sqlConnOut.Open()

<# GET Data From IAM #>

$sqlcmd = New-Object System.Data.SqlClient.SqlCommand
$sqlcmd.Connection = $sqlConnOut
$query = "SELECT CurrencyKey, CurrencyAlternateKey, CurrencyName FROM dbo.DimCurrency"
$sqlcmd.CommandText = $query

$adp = New-Object System.Data.SqlClient.SqlDataAdapter $sqlcmd

$data = New-Object System.Data.DataSet
$adp.Fill($data) | Out-Null

$ds = $data.Tables[0]


<# Clear IAM table in Car Booking DB #>

function Clean-LDAPUsersTable($OpenSQLConnection){

    $sqlCommand = New-Object System.Data.SqlClient.SqlCommand

    $sqlCommand.Connection = $OpenSQLConnection

    $sqlCommand.CommandText = "TRUNCATE TABLE [Testowa].[dbo].[DimCurrency]"

    try{
        $sqlCommand.ExecuteNonQuery()
    }
    catch{
        Write-Host "An error occurred:"
        Write-Host $_.ScriptStackTrace
    }
    

}

<# SAVE data function in Car Booking DB #>

function Do-IAMInsertRowByRow ($OpenSQLConnection, $IAMds) {

    $sqlCommand = New-Object System.Data.SqlClient.SqlCommand

    $sqlCommand.Connection = $OpenSQLConnection

    $sqlCommand.CommandText = "SET NOCOUNT ON; " +

        "SET IDENTITY_INSERT dbo.DimCurrency ON;" +

        "INSERT INTO dbo.DimCurrency (CurrencyKey,CurrencyAlternateKey,CurrencyName) " +

        "VALUES (@CurrencyKey,@CurrencyAlternateKey,@CurrencyName); " +

        "SET IDENTITY_INSERT dbo.DimCurrency OFF;"

 
    $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@CurrencyKey",[Data.SQLDBType]::Int))) | Out-Null

    $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@CurrencyAlternateKey",[Data.SQLDBType]::NChar, 3))) | Out-Null

    $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@CurrencyName",[Data.SQLDBType]::NVarChar, 50))) | Out-Null

    $sqlCommand.Transaction = $OpenSQLConnection.BeginTransaction()
   
    foreach ($Row in $IAMds.Rows) {

        # Here we set the values of the pre-existing parameters based on the $IAMds.Rows iterator

        $sqlCommand.Parameters[0].Value = $Row[0]

        $sqlCommand.Parameters[1].Value = $Row[1]

        $sqlCommand.Parameters[2].Value = $Row[2]

        # Run the query and get the scope ID back into $InsertedID

      
        try { 
            
            $sqlCommand.ExecuteNonQuery()
            
        }
        catch {
        Write-Host "An error occurred:"
        Write-Host $_.ScriptStackTrace

        $sqlCommand.Transaction.Rollback()
        
        Exit

        }

    }

    $sqlCommand.Transaction.Commit()

}
 <# Execute import data #>


$sqlConnection = New-Object System.Data.SqlClient.SqlConnection

$sqlConnection.ConnectionString = "Server=localhost;Database=Testowa;Integrated Security=True;"

$sqlConnection.Open()

 

# Quit if the SQL connection didn't open properly.

if ($sqlConnection.State -ne [Data.ConnectionState]::Open) {

    write-host "Connection to DB is not open."

    Exit

}


# Call the function that does the inserts and function for clean LDAPUsers.
Clean-LDAPUsersTable $sqlConnection
Do-IAMInsertRowByRow $sqlConnection $ds

# Close the connection.

if ($sqlConnection.State -eq [Data.ConnectionState]::Open) {

    $sqlConnection.Close()

}
if ($sqlConnOut -eq [Data.ConnectionState]::Open) {

    $sqlConnOut.Close()

}

