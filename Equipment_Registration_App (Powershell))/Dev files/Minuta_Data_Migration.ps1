# Import the required modules
Import-Module -Name ImportExcel
Import-Module -Name SQLite

# Set the paths to the Excel file and the SQLite database
$excelPath = "C:\Users\andresfelipe.perez\PycharmProjects\Equipment_Registration_App\Dev files\Minuta_computadores.xlsx"
$databasePath = "C:\Users\andresfelipe.perez\PycharmProjects\Equipment_Registration_App\Equipment_Registration_App\alm_hardware.db"

# Import the data from the Excel file
$data = Import-Excel -Path $excelPath

# Connect to the SQLite database
$connection = New-Object System.Data.SQLite.SQLiteConnection

$connection.ConnectionString = "Data Source=$databasePath;Version=3;"

$connection.Open()

# Create a command object
$command = $connection.CreateCommand()

# Loop through the rows of data and insert them into the database
foreach ($row in $data) {
    $query = "INSERT INTO Registration (serial_number, asset_tag, assigned_to, model, device_type, registration_date, registration_type, equipment_carrier, registration_verifier, additional_info) VALUES ('$($row.'Numero serial')', '$($row.'Codigo de activo')', '$($row.'Asignado a')', '$($row.'Modelo / Referencia')', '$($row.'Tipo de dispositivo')', '$($row.'Fecha y hora de registro')', '$($row.'Tipo de registro')', '$($row.'Portador del equipo')', '$($row.'Guarda verificador')', '$($row.'Informacion adicional')')"
    $command.CommandText = $query
    $command.ExecuteNonQuery()
}

# Close the connection to the database
$connection.Close()