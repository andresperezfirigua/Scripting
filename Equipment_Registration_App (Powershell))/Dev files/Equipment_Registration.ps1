#$global:databasePath = "C:\Equipment_Registration_App\alm_hardware.db"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Hide the console window
Add-Type -Name Window -Namespace ConsoleApp -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

public static void Hide()
{
    IntPtr hWnd = GetConsoleWindow();
    if (hWnd != IntPtr.Zero)
    {
        // 0 = Hide the window
        ShowWindow(hWnd, 0);
    }
}'
[ConsoleApp.Window]::Hide()

Import-Module -Name SQLite

$global:databasePath = "C:\Equipment_Registration_App\alm_hardware.db"

function Clean_Controls {
    $textboxSearch.Text = ""
    $radioIngreso.Checked = $false
    $radioRetiro.Checked = $false
}

function Add_Non_Asurion_Device {
    $timezone = [System.TimeZoneInfo]::FindSystemTimeZoneById("SA Pacific Standard Time")
    $date = [System.TimeZoneInfo]::ConvertTimeFromUtc((Get-Date).ToUniversalTime(), $timezone)
    $dateString = $date.ToString("yyyy-MM-dd HH:mm:ss")

    $tipoRegistro = ""

    if ($radioIngreso.Checked) {
        $tipoRegistro = "Ingreso"
    }
    elseif ($radioRetiro.Checked) {
        $tipoRegistro = "Retiro"
    }
    function ValidateTextbox {
        if ([string]::IsNullOrWhiteSpace($textboxSerial.Text)) {
            [System.Windows.Forms.MessageBox]::Show([System.Text.RegularExpressions.Regex]::Unescape("Indique el n\u00FAmero serial del equipo"), "Error", "OK", "Error")
            return $true
        } elseif ([string]::IsNullOrWhiteSpace($textboxModel.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Indique el modelo o referencia del equipo", "Error", "OK", "Error")
            return $true
        } elseif ([string]::IsNullOrWhiteSpace($textboxBearer.Text)) {
            [System.Windows.Forms.MessageBox]::Show([System.Text.RegularExpressions.Regex]::Unescape("Indique qui\u00E9n lleva el equipo"), "Error", "OK", "Error")
            return $true
        }
        return $false
    }
    # Create a form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = ""
    $form.Size = New-Object System.Drawing.Size(400, 280)
    $form.StartPosition = "CenterScreen"
    $form.ControlBox = $false
    $form.MaximizeBox = $false
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    # Create labels and textboxes
    $labelSerial = New-Object System.Windows.Forms.Label
    $labelSerial.Location = New-Object System.Drawing.Point(20, 20)
    $labelSerial.Size = New-Object System.Drawing.Size(100, 20)
    $labelSerial.Text = [System.Text.RegularExpressions.Regex]::Unescape("N\u00FAmero serial:")
    $form.Controls.Add($labelSerial)

    $textboxSerial = New-Object System.Windows.Forms.TextBox
    $textboxSerial.Location = New-Object System.Drawing.Point(150, 20)
    $textboxSerial.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($textboxSerial)
    $textboxSerial.Text = $textboxSearch.Text

    $labelModel = New-Object System.Windows.Forms.Label
    $labelModel.Location = New-Object System.Drawing.Point(20, 60)
    $labelModel.Size = New-Object System.Drawing.Size(100, 20)
    $labelModel.Text = "Modelo:"
    $form.Controls.Add($labelModel)

    $textboxModel = New-Object System.Windows.Forms.TextBox
    $textboxModel.Location = New-Object System.Drawing.Point(150, 60)
    $textboxModel.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($textboxModel)

    $labelType = New-Object System.Windows.Forms.Label
    $labelType.Location = New-Object System.Drawing.Point(20, 100)
    $labelType.Size = New-Object System.Drawing.Size(120, 20)
    $labelType.Text = "Tipo de dispositivo:"
    $form.Controls.Add($labelType)

    $dropdownType = New-Object System.Windows.Forms.ComboBox
    $dropdownType.Location = New-Object System.Drawing.Point(150, 100)
    $dropdownType.Size = New-Object System.Drawing.Size(200, 20)
    $dropdownType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $dropdownType.Items.Add("")
    $dropdownType.Items.Add("PC / Laptop")
    $dropdownType.Items.Add("Otro")
    $dropdownType.SelectedIndex = 0
    $form.Controls.Add($dropdownType)

    $labelBearer = New-Object System.Windows.Forms.Label
    $labelBearer.Location = New-Object System.Drawing.Point(20, 140)
    $labelBearer.Size = New-Object System.Drawing.Size(200, 20)
    $labelBearer.Text = "Nombre de quien lleva el equipo:"
    $form.Controls.Add($labelBearer)

    $textboxBearer = New-Object System.Windows.Forms.TextBox
    $textboxBearer.Location = New-Object System.Drawing.Point(220, 140)
    $textboxBearer.Size = New-Object System.Drawing.Size(130, 20)
    $form.Controls.Add($textboxBearer)

    $labelAdditional = New-Object System.Windows.Forms.Label
    $labelAdditional.Location = New-Object System.Drawing.Point(20, 180)
    $labelAdditional.Size = New-Object System.Drawing.Size(200, 20)
    $labelAdditional.Text = [System.Text.RegularExpressions.Regex]::Unescape("Informaci\u00F3n adicional:")
    $form.Controls.Add($labelAdditional)

    $textboxAdditional = New-Object System.Windows.Forms.TextBox
    $textboxAdditional.Location = New-Object System.Drawing.Point(220, 180)
    $textboxAdditional.Size = New-Object System.Drawing.Size(130, 20)
    $form.Controls.Add($textboxAdditional)

    # Create a button for bearer
    $buttonAddDevice = New-Object System.Windows.Forms.Button
    $buttonAddDevice.Location = New-Object System.Drawing.Point(90, 230)
    $buttonAddDevice.Size = New-Object System.Drawing.Size(200, 30)
    $buttonAddDevice.Text = "Registrar"
    $buttonAddDevice.Add_Click({
        if (-not (ValidateTextbox)){
            $newComputer = @()
            $newComputer += $textboxSerial.Text.ToUpper()
            $newComputer += "N/A"
            $newComputer += "N/A"
            $newComputer += $textboxModel.Text
            if ((-not ($dropdownType.SelectedItem -eq ""))) {
                $newComputer += $dropdownType.SelectedItem
            } else {
                [System.Windows.Forms.MessageBox]::Show("Seleccione el tipo de dispositivo", "Error", "OK", "Error")
                return
            }
            $newComputer += $dateString
            $newComputer += $tipoRegistro
            $newComputer += $textboxBearer.Text
            $newComputer += [System.Environment]::UserName.ToUpper()
            if ($null -eq $dropdownType.SelectedItem) {
                $newComputer += "N/A"
            } else {
                $newComputer += $textboxAdditional.Text
            }
            $form.Close()
            $form.Dispose()
            Add_Record_To_File $newComputer
        }
    })
    $form.Controls.Add($buttonAddDevice)
    # Show the form
    $form.ShowDialog() | Out-Null
}

function Display_Recent_Records {
    # Create connection to the SQLite database
    $connection = New-Object System.Data.SQLite.SQLiteConnection
    $connection.ConnectionString = "Data Source=$global:databasePath;Version=3;"
 
    $connection.Open()

    $command = $connection.CreateCommand()

    $today = Get-Date -Format 'yyyy-MM-dd'
    $query = "SELECT serial_number, asset_tag, assigned_to, model, registration_date, registration_type, equipment_carrier, registration_verifier FROM (SELECT serial_number, asset_tag, assigned_to, model, registration_date, registration_type, equipment_carrier, registration_verifier FROM Registration WHERE registration_date LIKE '$today%') ORDER BY registration_date ASC"
    
    # Set the command text to the SQL query
    $command.CommandText = $query

    $adapter = New-Object System.Data.SQLite.SQLiteDataAdapter($command)

    $connection.Close()

    return $adapter
}

function Find_Computer($searchValue) {
    $foundItem = @()

    $timezone = [System.TimeZoneInfo]::FindSystemTimeZoneById("SA Pacific Standard Time")
    $date = [System.TimeZoneInfo]::ConvertTimeFromUtc((Get-Date).ToUniversalTime(), $timezone)
    $dateString = $date.ToString("yyyy-MM-dd HH:mm:ss")

    $tipoRegistro = ""

    if ($radioIngreso.Checked) {
        $tipoRegistro = "Ingreso"
    }
    elseif ($radioRetiro.Checked) {
        $tipoRegistro = "Retiro"
    }

    # Create connection to the SQLite database
    $connection = New-Object System.Data.SQLite.SQLiteConnection
    $connection.ConnectionString = "Data Source=$global:databasePath;Version=3;"
 
    $connection.Open()

    $command = $connection.CreateCommand()

    # Set the SQL query to search for the computer by serial number or asset tag
    $query = "SELECT serial_number, asset_tag, assigned_to, model, device_type FROM Computers WHERE serial_number = '$searchValue' OR asset_tag = '$searchValue'"

    # Set the command text to the SQL query
    $command.CommandText = $query

    # Execute the SQL query and get the result set
    $reader = $command.ExecuteReader()

    # Check if the result set has any rows
    if ($reader.HasRows) {
        $row_count = 0
        # Loop through the rows in the result set
        while ($reader.Read()) {
            if ($row_count -eq 0){
                # Add the values in the row to the array
                for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                    $value = $reader.GetValue($i)
                    $stringValue = $value.ToString()
                    if (($i -eq 2) -and ($stringValue -eq "")) {
                        $foundItem += "No asignado"
                    } else {
                        $foundItem += $stringValue
                    }
                }
            } else {
                break
            }
            $row_count += 1
        }
    } else {
        return $null
    }

    # Add the current date/time to the next cell
    $foundItem += $dateString

    # Add the type of registration to the next cell
    $foundItem += $tipoRegistro

    # Add the name of the person that has the computer to the next cell
    $foundItem += $foundItem[2]

    # Add the person who verifies the registration to the next cell
    $foundItem += [System.Environment]::UserName.ToUpper()

    # Close the reader and the connection to the database
    $reader.Close()
    $connection.Close()

    return $foundItem
}

function Add_Record_To_File ($foundItem) {
    $tableName = "Registration"

    $connection = New-Object -TypeName System.Data.SQLite.SQLiteConnection -ArgumentList "Data Source=$global:databasePath;Version=3;"
    $connection.Open()

    $command = $connection.CreateCommand()
    $command.CommandText = "INSERT INTO $tableName (serial_number, asset_tag, assigned_to, model, device_type, registration_date, registration_type, equipment_carrier, registration_verifier, additional_info) VALUES (@serial_number, @asset_tag, @assigned_to, @model, @device_type, @registration_date, @registration_type, @equipment_carrier, @registration_verifier, @additional_info)"

    $param1 = New-Object -TypeName System.Data.SQLite.SQLiteParameter
    $param1.ParameterName = "@serial_number"
    $param1.Value = $foundItem[0]
    $command.Parameters.Add($param1)

    $param2 = New-Object -TypeName System.Data.SQLite.SQLiteParameter
    $param2.ParameterName = "@asset_tag"
    $param2.Value = $foundItem[1]
    $command.Parameters.Add($param2)

    $param3 = New-Object -TypeName System.Data.SQLite.SQLiteParameter
    $param3.ParameterName = "@assigned_to"
    $param3.Value = $foundItem[2]
    $command.Parameters.Add($param3)

    $param4 = New-Object -TypeName System.Data.SQLite.SQLiteParameter
    $param4.ParameterName = "@model"
    $param4.Value = $foundItem[3]
    $command.Parameters.Add($param4)

    $param5 = New-Object -TypeName System.Data.SQLite.SQLiteParameter
    $param5.ParameterName = "@device_type"
    $param5.Value = $foundItem[4]
    $command.Parameters.Add($param5)

    $param6 = New-Object -TypeName System.Data.SQLite.SQLiteParameter
    $param6.ParameterName = "@registration_date"
    $param6.Value = $foundItem[5]
    $command.Parameters.Add($param6)

    $param7 = New-Object -TypeName System.Data.SQLite.SQLiteParameter
    $param7.ParameterName = "@registration_type"
    $param7.Value = $foundItem[6]
    $command.Parameters.Add($param7)

    $param8 = New-Object -TypeName System.Data.SQLite.SQLiteParameter
    $param8.ParameterName = "@equipment_carrier"
    $param8.Value = $foundItem[7]
    $command.Parameters.Add($param8)

    $param9 = New-Object -TypeName System.Data.SQLite.SQLiteParameter
    $param9.ParameterName = "@registration_verifier"
    $param9.Value = $foundItem[8]
    $command.Parameters.Add($param9)

    $param10 = New-Object -TypeName System.Data.SQLite.SQLiteParameter
    $param10.ParameterName = "@additional_info"
    $param10.Value = $foundItem[9]
    $command.Parameters.Add($param10)

    $command.ExecuteNonQuery()
    $command.Parameters.Clear()
    
    $connection.Close()
}

function Get_Different_Bearer ($foundItem) {
    function ValidateBearer {
        if ([string]::IsNullOrWhiteSpace($textboxBearer.Text)) {
            [System.Windows.Forms.MessageBox]::Show([System.Text.RegularExpressions.Regex]::Unescape("Indique qui\u00E9n lleva el equipo"), "Error", "OK", "Error")
            return $true
        }
        return $false
    }
    # Create a form
    $formBearer = New-Object System.Windows.Forms.Form
    $formBearer.Text = ""
    $formBearer.Size = New-Object System.Drawing.Size(237, 178)
    $formBearer.StartPosition = "CenterScreen"
    $formBearer.ControlBox = $false
    $formBearer.MaximizeBox = $false
    $formBearer.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    # Create a label and text box for bearer
    $labelBearer = New-Object System.Windows.Forms.Label
    $labelBearer.Location = New-Object System.Drawing.Point(19, 50)
    $labelBearer.Size = New-Object System.Drawing.Size(200, 20)
    $labelBearer.Text = [System.Text.RegularExpressions.Regex]::Unescape("Qui\u00E9n lleva el equipo?")
    $formBearer.Controls.Add($labelBearer)

    $textboxBearer = New-Object System.Windows.Forms.TextBox
    $textboxBearer.Location = New-Object System.Drawing.Point(19, 80)
    $textboxBearer.Size = New-Object System.Drawing.Size(200, 20)
    $formBearer.Controls.Add($textboxBearer)

    # Create a button for bearer
    $buttonBearer = New-Object System.Windows.Forms.Button
    $buttonBearer.Location = New-Object System.Drawing.Point(19, 110)
    $buttonBearer.Size = New-Object System.Drawing.Size(200, 30)
    $buttonBearer.Text = "OK"
    $buttonBearer.Add_Click({
        if (-not (ValidateBearer)){
            $boxText = $textboxBearer.Text
            $cultureInfo = [System.Globalization.CultureInfo]::InvariantCulture
            $boxTextChanged = $cultureInfo.TextInfo.ToTitleCase($boxText.ToLower())
            $foundItem[7] = $boxTextChanged
            $formBearer.Close()
        }
    })
    $formBearer.Controls.Add($buttonBearer)
    # Show the form
    $formBearer.ShowDialog() | Out-Null
    $formBearer.Dispose()
    return $foundItem
}

function Record_Equipment {
    if ([string]::IsNullOrWhiteSpace($textboxSearch.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Ingrese un equipo para buscar", "Error", "OK", "Error")
        return
    }
    $searchToUpper = $textboxSearch.Text.ToUpper()
    $searchValue = $searchToUpper

    $foundItem = Find_Computer $searchValue

    if ($null -ne $foundItem){
        if (-not ($foundItem[2] -eq "No asignado")) {
            $bearer = $foundItem[7]
            $dialogResult = [System.Windows.Forms.MessageBox]::Show("Es $bearer quien tiene el equipo a registrar?", "Confirmation", "YesNoCancel", "Question")

            if ($dialogResult -eq "Yes") {
                Add_Record_To_File $foundItem
                # Display a message to the user indicating that the data was added successfully
                [System.Windows.Forms.MessageBox]::Show("Registro agregado exitosamente.", "", "OK", "Information")
                Clean_Controls
            }
            elseif ($dialogResult -eq "No") {
                $modifiedFoundItem = Get_Different_Bearer $foundItem

                Add_Record_To_File $modifiedFoundItem
                # Display a message to the user indicating that the data was added successfully
                [System.Windows.Forms.MessageBox]::Show("Registro agregado exitosamente.", "", "OK", "Information")
                Clean_Controls
            }
            elseif ($dialogResult -eq "Cancel") {
                Clean_Controls
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show([System.Text.RegularExpressions.Regex]::Unescape("Este equipo a\u00FAn no est\u00E1 asignado en inventario. De clic en OK para registrarlo"), "Error", "OK", "Error")
            $modifiedFoundItem = Get_Different_Bearer $foundItem

            Add_Record_To_File $modifiedFoundItem
            # Display a message to the user indicating that the data was added successfully
            [System.Windows.Forms.MessageBox]::Show("Registro agregado exitosamente.", "", "OK", "Information")
            Clean_Controls
        }
    } else {
        $add_non_asurion_device = [System.Windows.Forms.MessageBox]::Show("El equipo no fue encontrado.`n`n- Recuerda que solo se pueden registran computadores`n- Este puede ser un computador de Asurion pero en otro pais`n- Este computador puede no ser propiedad de Asurion`n`n Desea registrar este equipo?", "Confirmation", "YesNo", "Error")
        if ($add_non_asurion_device -eq "Yes") {
            Add_Non_Asurion_Device
            # Display a message to the user indicating that the data was added successfully
            [System.Windows.Forms.MessageBox]::Show("Registro agregado exitosamente.", "", "OK", "Information")
            Clean_Controls
        } elseif ($add_non_asurion_device -eq "No") {
            Clean_Controls
        }
    }

    $table.Clear()
    $adapter = Display_Recent_Records
    [void]$adapter.Fill($table)
    $dataGridView.DataSource = $table
    $dataGridView.Refresh()
    foreach ($col in $dataGridView.Columns) {
        $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    }
    $dataGridView.FirstDisplayedScrollingRowIndex = $dataGridView.Rows.Count - 1
}

# Load the Excel file into a variable
if (Test-Path $global:databasePath) {
    # Event handler for FormClosing event
    $handler_FormClosing = {
        # Force garbage collection to release memory
        [GC]::Collect()
    }

    # Create a form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Equipment Registration"
    $form.Size = New-Object System.Drawing.Size(1150, 510)
    $form.StartPosition = "CenterScreen"
    $form.MaximizeBox = $false
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    # Create a label and text box for search value
    $labelSearch = New-Object System.Windows.Forms.Label
    $labelSearch.Location = New-Object System.Drawing.Point(475, 20)
    $labelSearch.Size = New-Object System.Drawing.Size(200, 30)
    $labelSearch.Text = "Ingrese el equipo a buscar: `n(Se recomienda buscar por serial)"
    $form.Controls.Add($labelSearch)

    $textboxSearch = New-Object System.Windows.Forms.TextBox
    $textboxSearch.Location = New-Object System.Drawing.Point(475, 50)
    $textboxSearch.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($textboxSearch)

    $radioIngreso = New-Object System.Windows.Forms.RadioButton
    $radioIngreso.Location = New-Object System.Drawing.Point(475, 80)
    $radioIngreso.Size = New-Object System.Drawing.Size(200, 30)
    $radioIngreso.Text = "Ingreso de equipo"
    $form.Controls.Add($radioIngreso)

    $radioRetiro = New-Object System.Windows.Forms.RadioButton
    $radioRetiro.Location = New-Object System.Drawing.Point(475, 110)
    $radioRetiro.Size = New-Object System.Drawing.Size(200, 30)
    $radioRetiro.Text = "Retiro de equipo"
    $form.Controls.Add($radioRetiro)

    # Create a button for Registrar equipo
    $buttonRegistro = New-Object System.Windows.Forms.Button
    $buttonRegistro.Location = New-Object System.Drawing.Point(475, 150)
    $buttonRegistro.Size = New-Object System.Drawing.Size(200, 30)
    $buttonRegistro.Text = "Registrar"
    $form.add_FormClosing($handler_FormClosing)
    $buttonRegistro.Add_Click({
            if ($radioIngreso.Checked -or $radioRetiro.Checked) {
                Record_Equipment
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Indique si el equipo ingresa o sale del piso", "Error", "OK", "Error")
            }
        })
    $form.Controls.Add($buttonRegistro)

    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Location = New-Object System.Drawing.Point(25, 190)
    $dataGridView.Size = New-Object System.Drawing.Size(1085, 270)
    $dataGridView.ReadOnly = $true
    $dataGridView.AllowUserToResizeColumns = $false
    $dataGridView.AllowUserToResizeRows = $false
    $dataGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $dataGridView.RowHeadersWidthSizeMode = [System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode]::DisableResizing
    $dataGridView.RowHeadersWidth = 15
    $dataGridView.AllowUserToOrderColumns = $false
    $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

    $form.Controls.Add($dataGridView)

    $table = New-Object System.Data.DataTable

    $form.add_Load({
        $adapter = Display_Recent_Records
        [void]$adapter.Fill($table)
        $dataGridView.DataSource = $table
        $dataGridView.Refresh()
        foreach ($col in $dataGridView.Columns) {
            $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
        }
        $dataGridView.FirstDisplayedScrollingRowIndex = $dataGridView.Rows.Count - 1
    })

    # Show the form
    $form.ShowDialog() | Out-Null

} else {
    [System.Windows.Forms.MessageBox]::Show([System.Text.RegularExpressions.Regex]::Unescape("No se encontr\u00F3 base de datos"), "Error", "OK", "Error")
}