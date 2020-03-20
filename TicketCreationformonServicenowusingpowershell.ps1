Add-Type -AssemblyName system.windows.forms
Add-Type -AssemblyName system.drawing

$form1 = New-Object  system.windows.forms.form
$form1.text = "Ticket Creation Form"
$form1.size = New-Object System.Drawing.Size(400,400)
$form1.StartPosition = 'centerscreen'

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(5,12)
$label1.Size = New-Object System.Drawing.Size(135,30)
$label1.Text = 'Enter Requester ID:'
$form1.Controls.Add($label1)

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(200,10)
$textBox1.Size = New-Object System.Drawing.Size(150,80)
$form1.Controls.Add($textBox1)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(5,42)
$label2.Size = New-Object System.Drawing.Size(180,30)
$label2.Text = 'Enter Source:'
$form1.Controls.Add($label2)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(200,40)
$textBox2.Size = New-Object System.Drawing.Size(150,80)
$form1.Controls.Add($textBox2)

$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(5,72)
$label3.Size = New-Object System.Drawing.Size(180,30)
$label3.Text = 'Enter Contact Number:'
$form1.Controls.Add($label3)

$textBox3 = New-Object System.Windows.Forms.TextBox
$textBox3.Location = New-Object System.Drawing.Point(200,70)
$textBox3.Size = New-Object System.Drawing.Size(150,80)
$form1.Controls.Add($textBox3)

$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Point(5,102)
$label4.Size = New-Object System.Drawing.Size(180,30)
$label4.Text = 'Enter Summary:'
$form1.Controls.Add($label4)

$textBox4 = New-Object System.Windows.Forms.TextBox
$textBox4.Location = New-Object System.Drawing.Point(200,100)
$textBox4.Size = New-Object System.Drawing.Size(150,80)
$form1.Controls.Add($textBox4)

$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Point(5,132)
$label5.Size = New-Object System.Drawing.Size(180,30)
$label5.Text = 'Enter Incident status:'
$form1.Controls.Add($label5)

$textBox5 = New-Object System.Windows.Forms.TextBox
$textBox5.Location = New-Object System.Drawing.Point(200,130)
$textBox5.Size = New-Object System.Drawing.Size(150,80)
$form1.Controls.Add($textBox5)

$label6 = New-Object System.Windows.Forms.Label
$label6.Location = New-Object System.Drawing.Point(5,162)
$label6.Size = New-Object System.Drawing.Size(180,30)
$label6.Text = 'Enter Assignment Group:'
$form1.Controls.Add($label6)

$textBox6 = New-Object System.Windows.Forms.TextBox
$textBox6.Location = New-Object System.Drawing.Point(200,160)
$textBox6.Size = New-Object System.Drawing.Size(150,80)
$form1.Controls.Add($textBox6)

$label7 = New-Object System.Windows.Forms.Label
$label7.Location = New-Object System.Drawing.Point(5,192)
$label7.Size = New-Object System.Drawing.Size(180,30)
$label7.Text = 'Enter Category:'
$form1.Controls.Add($label7)

$arrayCategory=@("Cluster","Data Link","Database","Physical Server","Virtual Server","IP Address")
$comboBox1 = New-Object System.Windows.Forms.ComboBox 
$comboBox1.Location = New-Object System.Drawing.Size(200,190) 
$comboBox1.Size = New-Object System.Drawing.Size(150,80) 
$comboBox1.DropDownHeight = 200 
$form1.Controls.Add($comboBox1)

$label8 = New-Object System.Windows.Forms.Label
$label8.Location = New-Object System.Drawing.Point(5,222)
$label8.Size = New-Object System.Drawing.Size(180,30)
$label8.Text = 'Enter Sub Category:'
$form1.Controls.Add($label8)

$comboBox2 = New-Object System.Windows.Forms.ComboBox 
$comboBox2.Location = New-Object System.Drawing.Size(200,220) 
$comboBox2.Size = New-Object System.Drawing.Size(150,80) 
$comboBox2.Height = 80

foreach ($category in $arraycategory)
    {
    $comboBox1.Items.Add($category)                     
    } 

$comboBoxCompany_SelectedIndexChanged=
    {   
    If ($comboBox1.text -eq "Cluster") 
    {
    [void] $comboBox2.Items.Add("cluster")
    }
    elseif($comboBox1.text -eq "Data Link")
    {
    [void] $comboBox2.Items.Add("Data Link")
    }
    elseif($comboBox1.text -eq "Database")
    {
    [void] $comboBox2.Items.Add("DB2")
    [void] $comboBox2.Items.Add("MYSQL")
    [void] $comboBox2.Items.Add("Oracle")
    [void] $comboBox2.Items.Add("Sql Server")
    [void] $comboBox2.Items.Add("Sybase")
    }
    elseif($comboBox1.text -eq "Physical Server")
    {
    [void] $comboBox2.Items.Add("AS/400")
    [void] $comboBox2.Items.Add("Linux Server")
    [void] $comboBox2.Items.Add("Mac Server")
    [void] $comboBox2.Items.Add("Unix Server")
    [void] $comboBox2.Items.Add("Mainframe")
    }
    elseif($comboBox1.text -eq "Virtual Server")
    {
    [void] $comboBox2.Items.Add("Linux Server")
    [void] $comboBox2.Items.Add("Mac Server")
    [void] $comboBox2.Items.Add("Unix Server")
    [void] $comboBox2.Items.Add("Windows Server")
    }
    elseif($comboBox1.text -eq "IP Address")
    {
    [void] $comboBox2.Items.Add("IP Address")
    }
    $form1.Controls.Add($comboBox2)
    }
$comboBox1.add_SelectedIndexChanged($comboBoxCompany_SelectedIndexChanged)
    
$submitbutton = New-Object  system.windows.forms.button
$submitbutton.location = New-Object System.Drawing.Point(80,300)
$submitbutton.size = new-object system.drawing.size(75,23)
$submitbutton.text = 'SUBMIT'
$form1.AcceptButton = $submitButton
$form1.Controls.Add($submitButton)

$Resetbutton = New-Object  system.windows.forms.button
$Resetbutton.location = New-Object System.Drawing.Point(180,300)
$Resetbutton.size = new-object system.drawing.size(75,23)
$Resetbutton.text = 'RESET'
$form1.AcceptButton = $ResetButton
$form1.Controls.Add($ResetButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(280,300)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
#$form.CancelButton = $CancelButton
$form1.Controls.Add($CancelButton)


$submitbutton.Add_Click({ sendRequest; Submit })

function sendRequest()
    {
    $Body = @{
    caller_id = $textBox1.Text
    contact_type = $textBox2.Text
    u_contact_number = $textBox3.Text
    short_description = $textBox4.Text
    incident_state=$textBox5.Text
    assignment_group=$textBox6.Text
    category=$comboBox1.Text
    subcategory=$comboBox2.Text}
    $body=$Body | ConvertTo-Json
    $username="mydbdatasync2@hcl.com"
    $password=ConvertTo-SecureString 'Comnet@123' -AsPlainText  -Force
    $credential=New-Object System.Management.Automation.PSCredential -ArgumentList $username,$password
    [net.servicepointmanager]::SecurityProtocol=[net.securityProtocolType]::Tls12 
    $uri="https://hclgbpdev.service-now.com/api/now/table/incident"
    $result=Invoke-RestMethod -Uri $uri -Method Post -Body $body -Credential $credential -ContentType "application/json"
    $script:number=$result.result.number
    }

function Submit()
    {
    [System.Windows.Forms.MessageBox]::Show("Ticket ($number) have been created successfully !" , "SUBMIT SUCCESSFULLY")
    $textBox1.Text = ''
    $textBox2.Text = ''
    $textBox3.Text = ''
    $textBox4.Text = ''
    $textBox5.Text = ''
    $textBox6.Text = ''
    $comboBox1.Text = ''
    $comboBox2.Text = ''
    }
$Resetbutton.Add_Click({ Reset })

function Reset ()
    {
    [System.Windows.Forms.MessageBox]::Show("Reset the form!" , "RESET")
    $textBox1.Text = ''
    $textBox2.Text = ''
    $textBox3.Text = ''
    $textBox4.Text = ''
    $textBox5.Text = ''
    $textBox6.Text = ''
    $comboBox1.Text = ''
    $comboBox2.Text = ''
    }

function validateBoxes()
    {
    if($textBox1.Text -and $textBox2.Text -and $textBox3.Text -and $textBox4.Text -and $textBox5.Text -and $textBox6.Text -and $comboBox1.Text -and $comboBox2.Text)
    {
    $submitButton.Enabled = $true
    }
    else
    {
    $submitButton.Enabled=$false
    }
    }

$textBox1.Add_TextChanged({validateBoxes})
$textBox2.Add_TextChanged({validateBoxes})
$textBox3.Add_TextChanged({validateBoxes})
$textBox4.Add_TextChanged({validateBoxes})
$textBox5.Add_TextChanged({validateBoxes})
$textBox6.Add_TextChanged({validateBoxes})
$comboBox1.Add_TextChanged({validateBoxes})
$comboBox2.Add_TextChanged({validateBoxes})
$form1.ShowDialog()