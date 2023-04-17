$olapp = New-Object -ComObject Outlook.Application
# Открываем диалоговое окно для выбора файлов
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Multiselect = $true
$openFileDialog.ShowDialog() | Out-Null
# Запрос адресов получателей
$recipientsForm = New-Object System.Windows.Forms.Form
$recipientsForm.Text = "Введите адреса получателей"
$recipientsForm.Width = 300
$recipientsForm.Height = 150
$recipientsForm.FormBorderStyle = "FixedDialog"
$recipientsForm.StartPosition = "CenterScreen"
$recipientsForm.TopMost = $true
$recipientsLabel = New-Object System.Windows.Forms.Label
$recipientsLabel.Location = New-Object System.Drawing.Point(10, 20)
$recipientsLabel.Size = New-Object System.Drawing.Size(260, 20)
$recipientsLabel.Text = "Введите адреса получателей через запятую:"

$recipientsForm.Controls.Add($recipientsLabel)
$recipientsBox = New-Object System.Windows.Forms.TextBox
$recipientsBox.Location = New-Object System.Drawing.Point(10, 40)
$recipientsBox.Size = New-Object System.Drawing.Size(260, 20)
$recipientsForm.Controls.Add($recipientsBox)
$recipientsOkButton = New-Object System.Windows.Forms.Button
$recipientsOkButton.Location = New-Object System.Drawing.Point(10, 70)
$recipientsOkButton.Size = New-Object System.Drawing.Size(75, 23)
$recipientsOkButton.Text = "OK"
$recipientsOkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$recipientsForm.Controls.Add($recipientsOkButton)
$recipientsForm.AcceptButton = $recipientsOkButton
if ($recipientsForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $recipients = $recipientsBox.Text.Split(',')
}
# Запрос темы письма
$subjectForm = New-Object System.Windows.Forms.Form
$subjectForm.Text = "Введите тему письма"
$subjectForm.Width = 250
$subjectForm.Height = 150
$subjectForm.FormBorderStyle = "FixedDialog"
$subjectForm.StartPosition = "CenterScreen"
$subjectForm.TopMost = $true
$subjectLabel = New-Object System.Windows.Forms.Label
$subjectLabel.Location = New-Object System.Drawing.Point(10, 20)
$subjectLabel.Size = New-Object System.Drawing.Size(220, 20)
$subjectLabel.Text = "Введите текст для темы"
$subjectForm.Controls.Add($subjectLabel)
$subjectBox = New-Object System.Windows.Forms.TextBox
$subjectBox.Location = New-Object System.Drawing.Point(10, 40)
$subjectBox.Size = New-Object System.Drawing.Size(220, 20)
$subjectForm.Controls.Add($subjectBox)
$dateOkButton = New-Object System.Windows.Forms.Button
$dateOkButton.Location = New-Object System.Drawing.Point(10, 70)
$dateOkButton.Size = New-Object System.Drawing.Size(75, 23)
$dateOkButton.Text = "OK"
$dateOkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$subjectForm.Controls.Add($dateOkButton)
$subjectForm.AcceptButton = $dateOkButton
if ($subjectForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $subjectText = $subjectBox.Text
}
# Запрос даты
<# Add-Type -AssemblyName System.Windows.Forms
$dateForm = New-Object System.Windows.Forms.Form
$dateForm.Text = "Выберите дату"
$dateForm.Width = 250
$dateForm.Height = 200
$dateForm.FormBorderStyle = "FixedDialog"
$dateForm.StartPosition = "CenterScreen"
$dateForm.TopMost = $true
$dateOkButton = New-Object System.Windows.Forms.Button
$dateOkButton.Location = New-Object System.Drawing.Point(10, 150)
$dateOkButton.Size = New-Object System.Drawing.Size(75, 23)
$dateOkButton.Text = "OK"
$dateOkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$dateForm.Controls.Add($dateOkButton)
$dateForm.AcceptButton = $dateOkButton
$dateCancelButton = New-Object System.Windows.Forms.Button
$dateCancelButton.Location = New-Object System.Drawing.Point(90, 150)
$dateCancelButton.Size = New-Object System.Drawing.Size(75, 23)
$dateCancelButton.Text = "Cancel"
$dateCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$dateForm.Controls.Add($dateCancelButton)
$dateForm.CancelButton = $dateCancelButton

$monthCalendar = New-Object System.Windows.Forms.MonthCalendar
$monthCalendar.Location = New-Object System.Drawing.Point(10, 10)

$dateForm.Controls.Add($monthCalendar)
$result = $dateForm.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $date = $monthCalendar.SelectionStart.ToShortDateString()
}
elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
    [System.Windows.Forms.MessageBox]::Show("Отправления отменены")
    exit
} #>

# Отправка каждого письма
$sentCount = 0
foreach ($file in $openFileDialog.FileNames) {
    $mail = $olapp.CreateItem(0)
    # Добавляем выбранный файл во вложения
    $attachments = $mail.Attachments
    $attachments.Add($file, 1)
    # Добавляем сохраненные данные в письмо
    foreach ($recipient in $recipients) {
        $mail.Recipients.Add($recipient.Trim())
    }
    $name = [System.IO.Path]::GetFileNameWithoutExtension($file)
    $mail.Subject = "Отправка: " + $name + " " + $subjectText + " " + $date
    $mail.Body = $name
    $mail.Send()
    $sentCount++
}
# Окно с информацией о количестве отправленных сообщений
$infoForm = New-Object System.Windows.Forms.Form
$infoForm.Text = "Результат"
$infoForm.Width = 250
$infoForm.Height = 100
$infoForm.FormBorderStyle = "FixedDialog"
$infoForm.StartPosition = "CenterScreen"
$infoForm.TopMost = $true
$infoLabel = New-Object System.Windows.Forms.Label
$infoLabel.Location = New-Object System.Drawing.Point(10, 20)
$infoLabel.Size = New-Object System.Drawing.Size(220, 20)
$infoLabel.Text = "Отправлено сообщений: $sentCount"
$infoForm.Controls.Add($infoLabel)
$infoOkButton = New-Object System.Windows.Forms.Button
$infoOkButton.Location = New-Object System.Drawing.Point(10, 50)
$infoOkButton.Size = New-Object System.Drawing.Size(75, 23)
$infoOkButton.Text = "OK"
$infoForm.Controls.Add($infoOkButton)
$infoOkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$infoForm.AcceptButton = $infoOkButton
$infoForm.ShowDialog() | Out-Null