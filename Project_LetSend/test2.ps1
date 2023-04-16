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
# Запрос даты
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
$monthCalendar = New-Object System.Windows.Forms.MonthCalendar
$monthCalendar.Location = New-Object System.Drawing.Point(10, 10)
$dateForm.Controls.Add($monthCalendar)
if ($dateForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $date = $monthCalendar.SelectionStart.ToShortDateString()
}
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
    $mail.Subject = "Отправка: " + $name + " ARHIVE SCANNING HCP FROM " + $date
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