Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$olapp = New-Object -ComObject Outlook.Application
# Открываем диалоговое окно для выбора файлов
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Multiselect = $true
$openFileDialog.ShowDialog() | Out-Null
# Создаем форму
$form = New-Object System.Windows.Forms.Form
$form.TopMost = $true
$form.Text = 'Глобальная Гипер Отправка'
$form.Size = New-Object System.Drawing.Size(350,420)

# Создаем поле ввода темы письма
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Кусок Темы:'
$form.Controls.Add($label)

$subjectTextBox = New-Object System.Windows.Forms.TextBox
$subjectTextBox.Location = New-Object System.Drawing.Point(10,40)
$subjectTextBox.Size = New-Object System.Drawing.Size(280,20)
$form.Controls.Add($subjectTextBox)

# Создаем поле ввода адресов получателей
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,80)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Адресаты:'
$form.Controls.Add($label)

$recipientsTextBox = New-Object System.Windows.Forms.TextBox
$recipientsTextBox.Location = New-Object System.Drawing.Point(10,100)
$recipientsTextBox.Size = New-Object System.Drawing.Size(280,50)
#$recipientsTextBox.Multiline = $true
$form.Controls.Add($recipientsTextBox)

# Создаем календарь для выбора даты
$calendarLabel = New-Object System.Windows.Forms.Label
$calendarLabel.Location = New-Object System.Drawing.Point(10,130)
$calendarLabel.Size = New-Object System.Drawing.Size(200,20)
$calendarLabel.Text = 'Дата для темы:'
$form.Controls.Add($calendarLabel)

$calendar = New-Object System.Windows.Forms.MonthCalendar
$calendar.Location = New-Object System.Drawing.Point(10,150)
$calendar.MaxSelectionCount = 1
$form.Controls.Add($calendar)

# Создаем кнопку продолжения и отмены
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(30, 320)
$okButton.Size = New-Object System.Drawing.Size(80,23)
$okButton.Text = 'Продолжить'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(130, 320)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Палундра'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

# Отображаем форму и обрабатываем результат 
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    # Получаем данные из формы
    $subject = $subjectTextBox.Text
    $recipients = $recipientsTextBox.Text.Split(',')
    $date = $calendar.SelectionStart.ToShortDateString()

    # Ваш код для отправки письма
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
        $mail.Subject = "Отправка: " + $name + " " + $subject + " " + $date
        $mail.Body = $name
        $mail.Send()
        $sentCount++
    }
    # Окно с информацией о количестве отправленных сообщений
    $infoForm = New-Object System.Windows.Forms.Form
    $infoForm.Text = "Результат"
    $infoForm.Width = 250
    $infoForm.Height = 150
    $infoForm.FormBorderStyle = "FixedDialog"
    $infoForm.StartPosition = "CenterScreen"
    $infoForm.TopMost = $true
    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Location = New-Object System.Drawing.Point(10, 20)
    $infoLabel.Size = New-Object System.Drawing.Size(190, 20)
    $infoLabel.Text = "Отправлено: $sentCount"
    $infoForm.Controls.Add($infoLabel)
    $infoOkButton = New-Object System.Windows.Forms.Button
    $infoOkButton.Location = New-Object System.Drawing.Point(10, 50)
    $infoOkButton.Size = New-Object System.Drawing.Size(75, 23)
    $infoOkButton.Text = "OK"
    $infoForm.Controls.Add($infoOkButton)
    $infoOkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $infoForm.AcceptButton = $infoOkButton
    $infoForm.ShowDialog() | Out-Null
} else {
    [System.Windows.Forms.MessageBox]::Show("Протокол отмены Бззз Бззз")
    exit
}
