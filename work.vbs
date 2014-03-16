'Version 0.01-alfa of 16.03.14 - alex0rus

' Путь к Архиватору
const ZIP = "C:\Program Files\7-Zip\7z.exe"
' Общий лог
const AllLogFile   = "D:\Xasan\TZ\WORK\LOG\All.log"
' Рабочий лог
const WorkLogFile  = "D:\Xasan\TZ\WORK\LOG\Work.log"
' Лог ошибок
const ErrorLogFile = "D:\Xasan\TZ\WORK\LOG\Error.log"

Const ForReading   = 1   ' Read
Const ForWriting   = 2   ' Write
Const ForReWriting = 8   ' ReWrite

dim fso, AllLog, WorkLog, ErrorLog, CurrentFile
dim Folder, Files
dim i
dim WorkPath 'as string    ' Путь где идет разбор
dim AnswerPath 'as string  ' Путь ответов
dim SourcePath 'as string  ' Путь где надо проверить нахождение

WorkPath   = "D:\Xasan\TZ\WORK\WORK\"
AnswerPath = "D:\Xasan\TZ\Reports\ФСФМ\отправлено\"
SourcePath = "D:\Xasan\TZ\Reports\ФСФМ\отправлено\"

'==== ТЕЛО СКРИПТА =============================================================

Set fso = CreateObject("Scripting.FileSystemObject")

' Открываем логи или если нет создаем их
set AllLog  = fso.OpenTextFile(AllLogFile, ForReWriting, True)  
set WorkLog = fso.OpenTextFile(WorkLogFile, ForReWriting, True) 
set ErrLog  = fso.OpenTextFile(ErrorLogFile, ForReWriting, True) 

' Запишем что мы начали работать
AllLog.WriteLine "Запуск скрипта " & date & " " & time

' Первая итерация
Main_Circle(WorkPath)

' Запишем что мы закончили работать
AllLog.WriteLine "Финиш скрипта " & date & " " & time

AllLog.Close
WorkLog.Close
ErrLog.Close

'==== Main_Circle ==============================================================
sub Main_Circle(WorkPath)
' Обработка файлов в каталоге

  'Создаем объект Folder для для текущей папки
  set Folder=FSO.GetFolder(WorkPath)
 
  'Создаем коллекцию файлов
  set Files=Folder.Files
  
  ' Получим кол-во файлов
  k = (Files.Count)

  AllLog.WriteLine "Кол-во файлов в " & Folder & " - " & k

  ' начинаем перебор коллекции файлов
  For Each i In Files

    AllLog.WriteLine "Обработка файла " & i.Path  
    WorkLog.WriteLine "Обработка файла " & i.Name

    '
    if Check_ARJ(i.Path) then
      AllLog.WriteLine "Файл " & i.Path & " является архивом"
      WorkLog.WriteLine "Файл " & i.Name & " является архивом"  
    end if

  Next

end sub
'==== Check_ARJ ================================================================
function Check_ARJ(NameFile)
' Определяем Архив это или нет

  ' Получим первые два байта
  Set CurrentFile = fso.OpenTextFile(NameFile, ForReading)
  ss = CurrentFile.Read(2)
  CurrentFile.Close

  ' Если ARJ архив то первые два символа должны быть 60EA
  if hex(asc(left(ss, 1))) = "60" and hex(asc(right(ss, 1))) = "EA" Then
    Check_ARJ = True
  else
    Check_ARJ = False
  end if

end function
'====  ================================================================
