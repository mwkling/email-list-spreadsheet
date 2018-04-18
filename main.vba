Option Explicit

Sub do_all()
  Call import_all_csvs
  
  Dim startingList As Dictionary
  Dim suppressionsList As Dictionary
  
  Set startingList = get_starting_list()
  Set suppressionsList = get_suppressions_list()
  
  Call write_output_file(startingList, suppressionsList)
End Sub

Sub write_output_file(start As Dictionary, suppress As Dictionary)
  Dim domainsSuppress As Dictionary
  Dim out As String
  Dim email As Variant
  Dim fn As Long
  Dim i As Long
  Dim domain As String
  
  Set domainsSuppress = get_excluded_domains()
  out = Range("input_directory").Value + Range("out_file_name").Value
  fn = FreeFile
  i = 0

  Open out For Append As #fn
    Print #fn, "email"
    For Each email In start.Keys()
      domain = Split(email, "@")(1)
      
      If Not suppress.Exists(email) And Not domainsSuppress.Exists(domain) Then
        Print #fn, email
        i = i + 1
      End If
    Next
  Close #fn

  Application.ActiveWorkbook.Worksheets(1).Activate
  MsgBox Str(start.Count) & " initial emails" & Chr(13) & Str(i) & " written to file"
End Sub

Function get_excluded_domains() As Dictionary
  Dim i As Long
  Dim domainsSuppress As New Dictionary
  
  i = 0
  
  Do Until Range("exclude_domains").Offset(i, 0) = ""
    domainsSuppress(Range("exclude_domains").Offset(i, 0).Value) = 1
    i = i + 1
  Loop
  
  Set get_excluded_domains = domainsSuppress
End Function

Function get_suppressions_list() As Dictionary
  Dim cur As Worksheet
  Dim c As Long
  Dim r As Long
  Dim emailColumn As Long
  Dim emails As New Dictionary
  
  For Each cur In Worksheets
    If cur.Name <> "main" And Left(cur.Name, 7) <> "truejob" Then
      c = 1
      r = 2
      If InStr(cur.Cells(1, 1), "@") > 0 Then
        emailColumn = 1
        r = 1
      Else
        Do Until cur.Cells(1, c) = ""
          If LCase(Left(cur.Cells(1, c), 5)) = "email" Then
            emailColumn = c
          End If
          c = c + 1
        Loop
      End If
      
      Do Until cur.Cells(r, emailColumn) = ""
        ' for our purposes treat emails as case insensitive
        emails(LCase(cur.Cells(r, emailColumn).Value)) = 1
        r = r + 1
      Loop
      
    End If
  Next
  
  Set get_suppressions_list = emails
End Function

Function get_starting_list() As Dictionary
  Dim cur As Worksheet
  Dim c As Long
  Dim r As Long
  Dim key As Variant
  Dim emailColumns As New Dictionary
  Dim emails As New Dictionary
  
  For Each cur In Worksheets
    If Left(cur.Name, 7) = "truejob" Then
      c = 1
      Do Until cur.Cells(1, c) = ""
        If LCase(Left(cur.Cells(1, c), 5)) = "email" Then
          emailColumns.Add c, 1
        End If
        c = c + 1
      Loop
      r = 2
      Do Until cur.Cells(r, 1) = ""
        For Each key In emailColumns.Keys()
          If cur.Cells(r, key) <> "" Then
            emails.Add LCase(cur.Cells(r, key).Value), 1
            Exit For
          End If
        Next
        r = r + 1
      Loop
    End If
  Next
  
  Set get_starting_list = emails
End Function

Sub import_all_csvs()
  Dim s As String
  Dim totalSheets As Long
  Dim thisBook As String
  Dim baseDir As String
  
  baseDir = Range("input_directory").Value
  thisBook = Application.ActiveWorkbook.Name
  
  s = Dir(baseDir + "*.csv")
  
  Do Until s = ""
    If s <> Range("out_file_name").Value Then
      Workbooks.Open (baseDir + s)
      totalSheets = Workbooks(thisBook).Worksheets.Count
      Workbooks(s).Worksheets(1).Copy after:=Workbooks(thisBook).Worksheets(totalSheets)
      Workbooks(s).Close
    End If

    s = Dir
  Loop
End Sub

Sub reset_book()
  Application.DisplayAlerts = False
  Do Until Application.ActiveWorkbook.Worksheets.Count = 1
    Application.ActiveWorkbook.Worksheets(2).Delete
  Loop
  Application.DisplayAlerts = True
End Sub
