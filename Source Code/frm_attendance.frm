VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_attendance 
   Caption         =   "Smart Attendance Manager"
   ClientHeight    =   6960
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10860
   OleObjectBlob   =   "frm_attendance.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub btn_Asc_Click()

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("AttendanceDisplay")

        '''''' Sort Data
    ThisWorkbook.Sheets("Attendance Manager").Range("A2").Value = 1
    sh.UsedRange.Sort key1:=sh.Cells(1, Application.WorksheetFunction.Match(Me.cmb_Order_By.Value, sh.Range("1:1"), 0)), order1:=xlAscending, Header:=xlYes
     
End Sub

Private Sub btn_Desc_Click()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("AttendanceDisplay")
        '''''' Sort Data
    ThisWorkbook.Sheets("Attendance Manager").Range("A2").Value = 2
    sh.UsedRange.Sort key1:=sh.Cells(1, Application.WorksheetFunction.Match(Me.cmb_Order_By.Value, sh.Range("1:1"), 0)), order1:=xlDescending, Header:=xlYes
     
End Sub

 

Private Sub btn_Save_Click()
ThisWorkbook.Save
MsgBox "File saved successfuly"
End Sub

Private Sub cmdDeleteAC_Click()
    Dim wsAC As Worksheet
    Dim wsAttend As Worksheet
    Dim strAC As String
    Dim lngLastRow As Long
    Dim lngRow As Long
    '''''''''''''''''' Validation ''''''
    
    If Me.txt_AC_code.Value = "" Then
        MsgBox "Double click on any attendance code which you want to delete", vbCritical
        Exit Sub
    End If
    
    ' Delete the record from the AttendanceDisplay sheet
    Set wsAC = Sheets("AttendanceCodes")
    With wsAC
        .Activate
        lngLastRow = .Range("A1048576").End(xlUp).Row
        For lngRow = 2 To lngLastRow
            If .Cells(lngRow, "A") = Me.ListBox3.List(Me.ListBox3.ListIndex, 0) Then
                strAC = Me.ListBox3.List(Me.ListBox3.ListIndex, 0)
                .Cells(lngRow, "A").EntireRow.Delete
                Exit For
            End If
        Next
    End With
    
    ''''' Clear boxes
    Me.txt_AC_code.Value = ""
    Me.cmb_AC_attendance_Type.Value = ""
    Me.txt_AC_Remarks.Value = ""
 
    MsgBox "Attendance Code '" & strAC & "' has been deleted", vbInformation
    Call Attendance_Display_Listbox

End Sub

Private Sub cmb_EmployeeId_Change()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("EMPMaster")
    Dim emp_id As Variant
    
    If Me.cmb_EmployeeId.Value <> "" Then
        'On Error Resume Next
        If IsNumeric(Me.cmb_EmployeeId.Value) Then
            emp_id = CLng(Me.cmb_EmployeeId.Value)
        Else
            emp_id = Me.cmb_EmployeeId.Value
        End If
        
        
        Me.txt_EmpName.Value = Application.WorksheetFunction.VLookup(emp_id, sh.Range("A:C"), 2, 0)
        Me.txtSupervisor.Value = Application.WorksheetFunction.VLookup(emp_id, sh.Range("A:C"), 3, 0)
    End If
End Sub

Private Sub cmb_Filter_by_Change()
    If Me.cmb_Filter_by.Value = "ALL" And Me.txt_Search.Value <> "" Then
        Me.txt_Search.Value = ""
        Call Attendance_Display_Listbox
    End If
    
End Sub


Private Sub CommandButton1_Click()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Attendance")
    
    '''''''''''''''''' Validation ''''''
    If Me.cmb_EmployeeId.Value = "" Then
        MsgBox "Please select an Employee Id", vbCritical
        Exit Sub
    End If
    
    If Me.cmb_Attenandce_code.Value = "" Then
        MsgBox "Please select the attendance", vbCritical
        Exit Sub
    End If
    
    ''''' Check duplicates ''''
    
    If Application.WorksheetFunction.CountIfs(sh.Range("B:B"), Me.cmb_EmployeeId.Value, sh.Range("E:E"), Me.txt_Date.Value) > 0 Then
        MsgBox "Attendance already marked", vbCritical
        Exit Sub
    End If
     
   
    Dim lr As Long
    lr = Application.WorksheetFunction.CountA(sh.Range("A:A"))
    
    '''''' Add to worksheet
    sh.Range("A" & lr + 1).Value = "=ROW()-1"
    sh.Range("B" & lr + 1).Value = Me.cmb_EmployeeId.Value
    sh.Range("C" & lr + 1).Value = Me.txt_EmpName.Value
    sh.Range("D" & lr + 1).Value = Me.txtSupervisor.Value
    sh.Range("E" & lr + 1).Value = Me.txt_Date.Value
    sh.Range("F" & lr + 1).Value = Me.cmb_Attenandce_code.Value
    sh.Range("G" & lr + 1).Value = Now
       
    
    ''''' Clear boxes
    Me.cmb_EmployeeId.Value = ""
    Me.txt_EmpName.Value = ""
    Me.txtSupervisor.Value = ""
    Me.cmb_Attenandce_code.Value = ""
    
 
    MsgBox "Attendance has been Marked", vbInformation
    Call Attendance_Display_Listbox
 
      
    
End Sub

  
Private Sub CommandButton11_Click()
    
    If Me.ListBox2.ListIndex < 0 Then
        MsgBox "Please select the employees to mark attendance", vbCritical
        Exit Sub
    End If
    
    If Me.cmd_EM_Attendance_Code.Value = "" Then
        MsgBox "Please select the attendance", vbCritical
        Exit Sub
    End If
    
    Dim i As Integer
    Dim lr As Long
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Attendance")
    
    Dim Selected_Employee As Integer
    Dim Attendance_Marked As Integer
    
    Selected_Employee = 0
    Attendance_Marked = 0
    
    For i = 0 To Me.ListBox2.ListCount - 1
        If Me.ListBox2.Selected(i) = True Then
            Selected_Employee = Selected_Employee + 1
        
            ''''' Check duplicates ''''
            If Application.WorksheetFunction.CountIfs(sh.Range("B:B"), Me.ListBox2.List(i, 0), sh.Range("E:E"), Me.txt_EM_Date.Value) = 0 Then
                 lr = Application.WorksheetFunction.CountA(sh.Range("A:A"))
                    '''''' Add to worksheet
                    sh.Range("A" & lr + 1).Value = "=ROW()-1"
                    sh.Range("B" & lr + 1).Value = Me.ListBox2.List(i, 0)
                    sh.Range("C" & lr + 1).Value = Me.ListBox2.List(i, 1)
                    sh.Range("D" & lr + 1).Value = Me.ListBox2.List(i, 2)
                    sh.Range("E" & lr + 1).Value = Me.txt_EM_Date.Value
                    sh.Range("F" & lr + 1).Value = Me.cmd_EM_Attendance_Code.Value
                    sh.Range("G" & lr + 1).Value = Now
                    
                    Attendance_Marked = Attendance_Marked + 1
            End If
        End If
    Next i
    
    
    ''''' Clear boxes
    Me.cmd_EM_Attendance_Code.Value = ""
    
    MsgBox Attendance_Marked & " out of " & Selected_Employee & " attendance marked.", vbInformation
    Call Attendance_Display_Listbox
    
    
End Sub
 
Private Sub CommandButton2_Click()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Attendance")
    
    '''''''''''''''''' Validation ''''''
    
    If Me.txt_attendance_id.Value = "" Then
        MsgBox "Double click on any record which you want to update", vbCritical
        Exit Sub
    End If
    
    If Me.cmb_EmployeeId.Value = "" Then
        MsgBox "Please select an Employee Id", vbCritical
        Exit Sub
    End If
    
    If Me.cmb_Attenandce_code.Value = "" Then
        MsgBox "Please select the attendance", vbCritical
        Exit Sub
    End If
     
   
    Dim lr As Long
    lr = Application.WorksheetFunction.Match(CLng(Me.txt_attendance_id.Value), sh.Range("A:A"), 0)
    
    '''''' Add to worksheet
    sh.Range("A" & lr).Value = "=ROW()-1"
    sh.Range("B" & lr).Value = Me.cmb_EmployeeId.Value
    sh.Range("C" & lr).Value = Me.txt_EmpName.Value
    sh.Range("D" & lr).Value = Me.txtSupervisor.Value
    sh.Range("E" & lr).Value = Me.txt_Date.Value
    sh.Range("F" & lr).Value = Me.cmb_Attenandce_code.Value
    sh.Range("G" & lr).Value = Now
       
    
    ''''' Clear boxes
    Me.cmb_EmployeeId.Value = ""
    Me.txt_EmpName.Value = ""
    Me.txtSupervisor.Value = ""
    Me.cmb_Attenandce_code.Value = ""
    Me.txt_attendance_id.Value = ""
 
    MsgBox "Attendance has been updated", vbInformation
    Call Attendance_Display_Listbox
End Sub

Private Sub CommandButton3_Click()
    Dim nwb As Workbook
    Set nwb = Workbooks.Add
    
    ThisWorkbook.Sheets("AttendanceDisplay").UsedRange.Copy
    nwb.Sheets(1).Paste
    
End Sub

Private Sub CommandButton5_Click()
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("EMPMaster")
     
    ''''''' Validation ''''''
    
    If Me.txt_EM_id.Value = "" Then
        MsgBox "Please enter the Employee Id", vbCritical
        Exit Sub
    End If
    
    If Me.txt_EM_Name.Value = "" Then
        MsgBox "Please enter the Employee Name", vbCritical
        Exit Sub
    End If
    
    If Me.txt_EM_Supervisor.Value = "" Then
        MsgBox "Please enter the Supervisor Name", vbCritical
        Exit Sub
    End If
    
    '''''' Check Employee exits or not
    If Application.WorksheetFunction.CountIf(sh.Range("A:A"), txt_EM_id) = 0 Then
        MsgBox "This employee id does not exits to update", vbCritical
        Exit Sub
    End If


    Dim emp_id As Variant
    If IsNumeric(Me.txt_EM_id.Value) Then
        emp_id = CLng(Me.txt_EM_id.Value)
    Else
        emp_id = Me.txt_EM_id.Value
    End If
        
      
    Dim lr As Long
    lr = Application.WorksheetFunction.Match(emp_id, sh.Range("A:A"), 0)
     
    
    '''''' Add to worksheet
    sh.Range("A" & lr).Value = Me.txt_EM_id.Value
    sh.Range("B" & lr).Value = Me.txt_EM_Name.Value
    sh.Range("C" & lr).Value = Me.txt_EM_Supervisor.Value
    
    ''''' Clear boxes
    Me.txt_EM_id.Value = ""
    Me.txt_EM_Name.Value = ""
    Me.txt_EM_Supervisor.Value = ""
 
    MsgBox "Employee has been updated", vbInformation
    
    Call Employee_Master_Listbox
    Call Refresh_Data_List
    
End Sub

Private Sub CommandButton6_Click()
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("EMPMaster")
    
    ''''''' Validation ''''''
    If Me.txt_EM_id.Value = "" Then
        MsgBox "Please enter the Employee Id", vbCritical
        Exit Sub
    End If
    
    If Me.txt_EM_Name.Value = "" Then
        MsgBox "Please enter the Employee Name", vbCritical
        Exit Sub
    End If
    
    If Me.txt_EM_Supervisor.Value = "" Then
        MsgBox "Please enter the Supervisor Name", vbCritical
        Exit Sub
    End If
    
    ''''' Check duplicates ''''
    
    If Application.WorksheetFunction.CountIf(sh.Range("A:A"), Me.txt_EM_id.Value) > 0 Then
        MsgBox "Employee Id already exits", vbCritical
        Exit Sub
    End If
    
    Dim lr As Long
    lr = Application.WorksheetFunction.CountA(sh.Range("A:A"))
    
    '''''' Add to worksheet
    sh.Range("A" & lr + 1).Value = Me.txt_EM_id.Value
    sh.Range("B" & lr + 1).Value = Me.txt_EM_Name.Value
    sh.Range("C" & lr + 1).Value = Me.txt_EM_Supervisor.Value
    
    ''''' Clear boxes
    Me.txt_EM_id.Value = ""
    Me.txt_EM_Name.Value = ""
    Me.txt_EM_Supervisor.Value = ""
 
    MsgBox "Employee has been added", vbInformation
    Call Employee_Master_Listbox
    Call Refresh_Data_List
    
End Sub

Private Sub CommandButton7_Click()
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("AttendanceCodes")
    
   ''''''' Validation ''''''
 
    If Me.txt_AC_code.Value = "" Then
        MsgBox "Please enter the Attendance Code", vbCritical
        Exit Sub
    End If
    
    If Me.txt_AC_Remarks.Value = "" Then
        MsgBox "Please enter the remarks", vbCritical
        Exit Sub
    End If
    
    If Me.cmb_AC_attendance_Type.Value = "" Then
        MsgBox "Please select the attendance type", vbCritical
        Exit Sub
    End If
    
       
    '''''' Check attendance code exits or not
    If Application.WorksheetFunction.CountIf(sh.Range("A:A"), txt_AC_code) = 0 Then
        MsgBox "This attendance code does not exits to update", vbCritical
        Exit Sub
    End If
  
      
    Dim lr As Long
    lr = Application.WorksheetFunction.Match(txt_AC_code, sh.Range("A:A"), 0)
    
 
    
    '''''' Add to worksheet
    sh.Range("A" & lr).Value = Me.txt_AC_code.Value
    sh.Range("B" & lr).Value = Me.cmb_AC_attendance_Type.Value
    sh.Range("C" & lr).Value = Me.txt_AC_Remarks.Value
    
    ''''' Clear boxes
    Me.txt_AC_code.Value = ""
    Me.txt_AC_Remarks.Value = ""
 
    MsgBox "Attendance Code has been updated", vbInformation
    
    Call Attendance_Codes_Listbox
    Call Refresh_Data_List
End Sub

Private Sub CommandButton8_Click()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("AttendanceCodes")
    
    ''''''' Validation ''''''
    If Me.txt_AC_code.Value = "" Then
        MsgBox "Please enter the Attendance Code", vbCritical
        Exit Sub
    End If
    
    If Me.txt_AC_Remarks.Value = "" Then
        MsgBox "Please enter the remarks", vbCritical
        Exit Sub
    End If
    
    If Me.cmb_AC_attendance_Type.Value = "" Then
        MsgBox "Please select the attendance type", vbCritical
        Exit Sub
    End If
    
    ''''' Check duplicates ''''
    
    If Application.WorksheetFunction.CountIf(sh.Range("A:A"), Me.txt_AC_code.Value) > 0 Then
        MsgBox "Attendance Code already exits", vbCritical
        Exit Sub
    End If
    
    Dim lr As Long
    lr = Application.WorksheetFunction.CountA(sh.Range("A:A"))
    
    '''''' Add to worksheet
    sh.Range("A" & lr + 1).Value = Me.txt_AC_code.Value
    sh.Range("B" & lr + 1).Value = Me.cmb_AC_attendance_Type.Value
    sh.Range("C" & lr + 1).Value = Me.txt_AC_Remarks.Value
    
    ''''' Clear boxes
    Me.txt_AC_code.Value = ""
    Me.txt_AC_Remarks.Value = ""
 
    MsgBox "Attendance has been added", vbInformation
    
    Call Attendance_Codes_Listbox
    Call Refresh_Data_List
End Sub

Private Sub Image3_Click()
    Call Calendar.SelectedDate(Me.txt_Date)
End Sub

Sub Refresh_Data_List()

    Dim sh_EM As Worksheet
    Dim sh_AC As Worksheet
    
    Set sh_EM = ThisWorkbook.Sheets("EMPMaster")
    Set sh_AC = ThisWorkbook.Sheets("AttendanceCodes")
     
    
    Me.cmb_Attenandce_code.Clear
    Me.cmb_EmployeeId.Clear
    Me.cmd_EM_Attendance_Code.Clear
    
    ''''' Add Employee Id in combobox
    Me.cmb_EmployeeId.AddItem ""
    
    For i = 2 To Application.WorksheetFunction.CountA(sh_EM.Range("A:A"))
        If i > 1 Then
            Me.cmb_EmployeeId.AddItem sh_EM.Range("A" & i).Value
        End If
    Next i
    
    ''''' Add attendance code in combobox
    Me.cmb_Attenandce_code.AddItem ""
    Me.cmd_EM_Attendance_Code.AddItem ""
    For i = 2 To Application.WorksheetFunction.CountA(sh_AC.Range("A:A"))
        If i > 1 Then
            Me.cmb_Attenandce_code.AddItem sh_AC.Range("A" & i).Value
            Me.cmd_EM_Attendance_Code.AddItem sh_AC.Range("A" & i).Value
        End If
    Next i

End Sub
   
Private Sub Image5_Click()
    Call Calendar.SelectedDate(Me.txt_EM_Date)
End Sub

 

Private Sub Label19_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txt_attendance_id.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
    Me.cmb_EmployeeId.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
    Me.txt_Date.Value = VBA.Format(Me.ListBox1.List(Me.ListBox1.ListIndex, 4), "D-MMM-YYYY")
    Me.cmb_Attenandce_code.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 5)
     
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txt_EM_id.Value = Me.ListBox2.List(Me.ListBox2.ListIndex, 0)
    Me.txt_EM_Name.Value = Me.ListBox2.List(Me.ListBox2.ListIndex, 1)
    Me.txt_EM_Supervisor.Value = Me.ListBox2.List(Me.ListBox2.ListIndex, 2)
    
End Sub

 
Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txt_AC_code.Value = Me.ListBox3.List(Me.ListBox3.ListIndex, 0)
    Me.cmb_AC_attendance_Type.Value = Me.ListBox3.List(Me.ListBox3.ListIndex, 1)
    Me.txt_AC_Remarks.Value = Me.ListBox3.List(Me.ListBox3.ListIndex, 2)
End Sub
   
 


Private Sub cmdDeleteEmp_Click()
    Dim sh As Worksheet
    Dim wsEmpData As Worksheet
    Set sh = ThisWorkbook.Sheets("Attendance")
    Dim rngFound As Range
    Dim strEmpName As String
    
    '''''''''''''''''' Validation ''''''
    
    If Me.txt_attendance_id.Value = "" Then
        MsgBox "Double click on any record which you want to delete", vbCritical
        Exit Sub
    End If
    
    If Me.cmb_EmployeeId.Value = "" Then
        MsgBox "Please select an Employee Id", vbCritical
        Exit Sub
    End If
   
    Dim lr As Long
    lr = Application.WorksheetFunction.Match(CLng(Me.txt_attendance_id.Value), sh.Range("A:A"), 0)
    
    '''''' Delete employee from all worksheets
    If vbYes = MsgBox("Are you sure you want to delete employee '" & Me.txt_EmpName.Value _
                   & "' ?", vbYesNo + vbQuestion, "Deletion Confirmation") Then
        For Each wsEmpData In ThisWorkbook.Worksheets
            With wsEmpData
                Select Case .Name
                   Case "EMPMaster", "Attendance", "AttendanceDisplay"
                        ThisWorkbook.Sheets(.Name).Activate
                        ActiveSheet.Rows(1).Hidden = True

                        Set rngFound = .Cells.Find(What:=Me.txt_EmpName.Value, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
                            :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                            False, SearchFormat:=False)
                        FilterAndDeleteEmp Range(rngFound.Address)
                        ActiveSheet.Rows(1).Hidden = False
                End Select
           End With
       Next
       
    Else
        Me.cmb_EmployeeId.Value = ""
        Me.txt_EmpName.Value = ""
        Me.txtSupervisor.Value = ""
        Me.cmb_Attenandce_code.Value = ""
        Me.txt_attendance_id.Value = ""
        Exit Sub
    End If
    
    ''''' Clear boxes
    Me.cmb_EmployeeId.Value = ""
    strEmpName = Me.txt_EmpName.Value
    Me.txt_EmpName.Value = ""
    Me.txtSupervisor.Value = ""
    Me.cmb_Attenandce_code.Value = ""
    Me.txt_attendance_id.Value = ""
 
    MsgBox "Employee '" & strEmpName & "' has been deleted", vbInformation
    Call Attendance_Display_Listbox
    
End Sub


Private Sub cmdDeleteAttend_Click()
    Dim sh As Worksheet
    Dim wsAttendDisp As Worksheet
    Dim wsAttend As Worksheet
    Dim strEmpName As String
    Dim strEmpID As String
    Dim strDate As String
    Dim lngLastRow As Long
    Dim lngRow As Long
    '''''''''''''''''' Validation ''''''
    
    If Me.txt_attendance_id.Value = "" Then
        MsgBox "Double click on any record which you want to delete", vbCritical
        Exit Sub
    End If
    
    If Me.cmb_EmployeeId.Value = "" Then
        MsgBox "Please select an Employee Id", vbCritical
        Exit Sub
    End If
       
    ' Delete the record from the AttendanceDisplay sheet
    Set wsAttendDisp = Sheets("AttendanceDisplay")
    With wsAttendDisp
        .Activate
        lngLastRow = .Range("A1048576").End(xlUp).Row
        For lngRow = 2 To lngLastRow
            ' Check employee id and date
            If .Cells(lngRow, "B") = Me.ListBox1.List(Me.ListBox1.ListIndex, 1) And _
               VBA.Format(.Cells(lngRow, "G"), "D-MMM-YYYY") = VBA.Format(Me.ListBox1.List(Me.ListBox1.ListIndex, 4), "D-MMM-YYYY") Then
                strEmpName = Me.txt_EmpName.Value
                strDate = VBA.Format(.Cells(lngRow, "G"), "D-MMM-YYYY")
                strEmpID = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
                .Cells(lngRow, "A").EntireRow.Delete
                Exit For
            End If
        Next
    End With
    
    ' The record on the Attendance sheet also needs to be deleted
    Set wsAttend = Sheets("Attendance")
    With wsAttend
        .Activate
        lngLastRow = .Range("A1048576").End(xlUp).Row
        For lngRow = 2 To lngLastRow
            ' Check employee id and date
            If .Cells(lngRow, "B") = strEmpID And _
               VBA.Format(.Cells(lngRow, "G"), "D-MMM-YYYY") = strDate Then
                .Cells(lngRow, "A").EntireRow.Delete
                Exit For
            End If
        Next
    End With

    ''''' Clear boxes
    Me.cmb_EmployeeId.Value = ""
    Me.txt_EmpName.Value = ""
    Me.txtSupervisor.Value = ""
    Me.cmb_Attenandce_code.Value = ""
    Me.txt_attendance_id.Value = ""
 
    MsgBox "Attendance record for '" & strEmpName & "' on " & strDate & " has been deleted", vbInformation
    Call Attendance_Display_Listbox
    
End Sub

Private Sub txt_Search_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        Call Attendance_Display_Listbox
    End If
End Sub

Private Sub UserForm_Initialize()
 
    
 ''''' add list for filter by
    With Me.cmb_Filter_by
        .Clear
        .AddItem "ALL"
        .AddItem "EMP ID"
        .AddItem "EMP Name"
        .AddItem "Supervisor"
        .AddItem "Date"
        .AddItem "Attendance"
        .Value = "ALL"
    End With

 ''''' add list for order by
    With Me.cmb_Order_By
        .Clear
        .AddItem "EMP ID"
        .AddItem "EMP Name"
        .AddItem "Supervisor"
        .AddItem "Date"
        .AddItem "Attendance"
        .AddItem "Update Time"
        .Value = "Update Time"
    End With
      
 ''''' add list for attendance type for Attendance code page
    With Me.cmb_AC_attendance_Type
        .Clear
        .AddItem "Present"
        .AddItem "Absent"
        .AddItem "NCNS"
        .AddItem "Planned"
        .AddItem "Unplanned"
        .AddItem "Halfday"
        .AddItem "Other"
    End With
    
    ''''' default today's date
    Me.txt_Date.Value = Format(Date, "D-MMM-YYYY")
    Me.txt_EM_Date.Value = Format(Date, "D-MMM-YYYY")
    '''' intial calls
    Call Refresh_Data_List
    Call Employee_Master_Listbox
    Call Attendance_Codes_Listbox
    Call Attendance_Display_Listbox
     
End Sub
 
Sub Employee_Master_Listbox()

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("EMPMaster")
    
    Dim last_row As Long
    last_row = Application.WorksheetFunction.CountA(sh.Range("A:A"))
    If last_row = 1 Then last_row = 2
    
    With Me.ListBox2
        .ColumnHeads = True
        .ColumnCount = 3
        .ColumnWidths = "70,120,100"
        .RowSource = sh.Name & "!A2:C" & last_row
    End With
 
End Sub

Sub Attendance_Codes_Listbox()

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("AttendanceCodes")
    
    Dim last_row As Long
    last_row = Application.WorksheetFunction.CountA(sh.Range("A:A"))
    If last_row = 1 Then last_row = 2
    
    With Me.ListBox3
        .ColumnHeads = True
        .ColumnCount = 3
        .ColumnWidths = "50,100,150"
        .RowSource = sh.Name & "!A2:C" & last_row
    End With
 
End Sub

Sub Attendance_Display_Listbox()

    Dim dsh As Worksheet
    Set dsh = ThisWorkbook.Sheets("Attendance")
     
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("AttendanceDisplay")
    
    ''''' Copy data
    
    '''''' Filter data
    dsh.AutoFilterMode = False
    If Me.cmb_Filter_by.Value <> "ALL" And Me.txt_Search.Value <> "" Then
        dsh.UsedRange.AutoFilter Application.WorksheetFunction.Match(Me.cmb_Filter_by.Value, dsh.Range("1:1"), 0), Me.txt_Search.Value
    End If
    
    sh.Cells.ClearContents
    dsh.UsedRange.Copy
    sh.Range("A1").PasteSpecial xlPasteValues
    sh.Range("A1").PasteSpecial xlPasteFormats
    
    dsh.AutoFilterMode = False
    
    '''''' Sort Data
    If ThisWorkbook.Sheets("Attendance Manager").Range("A2").Value = 1 Then
        sh.UsedRange.Sort key1:=sh.Cells(1, Application.WorksheetFunction.Match(Me.cmb_Order_By.Value, sh.Range("1:1"), 0)), order1:=xlAscending, Header:=xlYes
    Else
        sh.UsedRange.Sort key1:=sh.Cells(1, Application.WorksheetFunction.Match(Me.cmb_Order_By.Value, sh.Range("1:1"), 0)), order1:=xlDescending, Header:=xlYes
    End If
    
    Dim last_row As Long
    last_row = Application.WorksheetFunction.CountA(sh.Range("A:A"))
    If last_row = 1 Then last_row = 2
    
    With Me.ListBox1
        .ColumnHeads = True
        .ColumnCount = 7
        .ColumnWidths = "0,50,100,100,70,70,100"
        .RowSource = sh.Name & "!A2:G" & last_row
    End With
 
End Sub

