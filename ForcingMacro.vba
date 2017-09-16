'************************************************************************
'*-----------------------------------------------------------------------
'*  Name： ThisWorkbook (Object Module)
'*-----------------------------------------------------------------------
'*  Descriptioin：forcing macro sample
'*-----------------------------------------------------------------------
'*  Copyright: HAYs  http://dev-clips.com , 2017 All Rights Reserved.
'*-----------------------------------------------------------------------
'*  <Update>
'*  Date        Version     Author     Memo
'*-----------------------------------------------------------------------
'*  2017.08.01  1.00        HAYs       New Release
'************************************************************************
' option
Option Explicit

'************************************************************************
'*  constant
'************************************************************************
Private Const PWD As String = "1234567890"

'************************************************************************
'*  Workbook event
'************************************************************************
'*-----------------------------------------------------------------------
'*　open
'*-----------------------------------------------------------------------
Private Sub Workbook_Open()
    Dim ws As Worksheet
    
    'fault ScreenUpdating
    Application.ScreenUpdating = False
    
    'unprotect book
    ThisWorkbook.Unprotect PWD
    
    'show sheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next
    
    'hide sheet (if necessary)
    Sheet3.Visible = xlSheetHidden
    Sheet4.Visible = xlSheetVeryHidden
    
    'hide macrosecurity sheet
    wsMacroSecurity.Visible = xlSheetVeryHidden
    
    'change save status
    ThisWorkbook.Saved = True
    
    'reset ScreenUpdating
    Application.ScreenUpdating = True
    
End Sub

'*-----------------------------------------------------------------------
'*　after save
'*-----------------------------------------------------------------------
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Dim saveScreenUpdating As Boolean
    Dim saveEnableEvents As Boolean
    Dim ws As Worksheet
    Dim ix() As Long
    Dim cnt As Long
    Dim i As Long
    
'#If DEBUG_MODE = 1 Then
'    Exit Sub
'#End If

    'save current status
    saveScreenUpdating = Application.ScreenUpdating
    saveEnableEvents = Application.EnableEvents
    
    'change property
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    '
    If Success Then
        
        'show macrosecurity sheet
        wsMacroSecurity.Visible = xlSheetVisible
        
        'hide visible sheet and save index
        For Each ws In ThisWorkbook.Worksheets
            
            'visible sheet
            If ws.Visible = xlSheetVisible Then
                'ignore macrosecurity sheet
                If Not ws Is wsMacroSecurity Then
                
                    'hide sheet
                    ws.Visible = xlSheetVeryHidden
                    'save index
                    ReDim Preserve ix(cnt) As Long
                    ix(cnt) = ws.Index
                    'increment counter
                    cnt = cnt + 1
                
                End If
            End If
            
        Next
        
        '
        If cnt Then
            
            'protect book
            ThisWorkbook.Protect PWD
                
            'save book
            ThisWorkbook.Save
            
            'unprotect book
            ThisWorkbook.Unprotect PWD
            
            'show visible sheet
            For i = 0 To cnt - 1
                ThisWorkbook.Worksheets(ix(i)).Visible = xlSheetVisible
            Next
            
            'hide macrosecurity sheet
            wsMacroSecurity.Visible = xlSheetVeryHidden
            
        End If
        
        'change save status
        ThisWorkbook.Saved = True
        
    End If

    'nomral end
    GoTo Fin
ErrTrap:

    '// err handling
    
Fin:

    'reset
    Application.ScreenUpdating = saveScreenUpdating
    Application.EnableEvents = saveEnableEvents
    
End Sub
