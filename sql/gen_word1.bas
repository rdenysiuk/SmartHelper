Attribute VB_Name = "gen_word1"
Option Explicit

Public Sub gen_WordDoc()
Dim appWord As Word.Application, _
    docWord As Word.Document, _
    rngCurrent As Word.Range, _
    objTable As Word.Table, _
    s As String, _
    iTBL_Rows As Integer ', _
    myfrmVP As frmVP

   
On Error GoTo Err_AccessToWord
    Set appWord = New Word.Application
    
    '�������� ��������
        Set docWord = appWord.Documents.Add()
        With docWord.PageSetup
            .TopMargin = CentimetersToPoints(10)
            .LeftMargin = CentimetersToPoints(2.5)
            .BottomMargin = CentimetersToPoints(1.5)
        End With
        appWord.Visible = True
    '--------------------------------------------
    '��������� ����� - ����
With docWord.Sections(1)
    .Footers(wdHeaderFooterPrimary).Range.Text = Date
    .Footers(wdHeaderFooterPrimary).PageNumbers.Add
    .Footers(wdHeaderFooterPrimary).Range.Select
    appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    appWord.Selection.Font.Name = "Times New Roman"
    appWord.Selection.Font.Size = 8
End With
    

With docWord.ActiveWindow
    .ActivePane.Close
    .View = wdPrintView
End With
   
    '����� ���������� �����
    Set rngCurrent = docWord.Range
    With rngCurrent
        .Collapse Direction:=wdCollapseEnd
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Text = "���������� �����"
        .Select
        .Font.Name = "Times New Roman"
        .Font.Size = 16
        .Font.Bold = True
        .InsertParagraphAfter
    End With
        
   
Set rngCurrent = docWord.Range
With rngCurrent
    .InsertParagraphAfter
    .Collapse Direction:=wdCollapseEnd
End With
        
'����� ��������
iTBL_Rows = 1
Set objTable = docWord.Tables.Add(Range:=rngCurrent, NumRows:=iTBL_Rows, NumColumns:=2)
        
With objTable
    .Borders.Enable = False '�� ������ �������
    .Rows.Height = 10
    .Columns.Width = 250
    .Cell(1, 1).Range.Text = "� " & frmVP.Text1 & " �� " & frmVP.dataF    '����� ���������
    ' ������� "����"
    Select Case frmVP.txtOblFrom.Text
        Case "68"
        .Cell(1, 2).Range.Text = "��������� ��������� ����� � ����������� ������"
        Case "22"
        .Cell(1, 2).Range.Text = "������� ��������� ��������� ����� ������ � " & frmVP.lblNameObl.Caption
    End Select
End With
    
    '������ �������� �� ������� ����, ����� - �� �����
    With objTable
        .Columns(2).Select
        appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        .Columns(1).Select
        appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    End With
        
    Set objTable = Nothing
    '-----------------------------
    
    '�� �������� �������� ���� �����
    Set rngCurrent = docWord.Range
    With rngCurrent
        .Collapse Direction:=wdCollapseEnd
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Text = "��� �������� ���������� ������� ������" & vbCrLf
        .Select
        .Font.Name = "Times New Roman"
        .Font.Size = 12
    End With
    '----------------------------------------
        
    '���� �����
    Set rngCurrent = docWord.Range
    With rngCurrent
        .InsertParagraphAfter
        .Collapse Direction:=wdCollapseEnd
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Text = "������� ��������� ��������� ����� ������ � ����������� ������ " _
                    & "������ ���������� ������� ������ ���������� ���� � ��'���� " _
                    & "� ����� ��������� ���� ����������:" & vbCrLf
        .Select
        .Font.Size = 14
        .ParagraphFormat.FirstLineIndent = CentimetersToPoints(1.25)
        .ParagraphFormat.Alignment = wdAlignParagraphJustify
    End With
    '----------------------------------------

    '�������� 2 �� 5
    Set rngCurrent = docWord.Range
    With rngCurrent
        .InsertParagraphAfter
        .Collapse Direction:=wdCollapseEnd
    End With
        
    Set objTable = docWord.Tables.Add(Range:=rngCurrent, NumRows:=iTBL_Rows + 1, NumColumns:=5)
    With objTable
        .Borders.Enable = True
        .Rows.Height = 10
        .Columns.Width = 60
        '������� �������
        .Cell(1, 1).Range.Text = "� �/�"
        .Cell(1, 2).Range.Text = "ϲ�"
        .Cell(1, 3).Range.Text = "�������, ����� �������"
        .Cell(1, 4).Range.Text = "����� ��������"
        .Cell(1, 5).Range.Text = "����� �����"
        .Cell(2, 1).Range.Text = "1"
        .Cell(2, 2).Range.Text = frmVP.txtPib 'ϲ�
        .Cell(2, 3).Range.Text = "����������� ���.," & Select_Raj_Hm(frmVP.txtRajFrom.Text) & _
                                vbCrLf & "(" & frmVP.txtOblFrom.Text & frmVP.txtRajFrom.Text & ")"
        .Cell(2, 4).Range.Text = "��������������� ���., ������������������ ����� (9999)"
        .Cell(2, 5).Range.Text = "68011296.1LS"
    End With

        objTable.Select
        appWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        appWord.Selection.Font.Size = 14
        appWord.Selection.MoveDown
       
        With objTable
        .AllowAutoFit = True
            .Columns(1).Width = 30
            .Columns(2).Width = 100
            .Columns(3).Width = 140
            .Columns(4).Width = 140
            .Columns(5).Width = 100
            .Columns.Select
            appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Rows(1).Select
            appWord.Selection.Font.Bold = True
            appWord.Selection.Font.Size = 13
            .Rows(2).Select
            appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        End With
        
        Set rngCurrent = docWord.Range
        With rngCurrent
            .InsertParagraphAfter
            .Collapse Direction:=wdCollapseEnd
        End With
                
        
        Set objTable = docWord.Tables.Add(Range:=rngCurrent, NumRows:=1, NumColumns:=2)
        With objTable
            .Borders.Enable = False
            .Rows.Height = 10
            .Columns.Width = 250
            .Cell(1, 1).Range.Text = "��������� ���������� ��������� � ����������� �� ������� ����� - " & _
                                    "��������� ����� � ������ ������� �����"
            .Cell(1, 2).Range.Text = "_______ �.�.�����"
            .Columns(2).Select
            appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            .Cell(1, 2).VerticalAlignment = wdCellAlignVerticalBottom
            .Columns(1).Width = 350
            .Columns(2).Width = 150
        End With
        
            objTable.Select
            appWord.Selection.Font.Size = 14
            appWord.Selection.MoveDown
        
        
        Set rngCurrent = docWord.Range
        
        With rngCurrent
            .InsertParagraphAfter
            .Collapse Direction:=wdCollapseEnd
        End With

        Set rngCurrent = docWord.Range
        
        With rngCurrent
            .InsertParagraphBefore
            .Collapse Direction:=wdCollapseEnd
            .ParagraphFormat.Alignment = wdAlignParagraphJustify
            .Text = "������� 75 20 62"
            .Select
            .Font.Name = "Times New Roman"
            .Font.Size = 14
         End With

        Set objTable = Nothing

        docWord.SaveAs "c:\send\out\6800test.doc"
    Set appWord = Nothing

L_Exit:

    Exit Sub

Err_AccessToWord:
'    AppActivate "Microsoft Access"
    Beep
    MsgBox "The Following Automation Error has occurred:" _
        & vbCrLf & Err.Description, vbCritical, "Automation Error!"
    Exit Sub
End Sub


