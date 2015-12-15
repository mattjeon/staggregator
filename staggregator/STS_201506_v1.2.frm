VERSION 5.00
Begin VB.Form STS 
   Caption         =   "STS"
   ClientHeight    =   4905
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command6 
      Caption         =   "�׸�����"
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   9
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   240
      TabIndex        =   8
      Text            =   "localhost;DATABASE=stock;"
      Top             =   4080
      Width           =   4695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����(������ü)"
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   7
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "�ϰ���ġ"
      Height          =   375
      Index           =   0
      Left            =   5280
      TabIndex        =   6
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����/�׸��ڵ�"
      Height          =   495
      Index           =   0
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����(����)"
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����(�����繫)"
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "����(����,������ü)"
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      Height          =   3120
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "STS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public conn As New ADODB.Connection
 Public rs   As New ADODB.Recordset
 Public sql  As String
 Public cnt  As Integer
 
Sub DB_Open()
    
    conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" + _
                              "SERVER=" + _
                              Text1.Text + _
                              "UID=root;PWD=jini7476; OPTION=16427;" + _
                              "STMT= set names euckr"
    conn.ConnectionTimeout = 30
    conn.Mode = adModeReadWrite
    
    conn.Open
    If conn.State = adStateClosed Then
        MsgBox ("DB���� ����")
        conn.Close
    End If
End Sub

Sub DB_Close()
    conn.Close
    
    If conn.State = adStateClosed Then
    Else
        conn.Close
    End If
    
    Set conn = Nothing
    Set rs = Nothing
End Sub


Sub Log(logmessage As String)
    'Dim fileName As String
    'Dim fileNum As Integer
    
    'fileName = "C:\sts.log"
    'fileNum = FreeFile
    
    'Open fileName For Append As fileNum
    '    Print #fileNum, Date; Time; Spc(5); logmessage
    'Close #fileNum
    
    'Debug.Print "[" & Now & "]" & logmessage
    
    Label1.Caption = "[" & Now & "]  " & logmessage
End Sub



'----------------------------------------------
'        ��ġ Ÿ�̸�
'----------------------------------------------
Private Sub Timer1_Timer()
    Static bStart As Boolean
   
    'If Hour(Now) = 17 Then
    '    If bStart = False Then
    '        bStart = True
     '       Call Company_Code
     '       Call Company_1
    '        Call Company_2
     '       Call Industry_1
     '   End If
   ' Else
   '     bStart = False
   ' End If
End Sub

'----------------------------------------------
'        �⺻�ڵ����� ��������
'----------------------------------------------
Private Sub Command1_Click(Index As Integer)
           
    Select Case (Index)
        '----------------------------------------------
        '�� �����ڵ�, �� �������� ok
        '----------------------------------------------
        
        Case 0
            Call Industry_Code
            Call Company_Code
            Call Theme_Code
            
    End Select
     
    
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If
    
    
End Sub



Private Sub Command2_Click(Index As Integer)
    Select Case (Index)
        '----------------------------------------------
        '���� Fundamental �������� - �������� ����.
        '----------------------------------------------
        Case 0
            Call Company_f
         
    End Select
    
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If
End Sub

Private Sub Command3_Click(Index As Integer)
    Select Case (Index)
        '----------------------------------------------
        '�����ְ�  OK
        '���� ������ü OK
        '----------------------------------------------
        Case 0
                Call Company_p
        
    End Select
    
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If
End Sub



Private Sub Command4_Click(Index As Integer)
    
    Select Case (Index)
        '----------------------------------------------
        '�����ְ� (StockChart)
        '������ü��
        '----------------------------------------------
        Case 0
            Call Company_i
    End Select
    
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If
End Sub


Private Sub Command5_Click(Index As Integer)
    
    Select Case (Index)
        Case 0
            Call Industry_pi
    
    End Select
    
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If

End Sub


Private Sub Command6_Click(Index As Integer)
    
    Select Case (Index)
        Case 0
        
            Call Theme_company
    
    End Select
    
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If

End Sub

Private Sub Command7_Click(Index As Integer)
    
    Select Case (Index)
        Case 0
                Call Industry_Code
                Call Company_Code
                Call Theme_Code
                Call Company_f
                Call Company_p
                Call Company_i
                Call Industry_pi
                Call Theme_company
        
    End Select
    
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If

End Sub


Sub Company_p()
            Call Log("-- ���� �ְ���������")
            Call DB_Open
            
            Dim arrList(3000) 'As Variant
            Dim i
            Dim codes As String
            
            sql = "SELECT COMPANY_CODE AS COMPANY_CODE, COMPANY_NAME AS COMPANY_NAME  "
            sql = sql & "FROM COMPANY "
            sql = sql & "WHERE 1=1 "
            rs.CursorLocation = adUseClient
            rs.Open sql, conn, adOpenStatic, adLockOptimistic
            
            i = 0
            While rs.EOF = False
                arrList(i) = rs("COMPANY_CODE")
                rs.MoveNext
                i = i + 1
            Wend
            
            rs.Close
            Set rs = Nothing
            Call DB_Close
            
            For i = 0 To UBound(arrList)
               codes = Trim(CStr(arrList(i)))
               If codes = "" Then
               Else
                Call Company_Price(codes)
               End If
            Next
            Call Log("-- ���� �ְ���������")
End Sub

Sub Company_i()
            Call Log("-- ���� �����ڼ�������")
            Call DB_Open
            
            Dim arrList(3000) 'As Variant
            Dim i
            Dim codes As String
            
            sql = "SELECT COMPANY_CODE AS COMPANY_CODE, COMPANY_NAME AS COMPANY_NAME  "
            sql = sql & "FROM COMPANY "
            sql = sql & "WHERE 1=1 "
            rs.CursorLocation = adUseClient
            rs.Open sql, conn, adOpenStatic, adLockOptimistic
            
            i = 0
            While rs.EOF = False
                arrList(i) = rs("COMPANY_CODE")
                rs.MoveNext
                i = i + 1
            Wend
            
            rs.Close
            Set rs = Nothing
            Call DB_Close
            
            For i = 0 To UBound(arrList)
               codes = Trim(CStr(arrList(i)))
               If codes = "" Then
               Else
                Call Company_Investors(codes)
               End If
            Next
            Call Log("-- ���� �����ڼ�������")
End Sub


Sub Company_f()
            Call Log("-- ���� �繫��������")
            Call DB_Open
            
            Dim arrList(3000) 'As Variant
            Dim i
            Dim codes As String
            
            sql = "SELECT COMPANY_CODE AS COMPANY_CODE, COMPANY_NAME AS COMPANY_NAME  "
            sql = sql & "FROM COMPANY "
            sql = sql & "WHERE 1=1 "
            rs.CursorLocation = adUseClient
            rs.Open sql, conn, adOpenStatic, adLockOptimistic
            
            i = 0
            While rs.EOF = False
                arrList(i) = rs("COMPANY_CODE")
                rs.MoveNext
                i = i + 1
            Wend
            
            rs.Close
            Set rs = Nothing
            Call DB_Close
            
            For i = 0 To UBound(arrList)
               codes = Trim(CStr(arrList(i)))
               If codes = "" Then
               Else
                    Call Company_Fundamental(codes)
               End If
            Next
            Call Log("-- ���� �繫��������")
End Sub



Sub Company_Price(codes As String)
            ' ���� �ְ�(StockWeek) default 10��ġ --------------------------------------------------------------------------------------------------------
            Dim codemgr    As New CpCodeMgr    '�ڵ�
            Dim stprice    As New StockWeek    '�ְ� ��¥ �˻��ȵ�
            Dim tempdate As String
            Dim tempdate2 As String
            Dim s, n  As String
            
            Call Log(codes & "[" & codemgr.CodeToName(codes) & "] �ְ� ������")
            
            Call DB_Open
    
            '�ֱ����� ��ȸ
            sql = "SELECT MAX(REG_DATE) AS REG_DATE FROM COMPANY_PRICE WHERE COMPANY_CODE = '" & codes & "' "
            rs.CursorLocation = adUseClient
            rs.Open sql, conn, adOpenStatic, adLockOptimistic
            
            If rs.EOF Or rs.BOF Then
                Debug.Print "DB�� ��ϵ��� �ʾҽ��ϴ�.[EOF]. ����� �����մϴ�. "
                tempdate = "19000101"
            Else
                If IsNull(rs("REG_DATE")) Then
                   Debug.Print "DB�� ��ϵ��� �ʾҽ��ϴ�.[NULL]. ����� �����մϴ�. "
                   tempdate = "19000101"
                Else
                    tempdate = rs("REG_DATE")
                    Debug.Print "�ֱ� DB ������ڴ� " & tempdate & "�Դϴ�. "
                End If
            End If
                
            rs.Close
            Set rs = Nothing
                
            List1.Clear
                          
            stprice.SetInputValue 0, codes '�����ڵ�
            stprice.BlockRequest
            bFind = False
                           
            s = "�ڵ�  ����  �ð�  ��   ����  ����  ���ϴ�� �ŷ���(����) �ŷ����  �����  ���κ��� ���κ��� ������ż���"
            List1.AddItem (s)
           
            n = stprice.GetHeaderValue(1)  '���Ű���
            cnt = 0
            
            For j = 0 To n - 1                  ' ���� ������ ����ŭ ������ ��
                    
                tempdate2 = stprice.GetDataValue(0, j)
                If tempdate >= tempdate2 Then 'DB���� ���ŵ����Ͱ� Ŀ����
                    Exit For
                ElseIf tempdate2 = 0 Then  '0���� �����°� ���͸�
                    Exit For
                Else
                    cnt = cnt + 1
                    s = " '" & codes & "', "                             ' �ڵ�
                    s = s & " '" & stprice.GetDataValue(0, j) & "', "    ' ����
                    s = s & " " & stprice.GetDataValue(1, j) & ", "    ' �ð�
                    s = s & " " & stprice.GetDataValue(2, j) & ", "    ' ��
                    s = s & " " & stprice.GetDataValue(3, j) & ", "    ' ����
                    s = s & " " & stprice.GetDataValue(4, j) & ", "    ' ����
                    s = s & " '" & stprice.GetDataValue(5, j) & "', "    ' ���ϴ��
                    s = s & " " & stprice.GetDataValue(6, j) & ", "    ' �ŷ���
                    s = s & " '" & stprice.GetDataValue(10, j) & "', "   ' �����
                    s = s & " '" & stprice.GetDataValue(11, j) & "', "   ' �����
                    s = s & " " & stprice.GetDataValue(7, j) & ", "       ' ���κ���
                    s = s & " '" & stprice.GetDataValue(9, j) & "', "    ' ���κ���
                    s = s & " " & stprice.GetDataValue(12, j) & " "   ' ������ż���
                    
                    sql = "INSERT INTO COMPANY_PRICE"
                    sql = sql & " (COMPANY_CODE, REG_DATE, OPEN_PRICE, HIGH_PRICE, LOW_PRICE, CLOSE_PRICE, "
                    sql = sql & "  YD_RATIO, VOLUME, UPDOWN_RATIO, UPDOWN_CODE, FOREIGNER_VOLUME, FOREIGNER_PORTION, ORG_VOLUME, "
                    sql = sql & "  INSERT_DT, INSERT_ID, UPDATE_DT, UPDATE_ID ) "
                    sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                    conn.Execute (sql)
                    
                    List1.AddItem (s)
               
                End If
            Next
            
            Do While stprice.Continue And bFind = False
                stprice.BlockRequest
                For j = 0 To n - 1
                    tempdate2 = stprice.GetDataValue(0, j)
                    If tempdate >= tempdate2 Then
                        Exit Do
                    ElseIf tempdate2 = 0 Then
                        Exit Do
                    Else
                    
                     cnt = cnt + 1
                     s = " '" & codes & "', "                             ' �ڵ�
                     s = s & " '" & stprice.GetDataValue(0, j) & "', "    ' ����
                     s = s & " " & stprice.GetDataValue(1, j) & ", "    ' �ð�
                     s = s & " " & stprice.GetDataValue(2, j) & ", "    ' ��
                     s = s & " " & stprice.GetDataValue(3, j) & ", "    ' ����
                     s = s & " " & stprice.GetDataValue(4, j) & ", "    ' ����
                     s = s & " '" & stprice.GetDataValue(5, j) & "', "    ' ���ϴ��
                     s = s & " " & stprice.GetDataValue(6, j) & ", "    ' �ŷ���
                     s = s & " '" & stprice.GetDataValue(10, j) & "', "   ' �����
                     s = s & " '" & stprice.GetDataValue(11, j) & "', "   ' �����
                     s = s & " " & stprice.GetDataValue(7, j) & ", "       ' ���κ���
                     s = s & " '" & stprice.GetDataValue(9, j) & "', "    ' ���κ���
                     s = s & " " & stprice.GetDataValue(12, j) & " "   ' ������ż���
                     
                     sql = "INSERT INTO COMPANY_PRICE"
                     sql = sql & " (COMPANY_CODE, REG_DATE, OPEN_PRICE, HIGH_PRICE, LOW_PRICE, CLOSE_PRICE, "
                     sql = sql & "  YD_RATIO, VOLUME, UPDOWN_RATIO, UPDOWN_CODE, FOREIGNER_VOLUME, FOREIGNER_PORTION, ORG_VOLUME, "
                     sql = sql & "  INSERT_DT, INSERT_ID, UPDATE_DT, UPDATE_ID ) "
                     sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                     conn.Execute (sql)
                     
                     List1.AddItem (s)
               
                    End If
                 Next
            Loop   ' DO WHILE lOOP End
            
            Call DB_Close
            Call Log(codes & "[" & codemgr.CodeToName(codes) & "] �ְ� ��������" & "[" & cnt & "] ")
                 
            Set stprice = Nothing
            Set codemgr = Nothing
End Sub

Sub Company_Investors(codes As String)
            
            '������ü�� ��ȸ  --------------------------------------------------------------------------------------------------------
            Dim stinvestor As New CpSvr7254 '������ü
            Dim codemgr    As New CpCodeMgr    '�ڵ�
            'Dim codes  As String
            
            Dim startdate As String
            Dim enddate  As String
            Dim tempdate As String
            Dim tempdate2 As String
            Dim s, n  As String
            
            Call Log(codes & "[" & codemgr.CodeToName(codes) & "] ������ü ������ ")
            
            Call DB_Open
            
            '�ֱ����� ��ȸ
            sql = "SELECT MAX(REG_DATE) AS REG_DATE FROM COMPANY_INVESTORS WHERE COMPANY_CODE = '" & codes & "' "
            rs.CursorLocation = adUseClient
    
            rs.Open sql, conn, adOpenStatic, adLockOptimistic
            
            If rs.EOF Or rs.BOF Then
                Debug.Print "DB�� ��ϵ��� �ʾҽ��ϴ�.[EOF]. ����� �����մϴ�. "
                tempdate = "19000101"
            Else
                If IsNull(rs("REG_DATE")) Then
                   Debug.Print "DB�� ��ϵ��� �ʾҽ��ϴ�.[NULL]. ����� �����մϴ�. "
                   tempdate = "19000101"
                Else
                   tempdate = rs("REG_DATE")
                   Debug.Print "�ֱ� DB ������ڴ� " & tempdate & "�Դϴ�. "
                End If
            End If
                
            rs.Close
            Set rs = Nothing
            
            List1.Clear
            
            startdate = "19900101"  '�ӽð�
            enddate = Format(Now, "yyyyMMdd") '�׽�Ʈ
    
            stinvestor.SetInputValue 0, codes '�����ڵ�
            stinvestor.SetInputValue 1, 6 '�Ⱓ���������
            stinvestor.SetInputValue 2, CLng(startdate) '��������
            stinvestor.SetInputValue 3, CLng(enddate) '��������
            stinvestor.SetInputValue 4, Asc("0") '���ż�
            stinvestor.SetInputValue 5, 0 '������ ��ü
            
            stinvestor.BlockRequest
            bFind = False
                            
            s = "����  ����   �ܱ���  �����  ��������   ����    ����   ����  ��Ÿ����   �����   ��Ÿ����  ��Ÿ����  ����ݵ�  ��������ü"
            List1.AddItem (s)
            
            cnt = 0
            
            For j = 0 To stinvestor.GetHeaderValue(1) - 1                  ' ���� ������ ����ŭ ������ ��
                tempdate2 = stinvestor.GetDataValue(0, j)
                If tempdate >= tempdate2 Then
                    Exit For
                ElseIf tempdate2 = 0 Then
                    Exit For
                Else
                    cnt = cnt + 1
                    s = " '" & codes & "', "                             ' �ڵ�
                    s = s & " '" & stinvestor.GetDataValue(0, j) & "', "    ' ����
                    s = s & " " & stinvestor.GetDataValue(1, j) & ", "    ' ����
                    s = s & " " & stinvestor.GetDataValue(2, j) & ", "    ' �ܱ���
                    s = s & " " & stinvestor.GetDataValue(3, j) & ", "    ' �����
                    s = s & " " & stinvestor.GetDataValue(4, j) & ", "    ' ��������
                    s = s & " " & stinvestor.GetDataValue(5, j) & ", "    ' ����
                    s = s & " " & stinvestor.GetDataValue(6, j) & ", "    ' ����
                    s = s & " " & stinvestor.GetDataValue(7, j) & ", "    ' ����
                    s = s & " " & stinvestor.GetDataValue(8, j) & ", "   ' ��Ÿ����
                    s = s & " " & stinvestor.GetDataValue(9, j) & ", "    ' �����
                    s = s & " " & stinvestor.GetDataValue(10, j) & ", "    ' ��Ÿ����13���� �ٲ�
                    s = s & " " & stinvestor.GetDataValue(11, j) & ", "    ' ��Ÿ����
                    s = s & " " & stinvestor.GetDataValue(12, j) & ", "    ' ����ݵ�
                    s = s & " " & stinvestor.GetDataValue(13, j) & " "    ' ��������ü
                
                    sql = "INSERT INTO COMPANY_INVESTORS "
                    sql = sql & " (COMPANY_CODE, REG_DATE, PERSONAL, FOREIGNER, ORG_SUM, FIN_INVEST, INSURANCE, INVEST_TRUST, "
                    sql = sql & "  BANK, ETC_FIN, PENSION, ETC_CORP,  ETC_FOREIGNER, PRIVATE_FUND, NATION, "
                    sql = sql & "  INSERT_DT, INSERT_ID, UPDATE_DT, UPDATE_ID ) "
                    sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                    conn.Execute (sql)
                    
                    List1.AddItem (s)
                
                End If
            Next
            
            Do While stinvestor.Continue And bFind = False
                stinvestor.BlockRequest
                
                For j = 0 To stinvestor.GetHeaderValue(1) - 1
                    tempdate2 = stinvestor.GetDataValue(0, j)
                    If tempdate >= tempdate2 Then
                        Exit Do
                    ElseIf tempdate2 = 0 Then
                        Exit Do
                    Else
                        cnt = cnt + 1
                        s = " '" & codes & "', "                             ' �ڵ�
                        s = s & " '" & stinvestor.GetDataValue(0, j) & "', "    ' ����
                        s = s & " " & stinvestor.GetDataValue(1, j) & ", "    ' ����
                        s = s & " " & stinvestor.GetDataValue(2, j) & ", "    ' �ܱ���
                        s = s & " " & stinvestor.GetDataValue(3, j) & ", "    ' �����
                        s = s & " " & stinvestor.GetDataValue(4, j) & ", "    ' ��������
                        s = s & " " & stinvestor.GetDataValue(5, j) & ", "    ' ����
                        s = s & " " & stinvestor.GetDataValue(6, j) & ", "    ' ����
                        s = s & " " & stinvestor.GetDataValue(7, j) & ", "    ' ����
                        s = s & " " & stinvestor.GetDataValue(8, j) & ", "   ' ��Ÿ����
                        s = s & " " & stinvestor.GetDataValue(9, j) & ", "    ' �����
                        s = s & " " & stinvestor.GetDataValue(10, j) & ", "    ' ��Ÿ����
                        s = s & " " & stinvestor.GetDataValue(11, j) & ", "    ' ��Ÿ����
                        s = s & " " & stinvestor.GetDataValue(12, j) & ", "    ' ����ݵ�
                        s = s & " " & stinvestor.GetDataValue(13, j) & " "    ' ����
                    
                        sql = "INSERT INTO COMPANY_INVESTORS "
                        sql = sql & " (COMPANY_CODE, REG_DATE, PERSONAL, FOREIGNER, ORG_SUM, FIN_INVEST, INSURANCE, INVEST_TRUST, "
                        sql = sql & "  BANK, ETC_FIN, PENSION, NATION,  ETC_FOREIGNER, PRIVATE_FUND, ETC_CORP, "
                        sql = sql & "  INSERT_DT, INSERT_ID, UPDATE_DT, UPDATE_ID ) "
                        sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                        conn.Execute (sql)
                    
                        List1.AddItem (s)
                    
                    End If
               Next
            Loop ' DO WHILE lOOP End
                             
            Call DB_Close
            Call Log(codes & "[" & codemgr.CodeToName(codes) & "] ������ü ����" & "[" & cnt & "] ")
                                                     
            Set stinvestor = Nothing
            Set codemgr = Nothing

End Sub


Sub Company_Fundamental(codes As String)
            '���� Fundamental �������� - �������� ����.
                    
            Dim stfdmt     As New MarketEye    'Fundamental
            Dim codemgr    As New CpCodeMgr    '�ڵ�
            Dim s  As String
            Dim items() As Long
            Dim value As Variant
            Dim hangmok
            Dim regdate As String
            regdate = Format(Now, "yyyyMMdd")
            
            Call DB_Open
            
            sql = "DELETE FROM COMPANY_FUNDAMENTAL WHERE COMPANY_CODE = '" & codes & "' AND REG_DATE = '" & regdate & "'"
            conn.Execute (sql)
            
            
            Call Log(codes & "[" & codemgr.CodeToName(codes) & "] �繫��ǥ ������")
            
            List1.Clear
            
            hangmok = "  20, 24, 67, 70,71,72,75,76,77,78,79,80,81,82,83,86,87,88,89,90,91,92,93,94,95,"
            hangmok = hangmok & "96,97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,"
            hangmok = hangmok & "112,124,125,126,127,128 "
            
            value = Split(hangmok, ",")   '��û�׸�
            ReDim items(0 To UBound(value))
            For i = 0 To UBound(value)
                items(i) = CLng(value(i))
            Next
            
            stfdmt.SetInputValue 0, items
            stfdmt.SetInputValue 1, codes '�����ڵ�
            stfdmt.BlockRequest
            
            value = stfdmt.GetHeaderValue(1) '�ʵ��
            
            cnt = 0
            For i = 0 To stfdmt.GetHeaderValue(2) - 1
                cnt = cnt + 1
                s = "'" & codes & "', '" & regdate & "' "
                For j = 0 To stfdmt.GetHeaderValue(0) - 1
                    value = stfdmt.GetDataValue(j, i)
                    s = s & " , '" & value & "' "
                Next
                
                sql = "INSERT INTO COMPANY_FUNDAMENTAL "
                sql = sql & " (Company_Code , REG_DATE, TOTAL_STOCKS, CONTRACT_STRENGTH, PER, EPS, CAPITAL, FACE_VALUE"
                sql = sql & " ,Y_DEBT_RATIO, Y_RETENTION_RATIO, Y_RETURN_OF_EQUITY, Y_SALES_RISE_RATIO, Y_ORDINARY_PROFIT_RISE_RATIO,Y_NET_PROFIT_RISE_RATIO"
                sql = sql & " ,INVESTOR_SENTIMENT, VR, FIVE_DAY_TURNOVER"
                sql = sql & " ,Y_SALES, Y_ORDINARY_PROFIT, Y_TERM_NET_PROFIT, Y_BPS, Y_BUSINESS_PROFIT_RISE_RATIO"
                sql = sql & " ,Y_BUSINESS_PROFIT, Y_RETURN_ON_NET_SALES, Y_ORDINARY_MARGIN, Y_TIME_INTEREST_EARNED, Y_SETTING_YYYYMM"
                sql = sql & " ,Q_BPS, Q_SALES_RISE_RATIO, Q_BUSINESS_PROFIT_RISE_RATIO, Q_ORDINARY_PROFIT_RISE_RATIO, Q_NET_PROFIT_RISE_RATIO"
                sql = sql & " ,Q_SALES, Q_BUSINESS_PROFIT, Q_ORDINARY_PROFIT, Q_TERM_NET_PROFIT, Q_RETURN_ON_NET_SALES"
                sql = sql & " ,Q_ORDINARY_MARGIN, Q_ROE, Q_TIME_INTEREST_EARNED, Q_RETENTION_RATIO, Q_DEBT_RATIO, Q_SETTING_YYYYMM"
                sql = sql & " ,BASIS,CFPS,EBITDA,DEBIT_BALANCE_RATIO,SHORT_STOCK_SELLING_VOL,SHORT_STOCK_SELLING_YM"
                sql = sql & " ,INSERT_DT,INSERT_ID,UPDATE_DT,UPDATE_ID )"
                sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                conn.Execute (sql)
                
                List1.AddItem s
                
            Next
            Call DB_Close
            Call Log(codes & "[" & codemgr.CodeToName(codes) & "] �繫��ǥ ��������" & "[" & cnt & "] ")
            
            Set stfdmt = Nothing
            Set codemgr = Nothing
            
End Sub


Sub Industry_pi()
            Call Log("-- ���� �ְ�/������ü ��������")
            
            Call DB_Open
            
            Dim arrList(2000) 'As Variant
            Dim i
            Dim codes As String
            
            sql = "SELECT INDUSTRY_CODE, INDUSTRY_NAME FROM INDUSTRY "
            sql = sql & " WHERE 1=1 "
            rs.CursorLocation = adUseClient
            rs.Open sql, conn, adOpenStatic, adLockOptimistic
            
            i = 0
            While rs.EOF = False
                arrList(i) = rs("INDUSTRY_CODE")
                rs.MoveNext
                i = i + 1
            Wend
            
            rs.Close
            Set rs = Nothing
            Call DB_Close
            
            For i = 0 To UBound(arrList)
               codes = Trim(CStr(arrList(i)))
               If codes = "" Then
               Else
                    Call Industry_Price(codes)
                    Call Industry_Investor(codes)
               End If
            Next
            Call Log("-- ���� �ְ�/������ü ��������")
            
End Sub

Sub Industry_Price(codes As String)
            
            Dim codemgr    As New CpCodeMgr    '�ڵ�
            Dim ujchart    As New CbGraph1
            Dim tempdate  As String
            Dim tempdate2 As String
            Dim s  As String
            
            Call Log(codes & "[" & codemgr.GetIndustryName(codes) & "] �ְ� ������")
            Call DB_Open
                
            '�ֱ����� ��ȸ
            sql = "SELECT MAX(REG_DATE) AS REG_DATE FROM INDUSTRY_PRICE WHERE INDUSTRY_CODE = '" & codes & "' "
            rs.CursorLocation = adUseClient
    
            rs.Open sql, conn, adOpenStatic, adLockOptimistic
            
            If rs.EOF Or rs.BOF Then
                Debug.Print "DB�� ��ϵ��� �ʾҽ��ϴ�.[EOF]. ����� �����մϴ�. "
                tempdate = "19000101"
            Else
                If IsNull(rs("REG_DATE")) Then
                   Debug.Print "DB�� ��ϵ��� �ʾҽ��ϴ�.[NULL]. ����� �����մϴ�. "
                   tempdate = "19000101"
                Else
                    tempdate = rs("REG_DATE")
                    Debug.Print "�ֱ� DB ������ڴ� " & tempdate & "�Դϴ�. "
                End If
            End If
                
            rs.Close
            Set rs = Nothing
                
            List1.Clear
           
           
            Dim items(0 To 37) As Long
            Dim a As Variant
           
            For i = 0 To 37
                items(i) = i
            Next
                                      
            ujchart.SetInputValue 0, "U" & codes '�����ڵ�
            ujchart.SetInputValue 1, Asc("D")   '��û���� 1�� �Ⱓ, 2����
            ujchart.SetInputValue 3, 2800   '����
            ujchart.SetInputValue 4, Asc("1")   '1. �����ְ�
                                              
            ujchart.BlockRequest
            bFind = False
                           
            s = "����  �ð�  ��   ����  ����  ���ϴ�� �ŷ���(����) �ŷ���� "
            List1.AddItem (s)
           
            n = ujchart.GetHeaderValue(3)  '���Ű���
                      
            cnt = 0
            
            For j = 0 To ujchart.GetHeaderValue(3) - 1
                tempdate2 = ujchart.GetDataValue(0, j)
                If tempdate >= tempdate2 Then
                    Exit For
                ElseIf tempdate2 = 0 Then
                    Exit For
                Else
                    cnt = cnt + 1
                    s = " '" & codes & "', "                             ' �ڵ�
                    s = s & " '" & ujchart.GetDataValue(0, j) & "', "    ' ����
                    s = s & " " & ujchart.GetDataValue(1, j) & ", "    ' �ð�
                    s = s & " " & ujchart.GetDataValue(2, j) & ", "    ' ��
                    s = s & " " & ujchart.GetDataValue(3, j) & ", "    ' ����
                    s = s & " " & ujchart.GetDataValue(4, j) & ", "    ' ����
                    s = s & " " & ujchart.GetDataValue(5, j) & "  "    ' �ŷ���
                    
                    sql = "INSERT INTO INDUSTRY_PRICE"
                    sql = sql & " (INDUSTRY_CODE, REG_DATE, OPEN_PRICE, HIGH_PRICE, LOW_PRICE, CLOSE_PRICE, VOLUME, "
                    sql = sql & "  INSERT_DT, INSERT_ID, UPDATE_DT, UPDATE_ID ) "
                    sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                    conn.Execute (sql)
                    
                    List1.AddItem (s)
               
                End If
            Next
           
            Do While ujchart.Continue And bFind = False

                ujchart.BlockRequest
               
                For j = 0 To ujchart.GetHeaderValue(3) - 1
                    tempdate2 = ujchart.GetDataValue(0, j)
                    If tempdate >= tempdate2 Then
                         Exit For
                    ElseIf tempdate2 = 0 Then
                         Exit For
                    Else
                         cnt = cnt + 1
                         s = " '" & codes & "', "                             ' �ڵ�
                         s = s & " '" & ujchart.GetDataValue(0, j) & "', "    ' ����
                         s = s & " " & ujchart.GetDataValue(1, j) & ", "    ' �ð�
                         s = s & " " & ujchart.GetDataValue(2, j) & ", "    ' ��
                         s = s & " " & ujchart.GetDataValue(3, j) & ", "    ' ����
                         s = s & " " & ujchart.GetDataValue(4, j) & ", "    ' ����
                         s = s & " " & ujchart.GetDataValue(5, j) & "  "    ' �ŷ���
                         
                         sql = "INSERT INTO INDUSTRY_PRICE"
                         sql = sql & " (INDUSTRY_CODE, REG_DATE, OPEN_PRICE, HIGH_PRICE, LOW_PRICE, CLOSE_PRICE, VOLUME, "
                         sql = sql & "  INSERT_DT, INSERT_ID, UPDATE_DT, UPDATE_ID ) "
                         sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                         conn.Execute (sql)
                         
                         List1.AddItem (s)
                    
                    End If
                 Next
             Loop ' DO WHILE lOOP End
             
             Call DB_Close
             Call Log(codes & "[" & codemgr.GetIndustryName(codes) & "] �ְ� ��������" & "[" & cnt & "] ")
             
             Set codemgr = Nothing    '�ڵ�
             Set ujchart = Nothing
            
End Sub


Sub Industry_Investor(codes As String)
        '������ü�� ��ȸ  --------------------------------------------------------------------------------------------------------
            Dim codemgr    As New CpCodeMgr    '�ڵ�
            Dim ujinvestor As New CpSvr7223 '������ü
            Dim tempdate2 As String
            Dim s, n  As String
            
            Call Log(codes & "[" & codemgr.CodeToName(codes) & "] ������ü ������")
            
            Call DB_Open
    
            '�ֱ����� ��ȸ
            'Set rs = New ADODB.Recordset
            sql = "SELECT MAX(REG_DATE) AS REG_DATE FROM INDUSTRY_INVESTORS WHERE INDUSTRY_CODE = '" & codes & "' "
            rs.CursorLocation = adUseClient
    
            rs.Open sql, conn, adOpenStatic, adLockOptimistic
            
            If rs.EOF Or rs.BOF Then
                Debug.Print "DB�� ��ϵ��� �ʾҽ��ϴ�.[EOF]. ����� �����մϴ�. "
                tempdate = "19000101"
            Else
                If IsNull(rs("REG_DATE")) Then
                   Debug.Print "DB�� ��ϵ��� �ʾҽ��ϴ�.[NULL]. ����� �����մϴ�. "
                   tempdate = "19000101"
                Else
                   tempdate = rs("REG_DATE")
                   Debug.Print "�ֱ� DB ������ڴ� " & tempdate & "�Դϴ�. "
                End If
            End If
                
            rs.Close
            Set rs = Nothing
            
            List1.Clear
            
            ujinvestor.SetInputValue 0, Asc("4") '���ں�
            ujinvestor.SetInputValue 1, codes    '�����ڵ�
            
            ujinvestor.BlockRequest
            bFind = False
                            
            s = "����  ����   �ܱ���  �����  ��������   ����    ����   ����  ��Ÿ����   �����   ��Ÿ����  ��Ÿ����  ����ݵ�  ��������ü"
            List1.AddItem (s)
            
            cnt = 0
            
            For j = 0 To ujinvestor.GetHeaderValue(1) - 1
                
                tempdate2 = Replace(ujinvestor.GetDataValue(0, j), Chr("47"), "")
                
                If tempdate >= tempdate2 Then
                    Exit For
                ElseIf tempdate2 = 0 Then
                    Exit For
                Else
                    cnt = cnt + 1
                    s = " '" & codes & "', "                             ' �ڵ�
                    s = s & " '" & tempdate2 & "', "    ' ����
                    s = s & " " & ujinvestor.GetDataValue(1, j) & ", "    ' ����
                    s = s & " " & ujinvestor.GetDataValue(2, j) & ", "    ' �ܱ���
                    s = s & " " & ujinvestor.GetDataValue(3, j) & ", "    ' �����
                    s = s & " " & ujinvestor.GetDataValue(4, j) & ", "    ' ��������
                    s = s & " " & ujinvestor.GetDataValue(5, j) & ", "    ' ����
                    s = s & " " & ujinvestor.GetDataValue(6, j) & ", "    ' ����
                    s = s & " " & ujinvestor.GetDataValue(7, j) & ", "    ' ����
                    s = s & " " & ujinvestor.GetDataValue(8, j) & ", "   ' ��Ÿ����
                    s = s & " " & ujinvestor.GetDataValue(9, j) & ", "    ' �����
                    s = s & " " & ujinvestor.GetDataValue(10, j) & ", "    ' ��Ÿ����
                    s = s & " " & ujinvestor.GetDataValue(11, j) & ", "    ' ��Ÿ����
                    s = s & " " & ujinvestor.GetDataValue(12, j) & ", "    ' ����ݵ�
                    s = s & " " & ujinvestor.GetDataValue(13, j) & " "    ' ����
                    
                    sql = "INSERT INTO INDUSTRY_INVESTORS "
                    sql = sql & " (INDUSTRY_CODE, REG_DATE, PERSONAL, FOREIGNER, ORG_SUM, FIN_INVEST, INSURANCE, INVEST_TRUST, "
                    sql = sql & "  BANK, ETC_FIN, PENSION, ETC_CORP, ETC_FOREIGNER, PRIVATE_FUND, NATION, "
                    sql = sql & "  INSERT_DT, INSERT_ID, UPDATE_DT, UPDATE_ID ) "
                    sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                    conn.Execute (sql)
                    
                    List1.AddItem (s)
                
                End If
            Next
            
            Do While ujinvestor.Continue And bFind = False
                ujinvestor.BlockRequest
                
                For j = 0 To ujinvestor.GetHeaderValue(1) - 1
                    tempdate2 = Replace(ujinvestor.GetDataValue(0, j), Chr("47"), "")
                    If tempdate >= tempdate2 Then
                        Exit Do
                    ElseIf tempdate2 = 0 Then
                        Exit Do
                    Else
                        cnt = cnt + 1
                        s = " '" & codes & "', "                             ' �ڵ�
                        s = s & " '" & tempdate2 & "', "    ' ����
                        s = s & " " & ujinvestor.GetDataValue(1, j) & ", "    ' ����
                        s = s & " " & ujinvestor.GetDataValue(2, j) & ", "    ' �ܱ���
                        s = s & " " & ujinvestor.GetDataValue(3, j) & ", "    ' �����
                        s = s & " " & ujinvestor.GetDataValue(4, j) & ", "    ' ��������
                        s = s & " " & ujinvestor.GetDataValue(5, j) & ", "    ' ����
                        s = s & " " & ujinvestor.GetDataValue(6, j) & ", "    ' ����
                        s = s & " " & ujinvestor.GetDataValue(7, j) & ", "    ' ����
                        s = s & " " & ujinvestor.GetDataValue(8, j) & ", "   ' ��Ÿ����
                        s = s & " " & ujinvestor.GetDataValue(9, j) & ", "    ' �����
                        s = s & " " & ujinvestor.GetDataValue(10, j) & ", "    ' ��Ÿ����
                        s = s & " " & ujinvestor.GetDataValue(11, j) & ", "    ' ��Ÿ����
                        s = s & " " & ujinvestor.GetDataValue(12, j) & ", "    ' ����ݵ�
                        s = s & " " & ujinvestor.GetDataValue(13, j) & " "    ' ����
                        
                        sql = "INSERT INTO INDUSTRY_INVESTORS "
                        sql = sql & " (INDUSTRY_CODE, REG_DATE, PERSONAL, FOREIGNER, ORG_SUM, FIN_INVEST, INSURANCE, INVEST_TRUST, "
                        sql = sql & "  BANK, ETC_FIN, PENSION, ETC_CORP, ETC_FOREIGNER, PRIVATE_FUND, NATION, "
                        sql = sql & "  INSERT_DT, INSERT_ID, UPDATE_DT, UPDATE_ID ) "
                        sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                        conn.Execute (sql)
                        
                        List1.AddItem (s)
                    
                    End If
               Next
            Loop ' DO WHILE lOOP End
            
            Call DB_Close
            Call Log(codes & "[" & codemgr.CodeToName(codes) & "] ������ü ��������" & "[" & cnt & "] ")
            
            
            Set codemgr = Nothing    '�ڵ�
            Set ujinvestor = Nothing '������ü
            
End Sub


Sub Theme_company()
            Call DB_Open
            
            Dim arrList(3000) 'As Variant
            Dim i
            Dim codes As String
            Dim regdate As String
            regdate = Format(Now, "yyyyMMdd")
            
            sql = "SELECT THEME_CODE AS THEME_CODE  "
            sql = sql & "FROM THEME "
            sql = sql & "WHERE REG_DATE = '" & regdate & "'"
            rs.CursorLocation = adUseClient
            rs.Open sql, conn, adOpenStatic, adLockOptimistic
            
            i = 0
            While rs.EOF = False
                arrList(i) = rs("THEME_CODE")
                rs.MoveNext
                i = i + 1
            Wend
            
            rs.Close
            Set rs = Nothing
            Call DB_Close
            
            
            Call DB_Open
            
            sql = "DELETE FROM THEME_COMPANY WHERE REG_DATE = '" & regdate & "'"
            conn.Execute (sql)

            Call DB_Close
            
            For i = 0 To UBound(arrList)
               codes = Trim(CStr(arrList(i)))
               If codes = "" Then
               Else
                Call Theme_company_1(codes)
               End If
            Next
            Call Log("-- �׸� ���� ��������")

End Sub

Sub Theme_company_1(themecode)
            '�׸� ���� �������� - �������� ����.
                    
            Dim incpsvr8561T    As New CpSvr8561T    '�ڵ�
            
            Dim s  As String
            Dim cp As String
            Dim regdate As String
            regdate = Format(Now, "yyyyMMdd")
            
            Call DB_Open
            
            Call Log("�׸� ���� ������")
            
            List1.Clear
            bFind = False
            
            incpsvr8561T.SetInputValue 0, themecode
            
            incpsvr8561T.BlockRequest
            
            n = incpsvr8561T.GetHeaderValue(1)
            
            For i = 0 To n - 1
               cp = Trim(CStr(incpsvr8561T.GetDataValue(0, i)))
               If cp = "" Then
               Else
                    s = "'" & themecode
                    s = s & "', '" & regdate
                    s = s & "', '" & cp & "' "
                    
                    sql = "INSERT INTO THEME_COMPANY "
                    sql = sql & " (THEME_CODE , REG_DATE, COMPANY_CODE"
                    sql = sql & " ,INSERT_DT,INSERT_ID,UPDATE_DT,UPDATE_ID )"
                    sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                    conn.Execute (sql)
                    
                    List1.AddItem s
               End If
                
            Next
            
            Do While incpsvr8561T.Continue And bFind = False
                incpsvr8561T.BlockRequest
                For j = 0 To n - 1
                    cp = Trim(CStr(incpsvr8561T.GetDataValue(0, j)))
                    If cp = "" Then
                    Else
                         s = "'" & themecode
                         s = s & "', '" & regdate
                         s = s & "', '" & cp & "' "
                     
                         sql = "INSERT INTO THEME_COMPANY "
                         sql = sql & " (THEME_CODE , REG_DATE, COMPANY_CODE"
                         sql = sql & " ,INSERT_DT,INSERT_ID,UPDATE_DT,UPDATE_ID )"
                         sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                         conn.Execute (sql)
                         
                         List1.AddItem s
                    End If
                 Next
            Loop   ' DO WHILE lOOP End
            
            Call DB_Close
            Call Log("�׸��ڵ�  ��������")
            
            Set incpsvr8561T = Nothing
            
End Sub





'------------------------- ���� �ڵ强 -------------------------------------------------------
Sub Industry_Code()
            Call Log("�� �����ڵ� ��������")
            Call DB_Open
            
            Dim codemgr As New CpCodeMgr
            Dim codes As Variant
            Dim s, n
            
            
            'KOSPI�ڵ�, �� ��������
            codes = codemgr.GetIndustryList()
            
            sql = "DELETE FROM INDUSTRY  "
            conn.Execute (sql)
            
                
            cnt = 0
            For i = LBound(codes) To UBound(codes)
                cnt = cnt + 1
                s = " '" & codes(i) & "', '"
                s = s & codemgr.CodeToName(codes(i)) & "', '" & codemgr.GetStockMarketKind(codes(i)) & "' "

                            
                sql = "INSERT INTO INDUSTRY"
                sql = sql & " (INDUSTRY_CODE, INDUSTRY_NAME, INDUSTRY_KIND,  "
                sql = sql & "  INSERT_DT, INSERT_ID, UPDATE_DT, UPDATE_ID ) "
                sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                'Debug.Print sql
                conn.Execute (sql)
                
                List1.AddItem s
            Next
            
            'KOSDAQ�ڵ�, �� ��������
            
            codes = codemgr.GetKosdaqIndustry1List()
            
                     
            For i = LBound(codes) To UBound(codes)
                cnt = cnt + 1
                s = " '" & codes(i) & "', '"
                s = s & codemgr.CodeToName(codes(i)) & "', '" & codemgr.GetStockMarketKind(codes(i)) & "' "

                            
                sql = "INSERT INTO INDUSTRY"
                sql = sql & " (INDUSTRY_CODE, INDUSTRY_NAME, INDUSTRY_KIND,  "
                sql = sql & "  INSERT_DT, INSERT_ID, UPDATE_DT, UPDATE_ID ) "
                sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                'Debug.Print sql
                conn.Execute (sql)
                
                List1.AddItem s
            Next
            
            'KOSDAQ�ڵ�, �� ��������
            codes = codemgr.GetKosdaqIndustry2List()
            
                   
             For i = LBound(codes) To UBound(codes)
                cnt = cnt + 1
                s = " '" & codes(i) & "', '"
                s = s & codemgr.CodeToName(codes(i)) & "', '" & codemgr.GetStockMarketKind(codes(i)) & "' "

                            
                sql = "INSERT INTO INDUSTRY"
                sql = sql & " (INDUSTRY_CODE, INDUSTRY_NAME, INDUSTRY_KIND,  "
                sql = sql & "  INSERT_DT, INSERT_ID, UPDATE_DT, UPDATE_ID ) "
                sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                'Debug.Print sql
                'MsgBox (sql)
                conn.Execute (sql)
                
                List1.AddItem s
            Next
            
            Set codemgr = Nothing
            Set codes = Nothing
            
            Call DB_Close
            Call Log("[" & cnt & "] �� �����ڵ� ������")

End Sub


Sub Company_Code()
            Call Log("�� �����ڵ� ��������")
            Call DB_Open
            
            Dim codemgr As New CpCodeMgr
            Dim codes As Variant
            Dim s, n
            
            'KOSPI�����ڵ�, �� ��������
            codes = codemgr.GetStockListByMarket(1)
                
            sql = "DELETE FROM COMPANY  "
            conn.Execute (sql)
            
            cnt = 0
            For i = LBound(codes) To UBound(codes)
                cnt = cnt + 1
                s = " '" & codes(i) & "', '"
                s = s & codemgr.CodeToName(codes(i)) & "', '" & codemgr.GetStockMarketKind(codes(i)) & "', '"
                s = s & codemgr.GetStockIndustryCode(codes(i)) & "', '" & codemgr.GetStockKospi200Kind(codes(i)) & "', '"
                s = s & codemgr.GetStockSupervisionKind(codes(i)) & "', '" & codemgr.GetStockControlKind(codes(i)) & "', '"
                s = s & codemgr.GetStockStatusKind(codes(i)) & "', '" & codemgr.GetStockLacKind(codes(i)) & "', '"
                s = s & codemgr.GetStockMarginRate(codes(i)) & "', '" & codemgr.GetStockMemeMin(codes(i)) & "', '"
                s = s & codemgr.GetStockCapital(codes(i)) & "', '" & codemgr.GetStockFiscalMonth(codes(i)) & "' "

                            
                sql = "INSERT INTO COMPANY"
                sql = sql & " (COMPANY_CODE, COMPANY_NAME, MARKET_KIND, INDUSTRY_CODE, KOSPI200, "
                sql = sql & " SUPERVISION, CONTROL_KIND, STATUS_KIND, LOCK_KIND,"
                sql = sql & "  MARGIN_RATE , MIN_TRADE_UNIT, CAPITAL, FISCAL_MONTH, "
                sql = sql & "  INSERT_DT, INSERT_ID, UPDATE_DT, UPDATE_ID ) "
                sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                'Debug.Print sql
                conn.Execute (sql)
                
                List1.AddItem s
            Next
            
            'KOSDAQ�����ڵ�, �� ��������
            
            codes = codemgr.GetStockListByMarket(2)
            
                     
            For i = LBound(codes) To UBound(codes)
                cnt = cnt + 1
                s = " '" & codes(i) & "', '"
                s = s & codemgr.CodeToName(codes(i)) & "', '" & codemgr.GetStockMarketKind(codes(i)) & "', '"
                s = s & codemgr.GetStockIndustryCode(codes(i)) & "', '" & codemgr.GetStockKospi200Kind(codes(i)) & "', '"
                s = s & codemgr.GetStockSupervisionKind(codes(i)) & "', '" & codemgr.GetStockControlKind(codes(i)) & "', '"
                s = s & codemgr.GetStockStatusKind(codes(i)) & "', '" & codemgr.GetStockLacKind(codes(i)) & "', '"
                s = s & codemgr.GetStockMarginRate(codes(i)) & "', '" & codemgr.GetStockMemeMin(codes(i)) & "', '"
                s = s & codemgr.GetStockCapital(codes(i)) & "', '" & codemgr.GetStockFiscalMonth(codes(i)) & "' "

                            
                sql = "INSERT INTO COMPANY"
                sql = sql & " (COMPANY_CODE, COMPANY_NAME, MARKET_KIND, INDUSTRY_CODE, KOSPI200, "
                sql = sql & " SUPERVISION, CONTROL_KIND, STATUS_KIND, LOCK_KIND,"
                sql = sql & "  MARGIN_RATE , MIN_TRADE_UNIT, CAPITAL, FISCAL_MONTH, "
                sql = sql & "  INSERT_DT, INSERT_ID, UPDATE_DT, UPDATE_ID ) "
                sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                'Debug.Print sql
                conn.Execute (sql)
                
                List1.AddItem s
            Next
            
            codes = codemgr.GetStockListByMarket(3)
            
                   
             For i = LBound(codes) To UBound(codes)
                cnt = cnt + 1
                s = " '" & codes(i) & "', '"
                s = s & Replace(codemgr.CodeToName(codes(i)), "'", "") & "', '" & codemgr.GetStockMarketKind(codes(i)) & "', '"
                s = s & codemgr.GetStockIndustryCode(codes(i)) & "', '" & codemgr.GetStockKospi200Kind(codes(i)) & "', '"
                s = s & codemgr.GetStockSupervisionKind(codes(i)) & "', '" & codemgr.GetStockControlKind(codes(i)) & "', '"
                s = s & codemgr.GetStockStatusKind(codes(i)) & "', '" & codemgr.GetStockLacKind(codes(i)) & "', '"
                s = s & codemgr.GetStockMarginRate(codes(i)) & "', '" & codemgr.GetStockMemeMin(codes(i)) & "', '"
                s = s & codemgr.GetStockCapital(codes(i)) & "', '" & codemgr.GetStockFiscalMonth(codes(i)) & "' "

                            
                sql = "INSERT INTO COMPANY"
                sql = sql & " (COMPANY_CODE, COMPANY_NAME, MARKET_KIND, INDUSTRY_CODE, KOSPI200, "
                sql = sql & " SUPERVISION, CONTROL_KIND, STATUS_KIND, LOCK_KIND,"
                sql = sql & "  MARGIN_RATE , MIN_TRADE_UNIT, CAPITAL, FISCAL_MONTH, "
                sql = sql & "  INSERT_DT, INSERT_ID, UPDATE_DT, UPDATE_ID ) "
                sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                'Debug.Print sql
                'MsgBox (sql)
                conn.Execute (sql)
                
                List1.AddItem s
            Next
            
            Set codemgr = Nothing
            Set codes = Nothing
            
            Call DB_Close
            Call Log("[" & cnt & "] �� �����ڵ� ������")

End Sub

Sub Theme_Code()
            '�׸��ڵ� �������� - �������� ����.
                    
            Dim incpsvr8561    As New CpSvr8561    '�ڵ�
            Dim s  As String
            Dim regdate As String
            regdate = Format(Now, "yyyyMMdd")
            
            Call DB_Open
            
            sql = "DELETE FROM THEME WHERE REG_DATE = '" & regdate & "'"
            conn.Execute (sql)
            
            
            Call Log("�׸� �ڵ� ������")
            
            List1.Clear
            incpsvr8561.BlockRequest
            bFind = False
            
            n = incpsvr8561.GetHeaderValue(0)
            
            For i = 0 To n - 1
                
                s = "'" & incpsvr8561.GetDataValue(0, i)
                s = s & "', '" & regdate
                s = s & "', '" & incpsvr8561.GetDataValue(2, i) & "' "
                
                sql = "INSERT INTO THEME "
                sql = sql & " (THEME_CODE , REG_DATE, THEME_NAME"
                sql = sql & " ,INSERT_DT,INSERT_ID,UPDATE_DT,UPDATE_ID )"
                sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                conn.Execute (sql)
                
                List1.AddItem s
                
            Next
            
            Do While incpsvr8561.Continue And bFind = False
                incpsvr8561.BlockRequest
                For j = 0 To n - 1
                    s = "'" & incpsvr8561.GetDataValue(0, j)
                    s = s & "', '" & regdate
                    s = s & "', '" & incpsvr8561.GetDataValue(2, j) & "' "
                
                    sql = "INSERT INTO THEME "
                    sql = sql & " (THEME_CODE , REG_DATE, THEME_NAME"
                    sql = sql & " ,INSERT_DT,INSERT_ID,UPDATE_DT,UPDATE_ID )"
                    sql = sql & " VALUES (" & s & ", NOW(), 'crawler', NOW(), 'crawler' )"
                    conn.Execute (sql)
                     
                    List1.AddItem (s)
               
                 Next
            Loop   ' DO WHILE lOOP End
            
            Call DB_Close
            Call Log("�׸��ڵ�  ��������")
            
            Set incpsvr8561 = Nothing
            
End Sub


