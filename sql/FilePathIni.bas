Attribute VB_Name = "Module3"
Option Explicit
Public DataBasePath As String
Public myADO As ADODB.Connection
Public myRS As ADODB.Recordset
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function PathFileIni() As String
PathFileIni = App.Path
    If Right(PathFileIni, 1) <> "\" Then PathFileIni = PathFileIni & "\" & "setting.ini"
End Function

'�������� ��� =========================================
Public Function Ftp_Obl() As String      ' �������� ���
Ftp_Obl = ReadINI("FtpConnect", "FtpObl", PathFileIni)
End Function
Public Function Kat_Obl() As String      ' ������� ��� ��������
Kat_Obl = ReadINI("FtpConnect", "OblKat", PathFileIni)
End Function
Public Function Login_Obl() As String      ' ���� �� ��� ��������
Login_Obl = ReadINI("FtpConnect", "LogObl", PathFileIni)
End Function
Public Function Pass_Obl() As String       ' ������ �� ��� ��������
Pass_Obl = ReadINI("FtpConnect", "PassObl", PathFileIni)
End Function
'======================================================


'�ȯ������ ��� ++++++++++++++++++++++++++++++++++++++++
Public Function Ftp_Kiev() As String      ' �������� ���
Ftp_Kiev = ReadINI("FtpConnect", "FtpKiev", PathFileIni)
End Function
Public Function Kat_Kiev() As String      ' ������� ��� ��������
Kat_Kiev = ReadINI("FtpConnect", "kievKat", PathFileIni)
End Function
Public Function Login_Kiev() As String     ' ���� �� ��� ��������
Login_Kiev = ReadINI("ftpconnect", "logkiev", PathFileIni)
End Function
Public Function Pass_Kiev() As String      ' ������ �� ��� ��������
Pass_Kiev = ReadINI("ftpconnect", "passkiev", PathFileIni)
End Function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++


Public Function Folder_DB_Raj() As String      '���� �� �� ����Ͳ�
Folder_DB_Raj = ReadINI("SendReport", "PathRpr", PathFileIni)
End Function

Public Function Folder_Post() As String      '���� ��������� ��������
Folder_Post = ReadINI("SendReport", "MailFolder", PathFileIni)
End Function

'���������� �� ����
Public Function ConnectToDataBase() As String
   Set myADO = New ADODB.Connection
   'myADO.Open "Provider=SQLNCLI;Server=172.198.3.250\SqlExpress;Database=SSPZ;Uid=sa;Pwd=sql13asopd;"
   myADO.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; password=sql13asopd; Initial Catalog=SSPZ;Data Source=172.198.3.250\SQLEXPRESS"
   Set myRS = New ADODB.Recordset

End Function
