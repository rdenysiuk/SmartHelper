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

'ОБЛАСНИЙ ФТП =========================================
Public Function Ftp_Obl() As String      ' Обласний ФТП
Ftp_Obl = ReadINI("FtpConnect", "FtpObl", PathFileIni)
End Function
Public Function Kat_Obl() As String      ' Каталог ФТП обласний
Kat_Obl = ReadINI("FtpConnect", "OblKat", PathFileIni)
End Function
Public Function Login_Obl() As String      ' логін на ФТП обласний
Login_Obl = ReadINI("FtpConnect", "LogObl", PathFileIni)
End Function
Public Function Pass_Obl() As String       ' пароль на ФТП обласний
Pass_Obl = ReadINI("FtpConnect", "PassObl", PathFileIni)
End Function
'======================================================


'КИЇВСЬКИЙ ФТП ++++++++++++++++++++++++++++++++++++++++
Public Function Ftp_Kiev() As String      ' Обласний ФТП
Ftp_Kiev = ReadINI("FtpConnect", "FtpKiev", PathFileIni)
End Function
Public Function Kat_Kiev() As String      ' Каталог ФТП обласний
Kat_Kiev = ReadINI("FtpConnect", "kievKat", PathFileIni)
End Function
Public Function Login_Kiev() As String     ' логін на ФТП київський
Login_Kiev = ReadINI("ftpconnect", "logkiev", PathFileIni)
End Function
Public Function Pass_Kiev() As String      ' пароль на ФТП київський
Pass_Kiev = ReadINI("ftpconnect", "passkiev", PathFileIni)
End Function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++


Public Function Folder_DB_Raj() As String      'ШЛЯХ ДО БД РАЙОНІВ
Folder_DB_Raj = ReadINI("SendReport", "PathRpr", PathFileIni)
End Function

Public Function Folder_Post() As String      'ШЛЯХ ПОШТОВОГО КАТАЛОГУ
Folder_Post = ReadINI("SendReport", "MailFolder", PathFileIni)
End Function

'підключення до бази
Public Function ConnectToDataBase() As String
   Set myADO = New ADODB.Connection
   'myADO.Open "Provider=SQLNCLI;Server=172.198.3.250\SqlExpress;Database=SSPZ;Uid=sa;Pwd=sql13asopd;"
   myADO.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; password=sql13asopd; Initial Catalog=SSPZ;Data Source=172.198.3.250\SQLEXPRESS"
   Set myRS = New ADODB.Recordset

End Function
