Sub automação_skip()
'Dimensionar as variaveis
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer
Dim y As Integer
Dim z As Integer
On Error Resume Next
Dim App As Object
Dim Connection As Object
Dim session As Object
Dim SapGuiAut As Object
Dim WScript As Object
Dim arq_qrip As Object
Dim fso
Dim file As String
Dim arquivo As String

'Atribuir valor para algumas variaveis
i = 2
j = 2
k = 2
l = 2
m = 2
y = 2
z = 2



'Criar integração com SAP
Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetScriptingEngine
Set Connection = App.Children(0)
Set session = Connection.Children(0)

If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If

If session.findById("wnd[0]").maximize Then

'Abre sistema SAP

    Dim SapGUI, Appl, WshShell, proc


    Set WshShell = CreateObject("WScript.Shell")

    Set proc = WshShell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")

    Application.Wait Now + TimeValue("0:00:5")
    Set SapGUI = GetObject("SAPGUI")

    Set Appl = SapGUI.GetScriptingEngine

    Set Connection = Appl.OpenConnection("PAG - North American AG Production (SSO)", True)

End If


Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetScriptingEngine
Set Connection = App.Children(0)
Set session = Connection.Children(0)

If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If


'Entrar na QA32 e baixa a planilha de SKIP, salva em um diretório padrão para a segunda parte da automação utilizar.

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nqa32"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtQL_ENSTD-LOW").Text = ""
session.findById("wnd[0]/usr/txtQL_MAX_R").Text = ""
session.findById("wnd[0]/usr/ctxtVARIANT").Text = "/skip"
session.findById("wnd[0]/usr/ctxtVARIANT").SetFocus
session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 5
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\f1p1xq8\Desktop"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "skip.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[11]").press
Application.Wait Now + TimeValue("0:00:5")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nqa32"
session.findById("wnd[0]").sendVKey 0

Workbooks("skip.XLSX").Worksheets("Sheet1").Activate
Application.Wait Now + TimeValue("0:00:1")
w.Save
Application.Wait Now + TimeValue("0:00:1")
Application.Quit


End Sub