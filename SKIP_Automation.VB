Sub automação_skip()

'Abrir o arquivo que foi gerado pela outra automação
Workbooks.Open ("C:\Users\f1p1xq8\Desktop\skip.xlsx")


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

'atribuir valor para alguns contadores
i = 2
j = 2
k = 2
l = 2
m = 2
y = 2
z = 2



'Criar conexão com SAP
Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetScriptingEngine
Set Connection = App.Children(0)
Set session = Connection.Children(0)

If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If

If session.findById("wnd[0]").maximize Then

'Se o SAP não estiver aberto, Abre sistema SAP

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


'Abre a lista de itens que foi gerada pela primeira automação e verifica se tem itens para liberar,
'se tiver continua o código, se não sai da sub

Workbooks("skip.XLSX").Worksheets("Sheet1").Activate
If Sheets("Sheet1").Cells(i, "B") = "" Then
Workbooks("skip.XLSX").Worksheets("Sheet1").Activate
Application.Wait Now + TimeValue("0:00:1")
Workbooks("skip.XLSX").Close SaveChanges:=False
Application.Wait Now + TimeValue("0:00:1")
Workbooks("Automation_Skip.xlsb").Worksheets("Planilha1").Activate
Application.Wait Now + TimeValue("0:00:1")
w.Save
Application.Wait Now + TimeValue("0:00:1")
Application.Quit
Workbooks("Automation_Skip.XLSX").Close SaveChanges:=False
Exit Sub
End If


'Verifica se as linhas tem fornecedor
i = 2
j = 2
While Sheets("Sheet1").Cells(i, "B") <> ""
           If Sheets("Sheet1").Cells(j, "C") = "" Then
            Application.Wait Now + TimeValue("0:00:1")
            Sheets("Sheet1").Rows(j).Select
            Application.Wait Now + TimeValue("0:00:1")
            Selection.Delete
             Application.Wait Now + TimeValue("0:00:1")
             i = i - 1
             j = j - 1
            End If

i = i + 1
j = j + 1

Wend

i = 2
j = 2

CONTINUE:

'Checar unidade de medida do item e colar na tabela


 Workbooks("skip.XLSX").Worksheets("Sheet1").Activate
While Sheets("Sheet1").Cells(i, "B") <> ""
    If Sheets("Sheet1").Cells(i, "B") = "" Then
    GoTo CONTINUE2
    End If
        If Sheets("Sheet1").Cells(i, "B") <> "" Then
            PN = Sheets("Sheet1").Cells(i, "B").Value
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmd04"
            session.findById("wnd[0]").sendVKey 0
            Application.Wait Now + TimeValue("0:00:1")
            session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").Text = PN
            session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").Text = "nw01"
            session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").SetFocus
            session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").caretPosition = 8
            session.findById("wnd[0]/tbar[0]/btn[0]").press
            Application.Wait Now + TimeValue("0:00:1")
            session.findById("wnd[0]/usr/subINCLUDE8XX:SAPMM61R:0800/ctxtRM61R-MEINH").SetFocus
            unidade = session.findById("wnd[0]/usr/subINCLUDE8XX:SAPMM61R:0800/ctxtRM61R-MEINH").Text
            Workbooks("skip.XLSX").Worksheets("Sheet1").Activate
            Sheets("Sheet1").Cells(k, "G").Value = unidade

        End If

i = i + 1
k = k + 1
Wend

i = 2
j = 2
k = 2
l = 2
m = 2

CONTINUE2:

'Colocar o primeiro digito do lote de controle para poder checar o grupo.
Application.Wait Now + TimeValue("0:00:1")
Workbooks("skip.XLSX").Worksheets("Sheet1").Activate
Range("H2").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-3],1)"
    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H500"), Type:=xlFillDefault

'Checar qual o grupo correto e inserir em nova coluna

 Workbooks("skip.XLSX").Worksheets("Sheet1").Activate
While Sheets("Sheet1").Cells(i, "B") <> ""
    If Sheets("Sheet1").Cells(i, "B") = "" Then
    GoTo CONTINUE3
    End If

'Checando entrada e transferencia de PC

        If Sheets("Sheet1").Cells(l, "H") = "1" And Sheets("Sheet1").Cells(k, "G") = "PC" Then
           Sheets("Sheet1").Cells(m, "I") = "215547"

        ElseIf Sheets("Sheet1").Cells(l, "H") = "8" And Sheets("Sheet1").Cells(k, "G") = "PC" Then
           Sheets("Sheet1").Cells(m, "I") = "186800"

'Checando entrada e transferencia de KG

        ElseIf Sheets("Sheet1").Cells(l, "H") = "1" And Sheets("Sheet1").Cells(k, "G") = "KG" Then
           Sheets("Sheet1").Cells(m, "I") = "116060"

        ElseIf Sheets("Sheet1").Cells(l, "H") = "8" And Sheets("Sheet1").Cells(k, "G") = "KG" Then
           Sheets("Sheet1").Cells(m, "I") = "195892"

'Checando entrada e transferencia de GAL

        ElseIf Sheets("Sheet1").Cells(l, "H") = "1" And Sheets("Sheet1").Cells(k, "G") = "GAL" Then
           Sheets("Sheet1").Cells(m, "I") = "185989"

        ElseIf Sheets("Sheet1").Cells(l, "H") = "8" And Sheets("Sheet1").Cells(k, "G") = "GAL" Then
           Sheets("Sheet1").Cells(m, "I") = "81951"

'Checando entrada e transferencia de L

        ElseIf Sheets("Sheet1").Cells(l, "H") = "1" And Sheets("Sheet1").Cells(k, "G") = "L" Then
           Sheets("Sheet1").Cells(m, "I") = "229319"

        ElseIf Sheets("Sheet1").Cells(l, "H") = "8" And Sheets("Sheet1").Cells(k, "G") = "L" Then
           Sheets("Sheet1").Cells(m, "I") = "229320"

'Checando entrada e transferencia de M

        ElseIf Sheets("Sheet1").Cells(l, "H") = "1" And Sheets("Sheet1").Cells(k, "G") = "M" Then
           Sheets("Sheet1").Cells(m, "I") = "81416"

        ElseIf Sheets("Sheet1").Cells(l, "H") = "8" And Sheets("Sheet1").Cells(k, "G") = "M" Then
           Sheets("Sheet1").Cells(m, "I") = "229323"

'Checando entrada e transferencia de M3

        ElseIf Sheets("Sheet1").Cells(l, "H") = "1" And Sheets("Sheet1").Cells(k, "G") = "M3" Then
           Sheets("Sheet1").Cells(m, "I") = "90046"

        ElseIf Sheets("Sheet1").Cells(l, "H") = "8" And Sheets("Sheet1").Cells(k, "G") = "M3" Then
           Sheets("Sheet1").Cells(m, "I") = "229325"

'Checando entrada e transferencia de ML

        ElseIf Sheets("Sheet1").Cells(l, "H") = "1" And Sheets("Sheet1").Cells(k, "G") = "ML" Then
           Sheets("Sheet1").Cells(m, "I") = "229326"

        ElseIf Sheets("Sheet1").Cells(l, "H") = "8" And Sheets("Sheet1").Cells(k, "G") = "ML" Then
           Sheets("Sheet1").Cells(m, "I") = "229327"

'Checando entrada e transferencia de PAC

        ElseIf Sheets("Sheet1").Cells(l, "H") = "1" And Sheets("Sheet1").Cells(k, "G") = "PAC" Then
           Sheets("Sheet1").Cells(m, "I") = "217209"

        ElseIf Sheets("Sheet1").Cells(l, "H") = "8" And Sheets("Sheet1").Cells(k, "G") = "PAC" Then
           Sheets("Sheet1").Cells(m, "I") = "229329"

'Checando entrada e transferencia de ROL

        ElseIf Sheets("Sheet1").Cells(l, "H") = "1" And Sheets("Sheet1").Cells(k, "G") = "ROL" Then
           Sheets("Sheet1").Cells(m, "I") = "229330"

        ElseIf Sheets("Sheet1").Cells(l, "H") = "8" And Sheets("Sheet1").Cells(k, "G") = "ROL" Then
           Sheets("Sheet1").Cells(m, "I") = "229331"

        End If
'Redefinir valores de algumas variaveis
i = i + 1
l = l + 1
k = k + 1
m = m + 1

Wend

CONTINUE3:
'Redefinir valores de algumas variaveis
i = 2
j = 2
k = 2
l = 2
m = 2

'Checar se o item está no grupo errado
 Workbooks("skip.XLSX").Worksheets("Sheet1").Activate
While Sheets("Sheet1").Cells(i, "B") <> ""
If Sheets("Sheet1").Cells(i, "B") = "" Then
GoTo CONTINUE4
End If
If Sheets("Sheet1").Cells(i, "B") <> "" Then
GRUPO = Sheets("Sheet1").Cells(m, "I")

    PN = Sheets("Sheet1").Cells(i, "B").Value
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nQP02"
    session.findById("wnd[0]").sendVKey 0
    Application.Wait Now + TimeValue("0:00:1")
    session.findById("wnd[0]/usr/ctxtRC27M-MATNR").Text = PN
    session.findById("wnd[0]/usr/ctxtRC27M-WERKS").Text = "nw01"
    session.findById("wnd[0]/usr/ctxtRC271-AENNR").Text = "ab0797nwqual"
    session.findById("wnd[0]/usr/ctxtRC271-AENNR").SetFocus
    session.findById("wnd[0]/usr/ctxtRC271-AENNR").caretPosition = 12
    session.findById("wnd[0]/tbar[1]/btn[6]").press
    session.findById("wnd[0]/tbar[1]/btn[6]").press
    Application.Wait Now + TimeValue("0:00:1")
    session.findById("wnd[1]/usr/tblSAPLCPSLTCTRL_2110/ctxtPLKO-PLNNR[0,0]").SetFocus
    If session.findById("wnd[1]/usr/tblSAPLCPSLTCTRL_2110/ctxtPLKO-PLNNR[0,0]").Text <> GRUPO Then
    GoTo EXCLUIR
    End If
EXCLUIR:
    session.findById("wnd[1]/usr/tblSAPLCPSLTCTRL_2110/ctxtPLKO-PLNNR[0,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[2]").press
    session.findById("wnd[0]/tbar[1]/btn[31]").press
    Application.Wait Now + TimeValue("0:00:1")
    VALOR = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,0]").Text
    FORNECEDOR = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,0]").Text
        While VALOR <> ""
        PN = Sheets("Sheet1").Cells(y, "B").Value
        FN = Sheets("Sheet1").Cells(z, "C").Value
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nQP02"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRC27M-MATNR").Text = PN
    session.findById("wnd[0]/usr/ctxtRC27M-WERKS").Text = "nw01"
    session.findById("wnd[0]/usr/ctxtRC271-AENNR").Text = "ab0797nwqual"
    session.findById("wnd[0]/usr/ctxtRC271-AENNR").SetFocus
    session.findById("wnd[0]/usr/ctxtRC271-AENNR").caretPosition = 12
    session.findById("wnd[0]/tbar[1]/btn[6]").press
    If session.findById("wnd[0]/tbar[1]/btn[6]").press = Error Then
    GoTo CONTINUE4
    End If
    Application.Wait Now + TimeValue("0:00:1")
    If session.findById("wnd[1]/usr/tblSAPLCPSLTCTRL_2110/ctxtPLKO-PLNNR[0,0]").SetFocus = Error Then
    GoTo CONTINUE4
    End If
            If VALOR <> PN And FORNECEDOR <> FN Then
                session.findById("wnd[1]/usr/tblSAPLCPSLTCTRL_2110/ctxtPLKO-PLNNR[0,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[2]").press
    If session.findById("wnd[0]/tbar[1]/btn[31]").press = Error Then
                GoTo CONTINUE4
                End If
                session.findById("wnd[1]/usr/tblSAPLCPSLTCTRL_2110/ctxtPLKO-PLNNR[0,0]").SetFocus
                If session.findById("wnd[0]/tbar[1]/btn[31]").press = Error Then
                   PN = Sheets("Sheet1").Cells(y, "B").Value
                    session.findById("wnd[0]").maximize
                    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nQP02"
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]/usr/ctxtRC27M-MATNR").Text = PN
                    session.findById("wnd[0]/usr/ctxtRC27M-WERKS").Text = "nw01"
                    session.findById("wnd[0]/usr/ctxtRC271-AENNR").Text = "ab0797nwqual"
                    session.findById("wnd[0]/usr/ctxtRC271-AENNR").SetFocus
                    session.findById("wnd[0]/usr/ctxtRC271-AENNR").caretPosition = 12
                    session.findById("wnd[0]/tbar[1]/btn[6]").press
                    session.findById("wnd[0]/tbar[1]/btn[6]").press
                    Application.Wait Now + TimeValue("0:00:1")
                    session.findById("wnd[1]/usr/tblSAPLCPSLTCTRL_2110/ctxtPLKO-PLNNR[0,0]").caretPosition = 4
                    session.findById("wnd[1]").sendVKey 2
                    session.findById("wnd[1]/usr/tblSAPLCPSLTCTRL_2110/ctxtPLKO-PLNNR[0,0]").SetFocus
                    session.findById("wnd[0]/tbar[1]/btn[31]").press
                    Application.Wait Now + TimeValue("0:00:1")
                Else:

                session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,0]").SetFocus
                VALOR = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,0]").Text
                            If VALOR = PN Then
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,0]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:2")

                            Else:
VALOR = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,0]").Text
VALOR1 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,1]").Text
VALOR2 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,2]").Text
VALOR3 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,3]").Text
VALOR4 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,4]").Text
VALOR5 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,5]").Text
VALOR6 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,6]").Text
VALOR7 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,7]").Text
VALOR8 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,8]").Text
VALOR9 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,9]").Text
VALOR10 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,10]").Text
VALOR11 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,11]").Text
VALOR12 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,12]").Text
VALOR13 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,13]").Text

FORNECEDOR = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,0]").Text
FORNECEDOR1 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,1]").Text
FORNECEDOR2 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,2]").Text
FORNECEDOR3 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,3]").Text
FORNECEDOR4 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,4]").Text
FORNECEDOR5 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,5]").Text
FORNECEDOR6 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,6]").Text
FORNECEDOR7 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,7]").Text
FORNECEDOR8 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,8]").Text
FORNECEDOR9 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,9]").Text
FORNECEDOR10 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,10]").Text
FORNECEDOR11 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,11]").Text
FORNECEDOR12 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,12]").Text
FORNECEDOR13 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,13]").Text
                            End If


                            While VALOR1 <> ""

                            If LAST = PN And FORNECEDOR <> FN And VALOR <> PN Then
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4


                            ElseIf VALOR = PN And FORNECEDOR = FN Then
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,0]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR = PN And FORNECEDOR <> FN And VALOR1 <> PN Then
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR1 = PN And FORNECEDOR1 = FN Then
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,1]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4


                            ElseIf VALOR1 = PN And FORNECEDOR <> FN And VALOR2 <> PN Then
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4


                            ElseIf VALOR2 = PN And FORNECEDOR2 = FN Then
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,2]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR2 = PN And FORNECEDOR <> FN And VALOR3 <> PN Then
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR3 = PN And FORNECEDOR3 = FN Then
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,3]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR3 = PN And FORNECEDOR <> FN And VALOR4 <> PN Then
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR4 = PN And FORNECEDOR4 = FN Then
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,4]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR4 = PN And FORNECEDOR <> FN And VALOR5 <> PN Then
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR5 = PN And FORNECEDOR5 = FN Then
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,5]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR5 = PN And FORNECEDOR <> FN And VALOR6 <> PN Then
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR6 = PN And FORNECEDOR6 = FN Then
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,6]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR6 = PN And FORNECEDOR <> FN And VALOR7 <> PN Then
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR7 = PN And FORNECEDOR7 = FN Then
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,7]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR7 = PN And FORNECEDOR <> FN And VALOR8 <> PN Then
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR8 = PN And FORNECEDOR8 = FN Then
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,8]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR8 = PN And FORNECEDOR <> FN And VALOR9 <> PN Then
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR9 = PN And FORNECEDOR9 = FN Then
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,9]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR9 = PN And FORNECEDOR <> FN And VALOR10 <> PN Then
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR10 = PN And FORNECEDOR10 = FN Then
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,10]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR10 = PN And FORNECEDOR <> FN And VALOR11 <> PN Then
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR11 = PN And FORNECEDOR11 = FN Then
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,11]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR11 = PN And FORNECEDOR <> FN And VALOR12 <> PN Then
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR12 = PN And FORNECEDOR12 = FN Then
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,12]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR12 = PN And FORNECEDOR <> FN And VALOR12 <> PN Then
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4

                            ElseIf VALOR13 = PN And FORNECEDOR13 = FN Then
                            LAST = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,13]").Text
                            session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,13]").SetFocus
                            session.findById("wnd[1]/tbar[0]/btn[14]").press
                            session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            Application.Wait Now + TimeValue("0:00:1")
                            GoTo CONTINUE4


                            Else:
                            session.findById("wnd[1]").sendVKey 82
VALOR = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,0]").Text
VALOR1 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,1]").Text
VALOR2 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,2]").Text
VALOR3 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,3]").Text
VALOR4 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,4]").Text
VALOR5 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,5]").Text
VALOR6 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,6]").Text
VALOR7 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,7]").Text
VALOR8 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,8]").Text
VALOR9 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,9]").Text
VALOR10 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,10]").Text
VALOR11 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,11]").Text
VALOR12 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,12]").Text
VALOR13 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,13]").Text

FORNECEDOR = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,0]").Text
FORNECEDOR1 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,1]").Text
FORNECEDOR2 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,2]").Text
FORNECEDOR3 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,3]").Text
FORNECEDOR4 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,4]").Text
FORNECEDOR5 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,5]").Text
FORNECEDOR6 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,6]").Text
FORNECEDOR7 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,7]").Text
FORNECEDOR8 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,8]").Text
FORNECEDOR9 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,9]").Text
FORNECEDOR10 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,10]").Text
FORNECEDOR11 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,11]").Text
FORNECEDOR12 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,12]").Text
FORNECEDOR13 = session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[5,13]").Text
                            End If
                            Wend

                    End If

 End If
STEP1:
 session.findById("wnd[0]/tbar[0]/btn[11]").press
 Application.Wait Now + TimeValue("0:00:1")
y = y + 1
z = z + 1


Wend
End If
CONTINUE4:
i = i + 1
m = m + 1

Wend

i = 2
j = 2
m = 2
y = 2
z = 2

'Incluir no Plano correto
INCLUSAO:
Workbooks("skip.XLSX").Worksheets("Sheet1").Activate
While Sheets("Sheet1").Cells(i, "B").Value <> ""
PN = Sheets("Sheet1").Cells(i, "B").Value
GRUPO = Sheets("Sheet1").Cells(m, "I").Value
FORNECEDOR = Sheets("Sheet1").Cells(j, "C").Value
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nQP02"
session.findById("wnd[0]").sendVKey 0
Application.Wait Now + TimeValue("0:00:1")
session.findById("wnd[0]/usr/ctxtRC27M-WERKS").Text = "NW01"
session.findById("wnd[0]/usr/ctxtRC271-PLNNR").Text = GRUPO
session.findById("wnd[0]/usr/ctxtRC271-AENNR").Text = "ab0797nwqual"
session.findById("wnd[0]/usr/ctxtRC271-AENNR").SetFocus
session.findById("wnd[0]/usr/ctxtRC271-AENNR").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[6]").press
session.findById("wnd[0]/tbar[1]/btn[6]").press
session.findById("wnd[0]/tbar[1]/btn[31]").press
Application.Wait Now + TimeValue("0:00:1")

If GRUPO = "186800" Then
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010").verticalScrollbar.Position = 7790
ElseIf GRUPO = "215547" Then
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010").verticalScrollbar.Position = 11800
ElseIf GRUPO = "116060" Then
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010").verticalScrollbar.Position = 39
ElseIf GRUPO = "195892" Then
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010").verticalScrollbar.Position = 69
End If
While session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/txtMAPL-PLNAL[0,1]").Text <> ""
session.findById("wnd[1]").sendVKey 82
Application.Wait Now + TimeValue("0:00:1")
Wend
If GRUPO <> "186800" Then
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/txtMAPL-PLNAL[0,1]").Text = "1"
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,1]").Text = PN
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-WERKS[3,1]").Text = "nw01"
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-LIFNR[5,1]").Text = FORNECEDOR
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-LIFNR[5,1]").SetFocus
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-LIFNR[5,1]").caretPosition = 0

ElseIf GRUPO = "186800" Then
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/txtMAPL-PLNAL[0,1]").Text = "3"
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-MATNR[2,1]").Text = PN
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-WERKS[3,1]").Text = "nw01"
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-LIFNR[5,1]").Text = FORNECEDOR
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-LIFNR[5,1]").SetFocus
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/ctxtMAPL-LIFNR[5,1]").caretPosition = 0
End If
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/txtMAPL-PLNAL[0,0]").SetFocus
session.findById("wnd[1]/usr/tblSAPLCZDITCTRL_4010/txtMAPL-PLNAL[0,0]").caretPosition = 1
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
If session.findById("wnd[1]").sendVKey Is Error Then
GoTo LK

Application.Wait Now + TimeValue("0:00:1")


End If
LK:
i = i + 1
m = m + 1
j = j + 1

Wend

i = 2

'Liberar efetivamente o SKIP
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nqa32"
session.findById("wnd[0]").sendVKey 0
Application.Wait Now + TimeValue("0:00:1")
session.findById("wnd[0]/usr/ctxtQL_ENSTD-LOW").Text = ""
session.findById("wnd[0]/usr/txtQL_MAX_R").Text = ""
session.findById("wnd[0]/usr/ctxtVARIANT").Text = "/skip"
session.findById("wnd[0]/usr/ctxtVARIANT").SetFocus
session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 5
session.findById("wnd[0]/tbar[1]/btn[8]").press
Application.Wait Now + TimeValue("0:00:10")
Workbooks("skip.XLSX").Worksheets("Sheet1").Activate
While Sheets("Sheet1").Cells(i, "B").Value <> ""
Workbooks("skip.XLSX").Worksheets("Sheet1").Activate
PN = Sheets("Sheet1").Cells(i, "B").Value
Application.Wait Now + TimeValue("0:00:1")
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = ""
If session.findById("wnd[0]/tbar[1]/btn[42]").press <> Error Then
Application.Wait Now + TimeValue("0:00:1")
If session.findById("wnd[0]/usr/tabsTABSTRIP/tabpVORG").Select = Error Then
Application.Wait Now + TimeValue("0:00:1")
Dim xWB As Workbook
Application.ScreenUpdating = False
For Each xWB In Application.Workbooks
    If Not (xWB Is Application.ActiveWorkbook) Then
        xWB.Close
    End If
Next
Application.ScreenUpdating = True
Application.Quit
Exit Sub
End If


session.findById("wnd[0]/usr/tabsTABSTRIP/tabpVORG/ssubSUBSCR_TABSTRIP:SAPLQPL1:0202/btnVORGSEL").press  'clica em plano de especificação
session.findById("wnd[1]/usr/sub:SAPLCPSE:2110/ctxtPLKO-PLNNR[0,0]").SetFocus
session.findById("wnd[1]/usr/sub:SAPLCPSE:2110/ctxtPLKO-PLNNR[0,0]").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/sub:SAPLCPSE:2110/ctxtPLKO-PLNNR[1,0]").SetFocus
session.findById("wnd[1]/usr/sub:SAPLCPSE:2110/ctxtPLKO-PLNNR[1,0]").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[2]").press
Application.Wait Now + TimeValue("0:00:1")


session.findById("wnd[0]/usr/tabsTABSTRIP/tabpSTIP").Select 'clica em amostra
session.findById("wnd[0]/usr/tabsTABSTRIP/tabpSTIP/ssubSUBSCR_TABSTRIP:SAPLQPL1:0203/btnDYNA").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[1]/btn[14]").press
Application.Wait Now + TimeValue("0:00:4")

End If
i = i + 1
Application.Wait Now + TimeValue("0:00:1")
Wend
On Error Resume Next

Workbooks("skip.XLSX").Close SaveChanges:=False

Application.Wait Now + TimeValue("0:00:10")
Workbooks("skip.XLSX").Close SaveChanges:=False
Application.ScreenUpdating = False
For Each xWB In Application.Workbooks
    If Not (xWB Is Application.ActiveWorkbook) Then
        xWB.Close
    End If
Next
Application.ScreenUpdating = True
Application.Quit



End Sub

