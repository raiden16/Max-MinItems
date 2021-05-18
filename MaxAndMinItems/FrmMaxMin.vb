Imports System.Drawing

Public Class FrmMaxMin

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private csFormUID As String
    Private stDocNum As String
    Friend Monto As Double

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Me.stDocNum = stDocNum
    End Sub

    'Private Property stRuta As String

    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function openForm(ByVal psDirectory As String)
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim WhsCode As String
        'Dim Monto As Integer

        Try

            oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            csFormUID = "tekMaxMin"
            '//CARGA LA FORMA
            If (loadFormXML(cSBOApplication, csFormUID, psDirectory + "\Forms\" + csFormUID + ".srf") <> 0) Then

                Err.Raise(-1, 1, "")
            End If
            '"que pedo"
            '--- Referencia de Forma
            setForm(csFormUID)

            WhsCode = AgregarLineas()

            '---- refresca forma
            coForm.Refresh()
            coForm.Visible = True

            coForm = cSBOApplication.Forms.Item("tekMaxMin")

            'coForm.Items.Item("3").Enabled = False
            coForm.Items.Item("2").Enabled = False

            Return WhsCode

        Catch ex As Exception
            If (ex.Message <> "") Then
                cSBOApplication.MessageBox("FrmtekDel. No se pudo iniciar la forma. " & ex.Message)
            End If
            Me.close()
        End Try
    End Function


    '//----- CIERRA LA VENTANA
    Public Function close() As Integer
        close = 0
        coForm.Close()
    End Function


    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function setForm(ByVal psFormUID As String) As Integer
        Try
            setForm = 0
            '//ESTABLECE LA REFERENCIA A LA FORMA
            coForm = cSBOApplication.Forms.Item(psFormUID)
            '//OBTIENE LA REFERENCIA A LOS USER DATA SOURCES
            setForm = getUserDataSources()
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmtekDel. Al referenciar a la forma. " & ex.Message)
            setForm = -1
        End Try
    End Function


    '//----- OBTIENE LA REFERENCIA A LOS USERDATASOURCES
    Private Function getUserDataSources() As Integer
        'Dim llIndice As Integer
        Try
            coForm.Freeze(True)
            getUserDataSources = 0
            '//SI YA EXISTEN LOS DATASOURCES, SOLO LOS ASOCIA
            If (coForm.DataSources.UserDataSources.Count() > 0) Then
            Else '//EN CASO DE QUE NO EXISTAN, LOS CREA
                getUserDataSources = bindUserDataSources()
            End If
            coForm.Freeze(False)
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmtekDel. Al referenciar los UserDataSources" & ex.Message)
            getUserDataSources = -1
        End Try
    End Function


    '//----- ASOCIA LOS USERDATA A ITEMS
    Private Function bindUserDataSources() As Integer
        Dim loText As SAPbouiCOM.EditText
        Dim loText2 As SAPbouiCOM.StaticText
        Dim loDS As SAPbouiCOM.UserDataSource
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim oGrid As SAPbouiCOM.Grid
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset

        Try

            oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            bindUserDataSources = 0

            stQueryH = "Select T0.""DfltsGroup"" from OUSR T0 where T0.""USER_CODE""='" & cSBOCompany.UserName & "'"
            oRecSetH.DoQuery(stQueryH)

            loDS = coForm.DataSources.UserDataSources.Add("dsHouse", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
            loText2 = coForm.Items.Item("2").Specific  'identifico mi caja de texto
            loText2.Caption = oRecSetH.Fields.Item("DfltsGroup").Value

            oGrid = coForm.Items.Item("3").Specific
            oDataTable = coForm.DataSources.DataTables.Add("Stock")
            oGrid.DataTable = oDataTable

        Catch ex As Exception
            cSBOApplication.MessageBox("FrmtekDel. Al crear los UserDataSources. " & ex.Message)
            bindUserDataSources = -1
        Finally
            loText = Nothing
            loDS = Nothing
            oDataTable = Nothing
            oGrid = Nothing
        End Try
    End Function


    '----- carga los procesos de carga
    Public Function AgregarLineas()
        Dim oGrid As SAPbouiCOM.Grid
        Dim stQuery, stQuery2 As String
        Dim oRecSet, oRecSet2 As SAPbobsCOM.Recordset
        Dim oCombo As SAPbouiCOM.ComboBoxColumn
        Dim WhsCode, User As String

        Try

            oGrid = coForm.Items.Item("3").Specific
            oGrid.DataTable.Clear()

            User = cSBOCompany.UserName
            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuery = "Select T1.""Warehouse"" as ""Stock"" from OUSR T0 Inner Join OUDG T1 on T0.""DfltsGroup""=T1.""Code"" where T0.""USER_CODE""='" & User & "'"
            oRecSet.DoQuery(stQuery)
            WhsCode = oRecSet.Fields.Item("Stock").Value

            oRecSet2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuery2 = "CALL CONSULTA_MAXANDMIN_ITEMS('" & WhsCode & "')"
            oGrid.DataTable.ExecuteQuery(stQuery2)

            For numfila As Integer = 0 To oGrid.Rows.Count - 1
                Dim valorFila As Integer = oGrid.GetDataTableRowIndex(numfila)
                If (valorFila <> -1) Then
                    If (oGrid.DataTable.GetValue("Stock", valorFila) < oGrid.DataTable.GetValue("Minimo de Stock", valorFila)) Then
                        oGrid.CommonSetting.SetCellBackColor(numfila + 1, 3, ColorTranslator.ToOle(Color.Red))
                    End If
                End If
            Next

            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(8).Editable = False

            Return WhsCode

        Catch ex As Exception

            MsgBox("FrmtekDel. fallo la carga previa de la forma AgregarLineas: " & ex.Message)

        Finally

            oGrid = Nothing

        End Try

    End Function

End Class
