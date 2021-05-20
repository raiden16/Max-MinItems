Imports System.Drawing

Public Class FrmMaxMin

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private csFormUID As String
    Private stDocNum As String
    Friend Monto As Double
    Friend sucursal As String

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

            stQueryH = "Select T0.""DfltsGroup"" from OUSR T0 where T0.""USER_CODE""='" & cSBOCompany.UserName & "'"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                Sucursal = oRecSetH.Fields.Item("DfltsGroup").Value

            End If

            If sucursal = "VALLEJO" Then

                csFormUID = "tekMaxMinValle"

            Else

                csFormUID = "tekMaxMin"

            End If

            '//CARGA LA FORMA
            If (loadFormXML(cSBOApplication, csFormUID, psDirectory + "\Forms\" + csFormUID + ".srf") <> 0) Then

                Err.Raise(-1, 1, "")

            End If
            '"que pedo"
            '--- Referencia de Forma
            setForm(csFormUID)

            If sucursal = "VALLEJO" Then

                cargarCombos()

            End If

            If sucursal <> "VALLEJO" Then

                WhsCode = AgregarLineas("tekMaxMin")

            Else

                WhsCode = "VALLEJO"

            End If

            '---- refresca forma
            coForm.Refresh()
            coForm.Visible = True

            If sucursal = "VALLEJO" Then

                coForm = cSBOApplication.Forms.Item("tekMaxMinValle")

            Else

                coForm = cSBOApplication.Forms.Item("tekMaxMin")

            End If

            'coForm.Items.Item("3").Enabled = False
            If sucursal <> "VALLEJO" Then

                coForm.Items.Item("2").Enabled = False

            End If

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
        Dim oCombo, oCombo2 As SAPbouiCOM.ComboBox

        Try

            oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            bindUserDataSources = 0

            stQueryH = "Select T0.""DfltsGroup"" from OUSR T0 where T0.""USER_CODE""='" & cSBOCompany.UserName & "'"
            oRecSetH.DoQuery(stQueryH)

            If sucursal <> "VALLEJO" Then

                loDS = coForm.DataSources.UserDataSources.Add("dsHouse", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
                loText2 = coForm.Items.Item("2").Specific  'identifico mi caja de texto
                loText2.Caption = oRecSetH.Fields.Item("DfltsGroup").Value

            End If

            oGrid = coForm.Items.Item("3").Specific
            oDataTable = coForm.DataSources.DataTables.Add("Stock")
            oGrid.DataTable = oDataTable

            If sucursal = "VALLEJO" Then

                loDS = coForm.DataSources.UserDataSources.Add("dsSucursal", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
                oCombo = coForm.Items.Item("7").Specific  'identifico mi combobox
                oCombo.DataBind.SetBound(True, "", "dsSucursal")   ' uno mi userdatasources a mi combobox

                loDS = coForm.DataSources.UserDataSources.Add("dsSucursaO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
                oCombo2 = coForm.Items.Item("8").Specific  'identifico mi combobox
                oCombo2.DataBind.SetBound(True, "", "dsSucursaO")   ' uno mi userdatasources a mi combobox

            End If

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


    Public Function cargarCombos()

        Dim oCombo, oCombo2 As SAPbouiCOM.ComboBox
        Dim oRecSet As SAPbobsCOM.Recordset

        Try
            cargarCombos = 0
            '--- referencia de combo 
            oCombo = coForm.Items.Item("7").Specific
            oCombo2 = coForm.Items.Item("8").Specific
            coForm.Freeze(True)
            '---- SI YA SE TIENEN VALORES, SE ELIMMINAN DEL COMBO
            If oCombo.ValidValues.Count > 0 Then
                Do
                    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                Loop While oCombo.ValidValues.Count > 0
            End If
            If oCombo2.ValidValues.Count > 0 Then
                Do
                    oCombo2.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                Loop While oCombo2.ValidValues.Count > 0
            End If
            '--- realizar consulta
            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSet.DoQuery("Select ""WhsCode"",""WhsName"" from OWHS order by 1")
            '---- cargamos resultado
            oRecSet.MoveFirst()
            Do While oRecSet.EoF = False
                oCombo.ValidValues.Add(oRecSet.Fields.Item(0).Value, oRecSet.Fields.Item(1).Value)
                oCombo2.ValidValues.Add(oRecSet.Fields.Item(0).Value, oRecSet.Fields.Item(1).Value)
                oRecSet.MoveNext()
            Loop
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oCombo2.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            coForm.Freeze(False)


        Catch ex As Exception
            coForm.Freeze(False)
            MsgBox("FrmtekDel. Fallo la carga previa del comboBox cargarComboChofi: " & ex.Message)
        Finally
            oCombo = Nothing
            oCombo2 = Nothing
            oRecSet = Nothing
        End Try
    End Function


    '----- carga los procesos de carga
    Public Function AgregarLineas(ByVal FormUID As String)
        Dim oGrid As SAPbouiCOM.Grid
        Dim stQuery, stQuery2 As String
        Dim oRecSet, oRecSet2 As SAPbobsCOM.Recordset
        Dim oCombo As SAPbouiCOM.ComboBoxColumn
        Dim WhsCode, User As String

        Try

            coForm = cSBOApplication.Forms.Item(FormUID)
            oGrid = coForm.Items.Item("3").Specific
            oGrid.DataTable.Clear()

            User = cSBOCompany.UserName
            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If FormUID = "tekMaxMin" Then

                stQuery = "Select T1.""Warehouse"" as ""Stock"" from OUSR T0 Inner Join OUDG T1 on T0.""DfltsGroup""=T1.""Code"" where T0.""USER_CODE""='" & User & "'"
                oRecSet.DoQuery(stQuery)
                WhsCode = oRecSet.Fields.Item("Stock").Value

                oRecSet2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                stQuery2 = "CALL CONSULTA_MAXANDMIN_ITEMS('" & WhsCode & "')"

            Else

                WhsCode = coForm.DataSources.UserDataSources.Item("dsSucursal").Value

                oRecSet2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                stQuery2 = "CALL CONSULTA_MAXANDMIN_ITEMSVALLE('" & coForm.DataSources.UserDataSources.Item("dsSucursal").Value & "','" & coForm.DataSources.UserDataSources.Item("dsSucursaO").Value & "')"
                oRecSet2.DoQuery(stQuery2)

            End If

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
