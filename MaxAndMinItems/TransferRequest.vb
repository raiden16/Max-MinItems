Public Class TransferRequest

    Private SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private SBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Dim oTransferRequest As SAPbobsCOM.StockTransfer
    Dim Duplicadas, Registradas, NExist As Integer

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        SBOApplication = oCatchingEvents.SBOApplication
        SBOCompany = oCatchingEvents.SBOCompany
    End Sub

    Public Sub addTransferRequest(ByVal FormUID As String, ByVal csDirectory As String, ByVal WhsCode As String)

        Dim oFrmMaxMin As FrmMaxMin
        Dim coForm As SAPbouiCOM.Form
        Dim oGrid As SAPbouiCOM.Grid
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim stQueryH, stQueryH2, stQueryH3 As String
        Dim oRecSetH, oRecSetH2, oRecSetH3 As SAPbobsCOM.Recordset
        'Dim DocNum, CardCode, DocCur, ItemCode, Quantity, Price, DiscPrcnt, TaxCode, WhsCode, Currency, ObjType, LineNum As String
        Dim TotalRequerido, Stock, Minimo, Resto, Requerido, Cajas As Double
        Dim ItemsStock As Integer
        Dim llError As Long
        Dim lsError As String
        Dim FromWhsCode As String

        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH3 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTransferRequest = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest)

        Try

            coForm = SBOApplication.Forms.Item(FormUID)
            oGrid = coForm.Items.Item("3").Specific
            oDataTable = oGrid.DataTable
            FromWhsCode = coForm.DataSources.UserDataSources.Item("dsSucursaO").Value


            'Revisar si hay articulos a surtir

            For i = 0 To oDataTable.Rows.Count - 1

                If oDataTable.GetValue("Solicitud Sugerida M2", i) > 0 Or oDataTable.GetValue("Solicitud Extra M2", i) > 0 Then

                    ItemsStock = ItemsStock + 1

                End If

            Next

            'Si hay articulos a surtir generar Traslado

            If ItemsStock > 0 Then

                oTransferRequest.DocDate = DateTime.Now.ToString("dd/MM/yyyy")
                oTransferRequest.DueDate = DateTime.Now.ToString("dd/MM/yyyy")
                oTransferRequest.ToWarehouse = WhsCode

                If WhsCode <> "001" Then

                    oTransferRequest.FromWarehouse = "001"

                Else

                    oTransferRequest.FromWarehouse = FromWhsCode

                End If

                For i = 0 To oDataTable.Rows.Count - 1

                    If oDataTable.GetValue("Solicitud Sugerida M2", i) > 0 Or oDataTable.GetValue("Solicitud Extra M2", i) > 0 Then

                        TotalRequerido = oDataTable.GetValue("Solicitud Sugerida M2", i) + oDataTable.GetValue("Solicitud Extra M2", i)

                        If WhsCode <> "001" Then

                            stQueryH = "CALL CONSULTA_STOCK_VALLEJO('" & oDataTable.GetValue("Artículo", i) & "')"
                            oRecSetH.DoQuery(stQueryH)

                        Else

                            stQueryH = "CALL CONSULTA_STOCK('" & oDataTable.GetValue("Artículo", i) & "','" & FromWhsCode & "')"
                            oRecSetH.DoQuery(stQueryH)

                        End If

                        If oRecSetH.RecordCount > 0 Then

                            oRecSetH.MoveFirst()

                            Stock = oRecSetH.Fields.Item("Stock Vallejo").Value
                            Minimo = oRecSetH.Fields.Item("Minimo de Stock Vallejo").Value
                            Resto = Stock - Minimo

                            If Resto > 0 And Resto > TotalRequerido Then

                                stQueryH2 = "CALL CONSULTA_REQUERIDO(" & TotalRequerido & ",'" & oDataTable.GetValue("Artículo", i) & "')"
                                oRecSetH2.DoQuery(stQueryH2)

                                If oRecSetH2.RecordCount > 0 Then

                                    oRecSetH2.MoveFirst()

                                    Requerido = oRecSetH2.Fields.Item("Requerido").Value
                                    Cajas = oRecSetH2.Fields.Item("Cajas").Value

                                    oTransferRequest.Lines.ItemCode = oDataTable.GetValue("Artículo", i)
                                    oTransferRequest.Lines.Quantity = Requerido
                                    oTransferRequest.Lines.WarehouseCode = WhsCode
                                    oTransferRequest.Lines.UserFields.Fields.Item("U_Requerido").Value = Requerido
                                    oTransferRequest.Lines.UserFields.Fields.Item("U_CajasReq").Value = Cajas

                                    If WhsCode <> "001" Then

                                        oTransferRequest.Lines.FromWarehouseCode = "001"

                                    Else

                                        oTransferRequest.Lines.FromWarehouseCode = FromWhsCode

                                    End If

                                    oTransferRequest.Lines.Add()

                                End If

                            ElseIf Resto > 0 And Resto < TotalRequerido Then

                                stQueryH3 = "CALL CONSULTA_REQUERIDO_DOWN(" & Resto & ",'" & oDataTable.GetValue("Artículo", i) & "')"
                                oRecSetH3.DoQuery(stQueryH3)

                                If oRecSetH3.RecordCount > 0 Then

                                    oRecSetH3.MoveFirst()

                                    Requerido = oRecSetH3.Fields.Item("Requerido").Value
                                    Cajas = oRecSetH3.Fields.Item("Cajas").Value

                                    oTransferRequest.Lines.ItemCode = oDataTable.GetValue("Artículo", i)
                                    oTransferRequest.Lines.Quantity = Requerido
                                    oTransferRequest.Lines.WarehouseCode = WhsCode
                                    oTransferRequest.Lines.UserFields.Fields.Item("U_Requerido").Value = Requerido
                                    oTransferRequest.Lines.UserFields.Fields.Item("U_CajasReq").Value = Cajas

                                    If WhsCode <> "001" Then

                                        oTransferRequest.Lines.FromWarehouseCode = "001"

                                    Else

                                        oTransferRequest.Lines.FromWarehouseCode = FromWhsCode

                                    End If

                                    oTransferRequest.Lines.Add()

                                End If

                            End If

                        End If

                    End If

                Next

                If oTransferRequest.Add() <> 0 Then

                        SBOCompany.GetLastError(llError, lsError)
                        Err.Raise(-1, 1, lsError)

                    Else

                        stQueryH3 = "Select ""DocNum"" from OWTQ where ""DocEntry""=" & SBOCompany.GetNewObjectKey()
                        oRecSetH3.DoQuery(stQueryH3)

                        If oRecSetH3.RecordCount > 0 Then

                            SBOApplication.MessageBox("La solicitud de traslado se creo con éxito con el número " & oRecSetH3.Fields.Item("DocNum").Value & ".")
                            coForm.Close()

                        End If

                    End If

                End If


        Catch ex As Exception

            SBOApplication.MessageBox("Error en el evento sobre Agregar facturas de anticpo. " & ex.Message)

        Finally

        End Try

    End Sub

End Class
