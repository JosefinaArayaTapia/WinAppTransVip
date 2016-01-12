Imports SAPbobsCOM
Imports System.Data.SqlClient

Public Class Form1

    Public oCompañia As SAPbobsCOM.Company

    Public Sap As SAPbouiCOM.SboGuiApi
    Public AppSap As SAPbouiCOM.Application

    Dim iError As Integer
    Dim sError As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If ConectaUIAPI() Then
            Me.Text = oCompañia.CompanyName
        Else

            If ConectaB1() = True Then
                Me.Text = oCompañia.CompanyName
            Else
                End
            End If
        End If

    End Sub

    Public Function ConectaUIAPI() As Boolean
        Dim strConnect As String
        Dim sErr As Long
        Dim msg As String = ""
        Try
            Dim NumOfParams As Integer
            NumOfParams = Environment.GetCommandLineArgs.Length
            If NumOfParams = 2 Then
                Try
                    strConnect = Environment.GetCommandLineArgs.GetValue(1)
                Catch ex As Exception
                    MessageBox.Show("No se Pudo Rescatar Argumento de Sesión de B1", "de Addon", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                End Try
            Else
                strConnect = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
            End If
            Sap = New SAPbouiCOM.SboGuiApi
            Sap.Connect(strConnect)
            oCompañia = New SAPbobsCOM.Company
            sErr = ConectarUI()
            If sErr <> 0 Then
                oCompañia.GetLastError(sErr, msg)
                MessageBox.Show(msg, "Error al Conectar UIAPI", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function ConectarUI() As Long
        Try
            Dim strCook As String, strCnn As String
            Dim Valor As Long

            AppSap = Sap.GetApplication
            strCook = oCompañia.GetContextCookie
            strCnn = AppSap.Company.GetConnectionContext(strCook)
            Valor = oCompañia.SetSboLoginContext(strCnn)
            Valor = oCompañia.Connect
            Return Valor

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return -1
        End Try

    End Function


    Public Function ConectaB1() As Boolean
        Try
            oCompañia = New SAPbobsCOM.Company
            oCompañia.Server = "192.168.9.102"
            oCompañia.CompanyDB = "ZBO_Prueba_Transvip"
            oCompañia.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            oCompañia.UserName = "jaat"
            oCompañia.Password = "kokoliko"
            oCompañia.DbUserName = "sa"
            oCompañia.DbPassword = "Sa123456"
            oCompañia.UseTrusted = False
            iError = oCompañia.Connect
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            End If
            Return True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
            Return False
        End Try
    End Function


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim oSocioNegocio As SAPbobsCOM.BusinessPartners
        MessageBox.Show("Dot Net Perls is awesome.")
        Try
            oSocioNegocio = oCompañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oSocioNegocio.CardCode = txtcardcode.Text
            oSocioNegocio.CardName = "Joe Garcia"
            oSocioNegocio.FederalTaxID = "1-1"
            ' oSocioNegocio.UserFields.Fields.Item("U_CVIP").Value = "Hola mundo"

            iError = oSocioNegocio.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim oSocioNegocio As SAPbobsCOM.BusinessPartners
        Try
            oSocioNegocio = oCompañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            If oSocioNegocio.GetByKey("joe") = True Then
                'oSocioNegocio.CardForeignName = Date.Now.ToString
                oSocioNegocio.FederalTaxID = "1-1"
                iError = oSocioNegocio.Update
                If iError <> 0 Then
                    oCompañia.GetLastError(iError, sError)
                    Throw New Exception(sError)
                End If
            End If



        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim oSocioNegocio As SAPbobsCOM.BusinessPartners
        Try
            oSocioNegocio = oCompañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            If oSocioNegocio.GetByKey("joe") = True Then

                iError = oSocioNegocio.Remove
                If iError <> 0 Then
                    oCompañia.GetLastError(iError, sError)
                    Throw New Exception(sError)
                End If
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim oRegistros As SAPbobsCOM.Recordset
        Try
            oRegistros = oCompañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRegistros.DoQuery("exec dbo.sp_transvip ")
            If oRegistros.RecordCount > 0 Then
                For i = 0 To oRegistros.RecordCount - 1
                    MessageBox.Show(oRegistros.Fields.Item("cardcode").Value & " " & oRegistros.Fields.Item("cardname").Value)
                    oRegistros.MoveNext()
                Next
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    ''''''''''''''''''''''''''''''CREA ORDEN DE VENTA'''''''''''''''''''''
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim oDoc As SAPbobsCOM.Documents
        Try
            oDoc = oCompañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
            'oDoc.Series = 4
            oDoc.CardCode = "10014146-9C"
            oDoc.DocDueDate = Date.Now.AddDays(1)
            oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

            oDoc.Lines.SetCurrentLine(0)
            oDoc.Lines.ItemDescription = "123456789"
            oDoc.Lines.AccountCode = "_SYS00000000008" 'efectivo pesos
            oDoc.Lines.UnitPrice = 44200
            oDoc.Lines.TaxCode = "IVA_EXE"

            'AGregar Lineas Al Documento

            'oDoc.Lines.Add()
            'oDoc.Lines.SetCurrentLine(1)
            'oDoc.Lines.ItemDescription = "chao mundo"
            'oDoc.Lines.AccountCode = "1-1-010-10-001"
            'oDoc.Lines.UnitPrice = 2000
            'oDoc.Lines.TaxCode = "IVA"

            iError = oDoc.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                MessageBox.Show(oCompañia.GetNewObjectKey)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim odoc As SAPbobsCOM.Documents
        Try
            odoc = oCompañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
            odoc.CardCode = "joe"
            odoc.DocDueDate = Date.Now.AddDays(1)
            odoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

            odoc.Lines.SetCurrentLine(0)
            odoc.Lines.BaseType = 17 'basado en orden de venta
            odoc.Lines.BaseEntry = 271 'docentry de la tabla ordr
            odoc.Lines.BaseLine = 0

            iError = odoc.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                MessageBox.Show(oCompañia.GetNewObjectKey)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim oDoc As SAPbobsCOM.Documents
        Try
            oDoc = oCompañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
            oDoc.Series = 56
            oDoc.CardCode = "joe"
            oDoc.DocDueDate = Date.Now
            oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items

            oDoc.Lines.SetCurrentLine(0)
            oDoc.Lines.ItemCode = "TR0001"
            oDoc.Lines.UnitPrice = 7000
            oDoc.Lines.TaxCode = "IVA_EXE"

            iError = oDoc.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                MessageBox.Show(oCompañia.GetNewObjectKey)
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim odoc As SAPbobsCOM.Documents
        Try
            odoc = oCompañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            odoc.CardCode = "joe"
            odoc.Indicator = "35"
            'odoc.DocDueDate = Date.Now.AddDays(1)
            odoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

            odoc.Lines.SetCurrentLine(0)
            odoc.Lines.BaseType = 15 'basado en la entrega 1
            odoc.Lines.BaseEntry = 251 'docentry de la tabla odln
            odoc.Lines.BaseLine = 0

            iError = odoc.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                MessageBox.Show(oCompañia.GetNewObjectKey)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim oDoc As SAPbobsCOM.Documents
        Try
            oDoc = oCompañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            oDoc.CardCode = "joe"
            oDoc.DocDueDate = Date.Now
            oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

            oDoc.Lines.SetCurrentLine(0)
            oDoc.Lines.ItemDescription = "servicio 1"
            oDoc.Lines.AccountCode = "1-1-010-10-002" 'columna acctcode de oact
            oDoc.Lines.UnitPrice = 1000
            oDoc.Lines.Currency = "CLP"
            oDoc.Lines.TaxCode = "IVA"

            iError = oDoc.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                MessageBox.Show(oCompañia.GetNewObjectKey)
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim oAsiento As SAPbobsCOM.JournalEntries


        Try
            oAsiento = oCompañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            oAsiento.Memo = "Creado desde la DI API"

            oAsiento.Lines.SetCurrentLine(0)
            oAsiento.Lines.AccountCode = "_SYS00000000008" 'Efectivo Pesos
            oAsiento.Lines.Debit = 15803540
            oAsiento.Lines.Add()

            oAsiento.Lines.SetCurrentLine(1)
            oAsiento.Lines.AccountCode = "_SYS00000000009" 'EFECTIVO MONEDA EXTRANJERA
            oAsiento.Lines.Debit = 0
            oAsiento.Lines.Add()


            oAsiento.Lines.SetCurrentLine(2)
            oAsiento.Lines.AccountCode = "_SYS00000000010" 'CHEQUES EN TRANSITO
            oAsiento.Lines.Debit = 69000
            oAsiento.Lines.Add()


            oAsiento.Lines.SetCurrentLine(3)
            oAsiento.Lines.AccountCode = "_SYS00000000332" 'Ingresos por Transportes
            oAsiento.Lines.Credit = 15803540 + 69000

            iError = oAsiento.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                MessageBox.Show(oCompañia.GetNewObjectKey)
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim oAsiento As SAPbobsCOM.JournalEntries
        Try
            oAsiento = oCompañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            oAsiento.Memo = "creado desde di api"

            oAsiento.Lines.SetCurrentLine(0)
            oAsiento.Lines.AccountCode = "1-1-010-10-001"
            oAsiento.Lines.FCCurrency = "USD"
            oAsiento.Lines.FCDebit = 1

            oAsiento.Lines.Add()
            oAsiento.Lines.SetCurrentLine(1)
            oAsiento.Lines.AccountCode = "1-1-030-00-000"
            oAsiento.Lines.FCCurrency = "USD"
            oAsiento.Lines.FCCredit = 1

            iError = oAsiento.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                MessageBox.Show(oCompañia.GetNewObjectKey)
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click

        oCompañia.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
        oCompañia.XMLAsString = False

        Dim oEsquema As SAPbobsCOM.Documents
        Try
            oEsquema = oCompañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
            If oEsquema.GetByKey(274) Then
                oEsquema.SaveXML("c:\orden_venta_original.xml")
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click

        Dim order As SAPbobsCOM.Documents
        Dim x, ii As Integer

        Try
            x = oCompañia.GetXMLelementCount("c:\orden_venta_add.xml")
            For ii = 0 To x - 1
                order = oCompañia.GetBusinessObjectFromXML("c:\orden_venta_add.xml", ii)

                iError = order.Add
                If iError <> 0 Then
                    oCompañia.GetLastError(iError, sError)
                    Throw New System.Exception(sError)
                End If
            Next


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim oSn As SAPbobsCOM.BusinessPartners
        Dim oDoc As SAPbobsCOM.Documents
        Dim identi As String
        Try
            oCompañia.StartTransaction()

            oSn = oCompañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oSn.CardCode = "PTRANS2"
            oSn.CardName = "prueba transaccion2"
            oSn.CardType = SAPbobsCOM.BoCardTypes.cCustomer
            oSn.FederalTaxID = "1-1"
            iError = oSn.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            End If

            oDoc = oCompañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            oDoc.CardCode = "PTRANS2"
            oDoc.DocDueDate = Date.Now

            oDoc.Lines.SetCurrentLine(0)
            oDoc.Lines.ItemCode = "A00001"
            oDoc.Lines.Quantity = 10
            oDoc.Lines.UnitPrice = 100

            'provocamos error
            oDoc.Lines.TaxCode = "Ivo"


            iError = oDoc.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                identi = oCompañia.GetNewObjectKey
                MsgBox(identi)
            End If

            If oCompañia.InTransaction Then
                oCompañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If


        Catch ex As Exception
            If oCompañia.InTransaction Then
                oCompañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Dim oPag As SAPbobsCOM.Payments
        Dim identi As String

        Try
            oPag = oCompañia.GetBusinessObject(BoObjectTypes.oIncomingPayments)
            oPag.DocType = BoRcptTypes.rCustomer
            oPag.CardCode = "8622-xx2"
            oPag.CheckAccount = "1-1-010-10-001"

            oPag.Invoices.SetCurrentLine(0)
            oPag.Invoices.DocEntry = 235
            oPag.Invoices.SumApplied = 75000

            'efectivo
            oPag.CashSum = 25000

            'detalle(cheque)
            'oPag.Checks.SetCurrentLine(0)
            'oPag.Checks.CheckSum = 12500
            'oPag.Checks.BankCode = "037"
            'oPag.Checks.CountryCode = "CL"
            'oPag.Checks.CheckNumber = 123456

            'oPag.Checks.Add()
            'oPag.Checks.SetCurrentLine(1)
            'oPag.Checks.CheckSum = 12500
            'oPag.Checks.BankCode = "037"
            'oPag.Checks.CountryCode = "CL"
            'oPag.Checks.CheckNumber = 123457

            oPag.CreditCards.SetCurrentLine(0)
            oPag.CreditCards.CreditCard = 1
            oPag.CreditCards.CreditSum = 50000
            oPag.CreditCards.CreditCardNumber = "14000"
            oPag.CreditCards.VoucherNum = "1023a"
            oPag.CreditCards.OwnerIdNum = "kkW"
            oPag.CreditCards.CardValidUntil = "2015-04-30"

            iError = oPag.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                identi = oCompañia.GetNewObjectKey
                MsgBox(identi)
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Dim oRegisConta As SAPbobsCOM.JournalEntries
        Try
            oRegisConta = oCompañia.GetBusinessObject(BoObjectTypes.oJournalEntries)
            If oRegisConta.GetByKey(1449) Then
                oRegisConta.StornoDate = oRegisConta.ReferenceDate.AddDays(1)
                iError = oRegisConta.Cancel
                If iError <> 0 Then
                    oCompañia.GetLastError(iError, sError)
                    Throw New Exception(sError)
                End If
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Dim oDoc As SAPbobsCOM.Documents
        Dim identi As String
        Try
            oDoc = oCompañia.GetBusinessObject(BoObjectTypes.oInventoryGenEntry)
            oDoc.Reference2 = "hola mundo"

            oDoc.Lines.SetCurrentLine(0)
            oDoc.Lines.ItemCode = "joe"
            oDoc.Lines.Quantity = 100

            oDoc.Lines.BatchNumbers.SetCurrentLine(0)
            oDoc.Lines.BatchNumbers.BatchNumber = "xxx3"
            oDoc.Lines.BatchNumbers.Quantity = 50

            oDoc.Lines.BatchNumbers.Add()
            oDoc.Lines.BatchNumbers.SetCurrentLine(1)
            oDoc.Lines.BatchNumbers.BatchNumber = "xxx2"
            oDoc.Lines.BatchNumbers.Quantity = 50

            iError = oDoc.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                identi = oCompañia.GetNewObjectKey
                MsgBox(identi)
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Dim oDoc As SAPbobsCOM.Documents
        Dim identi As String
        Try
            oDoc = oCompañia.GetBusinessObject(BoObjectTypes.oInventoryGenExit)
            oDoc.Reference2 = "chao mundo"

            oDoc.Lines.SetCurrentLine(0)
            oDoc.Lines.ItemCode = "joe"
            oDoc.Lines.Quantity = 100

            oDoc.Lines.BatchNumbers.SetCurrentLine(0)
            oDoc.Lines.BatchNumbers.BatchNumber = "xxx3"
            oDoc.Lines.BatchNumbers.Quantity = 50

            oDoc.Lines.BatchNumbers.Add()
            oDoc.Lines.BatchNumbers.SetCurrentLine(1)
            oDoc.Lines.BatchNumbers.BatchNumber = "xxx2"
            oDoc.Lines.BatchNumbers.Quantity = 50

            iError = oDoc.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                identi = oCompañia.GetNewObjectKey
                MsgBox(identi)
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Dim oTransferencia As SAPbobsCOM.StockTransfer
        Dim identi As String

        Try
            oTransferencia = oCompañia.GetBusinessObject(BoObjectTypes.oStockTransfer)
            oTransferencia.FromWarehouse = "01"

            oTransferencia.Lines.SetCurrentLine(0)
            oTransferencia.Lines.ItemCode = "joe"
            oTransferencia.Lines.WarehouseCode = "02"
            oTransferencia.Lines.Quantity = 2

            oTransferencia.Lines.BatchNumbers.SetCurrentLine(0)
            oTransferencia.Lines.BatchNumbers.BatchNumber = "xxx1"
            oTransferencia.Lines.BatchNumbers.Quantity = 2

            oTransferencia.Lines.Add()
            oTransferencia.Lines.SetCurrentLine(1)
            oTransferencia.Lines.ItemCode = "joe"
            oTransferencia.Lines.WarehouseCode = "03"
            oTransferencia.Lines.Quantity = 2

            oTransferencia.Lines.BatchNumbers.SetCurrentLine(0)
            oTransferencia.Lines.BatchNumbers.BatchNumber = "xxx2"
            oTransferencia.Lines.BatchNumbers.Quantity = 2

            iError = oTransferencia.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                identi = oCompañia.GetNewObjectKey
                MsgBox(identi)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Dim odoc As SAPbobsCOM.Documents
        Dim identi As String
        Try
            odoc = oCompañia.GetBusinessObject(BoObjectTypes.oPurchaseOrders)
            odoc.CardCode = "V70000"
            odoc.DocDueDate = Date.Now.AddDays(1)
            odoc.DocType = BoDocumentTypes.dDocument_Items

            odoc.Lines.SetCurrentLine(0)
            odoc.Lines.ItemCode = "joe"
            odoc.Lines.Quantity = 10

            iError = odoc.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                identi = oCompañia.GetNewObjectKey
                MsgBox(identi)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Dim tab As UserTablesMD
        Try
            tab = oCompañia.GetBusinessObject(BoObjectTypes.oUserTables)
            If tab.GetByKey("TTransVIP") = False Then
                tab.TableName = "TTransVIP"
                tab.TableDescription = "TTransVIP"
                Try
                    iError = tab.Add()
                    If iError <> 0 Then
                        Dim msg As String = ""
                        oCompañia.GetLastError(iError, sError)
                        MessageBox.Show("Imposible Crear tabla TTransVIP, " & sError, "SBO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Imposible Crear tabla TTransVIP, falla en la librería de datos de SAP", "SBO - Kabsa", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        Dim Campos As UserFieldsMD

        Try
            Campos = oCompañia.GetBusinessObject(BoObjectTypes.oUserFields)

            If Campos.GetByKey("@TTransVIP", 0) = False Then
                Campos.TableName = "TTransVIP"
                Campos.Name = "campo1"
                Campos.Description = "campo1"
                Campos.Type = BoFieldTypes.db_Alpha
                Campos.Size = 8
                Campos.Mandatory = BoYesNoEnum.tNO
                Try
                    iError = Campos.Add
                    If iError <> 0 Then
                        Dim msg As String = ""
                        oCompañia.GetLastError(iError, sError)
                        MessageBox.Show("Imposible crear campo 'campo1' tabla TTransVIP, " & sError, "SBO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Catch EX As Exception
                    MessageBox.Show("Imposible crear campo 'campo1' tabla TTransVIP, falla en la librería de datos de SAP", "SBO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

            If Campos.GetByKey("@TTransVIP", 1) = False Then
                Campos.TableName = "TTransVIP"
                Campos.Name = "campo2"
                Campos.Description = "campo2"
                Campos.Type = BoFieldTypes.db_Numeric
                Campos.Size = 8
                Campos.Mandatory = BoYesNoEnum.tNO
                Campos.DefaultValue = 1
                Try
                    iError = Campos.Add
                    If iError <> 0 Then
                        Dim msg As String = ""
                        oCompañia.GetLastError(iError, sError)
                        MessageBox.Show("Imposible crear campo 'campo2' tabla TTransVIP, " & sError, "SBO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Catch EX As Exception
                    MessageBox.Show("Imposible crear campo 'campo2' tabla TTransVIP, falla en la librería de datos de SAP", "SBO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(Campos)


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        Dim oDoc As SAPbobsCOM.Documents
        Dim identi As String

        Try
            oDoc = oCompañia.GetBusinessObject(BoObjectTypes.oPurchaseDeliveryNotes)
            oDoc.CardCode = "V70000"
            oDoc.DocDueDate = Date.Now

            oDoc.DocType = BoDocumentTypes.dDocument_Items

            oDoc.Lines.SetCurrentLine(0)
            oDoc.Lines.BaseType = 22
            oDoc.Lines.BaseEntry = 269
            oDoc.Lines.BaseLine = 0
            oDoc.Lines.Quantity = 11

            oDoc.Lines.BatchNumbers.SetCurrentLine(0)
            oDoc.Lines.BatchNumbers.BatchNumber = "xxx4"
            oDoc.Lines.BatchNumbers.Quantity = 10

            oDoc.Lines.BatchNumbers.Add()
            oDoc.Lines.BatchNumbers.SetCurrentLine(1)
            oDoc.Lines.BatchNumbers.BatchNumber = "xxx5"
            oDoc.Lines.BatchNumbers.Quantity = 1

            iError = oDoc.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                identi = oCompañia.GetNewObjectKey
                MsgBox(identi)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        Dim odoc As SAPbobsCOM.Documents
        Dim identi As String

        Try
            odoc = oCompañia.GetBusinessObject(BoObjectTypes.oPurchaseInvoices)
            odoc.CardCode = "V70000"
            odoc.DocDueDate = Date.Now
            odoc.DocType = BoDocumentTypes.dDocument_Items

            odoc.Lines.SetCurrentLine(0)
            odoc.Lines.BaseType = 20
            odoc.Lines.BaseEntry = 268
            odoc.Lines.BaseLine = 0

            iError = odoc.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                identi = oCompañia.GetNewObjectKey
                MsgBox(identi)
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        Dim odoc As SAPbobsCOM.Documents
        Dim identi As String
        Try
            odoc = oCompañia.GetBusinessObject(BoObjectTypes.oPurchaseQuotations)
            odoc.CardCode = "V70000"
            odoc.DocDueDate = Date.Now.AddDays(1)
            odoc.RequriedDate = Date.Now.AddMonths(1)

            odoc.DocType = BoDocumentTypes.dDocument_Items

            odoc.Lines.SetCurrentLine(0)
            odoc.Lines.ItemCode = "joe"
            odoc.Lines.Quantity = 10

            iError = odoc.Add
            If iError <> 0 Then
                oCompañia.GetLastError(iError, sError)
                Throw New Exception(sError)
            Else
                identi = oCompañia.GetNewObjectKey
                MsgBox(identi)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click

        Dim cn As New SqlConnection
        cn = DBConnect()
        cn.Open()
        If (cn.State.ToString = "Open") Then
            MessageBox.Show(cn.State.ToString)
        Else
            MessageBox.Show("Cerrar")
        End If
    End Sub

    Public Function DBConnect() As SqlConnection
        Dim connectionstring As String
        connectionstring = "Data Source=IP_CONEXION;Persist Security Info=True;User ID=USUARIO;Password=PASSWORD"
        Dim cn As New SqlConnection(connectionstring)
        Return cn
    End Function

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim odoc As SAPbobsCOM.JournalEntries

        Dim identi As String


    End Sub


End Class
