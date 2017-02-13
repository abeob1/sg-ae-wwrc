Imports System.IO
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web
Imports System.Text
Namespace AE_WWRC_V001

    Module modGSTAuditFile

        Function GSTAuditFile(ByVal dtFromDate As String, ByVal dtToDate As String, ByVal oCompany As SAPbobsCOM.Company, ByVal spath As String, ByRef sErrDesc As String) As Long

            Dim sFuncName As String = String.Empty
            Dim oDSAuditFile As New DataSet
            Dim sQueryString As String = String.Empty
            Dim sFilePath As String = String.Empty


            Try
                sFuncName = "GSTAuditFile"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

                sQueryString = "EXEC AE_SP001_AuditFileReport '" & GetDate(dtFromDate, oCompany) & "','" & GetDate(dtToDate, oCompany) & "'"

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetDataSet() for Fetching the Records from the DataBase  ", sFuncName)
                oDSAuditFile = GetDataSet(sQueryString, sErrDesc)

                If oDSAuditFile Is Nothing AndAlso sErrDesc = String.Empty Then
                    sErrDesc = "There is no Matching Records Found Based on Selection Criteria."
                    Throw New ArgumentException(sErrDesc)
                End If


                sFilePath = spath & "\GSTAuditFile_" & Now.ToString("yyyyMMddHHmm") & ".xml"

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GenerateXMLFile() for Generate the XML File Based on DataSets ", sFuncName)
                If GenerateXMLFile(oDSAuditFile, sFilePath, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                GSTAuditFile = RTN_SUCCESS
            Catch ex As Exception
                Call WriteToLogFile(ex.Message, sFuncName)
                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & ex.Message, sFuncName)
                GSTAuditFile = RTN_ERROR
                Exit Function
            End Try
        End Function

        Function GetDataSet(ByVal sQueryString As String, ByRef sErrDesc As String) As DataSet

            Dim sFuncName As String = String.Empty
            Dim connetionString As String = String.Empty
            Dim connection As New SqlConnection
            Dim adapter As New SqlDataAdapter
            Dim oDSResult As New DataSet
            Dim sSQLPassword As String = String.Empty
            Dim oRS As SAPbobsCOM.Recordset


            Try
                sFuncName = "GetDataSet"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

                oRS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                oRS.DoQuery("select Name from [@CRYSTALDETAILS] where Code ='sa'")

                sSQLPassword = oRS.Fields.Item("Name").Value

                If sSQLPassword = String.Empty Then
                    sErrDesc = "Please update SQL Password In the Master Table"
                    Throw New ArgumentException(sErrDesc)
                End If

                connetionString = "Data Source=" & p_oDICompany.Server & ";Initial Catalog=" & p_oDICompany.CompanyDB & ";User ID=" & p_oDICompany.DbUserName & ";Password=" & sSQLPassword & ""
                connection = New SqlConnection(connetionString)

                'sQueryString = "EXEC AE_SP001_AuditFileReport '20141122','20141231'"

                connection.Open()
                adapter = New SqlDataAdapter(sQueryString, connection)
                adapter.Fill(oDSResult)
                connection.Close()

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)

                GetDataSet = oDSResult
            Catch ex As Exception
                sErrDesc = ex.Message.ToString()
                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Call WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & ex.Message, sFuncName)
                GetDataSet = Nothing
            End Try

        End Function

        Function GenerateXMLFile(ByRef oDSAuditFile As DataSet, ByVal sFilePath As String, ByRef sErrDesc As String) As Long

            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "GenerateXMlFile"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

                'Assign the Table Names to the Datatables

                oDSAuditFile.DataSetName = "GSTAuditFile"

                oDSAuditFile.Tables(0).TableName = "Company"
                oDSAuditFile.Tables(1).TableName = "Purchase"
                oDSAuditFile.Tables(2).TableName = "Supply"
                oDSAuditFile.Tables(3).TableName = "LedgerEntry"
                oDSAuditFile.Tables(4).TableName = "Footer"


                ' Create XmlWriterSettings.
                Dim settings As XmlWriterSettings = New XmlWriterSettings()
                settings.Indent = True

                ' Create XmlWriter.
                Using writer As XmlWriter = XmlWriter.Create(sFilePath, settings)

                    ' Begin writing.
                    writer.WriteStartDocument()

                    writer.WriteStartElement(oDSAuditFile.DataSetName)
                    writer.WriteStartElement("Companies") ' Root.

                    '=================================================================  Loop for Company Data - Starting  ==================================================================
                    For iRow As Integer = 0 To oDSAuditFile.Tables(0).Rows.Count - 1

                        writer.WriteStartElement(oDSAuditFile.Tables(0).TableName)

                        writer.WriteElementString("BusinessName", oDSAuditFile.Tables(0).Rows(0)("BusinessName").ToString())
                        writer.WriteElementString("BusinessRN", oDSAuditFile.Tables(0).Rows(0)("BusinessRN").ToString())
                        writer.WriteElementString("GSTNumber", oDSAuditFile.Tables(0).Rows(0)("GSTNumber").ToString())
                        writer.WriteElementString("PeriodStart", oDSAuditFile.Tables(0).Rows(0)("PeriodStart").ToString())
                        writer.WriteElementString("PeriodEnd", oDSAuditFile.Tables(0).Rows(0)("PeriodEnd").ToString())
                        writer.WriteElementString("GAFCreationDate", oDSAuditFile.Tables(0).Rows(0)("GAFCreationDate").ToString())
                        writer.WriteElementString("ProductVersion", oDSAuditFile.Tables(0).Rows(0)("ProductVersion").ToString())
                        writer.WriteElementString("GAFVersion", oDSAuditFile.Tables(0).Rows(0)("GAFVersion").ToString())

                        writer.WriteEndElement()

                    Next

                    ' End document.
                    writer.WriteEndElement()

                    '=================================================================  Loop for Company Data - Ending ==================================================================


                    '=================================================================  Loop for Purchases Data -Starting ==================================================================
                    writer.WriteStartElement("Purchases") ' Root.

                    For iRow As Integer = 0 To oDSAuditFile.Tables(1).Rows.Count - 1

                        writer.WriteStartElement(oDSAuditFile.Tables(1).TableName)

                        writer.WriteElementString("SupplierName", oDSAuditFile.Tables(1).Rows(iRow)("SupplierName").ToString())
                        writer.WriteElementString("SupplierBRN", oDSAuditFile.Tables(1).Rows(iRow)("SupplierBRN").ToString())
                        writer.WriteElementString("InvoiceDate", oDSAuditFile.Tables(1).Rows(iRow)("InvoiceDate").ToString())
                        writer.WriteElementString("InvoiceNumber", oDSAuditFile.Tables(1).Rows(iRow)("InvoiceNumber").ToString())
                        writer.WriteElementString("ImportDeclarationNo", oDSAuditFile.Tables(1).Rows(iRow)("ImportDeclarationNo").ToString())
                        writer.WriteElementString("LineNumber", oDSAuditFile.Tables(1).Rows(iRow)("LineNumber").ToString())
                        writer.WriteElementString("ProductDescription", oDSAuditFile.Tables(1).Rows(iRow)("ProductDescription").ToString())
                        writer.WriteElementString("PurchaseValueMYR", oDSAuditFile.Tables(1).Rows(iRow)("PurchaseValueMYR").ToString())
                        writer.WriteElementString("GSTValueMYR", oDSAuditFile.Tables(1).Rows(iRow)("GSTValueMYR").ToString())
                        writer.WriteElementString("TaxCode", oDSAuditFile.Tables(1).Rows(iRow)("TaxCode").ToString())
                        writer.WriteElementString("FCYCode", oDSAuditFile.Tables(1).Rows(iRow)("FCYCode").ToString())
                        writer.WriteElementString("PurchaseFCY", oDSAuditFile.Tables(1).Rows(iRow)("PurchaseFCY").ToString())
                        writer.WriteElementString("GSTFCY", oDSAuditFile.Tables(1).Rows(iRow)("GSTFCY").ToString())

                        writer.WriteEndElement()

                    Next

                    writer.WriteEndElement()
                    '=================================================================  Loop for Purchases Data -Ending ==================================================================

                    '=================================================================  Loop for Supplies Data -Starting ==================================================================
                    writer.WriteStartElement("Supplies") ' Root.

                    For iRow As Integer = 0 To oDSAuditFile.Tables(2).Rows.Count - 1

                        writer.WriteStartElement(oDSAuditFile.Tables(2).TableName)

                        writer.WriteElementString("CustomerName", oDSAuditFile.Tables(2).Rows(iRow)("CustomerName").ToString())
                        writer.WriteElementString("CustomerBRN", oDSAuditFile.Tables(2).Rows(iRow)("CustomerBRN").ToString())
                        writer.WriteElementString("InvoiceDate", oDSAuditFile.Tables(2).Rows(iRow)("InvoiceDate").ToString())
                        writer.WriteElementString("InvoiceNumber", oDSAuditFile.Tables(2).Rows(iRow)("InvoiceNumber").ToString())
                        writer.WriteElementString("LineNumber", oDSAuditFile.Tables(2).Rows(iRow)("LineNumber").ToString())
                        writer.WriteElementString("ProductDescription", oDSAuditFile.Tables(2).Rows(iRow)("ProductDescription").ToString())
                        writer.WriteElementString("SupplyValueMYR", oDSAuditFile.Tables(2).Rows(iRow)("SupplyValueMYR").ToString())
                        writer.WriteElementString("GSTValueMYR", oDSAuditFile.Tables(2).Rows(iRow)("GSTValueMYR").ToString())
                        writer.WriteElementString("TaxCode", oDSAuditFile.Tables(2).Rows(iRow)("TaxCode").ToString())
                        writer.WriteElementString("Country", oDSAuditFile.Tables(2).Rows(iRow)("Country").ToString())
                        writer.WriteElementString("FCYCode", oDSAuditFile.Tables(2).Rows(iRow)("FCYCode").ToString())
                        writer.WriteElementString("SupplyFCY", oDSAuditFile.Tables(2).Rows(iRow)("SupplyFCY").ToString())
                        writer.WriteElementString("GSTFCY", oDSAuditFile.Tables(2).Rows(iRow)("GSTFCY").ToString())

                        writer.WriteEndElement()

                    Next

                    writer.WriteEndElement()

                    '=================================================================  Loop for Aupplies Data -Ending ==================================================================

                    '=================================================================  Loop for Ledger Data -Starting ==================================================================

                    Dim dBalanceAmount As Double = 0
                    Dim dDebit As Double = 0
                    Dim dCredit As Double = 0
                    Dim sAcctCode As String = String.Empty

                    writer.WriteStartElement("Ledger") ' Root.

                    For iRow As Integer = 0 To oDSAuditFile.Tables(3).Rows.Count - 1

                        writer.WriteStartElement(oDSAuditFile.Tables(3).TableName)

                        writer.WriteElementString("TransactionDate", oDSAuditFile.Tables(3).Rows(iRow)("TransactionDate").ToString())
                        writer.WriteElementString("AccountID", oDSAuditFile.Tables(3).Rows(iRow)("AccountID").ToString())
                        writer.WriteElementString("AccountName", oDSAuditFile.Tables(3).Rows(iRow)("AccountName").ToString())
                        writer.WriteElementString("TransactionDescription", oDSAuditFile.Tables(3).Rows(iRow)("TransactionDescription").ToString())
                        writer.WriteElementString("Name", oDSAuditFile.Tables(3).Rows(iRow)("Name").ToString())
                        writer.WriteElementString("TransactionID", oDSAuditFile.Tables(3).Rows(iRow)("TransactionID").ToString())
                        writer.WriteElementString("SourceDocumentID", oDSAuditFile.Tables(3).Rows(iRow)("SourceDocumentID").ToString())
                        writer.WriteElementString("SourceType", oDSAuditFile.Tables(3).Rows(iRow)("SourceType").ToString())
                        writer.WriteElementString("Debit", oDSAuditFile.Tables(3).Rows(iRow)("Debit").ToString())
                        writer.WriteElementString("Credit", oDSAuditFile.Tables(3).Rows(iRow)("Credit").ToString())

                        If iRow = 0 Then
                            dBalanceAmount = (CDbl(oDSAuditFile.Tables(3).Rows(iRow)("Balance").ToString()) + CDbl(oDSAuditFile.Tables(3).Rows(iRow)("Debit").ToString())) - CDbl(oDSAuditFile.Tables(3).Rows(iRow)("Credit").ToString())
                        Else
                            dBalanceAmount += (CDbl(oDSAuditFile.Tables(3).Rows(iRow)("Debit").ToString()) - CDbl((oDSAuditFile.Tables(3).Rows(iRow)("Credit").ToString())))
                        End If

                        If sAcctCode <> oDSAuditFile.Tables(3).Rows(iRow)("AccountID").ToString() And iRow <> 0 Then
                            ''dBalanceAmount = CDbl(oDSAuditFile.Tables(3).Rows(iRow)("Balance").ToString())
                            dBalanceAmount = (CDbl(oDSAuditFile.Tables(3).Rows(iRow)("Balance").ToString()) + CDbl(oDSAuditFile.Tables(3).Rows(iRow)("Debit").ToString())) - CDbl(oDSAuditFile.Tables(3).Rows(iRow)("Credit").ToString())
                        End If

                        writer.WriteElementString("Balance", dBalanceAmount)

                        sAcctCode = oDSAuditFile.Tables(3).Rows(iRow)("AccountID").ToString()

                        writer.WriteEndElement()

                    Next

                    writer.WriteEndElement()

                    '=================================================================  Loop for Ledger Data -Ending ==================================================================

                    '=================================================================  Loop for Footer Data -Starting ==================================================================

                    writer.WriteStartElement("Footers") ' Root.

                    For iRow As Integer = 0 To oDSAuditFile.Tables(4).Rows.Count - 1

                        writer.WriteStartElement(oDSAuditFile.Tables(4).TableName)

                        writer.WriteElementString("TotalPurchaseCount", oDSAuditFile.Tables(4).Rows(iRow)("TotalPurchaseCount").ToString())
                        writer.WriteElementString("TotalPurchaseAmount", oDSAuditFile.Tables(4).Rows(iRow)("TotalPurchaseAmount").ToString())
                        writer.WriteElementString("TotalPurchaseAmountGST", oDSAuditFile.Tables(4).Rows(iRow)("TotalPurchaseAmountGST").ToString())
                        writer.WriteElementString("TotalSupplyCount", oDSAuditFile.Tables(4).Rows(iRow)("TotalSupplyCount").ToString())
                        writer.WriteElementString("TotalSupplyAmount", oDSAuditFile.Tables(4).Rows(iRow)("TotalSupplyAmount").ToString())
                        writer.WriteElementString("TotalSupplyAmountGST", oDSAuditFile.Tables(4).Rows(iRow)("TotalSupplyAmountGST").ToString())
                        writer.WriteElementString("TotalLedgerCount", oDSAuditFile.Tables(4).Rows(iRow)("TotalLedgerCount").ToString())
                        writer.WriteElementString("TotalLedgerDebit", oDSAuditFile.Tables(4).Rows(iRow)("TotalLedgerDebit").ToString())
                        writer.WriteElementString("TotalLedgerCredit", oDSAuditFile.Tables(4).Rows(iRow)("TotalLedgerCredit").ToString())
                        writer.WriteElementString("TotalLedgerBalance", oDSAuditFile.Tables(4).Rows(iRow)("TotalLedgerBalance").ToString())
                        writer.WriteEndElement()

                    Next

                    writer.WriteEndElement()

                    '=================================================================  Loop for Purchases Data -Ending ==================================================================

                    writer.WriteEndDocument()

                    writer.Close()
                    writer.Flush()

                End Using



                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                GenerateXMLFile = RTN_SUCCESS

            Catch ex As Exception
                Call WriteToLogFile(ex.Message, sFuncName)
                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & ex.Message, sFuncName)
                GenerateXMLFile = RTN_ERROR
            End Try

        End Function

    End Module
End Namespace
