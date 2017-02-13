Option Explicit On
Imports SAPbouiCOM.Framework
Imports System.Windows.Forms

Namespace AE_WWRC_V001
    Public Class clsEventHandler
        Dim WithEvents SBO_Application As SAPbouiCOM.Application ' holds connection with SBO
        Dim p_oDICompany As New SAPbobsCOM.Company

        Public Sub New(ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company)
            Dim sFuncName As String = String.Empty
            Try
                sFuncName = "Class_Initialize()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO Application handle", sFuncName)
                SBO_Application = oApplication
                p_oDICompany = oCompany

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Call WriteToLogFile(exc.Message, sFuncName)
            End Try
        End Sub

        Public Function SetApplication(ByRef sErrDesc As String) As Long
            ' **********************************************************************************
            '   Function   :    SetApplication()
            '   Purpose    :    This function will be calling to initialize the default settings
            '                   such as Retrieving the Company Default settings, Creating Menus, and
            '                   Initialize the Event Filters
            '               
            '   Parameters :    ByRef sErrDesc AS string
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            ' **********************************************************************************
            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "SetApplication()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetMenus()", sFuncName)
                If SetMenus(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetFilters()", sFuncName)
                If SetFilters(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                SetApplication = RTN_SUCCESS
            Catch exc As Exception
                sErrDesc = exc.Message
                Call WriteToLogFile(exc.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                SetApplication = RTN_ERROR
            End Try
        End Function

        Private Function SetMenus(ByRef sErrDesc As String) As Long
            ' **********************************************************************************
            '   Function   :    SetMenus()
            '   Purpose    :    This function will be gathering to create the customized menu
            '               
            '   Parameters :    ByRef sErrDesc AS string
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            ' **********************************************************************************
            Dim sFuncName As String = String.Empty
            ' Dim oMenuItem As SAPbouiCOM.MenuItem
            Try
                sFuncName = "SetMenus()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                SetMenus = RTN_SUCCESS
            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                SetMenus = RTN_ERROR
            End Try
        End Function

        Private Function SetFilters(ByRef sErrDesc As String) As Long

            ' **********************************************************************************
            '   Function   :    SetFilters()
            '   Purpose    :    This function will be gathering to declare the event filter 
            '                   before starting the AddOn Application
            '               
            '   Parameters :    ByRef sErrDesc AS string
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            ' **********************************************************************************

            Dim oFilters As SAPbouiCOM.EventFilters
            Dim oFilter As SAPbouiCOM.EventFilter
            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "SetFilters()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing EventFilters object", sFuncName)
                oFilters = New SAPbouiCOM.EventFilters



                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding filters", sFuncName)
                SBO_Application.SetFilter(oFilters)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                SetFilters = RTN_SUCCESS
            Catch exc As Exception
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                SetFilters = RTN_ERROR
            End Try
        End Function

        Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
            ' **********************************************************************************
            '   Function   :    SBO_Application_AppEvent()
            '   Purpose    :    This function will be handling the SAP Application Event
            '               
            '   Parameters :    ByVal EventType As SAPbouiCOM.BoAppEventTypes
            '                       EventType = set the SAP UI Application Eveny Object        
            ' **********************************************************************************
            Dim sFuncName As String = String.Empty
            Dim sErrDesc As String = String.Empty
            Dim sMessage As String = String.Empty

            Try
                sFuncName = "SBO_Application_AppEvent()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                Select Case EventType
                    Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                        sMessage = String.Format("Please wait for a while to disconnect the AddOn {0} ....", System.Windows.Forms.Application.ProductName)
                        p_oSBOApplication.SetStatusBarMessage(sMessage, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        End
                End Select

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch ex As Exception
                sErrDesc = ex.Message
                WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                ShowErr(sErrDesc)
            Finally
                GC.Collect()  'Forces garbage collection of all generations.
            End Try
        End Sub

        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
            ' **********************************************************************************
            '   Function   :    SBO_Application_MenuEvent()
            '   Purpose    :    This function will be handling the SAP Menu Event
            '               
            '   Parameters :    ByRef pVal As SAPbouiCOM.MenuEvent
            '                       pVal = set the SAP UI MenuEvent Object
            '                   ByRef BubbleEvent As Boolean
            '                       BubbleEvent = set the True/False        
            ' **********************************************************************************
            ' Dim oForm As SAPbouiCOM.Form = Nothing
            Dim sErrDesc As String = String.Empty
            Dim sFuncName As String = String.Empty
            Dim oForm As SAPbouiCOM.Form = Nothing
            Try
                sFuncName = "SBO_Application_MenuEvent()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If Not p_oDICompany.Connected Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                    If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If

                If pVal.BeforeAction = False Then
                    Select Case pVal.MenuUID

                        Case "FSOA"
                            Try

                                ' ''Dim F_SOA As Form1
                                ' ''F_SOA = New Form1
                                ' ''F_SOA.Show() 
                                LoadFromXML("SOA.srf", SBO_Application)
                                oForm = p_oSBOApplication.Forms.Item("SOA")
                                oForm.Visible = True
                                oForm.Items.Item("Item_5").Specific.String = PostDate(p_oDICompany)
                                oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
                                oForm.ActiveItem = "BPFrom"
                                If Set_Conditions(oForm, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                Exit Try
                            Catch ex As Exception
                                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                            End Try
                            Exit Sub


                          Case "GST"
                            Try
                                LoadFromXML_AuditFile("GSTAuditFile.srf", SBO_Application)
                                oForm = SBO_Application.Forms.Item("GST")

                                oForm.Visible = True
                                Exit Try

                            Catch ex As Exception
                                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                            End Try
                            Exit Sub

                    End Select
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                BubbleEvent = False
                ShowErr(exc.Message)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                WriteToLogFile(Err.Description, sFuncName)
            End Try
        End Sub

        Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, _
                ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
            ' **********************************************************************************
            '   Function   :    SBO_Application_ItemEvent()
            '   Purpose    :    This function will be handling the SAP Menu Event
            '               
            '   Parameters :    ByVal FormUID As String
            '                       FormUID = set the FormUID
            '                   ByRef pVal As SAPbouiCOM.ItemEvent
            '                       pVal = set the SAP UI ItemEvent Object
            '                   ByRef BubbleEvent As Boolean
            '                       BubbleEvent = set the True/False        
            ' **********************************************************************************

            Dim sErrDesc As String = String.Empty
            Dim sFuncName As String = String.Empty
            Dim p_oDVJE As DataView = Nothing
            Dim oDTDistinct As DataTable = Nothing
            Dim oDTRowFilter As DataTable = Nothing

            Try
                sFuncName = "SBO_Application_ItemEvent()"
                ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If Not IsNothing(p_oDICompany) Then
                    If Not p_oDICompany.Connected Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                        If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                End If

                If pVal.BeforeAction = False Then

                    Select Case pVal.FormUID
                        Case "SOA"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                oCFLEvento = pVal
                                Dim sCFL_ID As String
                                sCFL_ID = oCFLEvento.ChooseFromListUID
                                Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.Item(FormUID)
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                                Try
                                    If oCFLEvento.BeforeAction = False Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        If pVal.ItemUID = "BPFrom" Then 'BP From
                                            oForm.Items.Item("Item_1").Specific.string = oDataTable.GetValue("CardName", 0)
                                            oForm.Items.Item("BPFrom").Specific.string = oDataTable.GetValue("CardCode", 0)
                                        End If
                                        If pVal.ItemUID = "BPTo" Then 'BP To
                                            oForm.Items.Item("Item_3").Specific.string = oDataTable.GetValue("CardName", 0)
                                            oForm.Items.Item("BPTo").Specific.string = oDataTable.GetValue("CardCode", 0)
                                        End If
                                    End If
                                Catch ex As Exception
                                End Try
                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                If pVal.ItemUID = "Item_9" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Try
                                        Dim oMAtrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_8").Specific
                                        Dim oCheck As SAPbouiCOM.CheckBox = oForm.Items.Item("Item_9").Specific
                                        Dim ocheckColumn As SAPbouiCOM.CheckBox

                                        If oCheck.Checked = True Then
                                            For mjs As Integer = 1 To oMAtrix.RowCount
                                                ocheckColumn = oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific
                                                ocheckColumn.Checked = True
                                            Next mjs
                                        Else
                                            For mjs As Integer = 1 To oMAtrix.RowCount
                                                ocheckColumn = oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific
                                                ocheckColumn.Checked = False
                                            Next mjs
                                        End If

                                    Catch ex As Exception
                                        BubbleEvent = False
                                        sErrDesc = ex.Message
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        ShowErr(sErrDesc)
                                    End Try
                                    Exit Sub
                                End If

                            End If
                            '=========================================  GST Audit File Events Start =======================================================================================
                        Case "GST"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                If pVal.ItemUID = "9" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    sFuncName = "'Browse' Button Click - ID '9'"
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling File Open Function", sFuncName)

                                    fillopen()

                                    oForm.Items.Item("8").Specific.string = p_sSelectedFilepath
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With Success File Open Function", sFuncName)
                                    ' oForm.Items.Item("Item_5").Specific.string = p_sSelectedFilepath
                                    Exit Sub

                                ElseIf pVal.ItemUID = "btnGenerat" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Dim sPath As String = oForm.Items.Item("8").Specific.string

                                    p_oSBOApplication.StatusBar.SetText("Please wait While Generating the XML File ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                                    If GSTAuditFile(oForm.Items.Item("txtFrmDate").Specific.string, oForm.Items.Item("txtToDate").Specific.string, p_oDICompany, sPath, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("XML File Generated Successfully. Folder Path :" & sPath, sFuncName)

                                    p_oSBOApplication.StatusBar.SetText("XML File Generated Successfully. Folder Path :" & sPath, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                                End If

                            End If

                            '=========================================  GST Audit File Events End =======================================================================================
                    End Select
                Else
                    Select Case pVal.FormUID
                        Case "SOA"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                If pVal.ItemUID = "Item_10" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Try
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HeaderValidation()", sFuncName)

                                        SBO_Application.SetStatusBarMessage("Validation Process Started ........!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        If HeaderValidation(oForm, sErrDesc) = 0 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        SBO_Application.SetStatusBarMessage("Loading Data ........!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Loading_AgingDetails()", sFuncName)
                                        If Loading_AgingDetails(oForm, SBO_Application, p_oDICompany, sErrDesc) = 0 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        oForm.Items.Item("Item_12").Specific.String = ""
                                        SBO_Application.SetStatusBarMessage("Loading Data Completed Successfully ........!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        Exit Sub
                                    Catch ex As Exception
                                        BubbleEvent = False
                                        sErrDesc = ex.Message
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        ShowErr(sErrDesc)
                                    End Try
                                End If

                                If pVal.ItemUID = "Item_11" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Dim sTargetFileName As String = String.Empty
                                    Dim sRptFileName As String = String.Empty
                                    Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                                    Dim oCheck As SAPbouiCOM.CheckBox = Nothing

                                    Try
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HeaderValidation()", sFuncName)
                                        SBO_Application.SetStatusBarMessage("Validating the Records .... !", SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                                        If RowValidation(oForm, SBO_Application, sErrDesc) = 0 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        oMatrix = oForm.Items.Item("Item_8").Specific
                                        sTargetFileName = "Statement of Account_" & Format(Now.Date, "dd-MM-yyyy") & ".pdf"
                                        sTargetFileName = System.Windows.Forms.Application.StartupPath & "\" & sTargetFileName
                                        sRptFileName = System.Windows.Forms.Application.StartupPath & "\Statement_of_Account_V2.rpt"

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF() ", sFuncName)

                                        For mjs As Integer = 1 To oMatrix.RowCount
                                            oCheck = oMatrix.Columns.Item("Col_0").Cells.Item(mjs).Specific
                                            If oCheck.Checked And oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String <> "Successfully Sent Email" Then
                                                oForm.Items.Item("Item_12").Specific.String = "Processing the BP -  " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String
                                                SBO_Application.SetStatusBarMessage("Exporting SOA to PDF .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                If ExportToPDF(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, _
                                                              System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), "2", sTargetFileName, sRptFileName, SBO_Application, p_oDICompany, sErrDesc) <> RTN_SUCCESS Then
                                                    oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Fail To Send"
                                                    Throw New ArgumentException(sErrDesc)
                                                End If
                                                SBO_Application.SetStatusBarMessage("Sending Email .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                If SendEmailNotification(sTargetFileName, oMatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String, sErrDesc) <> RTN_SUCCESS Then
                                                    oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Fail To Send"
                                                    Dim sErrMsg As String = sErrDesc
                                                    sErrDesc = ""
                                                    '' Throw New ArgumentException(sErrDesc)
                                                    SBO_Application.SetStatusBarMessage("Updating the Transaction in Log Table .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                    If TransactionLog(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, _
                                                                      System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), oMatrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.String, _
                                                                      oMatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String, "Fail To Send", sErrMsg, p_oDICompany, sErrDesc) = RTN_SUCCESS Then
                                                    End If
                                                Else
                                                    oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Successfully Sent Email"
                                                    SBO_Application.SetStatusBarMessage("Updating the Transaction in Log Table .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                    If TransactionLog(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, _
                                                                     System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), oMatrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.String, _
                                                                     oMatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String, "Successfully Sent Email", "", p_oDICompany, sErrDesc) = RTN_SUCCESS Then
                                                    End If

                                                End If
                                            End If
                                        Next mjs
                                        oForm.Items.Item("Item_12").Specific.String = "Email Processing is Completed ......... "
                                        SBO_Application.SetStatusBarMessage("Email Processing is Completed .........", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        Exit Sub
                                    Catch ex As Exception
                                        BubbleEvent = False
                                        sErrDesc = ex.Message
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        ShowErr(sErrDesc)
                                    End Try
                                End If

                            End If
                            '=========================================  GST Audit File Events Start =======================================================================================

                        Case "GST"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                If pVal.ItemUID = "btnGenerat" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Head Validation Function", sFuncName)

                                    p_oSBOApplication.StatusBar.SetText("Please wait While Validating the Date and File Path ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                                    If HeaderValidation_AuditFile(oForm, sErrDesc) = 0 Then
                                        BubbleEvent = False
                                        Exit Sub
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Header Validation Function", sFuncName)
                                        p_oSBOApplication.StatusBar.SetText("Completed Validation Function", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    End If
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Header Validation Function", sFuncName)
                                    p_oSBOApplication.StatusBar.SetText("Completed Validation Function ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    Exit Sub
                                End If
                            End If
                            '=========================================  GST Audit File Events End =======================================================================================
                    End Select
                End If


                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                BubbleEvent = False
                sErrDesc = exc.Message
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                WriteToLogFile(Err.Description, sFuncName)
                ShowErr(sErrDesc)
            End Try

        End Sub

        Sub AddMenuItems()
            Dim oMenus As SAPbouiCOM.Menus
            Dim oMenuItem As SAPbouiCOM.MenuItem
            oMenus = SBO_Application.Menus

            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = (SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams))
            oMenuItem = SBO_Application.Menus.Item("43520") 'Modules

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "SOA"
            oCreationPackage.String = "Customization"
            oCreationPackage.Enabled = True
            oCreationPackage.Position = -1

            oCreationPackage.Image = System.Windows.Forms.Application.StartupPath & "\SOA1.bmp"
            oMenus = oMenuItem.SubMenus

            Try
                'If the manu already exists this code will fail
                If Not p_oSBOApplication.Menus.Exists("SOA") Then
                    oMenus.AddEx(oCreationPackage)
                End If

            Catch
            End Try


            Try
                'Get the menu collection of the newly added pop-up item
                oMenuItem = SBO_Application.Menus.Item("SOA")
                oMenus = oMenuItem.SubMenus

                'Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "FSOA"
                oCreationPackage.String = "Statement of Account"

                If Not p_oSBOApplication.Menus.Exists("FSOA") Then
                    oMenus.AddEx(oCreationPackage)
                End If


                'Get the menu collection of the newly added pop-up item
                oMenuItem = SBO_Application.Menus.Item("SOA")
                oMenus = oMenuItem.SubMenus

                'Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GST"
                oCreationPackage.String = "Export GST Audit File"

                If Not p_oSBOApplication.Menus.Exists("GST") Then
                    oMenus.AddEx(oCreationPackage)
                End If

            Catch
                'Menu already exists
                SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End Try
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
End Namespace


