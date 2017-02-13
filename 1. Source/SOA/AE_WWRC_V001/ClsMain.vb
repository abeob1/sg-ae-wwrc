Imports SAPbouiCOM.Framework

Namespace AE_WWRC_V001
    Public Class ClsMain


        Private WithEvents SBO_Application As SAPbouiCOM.Application

        Sub New()
            SBO_Application = Application.SBO_Application
        End Sub

        Sub AddMenuItems()
            Dim oMenus As SAPbouiCOM.Menus
            Dim oMenuItem As SAPbouiCOM.MenuItem
            oMenus = Application.SBO_Application.Menus

            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = (Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams))
            oMenuItem = Application.SBO_Application.Menus.Item("43520") 'Modules

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "SOA"
            oCreationPackage.String = "Statement of Acccount"
            oCreationPackage.Enabled = True
            oCreationPackage.Position = -1

            oMenus = oMenuItem.SubMenus

            Try
                'If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage)
            Catch
            End Try


            Try
                'Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("SOA")
                oMenus = oMenuItem.SubMenus

                'Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "FSOA"
                oCreationPackage.String = "Statement of Account"
                oMenus.AddEx(oCreationPackage)
            Catch
                'Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub


        Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
            BubbleEvent = True

            Try
                If (pVal.BeforeAction And pVal.MenuUID = "SOA") Then
                    Dim activeForm As Form1
                    activeForm = New Form1
                    activeForm.Show()
                End If
            Catch ex As System.Exception
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "")
            End Try

        End Sub

    End Class
End Namespace