Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace AE_WWRC_V001
    <FormAttribute("AE_WWRC_V001.Form1", "SOA.b1f")>
    Friend Class Form1
        Inherits UserFormBase

        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.StaticText0 = CType(Me.GetItem("Item_0").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("BPFrom").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("Item_2").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("BPTo").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("Item_4").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("Item_5").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("Item_6").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("Item_7").Specific, SAPbouiCOM.EditText)
            Me.Matrix0 = CType(Me.GetItem("Item_8").Specific, SAPbouiCOM.Matrix)
            Me.CheckBox0 = CType(Me.GetItem("Item_9").Specific, SAPbouiCOM.CheckBox)
            Me.Button0 = CType(Me.GetItem("Item_10").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("Item_11").Specific, SAPbouiCOM.Button)
            Me.Button2 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.EditText5 = CType(Me.GetItem("Item_12").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText

        Private Sub OnCustomInitialize()

        End Sub
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents CheckBox0 As SAPbouiCOM.CheckBox
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Button2 As SAPbouiCOM.Button
        Private WithEvents EditText5 As SAPbouiCOM.EditText

    End Class
End Namespace
