VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRadiusConduit 
   Caption         =   "Conduit if Using Radius"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6150
   OleObjectBlob   =   "frmRadiusConduit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRadiusConduit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboTrayCableTypes2_Change()
If cboTrayCableTypes2.Text = "LM1000_Cat6_Riser" Then
Me.txtCableAreaValue2.Text = "0.040"
    ElseIf cboTrayCableTypes2.Text = "LM1000_Cat6_Plenum" Then
    Me.txtCableAreaValue2.Text = "0.040"
        ElseIf cboTrayCableTypes2.Text = "Clarity_Modular_Patch_Cords" Then
        Me.txtCableAreaValue2.Text = "0.066"
            ElseIf cboTrayCableTypes2.Text = "Cat5e_25_pair" Then
            Me.txtCableAreaValue2.Text = "0.169"
                ElseIf cboTrayCableTypes2.Text = "Corning_048T88-61180-A3" Then
                Me.txtCableAreaValue2.Text = "0.816"
                    ElseIf cboTrayCableTypes2.Text = "Corning_48T8F-31190-A1" Then
                    Me.txtCableAreaValue2.Text = ".0882"
                        ElseIf cboTrayCableTypes2.Text = "Corning_024T88-33180" Then
                        Me.txtCableAreaValue2.Text = "0.075"
                            ElseIf cboTrayCableTypes2.Text = "Corning_012T88-33180" Then
                            Me.txtCableAreaValue2.Text = "0.045"
                                ElseIf cboTrayCableTypes2.Text = "Corning_x757512QPNNDUxxxf" Then
                                Me.txtCableAreaValue2.Text = "0.022"
                                    ElseIf cboTrayCableTypes2.Text = "Corning_x757524OPNNDUxxxf" Then
                                    Me.txtCableAreaValue2.Text = "0.049"
                                        ElseIf cboTrayCableTypes2.Text = "Corning_x757548QPNNDUxxxf" Then
                                        Me.txtCableAreaValue2.Text = "0.070"
                                            ElseIf cboTrayCableTypes2.Text = "Sec._Composite" Then
                                            Me.txtCableAreaValue2.Text = "0.131"
                                                ElseIf cboTrayCableTypes2.Text = "Outdoor_fiber_ALTOS" Then
                                                Me.txtCableAreaValue2.Text = "0.180"
                                                    ElseIf cboTrayCableTypes2.Text = "Outdoor_fiber_FREEDM" Then
                                                    Me.txtCableAreaValue2.Text = "0.066"
End If
End Sub

Private Sub cmdHelp_Click()
msgbox "As a Cable Type is selected from the drop down menu, the value of it's area changes accordingly. Should your desired Cable/Cable Area not be there you may enter in your own custom Cable Area Value and it shall function in the same way."
End Sub

Private Sub txtCableAreaValue2_Change()
On Error Resume Next
If IsNumeric(Me.txtCableAreaValue2) Then
Me.txtQuantity2.Text = Int((Me.lblActualConArea2) / (Me.txtCableAreaValue2))
End If
If IsNumeric(Me.txtCableAreaValue2) Then
Me.txtDerated.Text = Int((Me.lblActualConArea2 * 0.85) / (Me.txtCableAreaValue2))
End If
End Sub

Private Sub txtRadius_Change()
If IsNumeric(Me.txtRadius) Then
Me.lblConArea2.Caption = ((Me.txtRadius) * (Me.txtRadius)) * 3.14
End If
If IsNumeric(Me.lblConArea2) Then
Me.lblActualConArea2.Caption = Me.lblConArea2 * 0.4
End If
If IsNumeric(Me.txtCableAreaValue2) Then
Me.txtQuantity2.Text = Int((Me.lblActualConArea2) / (Me.txtCableAreaValue2))
End If
If IsNumeric(Me.txtQuantity2) Then
Me.lblConFill.Caption = ((Me.txtQuantity2 * Me.txtCableAreaValue2) / (Me.lblActualConArea2) * 100)
End If
lblPercent.Visible = True
If IsNumeric(Me.txtCableAreaValue2) Then
Me.txtDerated.Text = Int((Me.lblActualConArea2 * 0.85) / (Me.txtCableAreaValue2))
End If
If IsNumeric(Me.txtDerated) Then
Me.lblDeFill.Caption = ((Me.txtDerated * Me.txtCableAreaValue2) / (Me.lblActualConArea2) * 100)
End If
lblPercent2.Visible = True
End Sub

Private Sub UserForm_initialize()
With cboTrayCableTypes2
.AddItem "LM1000_Cat6_Riser"
    .AddItem "LM1000_Cat6_Plenum"
        .AddItem "Clarity_Modular_Patch_Cords"
            .AddItem "Cat5e_25_pair"
                .AddItem "Corning_048T88-61180-A3"
                    .AddItem "Corning_48T8F-31190-A1"
                        .AddItem "Corning_024T88-33180"
                            .AddItem "Corning_012T88-33180"
                                .AddItem "Corning_x757512QPNNDUxxxf"
                                    .AddItem "Corning_x757524OPNNDUxxxf"
                                        .AddItem "Corning_x757548QPNNDUxxxf"
                                            .AddItem "Sec._Composite"
                                                .AddItem "Outdoor_fiber_ALTOS"
                                                    .AddItem "Outdoor_fiber_FREEDM"
End With
End Sub
