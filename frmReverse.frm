VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReverse 
   Caption         =   "Reverse Fill"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10440
   OleObjectBlob   =   "frmReverse.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCable_Change()
If cboCable.Text = "LM1000_Cat6_Riser" Then
Me.txtCableAreaValue.Text = "0.040"
    ElseIf cboCable.Text = "LM1000_Cat6_Plenum" Then
    Me.txtCableAreaValue.Text = "0.040"
        ElseIf cboCable.Text = "Clarity_Modular_Patch_Cords" Then
        Me.txtCableAreaValue.Text = "0.066"
            ElseIf cboCable.Text = "Cat5e_25_pair" Then
            Me.txtCableAreaValue.Text = "0.169"
                ElseIf cboCable.Text = "Corning_048T88-61180-A3" Then
                Me.txtCableAreaValue.Text = "0.816"
                    ElseIf cboCable.Text = "Corning_48T8F-31190-A1" Then
                    Me.txtCableAreaValue.Text = ".0882"
                        ElseIf cboCable.Text = "Corning_024T88-33180" Then
                        Me.txtCableAreaValue.Text = "0.075"
                            ElseIf cboCable.Text = "Corning_012T88-33180" Then
                            Me.txtCableAreaValue.Text = "0.045"
                                ElseIf cboCable.Text = "Corning_x757512QPNNDUxxxf" Then
                                Me.txtCableAreaValue.Text = "0.022"
                                    ElseIf cboCable.Text = "Corning_x757524OPNNDUxxxf" Then
                                    Me.txtCableAreaValue.Text = "0.049"
                                        ElseIf cboCable.Text = "Corning_x757548QPNNDUxxxf" Then
                                        Me.txtCableAreaValue.Text = "0.070"
                                            ElseIf cboCable.Text = "Security_Composite" Then
                                            Me.txtCableAreaValue.Text = "0.131"
                                                ElseIf cboCable.Text = "Outdoor_fiber_ALTOS" Then
                                                Me.txtCableAreaValue.Text = "0.180"
                                                    ElseIf cboCable.Text = "Outdoor_fiber_FREEDM" Then
                                                    Me.txtCableAreaValue.Text = "0.066"
End If
End Sub

Private Sub cboCable2_Change()
If cboCable2.Text = "LM1000_Cat6_Riser" Then
Me.txtCableAreaValue2.Text = "0.040"
    ElseIf cboCable2.Text = "LM1000_Cat6_Plenum" Then
    Me.txtCableAreaValue2.Text = "0.040"
        ElseIf cboCable2.Text = "Clarity_Modular_Patch_Cords" Then
        Me.txtCableAreaValue2.Text = "0.066"
            ElseIf cboCable2.Text = "Cat5e_25_pair" Then
            Me.txtCableAreaValue2.Text = "0.169"
                ElseIf cboCable2.Text = "Corning_048T88-61180-A3" Then
                Me.txtCableAreaValue2.Text = "0.816"
                    ElseIf cboCable2.Text = "Corning_48T8F-31190-A1" Then
                    Me.txtCableAreaValue2.Text = ".0882"
                        ElseIf cboCable2.Text = "Corning_024T88-33180" Then
                        Me.txtCableAreaValue2.Text = "0.075"
                            ElseIf cboCable2.Text = "Corning_012T88-33180" Then
                            Me.txtCableAreaValue2.Text = "0.045"
                                ElseIf cboCable2.Text = "Corning_x757512QPNNDUxxxf" Then
                                Me.txtCableAreaValue2.Text = "0.022"
                                    ElseIf cboCable2.Text = "Corning_x757524OPNNDUxxxf" Then
                                    Me.txtCableAreaValue2.Text = "0.049"
                                        ElseIf cboCable2.Text = "Corning_x757548QPNNDUxxxf" Then
                                        Me.txtCableAreaValue2.Text = "0.070"
                                            ElseIf cboCable2.Text = "Security_Composite" Then
                                            Me.txtCableAreaValue2.Text = "0.131"
                                                ElseIf cboCable2.Text = "Outdoor_fiber_ALTOS" Then
                                                Me.txtCableAreaValue2.Text = "0.180"
                                                    ElseIf cboCable2.Text = "Outdoor_fiber_FREEDM" Then
                                                    Me.txtCableAreaValue2.Text = "0.066"
End If
End Sub

Private Sub cboCable3_Change()
If cboCable3.Text = "LM1000_Cat6_Riser" Then
Me.txtCableAreaValue3.Text = "0.040"
    ElseIf cboCable3.Text = "LM1000_Cat6_Plenum" Then
    Me.txtCableAreaValue3.Text = "0.040"
        ElseIf cboCable3.Text = "Clarity_Modular_Patch_Cords" Then
        Me.txtCableAreaValue3.Text = "0.066"
            ElseIf cboCable3.Text = "Cat5e_25_pair" Then
            Me.txtCableAreaValue3.Text = "0.169"
                ElseIf cboCable3.Text = "Corning_048T88-61180-A3" Then
                Me.txtCableAreaValue3.Text = "0.816"
                    ElseIf cboCable3.Text = "Corning_48T8F-31190-A1" Then
                    Me.txtCableAreaValue3.Text = ".0882"
                        ElseIf cboCable3.Text = "Corning_024T88-33180" Then
                        Me.txtCableAreaValue3.Text = "0.075"
                            ElseIf cboCable3.Text = "Corning_012T88-33180" Then
                            Me.txtCableAreaValue3.Text = "0.045"
                                ElseIf cboCable3.Text = "Corning_x757512QPNNDUxxxf" Then
                                Me.txtCableAreaValue3.Text = "0.022"
                                    ElseIf cboCable3.Text = "Corning_x757524OPNNDUxxxf" Then
                                    Me.txtCableAreaValue3.Text = "0.049"
                                        ElseIf cboCable3.Text = "Corning_x757548QPNNDUxxxf" Then
                                        Me.txtCableAreaValue3.Text = "0.070"
                                            ElseIf cboCable3.Text = "Security_Composite" Then
                                            Me.txtCableAreaValue3.Text = "0.131"
                                                ElseIf cboCable3.Text = "Outdoor_fiber_ALTOS" Then
                                                Me.txtCableAreaValue3.Text = "0.180"
                                                    ElseIf cboCable3.Text = "Outdoor_fiber_FREEDM" Then
                                                    Me.txtCableAreaValue3.Text = "0.066"
End If
End Sub

Private Sub cboCable4_Change()
If cboCable4.Text = "LM1000_Cat6_Riser" Then
Me.txtCableAreaValue4.Text = "0.040"
    ElseIf cboCable4.Text = "LM1000_Cat6_Plenum" Then
    Me.txtCableAreaValue4.Text = "0.040"
        ElseIf cboCable4.Text = "Clarity_Modular_Patch_Cords" Then
        Me.txtCableAreaValue4.Text = "0.066"
            ElseIf cboCable4.Text = "Cat5e_25_pair" Then
            Me.txtCableAreaValue4.Text = "0.169"
                ElseIf cboCable4.Text = "Corning_048T88-61180-A3" Then
                Me.txtCableAreaValue4.Text = "0.816"
                    ElseIf cboCable4.Text = "Corning_48T8F-31190-A1" Then
                    Me.txtCableAreaValue4.Text = ".0882"
                        ElseIf cboCable4.Text = "Corning_024T88-33180" Then
                        Me.txtCableAreaValue4.Text = "0.075"
                            ElseIf cboCable4.Text = "Corning_012T88-33180" Then
                            Me.txtCableAreaValue4.Text = "0.045"
                                ElseIf cboCable4.Text = "Corning_x757512QPNNDUxxxf" Then
                                Me.txtCableAreaValue4.Text = "0.022"
                                    ElseIf cboCable4.Text = "Corning_x757524OPNNDUxxxf" Then
                                    Me.txtCableAreaValue4.Text = "0.049"
                                        ElseIf cboCable4.Text = "Corning_x757548QPNNDUxxxf" Then
                                        Me.txtCableAreaValue4.Text = "0.070"
                                            ElseIf cboCable4.Text = "Security_Composite" Then
                                            Me.txtCableAreaValue4.Text = "0.131"
                                                ElseIf cboCable4.Text = "Outdoor_fiber_ALTOS" Then
                                                Me.txtCableAreaValue4.Text = "0.180"
                                                    ElseIf cboCable4.Text = "Outdoor_fiber_FREEDM" Then
                                                    Me.txtCableAreaValue4.Text = "0.066"
End If
End Sub

Private Sub cboCable5_Change()
If cboCable5.Text = "LM1000_Cat6_Riser" Then
Me.txtCableAreaValue5.Text = "0.040"
    ElseIf cboCable5.Text = "LM1000_Cat6_Plenum" Then
    Me.txtCableAreaValue5.Text = "0.040"
        ElseIf cboCable5.Text = "Clarity_Modular_Patch_Cords" Then
        Me.txtCableAreaValue5.Text = "0.066"
            ElseIf cboCable5.Text = "Cat5e_25_pair" Then
            Me.txtCableAreaValue5.Text = "0.169"
                ElseIf cboCable5.Text = "Corning_048T88-61180-A3" Then
                Me.txtCableAreaValue5.Text = "0.816"
                    ElseIf cboCable5.Text = "Corning_48T8F-31190-A1" Then
                    Me.txtCableAreaValue5.Text = ".0882"
                        ElseIf cboCable5.Text = "Corning_024T88-33180" Then
                        Me.txtCableAreaValue5.Text = "0.075"
                            ElseIf cboCable5.Text = "Corning_012T88-33180" Then
                            Me.txtCableAreaValue5.Text = "0.045"
                                ElseIf cboCable5.Text = "Corning_x757512QPNNDUxxxf" Then
                                Me.txtCableAreaValue5.Text = "0.022"
                                    ElseIf cboCable5.Text = "Corning_x757524OPNNDUxxxf" Then
                                    Me.txtCableAreaValue5.Text = "0.049"
                                        ElseIf cboCable5.Text = "Corning_x757548QPNNDUxxxf" Then
                                        Me.txtCableAreaValue5.Text = "0.070"
                                            ElseIf cboCable5.Text = "Security_Composite" Then
                                            Me.txtCableAreaValue5.Text = "0.131"
                                                ElseIf cboCable5.Text = "Outdoor_fiber_ALTOS" Then
                                                Me.txtCableAreaValue5.Text = "0.180"
                                                    ElseIf cboCable5.Text = "Outdoor_fiber_FREEDM" Then
                                                    Me.txtCableAreaValue5.Text = "0.066"
End If
End Sub

Private Sub cboCable6_Change()
If cboCable6.Text = "LM1000_Cat6_Riser" Then
Me.txtCableAreaValue6.Text = "0.040"
    ElseIf cboCable6.Text = "LM1000_Cat6_Plenum" Then
    Me.txtCableAreaValue6.Text = "0.040"
        ElseIf cboCable6.Text = "Clarity_Modular_Patch_Cords" Then
        Me.txtCableAreaValue6.Text = "0.066"
            ElseIf cboCable6.Text = "Cat5e_25_pair" Then
            Me.txtCableAreaValue6.Text = "0.169"
                ElseIf cboCable6.Text = "Corning_048T88-61180-A3" Then
                Me.txtCableAreaValue6.Text = "0.816"
                    ElseIf cboCable6.Text = "Corning_48T8F-31190-A1" Then
                    Me.txtCableAreaValue6.Text = ".0882"
                        ElseIf cboCable6.Text = "Corning_024T88-33180" Then
                        Me.txtCableAreaValue6.Text = "0.075"
                            ElseIf cboCable6.Text = "Corning_012T88-33180" Then
                            Me.txtCableAreaValue6.Text = "0.045"
                                ElseIf cboCable6.Text = "Corning_x757512QPNNDUxxxf" Then
                                Me.txtCableAreaValue6.Text = "0.022"
                                    ElseIf cboCable6.Text = "Corning_x757524OPNNDUxxxf" Then
                                    Me.txtCableAreaValue6.Text = "0.049"
                                        ElseIf cboCable6.Text = "Corning_x757548QPNNDUxxxf" Then
                                        Me.txtCableAreaValue6.Text = "0.070"
                                            ElseIf cboCable6.Text = "Security_Composite" Then
                                            Me.txtCableAreaValue6.Text = "0.131"
                                                ElseIf cboCable6.Text = "Outdoor_fiber_ALTOS" Then
                                                Me.txtCableAreaValue6.Text = "0.180"
                                                    ElseIf cboCable6.Text = "Outdoor_fiber_FREEDM" Then
                                                    Me.txtCableAreaValue6.Text = "0.066"
End If
End Sub

Private Sub chkAdd1_Click()
If chkAdd1.Value = True Then
    txtInput2.Visible = True
    chkAdd2.Visible = True
    cboCable2.Visible = True
    lblCableArea2.Visible = True
    txtCableAreaValue2.Visible = True
        Else
        txtInput2.Visible = False
        chkAdd2.Visible = False
        cboCable2.Visible = False
        lblCableArea2.Visible = False
        txtCableAreaValue2.Visible = False
End If
End Sub

Private Sub chkAdd2_Click()
If chkAdd2.Value = True Then
    txtInput3.Visible = True
    chkAdd3.Visible = True
    cboCable3.Visible = True
    lblCableArea3.Visible = True
    txtCableAreaValue3.Visible = True
        Else
        txtInput3.Visible = False
        chkAdd3.Visible = False
        cboCable3.Visible = False
        lblCableArea3.Visible = False
        txtCableAreaValue3.Visible = False
End If
End Sub

Private Sub chkAdd3_Click()
If chkAdd3.Value = True Then
    txtInput4.Visible = True
    chkAdd4.Visible = True
    cboCable4.Visible = True
    lblCableArea4.Visible = True
    txtCableAreaValue4.Visible = True
        Else
        txtInput4.Visible = False
        chkAdd4.Visible = False
        cboCable4.Visible = False
        lblCableArea4.Visible = False
        txtCableAreaValue4.Visible = False
End If
End Sub

Private Sub chkAdd4_Click()
If chkAdd4.Value = True Then
    txtInput5.Visible = True
    chkAdd5.Visible = True
    cboCable5.Visible = True
    lblCableArea5.Visible = True
    txtCableAreaValue5.Visible = True
        Else
        txtInput5.Visible = False
        chkAdd5.Visible = False
        cboCable5.Visible = False
        lblCableArea5.Visible = False
        txtCableAreaValue5.Visible = False
End If
End Sub

Private Sub chkAdd5_Click()
If chkAdd5.Value = True Then
    txtInput6.Visible = True
    cboCable6.Visible = True
    lblCableArea6.Visible = True
    txtCableAreaValue6.Visible = True
        Else
        txtInput6.Visible = False
        cboCable6.Visible = False
        lblCableArea6.Visible = False
        txtCableAreaValue6.Visible = False
End If
End Sub

Private Sub cmdRHelp_Click()
msgbox "This calculator will show you fill percentage of certain tray/conduit sizes according to the number and type of cables you fill in. To add additional cables you may check the area's designated and be given an extra interface. Simply plug in your amount and cable to recieve your fill percentages. Should the case be that your cable may not be on the list you may fill in your own custom area accordingly."
End Sub

Private Sub txt1_25_Change()
If txt1_25.Value > 50 Then
    txt1_25.BackColor = &HFF&
        ElseIf txt1_25.Value < 41 Then
        txt1_25.BackColor = &HFF00&
            Else: txt1_25.BackColor = &HFFFF&
End If
End Sub

Private Sub txt1_5_Change()
If txt1_5.Value > 50 Then
    txt1_5.BackColor = &HFF&
        ElseIf txt1_5.Value < 41 Then
        txt1_5.BackColor = &HFF00&
            Else: txt1_5.BackColor = &HFFFF&
End If
End Sub

Private Sub txt2_5_Change()
If txt2_5.Value > 50 Then
    txt2_5.BackColor = &HFF&
        ElseIf txt2_5.Value < 41 Then
        txt2_5.BackColor = &HFF00&
            Else: txt2_5.BackColor = &HFFFF&
End If
End Sub

Private Sub txt2_Change()
If txt2.Value > 50 Then
    txt2.BackColor = &HFF&
        ElseIf txt2.Value < 41 Then
        txt2.BackColor = &HFF00&
            Else: txt2.BackColor = &HFFFF&
End If
End Sub

Private Sub txt3_Change()
If txt3.Value > 50 Then
    txt3.BackColor = &HFF&
        ElseIf txt3.Value < 41 Then
        txt3.BackColor = &HFF00&
            Else: txt3.BackColor = &HFFFF&
End If
End Sub

Private Sub txt4_Change()
If txt4.Value > 50 Then
    txt4.BackColor = &HFF&
        ElseIf txt4.Value < 41 Then
        txt4.BackColor = &HFF00&
            Else: txt4.BackColor = &HFFFF&
End If
End Sub


Private Sub txt1_Change()
If txt1.Value > 50 Then
    txt1.BackColor = &HFF&
        ElseIf txt1.Value < 41 Then
        txt1.BackColor = &HFF00&
            Else: txt1.BackColor = &HFFFF&
End If
End Sub

Private Sub txt2x6_Change()
If txt2x6.Value > 50 Then
    txt2x6.BackColor = &HFF&
        ElseIf txt2x6.Value < 41 Then
        txt2x6.BackColor = &HFF00&
            Else: txt2x6.BackColor = &HFFFF&
End If
End Sub

Private Sub txt4x12_Change()
If txt4x12.Value > 50 Then
    txt4x12.BackColor = &HFF&
        ElseIf txt4x12.Value < 41 Then
        txt4x12.BackColor = &HFF00&
            Else: txt4x12.BackColor = &HFFFF&
End If
End Sub

Private Sub txt4x18_Change()
If txt4x18.Value > 50 Then
    txt4x18.BackColor = &HFF&
        ElseIf txt4x18.Value < 41 Then
        txt4x18.BackColor = &HFF00&
            Else: txt4x18.BackColor = &HFFFF&
End If
End Sub

Private Sub txt4x24_Change()
If txt4x24.Value > 50 Then
    txt4x24.BackColor = &HFF&
        ElseIf txt4x24.Value < 41 Then
        txt4x24.BackColor = &HFF00&
            Else: txt4x24.BackColor = &HFFFF&
End If
End Sub

Private Sub txt4x6_Change()
If txt4x6.Value > 50 Then
    txt4x6.BackColor = &HFF&
        ElseIf txt4x6.Value < 41 Then
        txt4x6.BackColor = &HFF00&
            Else: txt4x6.BackColor = &HFFFF&
End If
End Sub

Private Sub txt4x8_Change()
If txt4x8.Value > 50 Then
    txt4x8.BackColor = &HFF&
        ElseIf txt4x8.Value < 41 Then
        txt4x8.BackColor = &HFF00&
            Else: txt4x8.BackColor = &HFFFF&
End If
End Sub

Private Sub txt6x12_Change()
If txt6x12.Value > 50 Then
    txt6x12.BackColor = &HFF&
        ElseIf txt6x12.Value < 41 Then
        txt6x12.BackColor = &HFF00&
            Else: txt6x12.BackColor = &HFFFF&
End If
End Sub

Private Sub txt6x18_Change()
If txt6x18.Value > 50 Then
    txt6x18.BackColor = &HFF&
        ElseIf txt6x18.Value < 41 Then
        txt6x18.BackColor = &HFF00&
            Else: txt6x18.BackColor = &HFFFF&
End If
End Sub

Private Sub txt6x24_Change()
If txt6x24.Value > 50 Then
    txt6x24.BackColor = &HFF&
        ElseIf txt6x24.Value < 41 Then
        txt6x24.BackColor = &HFF00&
            Else: txt6x24.BackColor = &HFFFF&
End If
End Sub

Private Sub txtCableAreaValue_Change()
On Error Resume Next
Call txtInput_Change
End Sub

Private Sub txtCableAreaValue2_Change()
On Error Resume Next
Call txtInput2_Change
End Sub

Private Sub txtCableAreaValue3_Change()
On Error Resume Next
Call txtInput3_Change
End Sub

Private Sub txtCableAreaValue4_Change()
On Error Resume Next
Call txtInput4_Change
End Sub

Private Sub txtCableAreaValue5_Change()
On Error Resume Next
Call txtInput5_Change
End Sub

Private Sub txtCableAreaValue6_Change()
On Error Resume Next
Call txtInput6_Change
End Sub

Private Sub txtInput_Change()
If Me.txtInput.Text = "" Then
    Me.txtInput.Value = 0
End If
If IsNumeric(Me.txtInput) Then
    Me.txt2x6.Text = Int(((txtInput * txtCableAreaValue) / (12)) * 100)
    Me.txt4x6.Text = Int(((txtInput * txtCableAreaValue) / (24)) * 100)
    Me.txt4x8.Text = Int(((txtInput * txtCableAreaValue) / (32)) * 100)
    Me.txt4x12.Text = Int(((txtInput * txtCableAreaValue) / (48)) * 100)
    Me.txt4x18.Text = Int(((txtInput * txtCableAreaValue) / (72)) * 100)
    Me.txt4x24.Text = Int(((txtInput * txtCableAreaValue) / (96)) * 100)
    Me.txt6x12.Text = Int(((txtInput * txtCableAreaValue) / (72)) * 100)
    Me.txt6x18.Text = Int(((txtInput * txtCableAreaValue) / (108)) * 100)
    Me.txt6x24.Text = Int(((txtInput * txtCableAreaValue) / (144)) * 100)
    Me.txt1.Text = Int(((txtInput * txtCableAreaValue) / (0.785)) * 100)
    Me.txt1_25.Text = Int(((txtInput * txtCableAreaValue) / (1.226)) * 100)
    Me.txt1_5.Text = Int(((txtInput * txtCableAreaValue) / (1.766)) * 100)
    Me.txt2.Text = Int(((txtInput * txtCableAreaValue) / (3.14)) * 100)
    Me.txt2_5.Text = Int(((txtInput * txtCableAreaValue) / (4.906)) * 100)
    Me.txt3.Text = Int(((txtInput * txtCableAreaValue) / (7.065)) * 100)
    Me.txt4.Text = Int(((txtInput * txtCableAreaValue) / (12.56)) * 100)
End If
End Sub

Private Sub txtInput2_Change()
If Me.txtInput2.Text = "" Then
    Me.txtInput2.Value = 0
End If
If Me.txtInput2.Visible = False Then
    Me.txtInput2.Value = 0
End If
If IsNumeric(Me.txtInput2) Then
    Me.txt2x6.Text = Int((((txtInput2 * txtCableAreaValue2) / (12)) * 100) + ((txtInput * txtCableAreaValue) / (12)) * 100)
    Me.txt4x6.Text = Int((((txtInput2 * txtCableAreaValue2) / (24)) * 100) + ((txtInput * txtCableAreaValue) / (24)) * 100)
    Me.txt4x8.Text = Int((((txtInput2 * txtCableAreaValue2) / (32)) * 100) + ((txtInput * txtCableAreaValue) / (32)) * 100)
    Me.txt4x12.Text = Int((((txtInput2 * txtCableAreaValue2) / (48)) * 100) + ((txtInput * txtCableAreaValue) / (48)) * 100)
    Me.txt4x18.Text = Int((((txtInput2 * txtCableAreaValue2) / (72)) * 100) + ((txtInput * txtCableAreaValue) / (72)) * 100)
    Me.txt4x24.Text = Int((((txtInput2 * txtCableAreaValue2) / (96)) * 100) + ((txtInput * txtCableAreaValue) / (96)) * 100)
    Me.txt6x12.Text = Int((((txtInput2 * txtCableAreaValue2) / (72)) * 100) + ((txtInput * txtCableAreaValue) / (72)) * 100)
    Me.txt6x18.Text = Int((((txtInput2 * txtCableAreaValue2) / (108)) * 100) + ((txtInput * txtCableAreaValue) / (108)) * 100)
    Me.txt6x24.Text = Int((((txtInput2 * txtCableAreaValue2) / (144)) * 100) + ((txtInput * txtCableAreaValue) / (144)) * 100)
    Me.txt1.Text = Int((((txtInput2 * txtCableAreaValue2) / (0.785)) * 100) + ((txtInput * txtCableAreaValue) / (0.785)) * 100)
    Me.txt1_25.Text = Int((((txtInput2 * txtCableAreaValue2) / (1.226)) * 100) + ((txtInput * txtCableAreaValue) / (1.226)) * 100)
    Me.txt1_5.Text = Int((((txtInput2 * txtCableAreaValue2) / (1.766)) * 100) + ((txtInput * txtCableAreaValue) / (1.766)) * 100)
    Me.txt2.Text = Int((((txtInput2 * txtCableAreaValue2) / (3.14)) * 100) + ((txtInput * txtCableAreaValue) / (3.14)) * 100)
    Me.txt2_5.Text = Int((((txtInput2 * txtCableAreaValue2) / (4.906)) * 100) + ((txtInput * txtCableAreaValue) / (4.906)) * 100)
    Me.txt3.Text = Int((((txtInput2 * txtCableAreaValue2) / (7.065)) * 100) + ((txtInput * txtCableAreaValue) / (7.065)) * 100)
    Me.txt4.Text = Int((((txtInput2 * txtCableAreaValue2) / (12.56)) * 100) + ((txtInput * txtCableAreaValue) / (12.56)) * 100)
End If
End Sub

Private Sub txtInput3_Change()
If Me.txtInput3.Text = "" Then
    Me.txtInput3.Value = 0
End If
If Me.txtInput3.Visible = False Then
    Me.txtInput3.Value = 0
End If
If IsNumeric(Me.txtInput3) Then
    Me.txt2x6.Text = Int(((((txtInput3 * txtCableAreaValue3) / (12)) * 100) + ((txtInput2 * txtCableAreaValue2) / (12)) * 100) + ((txtInput * txtCableAreaValue) / (12)) * 100)
    Me.txt4x6.Text = Int(((((txtInput3 * txtCableAreaValue3) / (24)) * 100) + ((txtInput2 * txtCableAreaValue2) / (24)) * 100) + ((txtInput * txtCableAreaValue) / (24)) * 100)
    Me.txt4x8.Text = Int(((((txtInput3 * txtCableAreaValue3) / (32)) * 100) + ((txtInput2 * txtCableAreaValue2) / (32)) * 100) + ((txtInput * txtCableAreaValue) / (32)) * 100)
    Me.txt4x12.Text = Int(((((txtInput3 * txtCableAreaValue3) / (48)) * 100) + ((txtInput2 * txtCableAreaValue2) / (48)) * 100) + ((txtInput * txtCableAreaValue) / (48)) * 100)
    Me.txt4x18.Text = Int(((((txtInput3 * txtCableAreaValue3) / (72)) * 100) + ((txtInput2 * txtCableAreaValue2) / (72)) * 100) + ((txtInput * txtCableAreaValue) / (72)) * 100)
    Me.txt4x24.Text = Int(((((txtInput3 * txtCableAreaValue3) / (96)) * 100) + ((txtInput2 * txtCableAreaValue2) / (96)) * 100) + ((txtInput * txtCableAreaValue) / (96)) * 100)
    Me.txt6x12.Text = Int(((((txtInput3 * txtCableAreaValue3) / (72)) * 100) + ((txtInput2 * txtCableAreaValue2) / (72)) * 100) + ((txtInput * txtCableAreaValue) / (72)) * 100)
    Me.txt6x18.Text = Int(((((txtInput3 * txtCableAreaValue3) / (108)) * 100) + ((txtInput2 * txtCableAreaValue2) / (108)) * 100) + ((txtInput * txtCableAreaValue) / (108)) * 100)
    Me.txt6x24.Text = Int(((((txtInput3 * txtCableAreaValue3) / (144)) * 100) + ((txtInput2 * txtCableAreaValue2) / (144)) * 100) + ((txtInput * txtCableAreaValue) / (144)) * 100)
    Me.txt1.Text = Int(((((txtInput3 * txtCableAreaValue3) / (0.785)) * 100) + ((txtInput2 * txtCableAreaValue2) / (0.785)) * 100) + ((txtInput * txtCableAreaValue) / (0.785)) * 100)
    Me.txt1_25.Text = Int(((((txtInput3 * txtCableAreaValue3) / (1.226)) * 100) + ((txtInput2 * txtCableAreaValue2) / (1.226)) * 100) + ((txtInput * txtCableAreaValue) / (1.226)) * 100)
    Me.txt1_5.Text = Int(((((txtInput3 * txtCableAreaValue3) / (1.766)) * 100) + ((txtInput2 * txtCableAreaValue2) / (1.766)) * 100) + ((txtInput * txtCableAreaValue) / (1.766)) * 100)
    Me.txt2.Text = Int(((((txtInput3 * txtCableAreaValue3) / (3.14)) * 100) + ((txtInput2 * txtCableAreaValue2) / (3.14)) * 100) + ((txtInput * txtCableAreaValue) / (3.14)) * 100)
    Me.txt2_5.Text = Int(((((txtInput3 * txtCableAreaValue3) / (4.906)) * 100) + ((txtInput2 * txtCableAreaValue2) / (4.906)) * 100) + ((txtInput * txtCableAreaValue) / (4.906)) * 100)
    Me.txt3.Text = Int(((((txtInput3 * txtCableAreaValue3) / (7.065)) * 100) + ((txtInput2 * txtCableAreaValue2) / (7.065)) * 100) + ((txtInput * txtCableAreaValue) / (7.065)) * 100)
    Me.txt4.Text = Int(((((txtInput3 * txtCableAreaValue3) / (12.56)) * 100) + ((txtInput2 * txtCableAreaValue2) / (12.56)) * 100) + ((txtInput * txtCableAreaValue) / (12.56)) * 100)
End If
End Sub

Private Sub txtInput4_Change()
If Me.txtInput4.Text = "" Then
    Me.txtInput4.Value = 0
End If
If Me.txtInput4.Visible = False Then
    Me.txtInput4.Value = 0
End If
If IsNumeric(Me.txtInput3) Then
    Me.txt2x6.Text = Int((((((txtInput4 * txtCableAreaValue4) / (12)) * 100) + ((txtInput3 * txtCableAreaValue3) / (12)) * 100) + ((txtInput2 * txtCableAreaValue2) / (12)) * 100) + ((txtInput * txtCableAreaValue) / (12)) * 100)
    Me.txt4x6.Text = Int((((((txtInput4 * txtCableAreaValue4) / (24)) * 100) + ((txtInput3 * txtCableAreaValue3) / (24)) * 100) + ((txtInput2 * txtCableAreaValue2) / (24)) * 100) + ((txtInput * txtCableAreaValue) / (24)) * 100)
    Me.txt4x8.Text = Int((((((txtInput4 * txtCableAreaValue4) / (32)) * 100) + ((txtInput3 * txtCableAreaValue3) / (32)) * 100) + ((txtInput2 * txtCableAreaValue2) / (32)) * 100) + ((txtInput * txtCableAreaValue) / (32)) * 100)
    Me.txt4x12.Text = Int((((((txtInput4 * txtCableAreaValue4) / (48)) * 100) + ((txtInput3 * txtCableAreaValue3) / (48)) * 100) + ((txtInput2 * txtCableAreaValue2) / (48)) * 100) + ((txtInput * txtCableAreaValue) / (48)) * 100)
    Me.txt4x18.Text = Int((((((txtInput4 * txtCableAreaValue4) / (72)) * 100) + ((txtInput3 * txtCableAreaValue3) / (72)) * 100) + ((txtInput2 * txtCableAreaValue2) / (72)) * 100) + ((txtInput * txtCableAreaValue) / (72)) * 100)
    Me.txt4x24.Text = Int((((((txtInput4 * txtCableAreaValue4) / (96)) * 100) + ((txtInput3 * txtCableAreaValue3) / (96)) * 100) + ((txtInput2 * txtCableAreaValue2) / (96)) * 100) + ((txtInput * txtCableAreaValue) / (96)) * 100)
    Me.txt6x12.Text = Int((((((txtInput4 * txtCableAreaValue4) / (72)) * 100) + ((txtInput3 * txtCableAreaValue3) / (72)) * 100) + ((txtInput2 * txtCableAreaValue2) / (72)) * 100) + ((txtInput * txtCableAreaValue) / (72)) * 100)
    Me.txt6x18.Text = Int((((((txtInput4 * txtCableAreaValue4) / (108)) * 100) + ((txtInput3 * txtCableAreaValue3) / (108)) * 100) + ((txtInput2 * txtCableAreaValue2) / (108)) * 100) + ((txtInput * txtCableAreaValue) / (108)) * 100)
    Me.txt6x24.Text = Int((((((txtInput4 * txtCableAreaValue4) / (144)) * 100) + ((txtInput3 * txtCableAreaValue3) / (144)) * 100) + ((txtInput2 * txtCableAreaValue2) / (144)) * 100) + ((txtInput * txtCableAreaValue) / (144)) * 100)
    Me.txt1.Text = Int((((((txtInput4 * txtCableAreaValue4) / (0.785)) * 100) + ((txtInput3 * txtCableAreaValue3) / (0.785)) * 100) + ((txtInput2 * txtCableAreaValue2) / (0.785)) * 100) + ((txtInput * txtCableAreaValue) / (0.785)) * 100)
    Me.txt1_25.Text = Int((((((txtInput4 * txtCableAreaValue4) / (1.226)) * 100) + ((txtInput3 * txtCableAreaValue3) / (1.226)) * 100) + ((txtInput2 * txtCableAreaValue2) / (1.226)) * 100) + ((txtInput * txtCableAreaValue) / (1.226)) * 100)
    Me.txt1_5.Text = Int((((((txtInput4 * txtCableAreaValue4) / (1.766)) * 100) + ((txtInput3 * txtCableAreaValue3) / (1.766)) * 100) + ((txtInput2 * txtCableAreaValue2) / (1.766)) * 100) + ((txtInput * txtCableAreaValue) / (1.766)) * 100)
    Me.txt2.Text = Int((((((txtInput4 * txtCableAreaValue4) / (3.14)) * 100) + ((txtInput3 * txtCableAreaValue3) / (3.14)) * 100) + ((txtInput2 * txtCableAreaValue2) / (3.14)) * 100) + ((txtInput * txtCableAreaValue) / (3.14)) * 100)
    Me.txt2_5.Text = Int((((((txtInput4 * txtCableAreaValue4) / (4.906)) * 100) + ((txtInput3 * txtCableAreaValue3) / (4.906)) * 100) + ((txtInput2 * txtCableAreaValue2) / (4.906)) * 100) + ((txtInput * txtCableAreaValue) / (4.906)) * 100)
    Me.txt3.Text = Int((((((txtInput4 * txtCableAreaValue4) / (7.065)) * 100) + ((txtInput3 * txtCableAreaValue3) / (7.065)) * 100) + ((txtInput2 * txtCableAreaValue2) / (7.065)) * 100) + ((txtInput * txtCableAreaValue) / (7.065)) * 100)
    Me.txt4.Text = Int((((((txtInput4 * txtCableAreaValue4) / (12.56)) * 100) + ((txtInput3 * txtCableAreaValue3) / (12.56)) * 100) + ((txtInput2 * txtCableAreaValue2) / (12.56)) * 100) + ((txtInput * txtCableAreaValue) / (12.56)) * 100)
End If
End Sub

Private Sub txtInput5_Change()
If Me.txtInput5.Text = "" Then
    Me.txtInput5.Value = 0
End If
If Me.txtInput5.Visible = False Then
    Me.txtInput5.Value = 0
End If
If IsNumeric(Me.txtInput3) Then
    Me.txt2x6.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (12)) * 100) + ((txtInput4 * txtCableAreaValue4) / (12)) * 100) + ((txtInput3 * txtCableAreaValue3) / (12)) * 100) + ((txtInput2 * txtCableAreaValue2) / (12)) * 100) + ((txtInput * txtCableAreaValue) / (12)) * 100)
    Me.txt4x6.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (24)) * 100) + ((txtInput4 * txtCableAreaValue4) / (24)) * 100) + ((txtInput3 * txtCableAreaValue3) / (24)) * 100) + ((txtInput2 * txtCableAreaValue2) / (24)) * 100) + ((txtInput * txtCableAreaValue) / (24)) * 100)
    Me.txt4x8.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (32)) * 100) + ((txtInput4 * txtCableAreaValue4) / (32)) * 100) + ((txtInput3 * txtCableAreaValue3) / (32)) * 100) + ((txtInput2 * txtCableAreaValue2) / (32)) * 100) + ((txtInput * txtCableAreaValue) / (32)) * 100)
    Me.txt4x12.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (48)) * 100) + ((txtInput4 * txtCableAreaValue4) / (48)) * 100) + ((txtInput3 * txtCableAreaValue3) / (48)) * 100) + ((txtInput2 * txtCableAreaValue2) / (48)) * 100) + ((txtInput * txtCableAreaValue) / (48)) * 100)
    Me.txt4x18.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (72)) * 100) + ((txtInput4 * txtCableAreaValue4) / (72)) * 100) + ((txtInput3 * txtCableAreaValue3) / (72)) * 100) + ((txtInput2 * txtCableAreaValue2) / (72)) * 100) + ((txtInput * txtCableAreaValue) / (72)) * 100)
    Me.txt4x24.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (96)) * 100) + ((txtInput4 * txtCableAreaValue4) / (96)) * 100) + ((txtInput3 * txtCableAreaValue3) / (96)) * 100) + ((txtInput2 * txtCableAreaValue2) / (96)) * 100) + ((txtInput * txtCableAreaValue) / (96)) * 100)
    Me.txt6x12.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (72)) * 100) + ((txtInput4 * txtCableAreaValue4) / (72)) * 100) + ((txtInput3 * txtCableAreaValue3) / (72)) * 100) + ((txtInput2 * txtCableAreaValue2) / (72)) * 100) + ((txtInput * txtCableAreaValue) / (72)) * 100)
    Me.txt6x18.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (108)) * 100) + ((txtInput4 * txtCableAreaValue4) / (108)) * 100) + ((txtInput3 * txtCableAreaValue3) / (108)) * 100) + ((txtInput2 * txtCableAreaValue2) / (108)) * 100) + ((txtInput * txtCableAreaValue) / (108)) * 100)
    Me.txt6x24.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (144)) * 100) + ((txtInput4 * txtCableAreaValue4) / (144)) * 100) + ((txtInput3 * txtCableAreaValue3) / (144)) * 100) + ((txtInput2 * txtCableAreaValue2) / (144)) * 100) + ((txtInput * txtCableAreaValue) / (144)) * 100)
    Me.txt1.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (0.785)) * 100) + ((txtInput4 * txtCableAreaValue4) / (0.785)) * 100) + ((txtInput3 * txtCableAreaValue3) / (0.785)) * 100) + ((txtInput2 * txtCableAreaValue2) / (0.785)) * 100) + ((txtInput * txtCableAreaValue) / (0.785)) * 100)
    Me.txt1_25.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (1.226)) * 100) + ((txtInput4 * txtCableAreaValue4) / (1.226)) * 100) + ((txtInput3 * txtCableAreaValue3) / (1.226)) * 100) + ((txtInput2 * txtCableAreaValue2) / (1.226)) * 100) + ((txtInput * txtCableAreaValue) / (1.226)) * 100)
    Me.txt1_5.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (1.766)) * 100) + ((txtInput4 * txtCableAreaValue4) / (1.766)) * 100) + ((txtInput3 * txtCableAreaValue3) / (1.766)) * 100) + ((txtInput2 * txtCableAreaValue2) / (1.766)) * 100) + ((txtInput * txtCableAreaValue) / (1.766)) * 100)
    Me.txt2.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (3.14)) * 100) + ((txtInput4 * txtCableAreaValue4) / (3.14)) * 100) + ((txtInput3 * txtCableAreaValue3) / (3.14)) * 100) + ((txtInput2 * txtCableAreaValue2) / (3.14)) * 100) + ((txtInput * txtCableAreaValue) / (3.14)) * 100)
    Me.txt2_5.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (4.906)) * 100) + ((txtInput4 * txtCableAreaValue4) / (4.906)) * 100) + ((txtInput3 * txtCableAreaValue3) / (4.906)) * 100) + ((txtInput2 * txtCableAreaValue2) / (4.906)) * 100) + ((txtInput * txtCableAreaValue) / (4.906)) * 100)
    Me.txt3.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (7.065)) * 100) + ((txtInput4 * txtCableAreaValue4) / (7.065)) * 100) + ((txtInput3 * txtCableAreaValue3) / (7.065)) * 100) + ((txtInput2 * txtCableAreaValue2) / (7.065)) * 100) + ((txtInput * txtCableAreaValue) / (7.065)) * 100)
    Me.txt4.Text = Int(((((((txtInput5 * txtCableAreaValue5) / (12.56)) * 100) + ((txtInput4 * txtCableAreaValue4) / (12.56)) * 100) + ((txtInput3 * txtCableAreaValue3) / (12.56)) * 100) + ((txtInput2 * txtCableAreaValue2) / (12.56)) * 100) + ((txtInput * txtCableAreaValue) / (12.56)) * 100)
End If
End Sub

Private Sub txtInput6_Change()
If Me.txtInput6.Text = "" Then
    Me.txtInput6.Value = 0
End If
If Me.txtInput6.Visible = False Then
    Me.txtInput6.Value = 0
End If
If IsNumeric(Me.txtInput3) Then
    Me.txt2x6.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (12)) * 100) + ((txtInput5 * txtCableAreaValue5) / (12)) * 100) + ((txtInput4 * txtCableAreaValue4) / (12)) * 100) + ((txtInput3 * txtCableAreaValue3) / (12)) * 100) + ((txtInput2 * txtCableAreaValue2) / (12)) * 100) + ((txtInput * txtCableAreaValue) / (12)) * 100)
    Me.txt4x6.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (24)) * 100) + ((txtInput5 * txtCableAreaValue5) / (24)) * 100) + ((txtInput4 * txtCableAreaValue4) / (24)) * 100) + ((txtInput3 * txtCableAreaValue3) / (24)) * 100) + ((txtInput2 * txtCableAreaValue2) / (24)) * 100) + ((txtInput * txtCableAreaValue) / (24)) * 100)
    Me.txt4x8.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (32)) * 100) + ((txtInput5 * txtCableAreaValue5) / (32)) * 100) + ((txtInput4 * txtCableAreaValue4) / (32)) * 100) + ((txtInput3 * txtCableAreaValue3) / (32)) * 100) + ((txtInput2 * txtCableAreaValue2) / (32)) * 100) + ((txtInput * txtCableAreaValue) / (32)) * 100)
    Me.txt4x12.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (48)) * 100) + ((txtInput5 * txtCableAreaValue5) / (48)) * 100) + ((txtInput4 * txtCableAreaValue4) / (48)) * 100) + ((txtInput3 * txtCableAreaValue3) / (48)) * 100) + ((txtInput2 * txtCableAreaValue2) / (48)) * 100) + ((txtInput * txtCableAreaValue) / (48)) * 100)
    Me.txt4x18.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (72)) * 100) + ((txtInput5 * txtCableAreaValue5) / (72)) * 100) + ((txtInput4 * txtCableAreaValue4) / (72)) * 100) + ((txtInput3 * txtCableAreaValue3) / (72)) * 100) + ((txtInput2 * txtCableAreaValue2) / (72)) * 100) + ((txtInput * txtCableAreaValue) / (72)) * 100)
    Me.txt4x24.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (96)) * 100) + ((txtInput5 * txtCableAreaValue5) / (96)) * 100) + ((txtInput4 * txtCableAreaValue4) / (96)) * 100) + ((txtInput3 * txtCableAreaValue3) / (96)) * 100) + ((txtInput2 * txtCableAreaValue2) / (96)) * 100) + ((txtInput * txtCableAreaValue) / (96)) * 100)
    Me.txt6x12.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (72)) * 100) + ((txtInput5 * txtCableAreaValue5) / (72)) * 100) + ((txtInput4 * txtCableAreaValue4) / (72)) * 100) + ((txtInput3 * txtCableAreaValue3) / (72)) * 100) + ((txtInput2 * txtCableAreaValue2) / (72)) * 100) + ((txtInput * txtCableAreaValue) / (72)) * 100)
    Me.txt6x18.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (108)) * 100) + ((txtInput5 * txtCableAreaValue5) / (108)) * 100) + ((txtInput4 * txtCableAreaValue4) / (108)) * 100) + ((txtInput3 * txtCableAreaValue3) / (108)) * 100) + ((txtInput2 * txtCableAreaValue2) / (108)) * 100) + ((txtInput * txtCableAreaValue) / (108)) * 100)
    Me.txt6x24.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (144)) * 100) + ((txtInput5 * txtCableAreaValue5) / (144)) * 100) + ((txtInput4 * txtCableAreaValue4) / (144)) * 100) + ((txtInput3 * txtCableAreaValue3) / (144)) * 100) + ((txtInput2 * txtCableAreaValue2) / (144)) * 100) + ((txtInput * txtCableAreaValue) / (144)) * 100)
    Me.txt1.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (0.785)) * 100) + ((txtInput5 * txtCableAreaValue5) / (0.785)) * 100) + ((txtInput4 * txtCableAreaValue4) / (0.785)) * 100) + ((txtInput3 * txtCableAreaValue3) / (0.785)) * 100) + ((txtInput2 * txtCableAreaValue2) / (0.785)) * 100) + ((txtInput * txtCableAreaValue) / (0.785)) * 100)
    Me.txt1_25.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (1.226)) * 100) + ((txtInput5 * txtCableAreaValue5) / (1.226)) * 100) + ((txtInput4 * txtCableAreaValue4) / (1.226)) * 100) + ((txtInput3 * txtCableAreaValue3) / (1.226)) * 100) + ((txtInput2 * txtCableAreaValue2) / (1.226)) * 100) + ((txtInput * txtCableAreaValue) / (1.226)) * 100)
    Me.txt1_5.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (1.766)) * 100) + ((txtInput5 * txtCableAreaValue5) / (1.766)) * 100) + ((txtInput4 * txtCableAreaValue4) / (1.766)) * 100) + ((txtInput3 * txtCableAreaValue3) / (1.766)) * 100) + ((txtInput2 * txtCableAreaValue2) / (1.766)) * 100) + ((txtInput * txtCableAreaValue) / (1.766)) * 100)
    Me.txt2.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (3.14)) * 100) + ((txtInput5 * txtCableAreaValue5) / (3.14)) * 100) + ((txtInput4 * txtCableAreaValue4) / (3.14)) * 100) + ((txtInput3 * txtCableAreaValue3) / (3.14)) * 100) + ((txtInput2 * txtCableAreaValue2) / (3.14)) * 100) + ((txtInput * txtCableAreaValue) / (3.14)) * 100)
    Me.txt2_5.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (4.906)) * 100) + ((txtInput5 * txtCableAreaValue5) / (4.906)) * 100) + ((txtInput4 * txtCableAreaValue4) / (4.906)) * 100) + ((txtInput3 * txtCableAreaValue3) / (4.906)) * 100) + ((txtInput2 * txtCableAreaValue2) / (4.906)) * 100) + ((txtInput * txtCableAreaValue) / (4.906)) * 100)
    Me.txt3.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (7.065)) * 100) + ((txtInput5 * txtCableAreaValue5) / (7.065)) * 100) + ((txtInput4 * txtCableAreaValue4) / (7.065)) * 100) + ((txtInput3 * txtCableAreaValue3) / (7.065)) * 100) + ((txtInput2 * txtCableAreaValue2) / (7.065)) * 100) + ((txtInput * txtCableAreaValue) / (7.065)) * 100)
    Me.txt4.Text = Int((((((((txtInput6 * txtCableAreaValue6) / (12.56)) * 100) + ((txtInput5 * txtCableAreaValue5) / (12.56)) * 100) + ((txtInput4 * txtCableAreaValue4) / (12.56)) * 100) + ((txtInput3 * txtCableAreaValue3) / (12.56)) * 100) + ((txtInput2 * txtCableAreaValue2) / (12.56)) * 100) + ((txtInput * txtCableAreaValue) / (12.56)) * 100)
End If
End Sub

Private Sub UserForm_initialize()
With cboCable
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
                                            .AddItem "Security_Composite"
                                                .AddItem "Outdoor_fiber_ALTOS"
                                                    .AddItem "Outdoor_fiber_FREEDM"
End With
With cboCable2
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
                                            .AddItem "Security_Composite"
                                                .AddItem "Outdoor_fiber_ALTOS"
                                                    .AddItem "Outdoor_fiber_FREEDM"
End With
With cboCable3
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
                                            .AddItem "Security_Composite"
                                                .AddItem "Outdoor_fiber_ALTOS"
                                                    .AddItem "Outdoor_fiber_FREEDM"
End With
With cboCable4
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
                                            .AddItem "Security_Composite"
                                                .AddItem "Outdoor_fiber_ALTOS"
                                                    .AddItem "Outdoor_fiber_FREEDM"
End With
With cboCable5
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
                                            .AddItem "Security_Composite"
                                                .AddItem "Outdoor_fiber_ALTOS"
                                                    .AddItem "Outdoor_fiber_FREEDM"
End With
With cboCable6
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
                                            .AddItem "Security_Composite"
                                                .AddItem "Outdoor_fiber_ALTOS"
                                                    .AddItem "Outdoor_fiber_FREEDM"
End With
End Sub

