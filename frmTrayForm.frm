VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTrayForm 
   Caption         =   "Tray Form"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6585
   OleObjectBlob   =   "frmTrayForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTrayForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboBends_Change()
If cboBends.Text = "1" Then
Me.txtDerated.Text = (Me.lblActualConArea * 0.85) / (Me.txtCableAreaValue)
    ElseIf cboBends.Text = "2" Then
    Me.txtDerated.Text = (Me.lblActualConArea * 0.7) / (Me.txtCableAreaValue)
        ElseIf cboBends.Text = "3" Then
        Me.txtDerated.Text = (Me.lblActualConArea * 0.55) / (Me.txtCableAreaValue)
End If
On Error Resume Next
If IsNumeric(Me.txtDerated) Then
Me.lblFillQtyDeRated.Caption = ((Me.txtDerated * Me.txtCableAreaValue) / (Me.lblActualConArea) * 100)
End If
lblDPercent.Visible = True
End Sub

Private Sub cboConduitChoices_Change()
If cboConduitChoices.Text = "1""" Then
Me.lblConArea.Caption = "0.785"
    ElseIf cboConduitChoices.Text = "1.25""" Then
    Me.lblConArea.Caption = "1.226"
        ElseIf cboConduitChoices.Text = "1.5""" Then
        Me.lblConArea.Caption = "1.766"
            ElseIf cboConduitChoices.Text = "2""" Then
            Me.lblConArea.Caption = "3.14"
                ElseIf cboConduitChoices.Text = "2.5""" Then
                Me.lblConArea.Caption = "4.906"
                    ElseIf cboConduitChoices.Text = "3""" Then
                    Me.lblConArea.Caption = "7.065"
                        ElseIf cboConduitChoices.Text = "4""" Then
                        Me.lblConArea.Caption = "12.56"
End If
If IsNumeric(Me.lblConArea) Then
    Me.lblActualConArea.Caption = Me.lblConArea * 0.4
End If
If IsNumeric(Me.txtCableAreaValue) Then
Me.txtQuantity.Text = Int((Me.lblActualConArea) / (Me.txtCableAreaValue))
End If
lblConVal.Visible = True
lblTrayVal.Visible = False
On Error Resume Next
If IsNumeric(Me.txtQuantity) Then
Me.lblConFillRatio.Caption = ((Me.txtQuantity * Me.txtCableAreaValue) / (Me.lblActualConArea) * 100)
End If
lblConFillPercent.Visible = True
End Sub

Private Sub cboTrayCableTypes_Change()
If cboTrayCableTypes.Text = "LM1000_Cat6_Riser" Then
Me.txtCableAreaValue.Text = "0.040"
    ElseIf cboTrayCableTypes.Text = "LM1000_Cat6_Plenum" Then
    Me.txtCableAreaValue.Text = "0.040"
        ElseIf cboTrayCableTypes.Text = "Clarity_Modular_Patch_Cords" Then
        Me.txtCableAreaValue.Text = "0.066"
            ElseIf cboTrayCableTypes.Text = "Cat5e_25_pair" Then
            Me.txtCableAreaValue.Text = "0.169"
                ElseIf cboTrayCableTypes.Text = "Corning_048T88-61180-A3" Then
                Me.txtCableAreaValue.Text = "0.816"
                    ElseIf cboTrayCableTypes.Text = "Corning_48T8F-31190-A1" Then
                    Me.txtCableAreaValue.Text = ".0882"
                        ElseIf cboTrayCableTypes.Text = "Corning_024T88-33180" Then
                        Me.txtCableAreaValue.Text = "0.075"
                            ElseIf cboTrayCableTypes.Text = "Corning_012T88-33180" Then
                            Me.txtCableAreaValue.Text = "0.045"
                                ElseIf cboTrayCableTypes.Text = "Corning_x757512QPNNDUxxxf" Then
                                Me.txtCableAreaValue.Text = "0.022"
                                    ElseIf cboTrayCableTypes.Text = "Corning_x757524OPNNDUxxxf" Then
                                    Me.txtCableAreaValue.Text = "0.049"
                                        ElseIf cboTrayCableTypes.Text = "Corning_x757548QPNNDUxxxf" Then
                                        Me.txtCableAreaValue.Text = "0.070"
                                            ElseIf cboTrayCableTypes.Text = "Security_Composite" Then
                                            Me.txtCableAreaValue.Text = "0.131"
                                                ElseIf cboTrayCableTypes.Text = "Outdoor_fiber_ALTOS" Then
                                                Me.txtCableAreaValue.Text = "0.180"
                                                    ElseIf cboTrayCableTypes.Text = "Outdoor_fiber_FREEDM" Then
                                                    Me.txtCableAreaValue.Text = "0.066"
End If
End Sub

Private Sub cboTrayChoices_Change()
If cboTrayChoices.Text = "2''x6''" Then
Me.lblTrayArea.Caption = "12"
    ElseIf cboTrayChoices.Text = "4''x6''" Then
    Me.lblTrayArea.Caption = "24"
        ElseIf cboTrayChoices.Text = "4''x8''" Then
        Me.lblTrayArea.Caption = "32"
            ElseIf cboTrayChoices.Text = "4''x12''" Then
            Me.lblTrayArea.Caption = "48"
                ElseIf cboTrayChoices.Text = "4''x18''" Then
                Me.lblTrayArea.Caption = "72"
                    ElseIf cboTrayChoices.Text = "4''x24''" Then
                    Me.lblTrayArea.Caption = "96"
                        ElseIf cboTrayChoices.Text = "6''x12''" Then
                        Me.lblTrayArea.Caption = "72"
                            ElseIf cboTrayChoices.Text = "6''x18''" Then
                            Me.lblTrayArea.Caption = "108"
                                ElseIf cboTrayChoices.Text = "6''x24''" Then
                                Me.lblTrayArea.Caption = "144"
End If
If IsNumeric(Me.lblTrayArea) Then
    Me.lblTrayActualArea.Caption = Me.lblTrayArea * 0.4
End If
If IsNumeric(Me.txtCableAreaValue) Then
Me.txtQuantity.Text = (Me.lblTrayActualArea) / (Me.txtCableAreaValue)
End If
lblTrayVal.Visible = True
lblConVal.Visible = False
On Error Resume Next
If IsNumeric(Me.txtQuantity) Then
Me.lblFillRatio.Caption = ((Me.txtQuantity * Me.txtCableAreaValue) / (Me.lblTrayActualArea) * 100)
End If
lblFillPercent.Visible = True
End Sub

Private Sub cmdCableHelp_Click()
msgbox "In order to change the cable in which the measurements are being calculated you must select one from the drop down box. Additionally you may input your own custom cable area in the case that your desired cable/cable area is not on the list."
End Sub

Private Sub cmdCHelp_Click()
msgbox "Select a common diameter of a Conduit or fill in a custom measurement if using diameter. If using radius select the button to the top-right corner of the screen promting radius. After selecting or filling your desired Conduit measurement the area should be calculated within the ''Area'' value. Usable area should also be given, being forty percent of the calculated area, as this represents the maximum fill potential. If/when desired to acquire the fill ratio simply make sure your Quantity box is showing Conduit value's and then click the ''Conduit Fill Ratio'' Button."
End Sub

Private Sub cmdRadius_Click()
frmRadiusConduit.Show
End Sub

Private Sub cmdTHelp_Click()
msgbox "Please select a common tray size or fill in your own custom sizes. The area of your measurements should be given within the ''Area'' output value. Within the ''Usable Area'' value will be forty percent of the given area, as this is the required maximum fill potential. If/when so desired to acquire the fill ratio simply make sure the Quantity box is show Tray Quantities and click the ''Tray Fill Ratio'' button."
End Sub

Private Sub txtCableAreaValue_Change()
On Error Resume Next
If IsNumeric(Me.txtCableAreaValue) Then
Me.txtQuantity.Text = (Me.lblTrayActualArea) / (Me.txtCableAreaValue)
End If
End Sub

Private Sub txtDiameter_Change()
If IsNumeric(Me.txtDiameter) Then
Me.lblConArea.Caption = (Me.txtDiameter / 2) * (Me.txtDiameter / 2) * 3.14
End If
If IsNumeric(Me.lblConArea) Then
    Me.lblActualConArea.Caption = Me.lblConArea * 0.4
End If
If IsNumeric(Me.lblConArea) Then
    Me.lblActualConArea.Caption = Me.lblConArea * 0.4
End If
If IsNumeric(Me.txtCableAreaValue) Then
Me.txtQuantity.Text = Int((Me.lblActualConArea) / (Me.txtCableAreaValue))
End If
lblConVal.Visible = True
lblTrayVal.Visible = False
If IsNumeric(Me.txtCableAreaValue) Then
Me.txtDerated.Text = (Me.lblActualConArea * 0.85) / (Me.txtCableAreaValue)
End If
End Sub

Private Sub txtLength_Change()
If IsNumeric(Me.txtLength) And IsNumeric(Me.txtWidth) Then
Me.lblTrayArea.Caption = Me.txtLength * Me.txtWidth
End If
If IsNumeric(Me.lblTrayArea) Then
    Me.lblTrayActualArea.Caption = Me.lblTrayArea * 0.4
End If
If IsNumeric(Me.txtCableAreaValue) Then
Me.txtQuantity.Text = (Me.lblTrayActualArea) / (Me.txtCableAreaValue)
End If
lblTrayVal.Visible = True
lblConVal.Visible = False
End Sub

Private Sub txtWidth_Change()
Call txtLength_Change
End Sub

Private Sub UserForm_initialize()
With cboBends
.AddItem "0"
    .AddItem "1"
        .AddItem "2"
            .AddItem "3"
End With
With cboTrayCableTypes
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
With cboConduitChoices
.AddItem "1"""
    .AddItem "1.25"""
        .AddItem "1.5"""
            .AddItem "2"""
                .AddItem "2.5"""
                    .AddItem "3"""
                        .AddItem "4"""
End With
With cboTrayChoices
.AddItem "2''x6''"
    .AddItem "4''x6''"
        .AddItem "4''x8''"
            .AddItem "4''x12''"
                .AddItem "4''x18''"
                    .AddItem "4''x24''"
                        .AddItem "6''x12''"
                            .AddItem "6''x18''"
                                .AddItem "6''x24''"
End With
End Sub
