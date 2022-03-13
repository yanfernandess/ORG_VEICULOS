VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FRM_MENU_PRINCIPAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub CMD_ALUGUEL_Click()
    DoCmd.OpenForm "FRM_ALUGUEL"
End Sub

Private Sub CMD_FUNCIONARIO_Click()
    DoCmd.OpenForm "FRM_FUNCIONARIO"
End Sub

Private Sub CMD_LOCADORA_Click()
    DoCmd.OpenForm "FRM_LOCADORA"
End Sub

Private Sub CMD_PAGAMENTOS_Click()
    DoCmd.OpenForm "FRM_PAGAMENTOS"
End Sub

Private Sub CMD_VEICULOS_Click()
    DoCmd.OpenForm "FRM_VEICULOS"
End Sub
