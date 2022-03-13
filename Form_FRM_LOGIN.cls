VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FRM_LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CMD_LOGAR_Click()
    Dim CNN As ADODB.Connection
    Dim RS As New ADODB.Recordset
    
    Set CNN = CurrentProject.Connection
    
    RS.Open "SELECT * FROM TBL_LOGIN", CNN, adOpenKeyset, adLockOptimistic
    RS.AddNew
    RS!LOGIN = LOGIN
    RS!SENHA = SENHA
    RS.Update
    
    LOGIN = Empty
    SENHA = Empty
    
    RS.Close
    CNN.Close
    
    DoCmd.OpenForm "FRM_MENU_PRINCIPAL"
    
End Sub
