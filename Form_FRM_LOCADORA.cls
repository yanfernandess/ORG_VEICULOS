VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FRM_LOCADORA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CMD_CADASTRAR_LOCADORA_Click()
    Dim CNN As ADODB.Connection
    Dim RS As New ADODB.Recordset
    
    Set CNN = CurrentProject.Connection
    
    RS.Open "SELECT * FROM TBL_LOCADORA", CNN, adOpenKeyset, adLockOptimistic
    RS.AddNew
    RS!STR_NOME = STR_NOME
    RS!STR_END = STR_END
    RS.Update
    
    STR_NOME = Empty
    STR_END = Empty
    
    RS.Close
    CNN.Close
    
    
    
End Sub
