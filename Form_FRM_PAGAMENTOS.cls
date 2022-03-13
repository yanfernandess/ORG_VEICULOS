VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FRM_PAGAMENTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CMD_PAGAGAMENTO_Click()
    Dim CNN As ADODB.Connection
    Dim RS As New ADODB.Recordset
    
    Set CNN = CurrentProject.Connection
    
    RS.Open "SELECT * FROM TBL_PAGAMENTOS", CNN, adOpenKeyset, adLockOptimistic
    RS.AddNew
    RS!NUM_ID_CHECK_IN = NUM_ID_CHECK_IN
    RS!DT_PAGTO = DT_PAGTO
    RS!VLR_PAGTO = VLR_PAGTO
    RS.Update
    
    NUM_ID_CHECK_IN = Empty
    DT_PAGTO = Empty
    VLR_PAGTO = Empty
    
    RS.Close
    CNN.Close
    
End Sub
