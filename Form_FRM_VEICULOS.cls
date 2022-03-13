VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FRM_VEICULOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CMD_CADASTRAR_VEICULOS_Click()
    Dim CNN As ADODB.Connection
    Dim RS As New ADODB.Recordset
    
    Set CNN = CurrentProject.Connection
    
    RS.Open "SELECT * FROM TBL_VEICULOS", CNN, adOpenKeyset, adLockOptimistic
    RS.AddNew
    RS!NUM_ID_LOC = NUM_ID_LOC.Column(0)
    RS!STR_MARCA = STR_MARCA
    RS!STR_PLACA = STR_PLACA
    RS!VLR_DIARIA = VLR_DIARIA
    RS.Update
    
    NUM_ID_LOC = Empty
    STR_MARCA = Empty
    STR_PLACA = Empty
    VLR_DIARIA = Empty
            
            
    RS.Close
    CNN.Close
    
End Sub
