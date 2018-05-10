Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private mstrSearchType As String
Private mstrSearchTable As String

Property Let SearchType(data As String)
    mstrSearchType = data
End Property

Property Get SearchType() As String
    SearchType = mstrSearchType
End Property

Property Let SearchTable(data As String)
    mstrSearchTable = data
End Property

Property Get SearchTable() As String
    SearchTable = mstrSearchTable
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    'Ctrl Shift C is the equivalent to double clicking the CnlyClaimNum or ICN field
    If KeyCode = vbKeyC And Shift = acShiftMask + acCtrlMask Then
        For i = 0 To Me.Controls.Count - 1
            If Me.Controls(i).ControlType = acTextBox Then
                If Me.Controls(i).ControlSource = "CnlyClaimNum" Then
                    OpenScreenByField Me.Controls(i)
                End If
            End If
        Next
    ElseIf KeyCode = vbKeyX And Shift = acShiftMask Then

        If TypeOf Me.RecordSet Is ADODB.RecordSet Then
            
        Else
            Call ExportRsToExcel(Me.RecordSet)
        End If
    
    End If
End Sub

Private Sub Text1_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text2_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub

Private Sub Text3_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub

Private Sub Text4_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub

Private Sub Text5_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub

Private Sub Text6_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text7_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text8_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text9_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text10_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text11_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text12_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text13_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text14_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text15_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text16_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text17_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text18_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text19_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text20_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text21_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text22_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text23_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text24_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text25_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text26_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text27_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text28_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text29_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text30_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text31_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text32_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text33_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text34_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text35_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text36_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text37_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text38_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text39_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text40_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text41_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text42_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text43_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text44_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text45_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text46_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text47_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text48_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text49_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text50_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text51_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text52_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text53_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text54_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text55_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text56_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text57_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text58_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text59_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text60_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text61_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text62_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text63_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text64_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text65_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text66_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text67_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text68_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text69_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text70_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text71_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text72_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text73_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text74_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text75_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text76_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text77_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text78_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text79_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text80_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text81_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text82_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text83_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text84_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text85_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text86_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text87_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text88_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text89_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text90_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text91_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text92_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text93_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text94_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text95_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text96_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text97_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text98_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text99_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text100_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text101_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text102_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text103_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text104_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text105_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text106_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text107_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text108_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text109_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text110_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text111_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text112_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text113_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text114_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text115_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text116_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text117_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text118_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text119_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text120_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text121_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text122_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text123_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text124_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text125_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text126_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text127_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text128_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text129_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text130_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text131_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text132_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text133_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text134_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text135_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text136_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text137_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text138_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text139_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text140_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text141_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text142_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text143_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text144_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text145_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text146_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text147_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text148_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text149_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text150_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text151_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text152_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text153_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text154_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text155_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text156_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text157_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text158_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text159_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text160_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text161_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text162_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text163_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text164_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text165_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text166_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text167_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text168_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text169_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text170_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text171_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text172_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text173_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text174_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text175_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text176_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text177_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text178_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text179_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text180_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text181_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text182_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text183_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text184_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text185_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text186_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text187_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text188_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text189_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text190_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text191_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text192_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text193_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text194_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text195_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text196_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text197_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text198_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text199_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text200_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text201_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text202_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text203_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text204_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text205_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text206_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text207_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text208_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text209_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text210_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text211_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text212_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text213_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text214_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text215_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text216_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text217_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text218_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text219_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text220_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text221_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text222_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text223_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text224_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text225_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text226_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text227_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text228_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text229_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text230_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text231_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text232_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text233_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text234_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text235_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text236_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text237_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text238_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text239_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text240_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text241_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text242_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text243_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text244_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text245_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text246_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text247_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text248_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text249_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text250_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub

Private Sub Text251_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text252_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text253_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text254_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text255_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text256_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text257_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text258_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text259_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text260_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text261_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text262_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text263_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text264_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text265_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text266_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text267_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text268_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text269_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text270_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text271_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text272_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text273_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text274_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text275_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text276_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text277_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text278_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text279_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text280_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text281_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text282_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text283_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text284_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text285_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text286_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text287_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text288_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text289_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text290_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text291_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text292_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text293_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text294_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text295_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text296_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text297_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text298_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text299_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub
Private Sub Text300_DblClick(Cancel As Integer)

    OpenScreenByField Me.ActiveControl

End Sub

Sub OpenScreenByField(CallingTextBox As TextBox)
'this form will perform an action associated with the double clicked field
'since this form will always represent a datasheet view of a report table from SQL the fields will always be textboxes

On Error GoTo ErrorTrap

    'since double click selects the whole field I will keep the same functionality
    'which wont work if the ctrl shift C combination is pressed
    If Me.ActiveControl = CallingTextBox Then
        CallingTextBox.SelStart = 0
        CallingTextBox.SelLength = Len(CallingTextBox.Value)
    End If
    
    'a different action depending of the field associated with the double clicked field
    Select Case CallingTextBox.ControlSource
        Case "CnlyClaimNum" 'Open Claim form
            Navigate "frm_RPT_AccessForm", "AUDITCLM", "DblClick", CallingTextBox.Value
        Case "ICN" 'since this is also a claim number I will pass the value from that field. If cnlyclaimnum does not exist it will triger an error
            Navigate "frm_RPT_AccessForm", "AUDITCLM", "DblClick", Me("CnlyClaimNum")
    End Select
    
    Exit Sub
    
ErrorTrap:
    
    'If CnlyClaimNum field does not exist in the table loaded into this form then an error occurs
    If Err.Number = 2465 Then
        MsgBox "A CnlyClaimNum MUST exist in the report for this feature to work.", vbCritical, "DblClick to Open Claim"
    Else
        MsgBox "Error: " & Err.Number & ", " & Err.Description, vbCritical, "DblClick to Open Claim"
        Resume Next
    End If
    
End Sub
