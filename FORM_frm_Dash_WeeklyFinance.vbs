Version =20
VersionRequired =20
Checksum =-1404227343
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9420
    DatasheetFontHeight =10
    ItemSuffix =22
    Left =7155
    Top =1485
    Right =16545
    Bottom =9990
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf4ab1282351be440
    End
    RecordSource ="SELECT * FROM v_REPORT_PAYER_TRENDS ORDER BY ROWID DESC"
    Caption ="v_REPORT_PAYER_TRENDS"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    NoSaveCTIWhenDisabled =1
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =780
            Name ="FormHeader"
            GUID = Begin
                0x81f05abc5ed3ee409176e54ea54d7c05
            End
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =540
                    Width =1320
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="reportwk_Label"
                    Caption ="Week"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x049a84e7cfc98c4c9ca92cd4d272fab3
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1500
                    Top =540
                    Width =1500
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="SentAmt_Label"
                    Caption ="Sent"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0xb4e1aecd20a9e742822f8800efd61a20
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3960
                    Top =540
                    Width =1680
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Arsetupamt_Label"
                    Caption ="Setup"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x46a2fd234e183c47927259af6ea91f7a
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =6540
                    Top =540
                    Width =1860
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="collectedamt_Label"
                    Caption ="Collected"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0xcd77e5539acbdd49963c45264182ea45
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3060
                    Top =540
                    Width =780
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label16"
                    Caption ="Sent"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x351b2d1f7a29e9438dbc86996157a787
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5700
                    Top =540
                    Width =780
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label18"
                    Caption ="Setup"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x2d6161ba5cbc2347aec32195067ae8d7
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =8520
                    Top =540
                    Width =900
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label20"
                    Caption ="Collected"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x1759a9c9c5381342bcd6ef63b72adaa6
                    End
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =-15
                    Width =9405
                    Height =420
                    FontSize =16
                    FontWeight =700
                    BackColor =16764057
                    ForeColor =0
                    Name ="Label22"
                    Caption ="Weekly Financial"
                    FontName ="Calibri"
                    GUID = Begin
                        0x48357129e948ee49ac71dee06918c55a
                    End
                End
            End
        End
        Begin Section
            Height =420
            Name ="Detail"
            GUID = Begin
                0x02d03e65aa9bbd4fbd4b7534f318bfdc
            End
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =1320
                    Height =255
                    ColumnWidth =2310
                    FontSize =9
                    Name ="reportwk"
                    ControlSource ="reportwk"
                    FontName ="Calibri"
                    GUID = Begin
                        0x54fa7e929886c345b368403cc580a49e
                    End

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1500
                    Top =60
                    Width =1500
                    Height =255
                    ColumnWidth =2310
                    FontSize =10
                    TabIndex =1
                    Name ="SentAmt"
                    ControlSource ="SentAmt"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Calibri"
                    GUID = Begin
                        0x16c698f49ed1dc4f8bcdb9779147d1bd
                    End

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3960
                    Top =60
                    Width =1680
                    Height =255
                    ColumnWidth =2310
                    FontSize =10
                    TabIndex =2
                    Name ="Arsetupamt"
                    ControlSource ="Arsetupamt"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Calibri"
                    GUID = Begin
                        0x007d92ef0a221e4a81204909a05d5f8e
                    End

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6540
                    Top =60
                    Width =1860
                    Height =255
                    ColumnWidth =2310
                    FontSize =10
                    TabIndex =3
                    Name ="collectedamt"
                    ControlSource ="collectedamt"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Calibri"
                    GUID = Begin
                        0x73a5b31175d3874b8c85f8316b121ed9
                    End

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3060
                    Top =60
                    Width =780
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="Text17"
                    ControlSource ="=([SentAmt]-[PreviousSent])/[PreviousSent]"
                    Format ="Percent"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001010000 ,
                        0xff000000ffffff00000000000400000003000000050000000101000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End
                    GUID = Begin
                        0x3000b7eb444f224eb1f26a808dcd5003
                    End

                    ConditionalFormat14 = Begin
                        0x010002000000000000000500000001010000ff000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000004000000010100 ,
                        0x0000800000ffffff000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5700
                    Top =60
                    Width =780
                    Height =255
                    FontSize =10
                    TabIndex =5
                    Name ="Text19"
                    ControlSource ="=([ARSetupAmt]-[PreviousSetup])/[PreviousSetup]"
                    Format ="Percent"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001010000 ,
                        0xff000000ffffff00000000000400000003000000050000000101000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End
                    GUID = Begin
                        0xc01b17ba555bd0439e5cba7cd5801fe3
                    End

                    ConditionalFormat14 = Begin
                        0x010002000000000000000500000001010000ff000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000004000000010100 ,
                        0x0000800000ffffff000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8520
                    Top =60
                    Width =840
                    Height =255
                    FontSize =10
                    TabIndex =6
                    Name ="Text21"
                    ControlSource ="=([CollectedAmt]-[PreviousCollected])/[PreviousCollected]"
                    Format ="Percent"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000004000000000000000200000001010000 ,
                        0x00800000ffffff000000000005000000030000000500000001010000ff000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End
                    GUID = Begin
                        0x3d236184a8ac28488b01467732db2664
                    End

                    ConditionalFormat14 = Begin
                        0x01000200000000000000040000000101000000800000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000005000000010100 ,
                        0x00ff000000ffffff000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
            GUID = Begin
                0x428f1a15fc05504fa2b07b54b70b1afc
            End
        End
    End
End
