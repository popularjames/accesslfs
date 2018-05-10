Version =20
VersionRequired =20
Checksum =483106053
Begin Form
    AutoCenter = NotDefault
    PictureTiling = NotDefault
    AllowAdditions = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    Width =10778
    RowHeight =360
    ItemSuffix =16
    Right =20250
    Bottom =13470
    HelpContextId =1126
    DatasheetGridlinesColor =12632256
    PaintPalette = Begin
        0x00030100ffffff0000000000
    End
    RecSrcDt = Begin
        0x87af9c42776be340
    End
    GUID = Begin
        0x918ec9de6da02042ab27aee527d5ea88
    End
    RecordSource ="SELECT CnlyVendorNotes.* FROM CnlyVendorNotes ORDER BY CnlyVendorNotes.DateCreat"
        "ed DESC; "
    Caption ="Vendor Notes"
    HelpFile ="C:\\WINDOWS\\Help\\Decipher2.chm"
    DatasheetFontName ="Arial"
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    NoSaveCTIWhenDisabled =1
    Begin
        Begin Label
            FontSize =9
            FontWeight =700
            BackColor =13209
            ForeColor =16777215
            FontName ="Arial"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            BorderColor =13209
        End
        Begin Line
            OldBorderStyle =2
            BorderWidth =2
            BorderColor =13209
        End
        Begin CommandButton
            FontSize =9
            ForeColor =13209
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderWidth =3
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =13209
        End
        Begin CheckBox
            BorderWidth =3
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =13209
        End
        Begin OptionGroup
            BorderLineStyle =0
            BorderColor =13209
        End
        Begin BoundObjectFrame
            BorderLineStyle =0
            BorderColor =13209
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            FontSize =9
            BorderColor =13209
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin ListBox
            BorderLineStyle =0
            FontSize =9
            BorderColor =13209
            FontName ="Arial"
        End
        Begin ComboBox
            BorderLineStyle =0
            FontSize =9
            BorderColor =13209
            FontName ="Arial"
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =13209
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =13209
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =210
            BackColor =-2147483633
            Name ="FormHeader"
            GUID = Begin
                0x26177df0b43b364daa029af313b4497f
            End
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Width =6765
                    Height =210
                    ForeColor =0
                    HelpContextId =1126
                    Name ="Notes_Label"
                    Caption ="Note"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x929dc3e0afb96a4881a16d0e31cda4aa
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    Left =6773
                    Width =825
                    Height =210
                    ForeColor =0
                    HelpContextId =1126
                    Name ="DateCreated_Label"
                    Caption ="Date"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x8a1c21994e037f43a20e8aa0291ffe54
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    Left =7605
                    Width =1200
                    Height =210
                    ForeColor =0
                    HelpContextId =1126
                    Name ="Host_Label"
                    Caption ="Computer"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x879a4f9eb42719458b94c05f9bf12679
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =8820
                    Width =1200
                    Height =210
                    ForeColor =0
                    HelpContextId =1126
                    Name ="UserName_Label"
                    Caption ="User"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0xad84eb1f5dd85e40b1cfdc0387c13403
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =10035
                    Width =690
                    Height =210
                    ForeColor =0
                    HelpContextId =1126
                    Name ="Auditor_Label"
                    Caption ="Auditor"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x253560d787ff104ebedad9b4dc7c68bd
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =255
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0x6b9770cf802a5b4cbffad8e02be1664f
            End
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =6765
                    Height =255
                    ColumnWidth =4455
                    FontSize =8
                    HelpContextId =1126
                    Name ="Notes"
                    ControlSource ="Notes"
                    GUID = Begin
                        0xccc5f04f17164c4181f08e66b60821a9
                    End

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6773
                    Width =825
                    Height =255
                    ColumnWidth =1245
                    FontSize =8
                    TabIndex =1
                    HelpContextId =1126
                    Name ="Date"
                    ControlSource ="DateCreated"
                    Format ="mm/dd/yy"
                    GUID = Begin
                        0x7096ce1cdec12f44a072596832d25e5b
                    End

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =7605
                    Width =1200
                    Height =255
                    ColumnWidth =3000
                    FontSize =8
                    TabIndex =2
                    HelpContextId =1126
                    Name ="Host"
                    ControlSource ="Host"
                    GUID = Begin
                        0x8dc1c0108087bf45bf7add2d54b5ab63
                    End

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8820
                    Width =1200
                    Height =255
                    ColumnWidth =2700
                    FontSize =8
                    TabIndex =3
                    HelpContextId =1126
                    Name ="User"
                    ControlSource ="UserName"
                    GUID = Begin
                        0x72622d56ddea6b41a77046af9b3f86f9
                    End

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10035
                    Width =690
                    Height =255
                    ColumnWidth =2310
                    FontSize =8
                    TabIndex =4
                    HelpContextId =1126
                    Name ="Auditor"
                    ControlSource ="Auditor"
                    GUID = Begin
                        0x4cd8597820a4b54285d0f29288f3cce1
                    End

                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
            GUID = Begin
                0x1754357c2d33a54aba9ab66959d7974b
            End
        End
    End
End
