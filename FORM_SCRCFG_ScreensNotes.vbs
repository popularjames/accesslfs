Version =20
VersionRequired =20
Checksum =-755910577
Begin Form
    AutoCenter = NotDefault
    AllowAdditions = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6300
    ItemSuffix =8
    Right =10560
    Bottom =13470
    HelpContextId =1129
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2282d33ff10ee440
    End
    GUID = Begin
        0x9345ea81069ec24983a7717fd63c5709
    End
    RecordSource ="SELECT SCR_ScreensNotes.* FROM SCR_ScreensNotes ORDER BY SCR_ScreensNotes.NoteDa"
        "te DESC; "
    Caption ="CnlyScreensNotes1"
    HelpFile ="C:\\WINDOWS\\Help\\Decipher2.chm"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    DatasheetForeColor12 =33554432
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
            GUID = Begin
                0x52009deb38a45a4d93d2ebc5a589786e
            End
        End
        Begin Section
            Height =2550
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0xa4f1ac5015f2c441ab650861abec1dec
            End
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =120
                    Width =4560
                    Height =840
                    ColumnWidth =3870
                    HelpContextId =1129
                    Name ="NoteText"
                    ControlSource ="NoteText"
                    GUID = Begin
                        0xc952b4662644954dafd0a516daa23f3b
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1560
                            Height =255
                            HelpContextId =1129
                            Name ="NoteText_Label"
                            Caption ="Note"
                            GUID = Begin
                                0xb098bced7e174146ad02bd527e14d493
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1080
                    Width =1035
                    Height =255
                    ColumnWidth =1530
                    TabIndex =1
                    HelpContextId =1129
                    Name ="NoteDate"
                    ControlSource ="NoteDate"
                    Format ="mm/dd/yyyy"
                    GUID = Begin
                        0x0d4d477badeb9a4a9912449a31d53497
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1080
                            Width =1560
                            Height =255
                            HelpContextId =1129
                            Name ="NoteDate_Label"
                            Caption ="Date"
                            GUID = Begin
                                0xd32d2008c35a0e4ab05606e16a9965fa
                            End
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1440
                    Width =4560
                    Height =450
                    ColumnWidth =1320
                    TabIndex =2
                    HelpContextId =1129
                    Name ="Computer"
                    ControlSource ="Computer"
                    GUID = Begin
                        0xf61796038cf33840a0b2f8721ca0f5aa
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1440
                            Width =1560
                            Height =255
                            HelpContextId =1129
                            Name ="Computer_Label"
                            Caption ="Computer"
                            GUID = Begin
                                0xf53b9ba4c23e2a4690e8827a819e9b1b
                            End
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1980
                    Width =4560
                    Height =450
                    ColumnWidth =1845
                    TabIndex =3
                    HelpContextId =1129
                    Name ="UserName"
                    ControlSource ="UserName"
                    GUID = Begin
                        0xa26df760b5b92b40941d4f36fdde23d1
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1980
                            Width =1560
                            Height =255
                            HelpContextId =1129
                            Name ="UserName_Label"
                            Caption ="UserName"
                            GUID = Begin
                                0x4f8b9ec08cdf7c4c98a9f81ddb433e18
                            End
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
            GUID = Begin
                0x5962cfaeaf93d049abdfee53fe062b5b
            End
        End
    End
End
