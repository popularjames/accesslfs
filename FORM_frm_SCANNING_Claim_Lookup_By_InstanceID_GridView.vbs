Version =20
VersionRequired =20
Checksum =-1489020590
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7920
    DatasheetFontHeight =10
    ItemSuffix =22
    Left =960
    Top =2025
    Right =16770
    Bottom =8280
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x69f258a48f93e340
    End
    RecordSource ="SELECT * FROM v_SCANNING_Claim_Lookup_By_InstanceID WHERE 1=2; "
    Caption ="Claim Lookup by Name"
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
            GUID = Begin
                0xf85da4d96031964fb301e6d2a5b20c30
            End
        End
        Begin Section
            Height =3570
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0x83819ab3f26db449bc8b7b620f6ce5e6
            End
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =120
                    Width =2310
                    Height =255
                    ColumnWidth =3375
                    Name ="CnlyClaimNum"
                    ControlSource ="CnlyClaimNum"
                    GUID = Begin
                        0x659386cba08c6f468f8d43af7807947f
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1020
                            Height =255
                            FontWeight =700
                            Name ="CnlyClaimNum_Label"
                            Caption ="CnlyClaimNum"
                            GUID = Begin
                                0xa1c0b6d758c76b4c882643a51c79a4a1
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =480
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    TabIndex =2
                    Name ="LastName"
                    ControlSource ="LastName"
                    GUID = Begin
                        0x26e695ab6d45ac4cb07c94c6a25d0571
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =1020
                            Height =255
                            FontWeight =700
                            Name ="LastName_Label"
                            Caption ="LastName"
                            GUID = Begin
                                0x938479ff3ab1fc428d7a485ed2659ff2
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =840
                    Width =2310
                    Height =255
                    ColumnWidth =1080
                    TabIndex =3
                    Name ="FirstName"
                    ControlSource ="FirstName"
                    GUID = Begin
                        0xcaf115d45375904aa9d9d510184d9df2
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =840
                            Width =1020
                            Height =255
                            FontWeight =700
                            Name ="FirstName_Label"
                            Caption ="FirstName"
                            GUID = Begin
                                0x20a7432393c4864a98477c7f85b09650
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =1200
                    Width =2820
                    Height =255
                    ColumnWidth =2370
                    TabIndex =6
                    Name ="PatCtlNum"
                    ControlSource ="PatCtlNum"
                    GUID = Begin
                        0xe4e8dfde587fdc40a645b5605ca19bb2
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1200
                            Width =1020
                            Height =255
                            FontWeight =700
                            Name ="PatCtlNum_Label"
                            Caption ="PatCtlNum"
                            GUID = Begin
                                0x87c2775d0af98340b532eb4a3df5ae0b
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =1560
                    Width =1035
                    Height =255
                    ColumnWidth =1230
                    TabIndex =5
                    Name ="BeneBirthDt"
                    ControlSource ="BeneBirthDt"
                    GUID = Begin
                        0xcc783ee5b3cf3e4fafaacaece4263c3e
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1560
                            Width =1020
                            Height =255
                            FontWeight =700
                            Name ="BeneBirthDt_Label"
                            Caption ="BeneBirthDt"
                            GUID = Begin
                                0x480bccc02804dd47bc56e154d36f3641
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =1920
                    Width =1035
                    Height =255
                    ColumnWidth =1275
                    TabIndex =7
                    Name ="IPAdmitDate"
                    ControlSource ="IPAdmitDate"
                    GUID = Begin
                        0xc819ef3d4b3efa468952fd54f5bfe169
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1920
                            Width =1020
                            Height =255
                            FontWeight =700
                            Name ="IPAdmitDate_Label"
                            Caption ="IPAdmitDate"
                            GUID = Begin
                                0xb4c1b23cdccea94baf0ec6979b86084a
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =2280
                    Width =2310
                    Height =255
                    ColumnWidth =1785
                    TabIndex =8
                    Name ="CnlyProvID"
                    ControlSource ="CnlyProvID"
                    GUID = Begin
                        0xccce1960aed641498c3c5b7a842d512b
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2280
                            Width =1020
                            Height =255
                            FontWeight =700
                            Name ="CnlyProvID_Label"
                            Caption ="CnlyProvID"
                            GUID = Begin
                                0x6bcb190d1aae5542bbf61a83782be244
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =2640
                    Width =1185
                    Height =255
                    ColumnWidth =1185
                    TabIndex =9
                    Name ="ProvNum"
                    ControlSource ="ProvNum"
                    GUID = Begin
                        0xeab3242aef77de4d919f509abb00fa85
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2640
                            Width =1020
                            Height =255
                            FontWeight =700
                            Name ="ProvNum_Label"
                            Caption ="ProvNum"
                            GUID = Begin
                                0xf042f733026f604ebaee553e6a68e988
                            End
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =3000
                    Width =2820
                    Height =450
                    ColumnWidth =3000
                    TabIndex =10
                    Name ="ProvName"
                    ControlSource ="ProvName"
                    GUID = Begin
                        0x71fc87bb510ecb45a7ad9734590bbbe1
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3000
                            Width =1020
                            Height =255
                            FontWeight =700
                            Name ="ProvName_Label"
                            Caption ="ProvName"
                            GUID = Begin
                                0xb186cb047883044592c73e85507f2659
                            End
                        End
                    End
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    EnterKeyBehavior = NotDefault
                    IsHyperlink = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5100
                    Top =120
                    Width =2760
                    Height =840
                    ColumnWidth =1515
                    TabIndex =11
                    ForeColor =1279872587
                    Name ="ImageLink"
                    ControlSource ="ImageLink"
                    GUID = Begin
                        0xa79e70188c5cca4686dee1a3f19f251b
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4020
                            Top =120
                            Width =1020
                            Height =255
                            Name ="ImageLink_Label"
                            Caption ="ImageLink"
                            GUID = Begin
                                0x45f42a3a5f0b9f4f90e7d43fede8e222
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5880
                    Top =1440
                    ColumnWidth =2205
                    TabIndex =1
                    Name ="ICN"
                    ControlSource ="ICN"
                    GUID = Begin
                        0x8993ae2008ff0a4894f2b7d6959c9aac
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4440
                            Top =1440
                            Width =720
                            Height =240
                            FontWeight =700
                            Name ="Label20"
                            Caption ="ICN"
                            GUID = Begin
                                0x1e3fb31b3658cd458ff2586460c2de70
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6240
                    Top =2160
                    ColumnWidth =1755
                    TabIndex =4
                    Name ="CAN"
                    ControlSource ="CAN"
                    GUID = Begin
                        0xb942c9933d580745b22e79a2c886fa50
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4440
                            Top =2160
                            Width =1140
                            Height =240
                            FontWeight =700
                            Name ="Label21"
                            Caption ="Medicare#"
                            GUID = Begin
                                0x6b9a2f45239e5846974f63808bbc34e3
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
                0x9bfd7842284489418b98fce12a04c350
            End
        End
    End
End
