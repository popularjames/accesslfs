Version =20
VersionRequired =20
Checksum =2039615724
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
    ItemSuffix =20
    Left =1665
    Top =2955
    Right =17235
    Bottom =9210
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x73ffb95fb477e340
    End
    RecordSource ="SELECT * FROM v_SCANNING_Claim_Lookup_ByName WHERE 1=2; "
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
                0x719ecbf764a7dc4d9213234f2513677f
            End
        End
        Begin Section
            Height =3570
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0x4e19d9e4589240419bb36cc871551999
            End
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =120
                    Width =2310
                    Height =255
                    ColumnWidth =3330
                    Name ="CnlyClaimNum"
                    ControlSource ="CnlyClaimNum"
                    GUID = Begin
                        0x6262c19c4367334c8959e48471c2dccb
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
                                0xcb57cffece52e249b4048792225ece36
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
                    TabIndex =1
                    Name ="LastName"
                    ControlSource ="LastName"
                    GUID = Begin
                        0xc49acc05c86dc4499f6c0ac8f10d42ee
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
                                0xa390846a2739c0418575691336b0b2d9
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
                    TabIndex =2
                    Name ="FirstName"
                    ControlSource ="FirstName"
                    GUID = Begin
                        0x56950c82e95289468116696bd2ec3b9b
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
                                0x0102b084ddfcf04c8e3de841d62ebacb
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
                    ColumnWidth =1440
                    TabIndex =3
                    Name ="PatCtlNum"
                    ControlSource ="PatCtlNum"
                    GUID = Begin
                        0xf0e0ab5266f305499aaa70815c4645f3
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
                                0xaf5ff81056cb594ea9d8248c7982085a
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
                    TabIndex =4
                    Name ="BeneBirthDt"
                    ControlSource ="BeneBirthDt"
                    GUID = Begin
                        0x211ab8be5fab0e488b9d3d9e50bd9a87
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
                                0xa512e27755407d4f9902bc9df55402ed
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
                    TabIndex =5
                    Name ="IPAdmitDate"
                    ControlSource ="IPAdmitDate"
                    GUID = Begin
                        0xc373c60b367c304e812d1c70d1dedf4a
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
                                0x41f32e1f8b99c74c9543762371544771
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
                    ColumnWidth =1230
                    TabIndex =6
                    Name ="CnlyProvID"
                    ControlSource ="CnlyProvID"
                    GUID = Begin
                        0x784aac542621b547b62e4fa1ef041d12
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
                                0x65771c764e79f942ab2e72b3dece0806
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
                    TabIndex =7
                    Name ="ProvNum"
                    ControlSource ="ProvNum"
                    GUID = Begin
                        0xc7090dda1085d440ac52e17cca0927f4
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
                                0x02e67aa0f2f55c43a97775a7095fef52
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
                    TabIndex =8
                    Name ="ProvName"
                    ControlSource ="ProvName"
                    GUID = Begin
                        0x0f4d1830838fe34c8f5b8d5c3b64eddc
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
                                0x3bcf5415e2669647ae09ecb7081cff3f
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
                    TabIndex =9
                    ForeColor =1279872587
                    Name ="ImageLink"
                    ControlSource ="ImageLink"
                    GUID = Begin
                        0x137b6d9cb65d0a45bd63b509a21e6ece
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
                                0x06e4fb03a435a643b1cb7ea7059c87c2
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
                0xb844e7c370b0cd4d9a8826c46d4fa935
            End
        End
    End
End
