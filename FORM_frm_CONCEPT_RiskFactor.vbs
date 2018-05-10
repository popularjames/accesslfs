Version =20
VersionRequired =20
Checksum =-916776849
Begin Form
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =270
    Top =990
    Right =9255
    Bottom =5910
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa19c8f8dce99e340
    End
    GUID = Begin
        0xf561b48c924d6a4bb9e5f8b295147e67
    End
    RecordSource ="SELECT * FROM CONCEPT_Hdr WHERE conceptID in (\"CM_C0024\",\"CM_C0027\",\"CM_C00"
        "88\",\"CM_C0135\",\"CM_C0162\",\"CM_C0300\",\"CM_C0302\",\"CM_C0303\",\"CM_C0306"
        "\"); "
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    NoSaveCTIWhenDisabled =1
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin FormHeader
            Height =240
            BackColor =-2147483633
            Name ="FormHeader"
            GUID = Begin
                0xa0b27a56d0ba07469821a9566d1ee63b
            End
            Begin
                Begin Label
                    OverlapFlags =85
                    Width =900
                    Height =240
                    Name ="Label0"
                    Caption ="ConceptID:"
                    GUID = Begin
                        0x4bfde8b7c9da164ea35efef781d0573d
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =1440
                    Width =885
                    Height =240
                    Name ="Label1"
                    Caption ="RiskFactor:"
                    GUID = Begin
                        0x3952561c70b9b8469e832c6bd94e79e0
                    End
                End
            End
        End
        Begin Section
            Height =240
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0xbd2bf6f94f7c624898454ed436011950
            End
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =1140
                    Name ="ConceptID"
                    ControlSource ="ConceptID"
                    GUID = Begin
                        0x7065ce0a39941d41aa4eb82fd10bb703
                    End

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1440
                    TabIndex =1
                    Name ="RiskFactor"
                    ControlSource ="RiskFactor"
                    GUID = Begin
                        0xcb33319fc43b614da65d1a5c87492e09
                    End

                End
            End
        End
        Begin FormFooter
            Height =360
            BackColor =-2147483633
            Name ="FormFooter"
            GUID = Begin
                0x36402e701fcd9e4493c5d33dff150ee0
            End
        End
    End
End
