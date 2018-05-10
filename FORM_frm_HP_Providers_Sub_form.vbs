Version =20
VersionRequired =20
Checksum =-319597977
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    ScrollBars =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7920
    ItemSuffix =40
    Left =8805
    Top =5145
    Right =19875
    Bottom =7155
    DatasheetGridlinesColor =12632256
    Filter ="Connolly Provider ID"
    RecSrcDt = Begin
        0x9177706c4f21e440
    End
    RecordSource ="select * from hp_providers where provnum = ''"
    Caption ="HP_Providers"
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
                0x55df62e445b1dd4993705b3cbb611072
            End
        End
        Begin Section
            Height =4110
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0x7923296a6d56494685e8bc412a1b215c
            End
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =900
                    Top =120
                    Width =2100
                    Height =255
                    ColumnWidth =2310
                    Name ="ProvNum"
                    ControlSource ="ProvNum"
                    GUID = Begin
                        0x5861ef6c429ba54497f5803c4c46ed10
                    End

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =60
                            Top =120
                            Width =1260
                            Height =255
                            Name ="ProvNum_Label"
                            Caption ="Provider Number"
                            GUID = Begin
                                0xc3ea81f136780a408ea6ad9ad0edef12
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =900
                    Top =480
                    Width =2100
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="CnlyProvID"
                    ControlSource ="CnlyProvID"
                    GUID = Begin
                        0x1af535b149e7904293ebf76013c68655
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =780
                            Height =255
                            Name ="CnlyProvID_Label"
                            Caption ="CnlyProvID"
                            GUID = Begin
                                0x2488e9426535244dac498d00f2d6d48b
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =900
                    Top =840
                    Width =2100
                    Height =255
                    ColumnWidth =2310
                    TabIndex =2
                    Name ="HP_ProvID"
                    ControlSource ="HP_ProvID"
                    GUID = Begin
                        0xe27e86e16fd18f49b4efdc64e0437b58
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =840
                            Width =780
                            Height =255
                            Name ="HP_ProvID_Label"
                            Caption ="Healthport Provider ID"
                            GUID = Begin
                                0xeb95961c8a45624f99b9262fffc42167
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =900
                    Top =1560
                    Width =2100
                    Height =450
                    ColumnWidth =4170
                    TabIndex =4
                    Name ="ProvName"
                    ControlSource ="ProvName"
                    GUID = Begin
                        0x5a099e921d766e488443760d6209f554
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1560
                            Width =780
                            Height =255
                            Name ="ProvName_Label"
                            Caption ="ProvName"
                            GUID = Begin
                                0xa6bf3c975dd99c41a15415ac26504614
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =900
                    Top =2100
                    Width =2100
                    Height =255
                    ColumnWidth =2310
                    TabIndex =5
                    Name ="ProvType"
                    ControlSource ="ProvType"
                    GUID = Begin
                        0x22caf960aaca4d44ab0c2cfe38b72532
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2100
                            Width =780
                            Height =255
                            Name ="ProvType_Label"
                            Caption ="Provider Type"
                            GUID = Begin
                                0x29e48a6d4fe3d2429d74148240447484
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =900
                    Top =2460
                    Width =2100
                    Height =450
                    ColumnWidth =3000
                    TabIndex =6
                    Name ="Addr01"
                    ControlSource ="Addr01"
                    GUID = Begin
                        0x882be3c49be8274a96aa9b657b7011fe
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2460
                            Width =780
                            Height =255
                            Name ="Addr01_Label"
                            Caption ="Addr01"
                            GUID = Begin
                                0xe93583ffe0b83445b0297f7afa8e9f52
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =900
                    Top =3000
                    Width =2100
                    Height =450
                    ColumnWidth =3000
                    TabIndex =7
                    Name ="Addr02"
                    ControlSource ="Addr02"
                    GUID = Begin
                        0x66cf361a537f384db4535f616acea23b
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3000
                            Width =780
                            Height =255
                            Name ="Addr02_Label"
                            Caption ="Addr02"
                            GUID = Begin
                                0x8cf822e01003ee4cb9d2c2b3bb1dc3c9
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =900
                    Top =3540
                    Width =2100
                    Height =450
                    ColumnWidth =3000
                    TabIndex =8
                    Name ="Addr03"
                    ControlSource ="Addr03"
                    GUID = Begin
                        0x357d4285c677aa42b4637560f33ddc2a
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3540
                            Width =780
                            Height =255
                            Name ="Addr03_Label"
                            Caption ="Addr03"
                            GUID = Begin
                                0xe16876b9d1e4ca45ada33cb41ebdad01
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =120
                    Width =2100
                    Height =450
                    ColumnWidth =3000
                    TabIndex =9
                    Name ="City"
                    ControlSource ="City"
                    GUID = Begin
                        0x343637f126c45c4fa6b67525a1e3a142
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3060
                            Top =120
                            Width =780
                            Height =255
                            Name ="City_Label"
                            Caption ="City"
                            GUID = Begin
                                0xdb368f0512a8b3469f1a6f2dca3919e3
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =660
                    Width =390
                    Height =255
                    ColumnWidth =390
                    TabIndex =10
                    Name ="State"
                    ControlSource ="State"
                    GUID = Begin
                        0x06005411c525b446a2fd6c0b943b1620
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3060
                            Top =660
                            Width =780
                            Height =255
                            Name ="State_Label"
                            Caption ="State"
                            GUID = Begin
                                0x975788d6c2498644b205887379961228
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =1020
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =11
                    Name ="Zip"
                    ControlSource ="Zip"
                    GUID = Begin
                        0x52a054d4616eb0488539ae2dc614c7da
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3060
                            Top =1020
                            Width =780
                            Height =255
                            Name ="Zip_Label"
                            Caption ="Zip"
                            GUID = Begin
                                0xfceb82ec65cca648960f424a88492398
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =1380
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =12
                    Name ="Phone"
                    ControlSource ="Phone"
                    GUID = Begin
                        0x78bad1d368c2794588519b433c70d5b5
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3060
                            Top =1380
                            Width =780
                            Height =255
                            Name ="Phone_Label"
                            Caption ="Phone"
                            GUID = Begin
                                0x6c6677ce9cb2f449942c71c1ff2be5df
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =1740
                    Width =2100
                    Height =255
                    ColumnWidth =2310
                    TabIndex =13
                    Name ="County"
                    ControlSource ="County"
                    GUID = Begin
                        0xd1fde0c796af0b46bd66e915cf2de77d
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3060
                            Top =1740
                            Width =780
                            Height =255
                            Name ="County_Label"
                            Caption ="County"
                            GUID = Begin
                                0xe332339e364c1a4fafc372d0b4d1489f
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =2100
                    Width =2100
                    Height =450
                    ColumnWidth =3000
                    TabIndex =14
                    Name ="ProvOwner"
                    ControlSource ="ProvOwner"
                    GUID = Begin
                        0xb7fe404a9bac7045a9371073139cc6c0
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3060
                            Top =2100
                            Width =780
                            Height =255
                            Name ="ProvOwner_Label"
                            Caption ="Provider Owner"
                            GUID = Begin
                                0x9ba3ae1db66db248be6a2da55a4cf084
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =2640
                    Width =900
                    Height =255
                    ColumnWidth =1800
                    TabIndex =15
                    Name ="EmergencySvcInd"
                    ControlSource ="EmergencySvcInd"
                    GUID = Begin
                        0xd42ce363f6c2d14595ec97ce6a1d00e0
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3060
                            Top =2640
                            Width =780
                            Height =255
                            Name ="EmergencySvcInd_Label"
                            Caption ="EmergencySvcInd"
                            GUID = Begin
                                0x1daa3ba0b2d290469f9749b29a5191ec
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =3000
                    Width =225
                    Height =255
                    ColumnWidth =870
                    TabIndex =16
                    Name ="HCA"
                    ControlSource ="HCA"
                    GUID = Begin
                        0x8d267ae0b3821e459c53ec1e41bc0938
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3060
                            Top =3000
                            Width =780
                            Height =255
                            Name ="HCA_Label"
                            Caption ="HCA"
                            GUID = Begin
                                0xe93cfa6ba6d0a54c85601fcf9f70ae50
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =3360
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =17
                    Name ="NoteID"
                    ControlSource ="NoteID"
                    GUID = Begin
                        0xc127c5c53f50564f88d7f22cb056ea7c
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3060
                            Top =3360
                            Width =780
                            Height =255
                            Name ="NoteID_Label"
                            Caption ="NoteID"
                            GUID = Begin
                                0xabe39991e99c9446adf81b79e81eff27
                            End
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =900
                    Top =1200
                    Width =1560
                    ColumnWidth =1560
                    TabIndex =3
                    GUID = Begin
                        0x6c5169fb37c8ee45abe307948f02e94c
                    End
                    Name ="CurStatus"
                    ControlSource ="CurStatus"
                    RowSourceType ="Value List"
                    RowSource ="ACTIVE;INACTIVE"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1200
                            Width =780
                            Height =255
                            Name ="CurStatus_Label"
                            Caption ="Status"
                            GUID = Begin
                                0x570e7355b7c0b04490b25fff51a8ada1
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
                0x8d214e369a45cf4aafc2e868070c1d28
            End
        End
    End
End
