Version =20
VersionRequired =20
Checksum =-2117015392
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7920
    DatasheetFontHeight =9
    ItemSuffix =60
    Left =7695
    Top =2145
    Right =15045
    Bottom =9990
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9ccfba937070e440
    End
    Caption ="SCANNING_Quick_Image_Log"
    DatasheetFontName ="Calibri"
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
                0x35c3aa00e7d5434bbad4e013ef693a8e
            End
        End
        Begin Section
            Height =4275
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0x8857b6805009d142a33ea5d6030ab177
            End
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =120
                    Width =1860
                    Height =255
                    ColumnWidth =0
                    Name ="SessionID"
                    ControlSource ="SessionID"
                    GUID = Begin
                        0x4af92e387a3db743a027c13a06fd45d8
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =660
                            Height =255
                            Name ="SessionID_Label"
                            Caption ="SessionID"
                            GUID = Begin
                                0xa8dea0617c5cf540bea7764500fec7de
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =480
                    Width =1860
                    Height =255
                    ColumnWidth =2295
                    TabIndex =1
                    Name ="UserID"
                    ControlSource ="UserID"
                    GUID = Begin
                        0x910e252de315a6408b5e2296fbe4c065
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =660
                            Height =255
                            Name ="UserID_Label"
                            Caption ="UserID"
                            GUID = Begin
                                0xff20fe1053986840921a051cf4fd0ff0
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =840
                    Width =900
                    Height =255
                    ColumnWidth =1035
                    TabIndex =2
                    Name ="CreateDt"
                    ControlSource ="CreateDt"
                    GUID = Begin
                        0x8da9b06c92d1b447b89f08b5be16ce68
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =840
                            Width =660
                            Height =255
                            Name ="CreateDt_Label"
                            Caption ="CreateDt"
                            GUID = Begin
                                0xbe26cb097a743a45b6a522140076bc1f
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =1200
                    Width =1860
                    Height =255
                    ColumnWidth =2310
                    TabIndex =3
                    Name ="RequestNum"
                    ControlSource ="RequestNum"
                    GUID = Begin
                        0x1e052e6bea88664c84c0ef0183c5ee5f
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1200
                            Width =660
                            Height =255
                            Name ="RequestNum_Label"
                            Caption ="RequestNum"
                            GUID = Begin
                                0x2453be312f35a344bb3d2061e6e4b344
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =1560
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =4
                    Name ="SeqNo"
                    ControlSource ="SeqNo"
                    GUID = Begin
                        0x068586c9e817f74ab30aa28bf5fe465a
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1560
                            Width =660
                            Height =255
                            Name ="SeqNo_Label"
                            Caption ="SeqNo"
                            GUID = Begin
                                0x62d6350e969c5448bb44594c8b724e0a
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =1920
                    Width =1860
                    Height =255
                    ColumnWidth =2310
                    TabIndex =5
                    Name ="CnlyClaimNum"
                    ControlSource ="CnlyClaimNum"
                    GUID = Begin
                        0x1bc06cdb5e03de4aa56bf260a75db507
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1920
                            Width =660
                            Height =255
                            Name ="CnlyClaimNum_Label"
                            Caption ="CnlyClaimNum"
                            GUID = Begin
                                0xa14ff2df17a3bd408b9b98d9ea9243c5
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =2280
                    Width =1860
                    Height =255
                    ColumnWidth =2310
                    TabIndex =6
                    Name ="ICN"
                    ControlSource ="ICN"
                    GUID = Begin
                        0x294b0f4106e29948a0b93ca08467082f
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2280
                            Width =660
                            Height =255
                            Name ="ICN_Label"
                            Caption ="ICN"
                            GUID = Begin
                                0xbd8e36198c1ba348b1c75f611466f3f8
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =2640
                    Width =1860
                    Height =255
                    ColumnWidth =2310
                    TabIndex =7
                    Name ="CnlyProvId"
                    ControlSource ="CnlyProvId"
                    GUID = Begin
                        0x220d41e9c86fbb4790780fdb461e2fec
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2640
                            Width =660
                            Height =255
                            Name ="CnlyProvId_Label"
                            Caption ="CnlyProvId"
                            GUID = Begin
                                0xe45e70cbb11a96438dcfe4bcfbcb4b47
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =3000
                    Width =1860
                    Height =255
                    ColumnWidth =2310
                    TabIndex =8
                    Name ="ProvNum"
                    ControlSource ="ProvNum"
                    GUID = Begin
                        0xc4b7445052c7174b82c4edfee03a3879
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3000
                            Width =660
                            Height =255
                            Name ="ProvNum_Label"
                            Caption ="ProvNum"
                            GUID = Begin
                                0x14aa0888f1386049baf39ad802d226f1
                            End
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =3360
                    Width =1860
                    Height =450
                    ColumnWidth =2970
                    TabIndex =9
                    Name ="ProvName"
                    ControlSource ="ProvName"
                    GUID = Begin
                        0x050af504d0f6e742899387acac5750de
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3360
                            Width =660
                            Height =255
                            Name ="ProvName_Label"
                            Caption ="ProvName"
                            GUID = Begin
                                0xf21c4a392a45354c9007ef401e5a2839
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =3900
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =10
                    Name ="LetterReqDt"
                    ControlSource ="LetterReqDt"
                    GUID = Begin
                        0x78f5da2f44c8084f84e3da5cb06e04d7
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3900
                            Width =660
                            Height =255
                            Name ="LetterReqDt_Label"
                            Caption ="LetterReqDt"
                            GUID = Begin
                                0x58d5957790785d42948f265afe41af65
                            End
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3420
                    Top =120
                    Width =1860
                    Height =450
                    ColumnWidth =3000
                    TabIndex =11
                    Name ="MemberName"
                    ControlSource ="MemberName"
                    GUID = Begin
                        0x7116fc1190f13e4db2c721f63f9a9c81
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Top =120
                            Width =660
                            Height =255
                            Name ="MemberName_Label"
                            Caption ="MemberName"
                            GUID = Begin
                                0x728cb4b869e0dd4caf24ee7897028751
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3420
                    Top =660
                    Width =1035
                    Height =255
                    ColumnWidth =1050
                    TabIndex =12
                    Name ="DOB"
                    ControlSource ="DOB"
                    GUID = Begin
                        0x8268ebd8bf822b408e00e02140b60085
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Top =660
                            Width =660
                            Height =255
                            Name ="DOB_Label"
                            Caption ="DOB"
                            GUID = Begin
                                0x02fd10195e82df4a8f90b9101a768bb6
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3420
                    Top =1020
                    Width =1860
                    Height =255
                    ColumnWidth =2310
                    TabIndex =13
                    Name ="MedicalRecordNum"
                    ControlSource ="MedicalRecordNum"
                    GUID = Begin
                        0xa4fd79843ca91c4e8b43ea8d5ca599eb
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Top =1020
                            Width =660
                            Height =255
                            Name ="MedicalRecordNum_Label"
                            Caption ="MedicalRecordNum"
                            GUID = Begin
                                0x64be9bb35359b6468545efbec59fcc52
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3420
                    Top =1380
                    Width =1860
                    Height =255
                    ColumnWidth =2910
                    TabIndex =14
                    Name ="PatCtlNum"
                    ControlSource ="PatCtlNum"
                    GUID = Begin
                        0x2c444bb694368c4fa7c39a11491f5be9
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Top =1380
                            Width =660
                            Height =255
                            Name ="PatCtlNum_Label"
                            Caption ="PatCtlNum"
                            GUID = Begin
                                0xf1e4f113bc11794b87cda6b9ed2bf58c
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3420
                    Top =1740
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =15
                    Name ="AdmitDate"
                    ControlSource ="AdmitDate"
                    GUID = Begin
                        0x4844466e992e404f95c2f0e1cd76010d
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Top =1740
                            Width =660
                            Height =255
                            Name ="AdmitDate_Label"
                            Caption ="AdmitDate"
                            GUID = Begin
                                0xf79ce9fce8bd3840b35bec7af6cb9e80
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3420
                    Top =2100
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =16
                    Name ="DischargedDate"
                    ControlSource ="DischargedDate"
                    GUID = Begin
                        0x2db2d624f6622e40825e8f4ed4ac62dc
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Top =2100
                            Width =660
                            Height =255
                            Name ="DischargedDate_Label"
                            Caption ="DischargedDate"
                            GUID = Begin
                                0xdbf64ea2be13474fb72fbc6b97e63cf6
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3420
                    Top =2460
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =17
                    Name ="ClmFromDt"
                    ControlSource ="ClmFromDt"
                    GUID = Begin
                        0x6f522edb6d6da3438291e1a41ef9db69
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Top =2460
                            Width =660
                            Height =255
                            Name ="ClmFromDt_Label"
                            Caption ="ClmFromDt"
                            GUID = Begin
                                0x187f0466ee05ef46b82f93fa93d5aad0
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3420
                    Top =2820
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =18
                    Name ="ClmThruDt"
                    ControlSource ="ClmThruDt"
                    GUID = Begin
                        0x7a716d9ffd9d4e4aa19b754a43b4d538
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Top =2820
                            Width =660
                            Height =255
                            Name ="ClmThruDt_Label"
                            Caption ="ClmThruDt"
                            GUID = Begin
                                0xa737489a9b7bb14289740bb9f4b87548
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3420
                    Top =3180
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =19
                    Name ="AuditNum"
                    ControlSource ="AuditNum"
                    GUID = Begin
                        0x1ae2b0dc7da2984c80d15f8250a7e4f9
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Top =3180
                            Width =660
                            Height =255
                            Name ="AuditNum_Label"
                            Caption ="AuditNum"
                            GUID = Begin
                                0xa172600d12f80a4893cac59a4697cc1e
                            End
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3420
                    Top =3540
                    Width =1860
                    Height =450
                    ColumnWidth =3000
                    TabIndex =20
                    Name ="ImageName"
                    ControlSource ="ImageName"
                    GUID = Begin
                        0xa996736e1381594688356a35af200f51
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Top =3540
                            Width =660
                            Height =255
                            Name ="ImageName_Label"
                            Caption ="ImageName"
                            GUID = Begin
                                0x9009a060f0ad5748a8f2dff1279926d9
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6060
                    Top =120
                    Width =900
                    Height =255
                    ColumnWidth =1860
                    TabIndex =21
                    Name ="ImageTypeDisplay"
                    ControlSource ="ImageTypeDisplay"
                    GUID = Begin
                        0xea854282990a8340a4b7e1a774ecf607
                    End

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5340
                            Top =120
                            Width =1380
                            Height =255
                            Name ="ImageType_Label"
                            Caption ="ImageTypeDisplay"
                            GUID = Begin
                                0x0b5621793657b242a5d7ca3a05e72a53
                            End
                            LayoutCachedLeft =5340
                            LayoutCachedTop =120
                            LayoutCachedWidth =6720
                            LayoutCachedHeight =375
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6060
                    Top =480
                    Width =1035
                    Height =255
                    ColumnWidth =1845
                    TabIndex =22
                    Name ="ScannedDt"
                    ControlSource ="ScannedDt"
                    GUID = Begin
                        0x97a521f7b47d10439ec4c616c36ec80b
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5340
                            Top =480
                            Width =660
                            Height =255
                            Name ="ScannedDt_Label"
                            Caption ="ScannedDt"
                            GUID = Begin
                                0x3997a1e27ed5f34081070c4ee242e92d
                            End
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6060
                    Top =840
                    Width =1800
                    Height =450
                    ColumnWidth =3000
                    TabIndex =23
                    Name ="ClientFileName"
                    ControlSource ="ClientFileName"
                    GUID = Begin
                        0xd85c73e854a376488f195326d6d76dc2
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5340
                            Top =840
                            Width =660
                            Height =255
                            Name ="ClientFileName_Label"
                            Caption ="ClientFileName"
                            GUID = Begin
                                0xa49015a520bbf345817f5a3b6eb24f53
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6060
                    Top =1380
                    Width =1560
                    Height =255
                    ColumnWidth =1560
                    TabIndex =24
                    Name ="ReceivedMeth"
                    ControlSource ="ReceivedMeth"
                    GUID = Begin
                        0x5e35d00d0b7a1044b6a8f6f379bfa99a
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5340
                            Top =1380
                            Width =660
                            Height =255
                            Name ="ReceivedMeth_Label"
                            Caption ="ReceivedMeth"
                            GUID = Begin
                                0x091291fd33098043a9ed285aff89c773
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6060
                    Top =1740
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =25
                    Name ="ReceivedDt"
                    ControlSource ="ReceivedDt"
                    GUID = Begin
                        0x4688e0c49f6714489d667123788bf17d
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5340
                            Top =1740
                            Width =660
                            Height =255
                            Name ="ReceivedDt_Label"
                            Caption ="ReceivedDt"
                            GUID = Begin
                                0x6584a325f599454484740ac944a16bf9
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6060
                    Top =2100
                    Width =1800
                    Height =255
                    ColumnWidth =2310
                    TabIndex =26
                    Name ="Carrier"
                    ControlSource ="Carrier"
                    GUID = Begin
                        0x669b639d8e12c942aa7c43ae92c54f1e
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5340
                            Top =2100
                            Width =660
                            Height =255
                            Name ="Carrier_Label"
                            Caption ="Carrier"
                            GUID = Begin
                                0x45fc1fa640f2674da864fd09d12e2f5f
                            End
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6060
                    Top =2460
                    Width =1800
                    Height =450
                    ColumnWidth =3000
                    TabIndex =27
                    Name ="TrackingNum"
                    ControlSource ="TrackingNum"
                    GUID = Begin
                        0xe0082b8c74af61428dfa0a5757784f9b
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5340
                            Top =2460
                            Width =660
                            Height =255
                            Name ="TrackingNum_Label"
                            Caption ="TrackingNum"
                            GUID = Begin
                                0x837ac769e2e0a4469c450332337e870e
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6060
                    Top =3000
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =28
                    Name ="AccountID"
                    ControlSource ="AccountID"
                    GUID = Begin
                        0xfc4c69d7d04c4840a91b10df1e591483
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5340
                            Top =3000
                            Width =660
                            Height =255
                            Name ="AccountID_Label"
                            Caption ="AccountID"
                            GUID = Begin
                                0xc8c11133d6fcb0419054bbf4bdf3142f
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6060
                    Top =3360
                    Width =225
                    Height =255
                    ColumnWidth =225
                    TabIndex =29
                    Name ="ImportFlag"
                    ControlSource ="ImportFlag"
                    GUID = Begin
                        0xb3e9298359d8a644bbbd074d54c1e354
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5340
                            Top =3360
                            Width =660
                            Height =255
                            Name ="ImportFlag_Label"
                            Caption ="ImportFlag"
                            GUID = Begin
                                0xc0846a4e8b242a40acddd5b880682b32
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
                0xc948a639999e6042b73a457f5149e341
            End
        End
    End
End
