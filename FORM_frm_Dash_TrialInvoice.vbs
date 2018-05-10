Version =20
VersionRequired =20
Checksum =425378196
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =0
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9360
    DatasheetFontHeight =10
    ItemSuffix =25
    Left =12855
    Top =2220
    Right =22320
    Bottom =4665
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa097c6965519e440
    End
    RecordSource ="RPT_R0102"
    Caption ="RPT_R0102"
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
        Begin Line
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
            Height =1320
            Name ="FormHeader"
            GUID = Begin
                0x91468dc62a5923439105a317f81f0a3a
            End
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =60
                    Top =1080
                    Width =600
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="AuditNum_Label"
                    Caption ="Audit"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0xb6a9ccbb6694c245a20d4086d8361fc0
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =720
                    Top =1080
                    Width =3000
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="AuditDesc_Label"
                    Caption ="AuditDesc"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x1dc756b7b4efee49b877ddb9299ba0ec
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =3780
                    Top =1080
                    Width =1500
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="RACFee_Label"
                    Caption ="Fee"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x533e600174a48f45a4a89d8a9b429d2b
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =5340
                    Top =1080
                    Width =1500
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="TransactionAmt_Label"
                    Caption ="Transaction"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x2154f15b5226054c9b0d4d8430f6f775
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =6900
                    Top =1080
                    Width =690
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label12"
                    Caption ="OP/UP"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0xd346b4cd0e6e1747ac7722620be70d4d
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =7740
                    Top =1080
                    Width =1560
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label14"
                    Caption ="Type"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0xa9e127930a94b2418cb8aefe98459c54
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1380
                    Top =720
                    Width =2820
                    Height =255
                    FontSize =10
                    FontWeight =700
                    ForeColor =255
                    Name ="Text16"
                    ControlSource ="=Sum([RACFee])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Calibri"
                    GUID = Begin
                        0xe0c11ef788fb5449ad75358c5c2aa0a6
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =60
                            Top =720
                            Width =855
                            Height =240
                            FontSize =10
                            FontWeight =700
                            Name ="Label19"
                            Caption ="Fee"
                            FontName ="Calibri"
                            Tag ="DetachedLabel"
                            GUID = Begin
                                0x5d657b5cd197fd419b146e176eb6ec7c
                            End
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1380
                    Top =420
                    Width =2820
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Text17"
                    ControlSource ="=Sum([TransactionAmt])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Calibri"
                    GUID = Begin
                        0x8978615c9e4a5f4e9ecae33b86c62a9d
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =60
                            Top =420
                            Width =1185
                            Height =240
                            FontSize =10
                            FontWeight =700
                            Name ="Label20"
                            Caption ="Transactions"
                            FontName ="Calibri"
                            Tag ="DetachedLabel"
                            GUID = Begin
                                0xa70a5d38f8dabd44982c9b8e404f23c6
                            End
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    TextAlign =2
                    Width =9360
                    Height =360
                    FontSize =16
                    FontWeight =700
                    BackColor =16764057
                    BorderColor =4210752
                    ForeColor =0
                    Name ="Label22"
                    Caption ="TRIAL INVOICE"
                    FontName ="Calibri"
                    GUID = Begin
                        0xddbb40145da9bd42a8b3988346f6f40e
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7320
                    Top =60
                    Width =2040
                    Height =255
                    FontSize =10
                    TabIndex =2
                    ForeColor =0
                    Name ="ReportLastRunDt"
                    ControlSource ="ReportLastRunDt"
                    FontName ="Calibri"
                    GUID = Begin
                        0x3fd6e0a75ffc924d83bc361f0b8d12ec
                    End

                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    OverlapFlags =87
                    Top =1320
                    Width =9360
                    Name ="Line23"
                    GUID = Begin
                        0x980084c2b611974f9cab0720b00b8e33
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    OverlapFlags =85
                    Top =1020
                    Width =9360
                    Name ="Line24"
                    GUID = Begin
                        0x4aee8538ef706449ba40c01831f67335
                    End
                End
            End
        End
        Begin Section
            Height =255
            Name ="Detail"
            GUID = Begin
                0xcd8426a0ca157444a45f93a9641ce592
            End
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    BorderWidth =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Width =600
                    Height =255
                    ColumnWidth =900
                    FontSize =10
                    Name ="AuditNum"
                    ControlSource ="AuditNum"
                    FontName ="Calibri"
                    GUID = Begin
                        0x1e6fd85991182b4d9d253958510864a4
                    End

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    SpecialEffect =0
                    OldBorderStyle =0
                    BorderWidth =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Width =3000
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =1
                    Name ="AuditDesc"
                    ControlSource ="AuditDesc"
                    FontName ="Calibri"
                    GUID = Begin
                        0xf720148aafbf514f90c1c7204a813cd6
                    End

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    BorderWidth =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3780
                    Width =1500
                    Height =255
                    ColumnWidth =2310
                    FontSize =10
                    TabIndex =2
                    Name ="RACFee"
                    ControlSource ="RACFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Calibri"
                    GUID = Begin
                        0xa6ff610f238438489d1fdb1155a223c6
                    End

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    BorderWidth =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5340
                    Width =1500
                    Height =255
                    ColumnWidth =2310
                    FontSize =10
                    TabIndex =3
                    Name ="TransactionAmt"
                    ControlSource ="TransactionAmt"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Calibri"
                    GUID = Begin
                        0x3d991be6e2c5db4fa9a3cc992b985f62
                    End

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    BorderWidth =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6900
                    Width =600
                    Height =255
                    FontSize =10
                    TabIndex =4
                    Name ="Text13"
                    ControlSource ="OP_UP_Indicator"
                    FontName ="Calibri"
                    GUID = Begin
                        0x1cfeac2a5d251d4aa57aa6f4bd9dedc8
                    End

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    BorderWidth =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7740
                    Width =1560
                    Height =255
                    FontSize =10
                    TabIndex =5
                    Name ="Text15"
                    ControlSource ="TabName"
                    FontName ="Calibri"
                    GUID = Begin
                        0x166832a96033e340b9707ec8ac3ff3ed
                    End

                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            GUID = Begin
                0x731cd4fd65bbe742a56dcabc396761a5
            End
        End
    End
End
