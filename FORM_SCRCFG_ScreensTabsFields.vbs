Version =20
VersionRequired =20
Checksum =-1089922282
Begin Form
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    ScrollBars =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6360
    DatasheetFontHeight =10
    ItemSuffix =10
    Left =5790
    Top =3675
    Right =12555
    Bottom =9750
    HelpContextId =1129
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x58154326ee0ee440
    End
    GUID = Begin
        0x66832658c4a0f842bfc90cd130003ad1
    End
    RecordSource ="SCR_ScreensTabsFields"
    Caption ="CcaCfgScrTableFields"
    HelpFile ="C:\\WINDOWS\\Help\\Decipher2.chm"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    AllowLayoutView =0
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
            SpecialEffect =3
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
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
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
            Height =240
            BackColor =10187799
            Name ="FormHeader"
            GUID = Begin
                0x00fa25ebe2fc2a48bf57a687a9863a13
            End
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =60
                    Width =3060
                    Height =240
                    FontWeight =700
                    ForeColor =16777215
                    HelpContextId =1129
                    Name ="FieldName Label"
                    Caption ="Master Field"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="FieldName_Label"
                    GUID = Begin
                        0x4874b4e3d4d4544bb1b8d1e03c335045
                    End
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =3120
                    Width =3060
                    Height =240
                    FontWeight =700
                    ForeColor =16777215
                    HelpContextId =1129
                    Name ="Label8"
                    Caption ="Child Field"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x5f756592f52dfd4eb0a2319c5f013ab8
                    End
                End
            End
        End
        Begin Section
            Height =300
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0x84551ab3cb92bc41a36545ea07ea3028
            End
            Begin
                Begin ComboBox
                    RowSourceTypeInt =10
                    OverlapFlags =93
                    Left =60
                    Top =30
                    Width =3060
                    ColumnWidth =2310
                    HelpContextId =1129
                    GUID = Begin
                        0x99cde0c1dbd34e4dba252a4162b7a34a
                    End
                    Name ="MasterField"
                    ControlSource ="MasterField"
                    RowSourceType ="Field List"
                    RowSource ="ScrApData"

                End
                Begin ComboBox
                    RowSourceTypeInt =10
                    OverlapFlags =87
                    Left =3120
                    Top =30
                    Width =3060
                    TabIndex =1
                    HelpContextId =1129
                    GUID = Begin
                        0x7f5f58cc71c36749924d6bada1362c7b
                    End
                    Name ="ChildField"
                    ControlSource ="ChildField"
                    RowSourceType ="Field List"
                    RowSource ="ScrApData"

                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
            GUID = Begin
                0xf1f895b47cedf24ab4ddd46a0e016800
            End
        End
    End
End
