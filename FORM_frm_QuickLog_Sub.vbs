Version =20
VersionRequired =20
Checksum =-1387472537
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6180
    DatasheetFontHeight =9
    ItemSuffix =6
    Left =9795
    Top =1560
    Right =16095
    Bottom =9450
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb5e47bead71ee440
    End
    RecordSource ="tbl_QuickLog"
    Caption ="tbl_QuickLog"
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
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            GUID = Begin
                0xadbfb99442140a48b0cc4b5c0e1cd652
            End
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =45
                    Top =60
                    Width =1620
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="FieldName_Label"
                    Caption ="Field Name"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x7f826d7eb7f82243966da545792a234b
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3045
                    Top =60
                    Width =2610
                    Height =255
                    FontSize =9
                    FontWeight =700
                    Name ="Value_Label"
                    Caption ="Filter"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0xe96d11481e40e243a52883b8772bdd4d
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1830
                    Top =60
                    Width =870
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label4"
                    Caption ="Criteria"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                    GUID = Begin
                        0x240b4951b883a4438aac7a62ac870b9c
                    End
                End
            End
        End
        Begin Section
            Height =240
            Name ="Detail"
            GUID = Begin
                0xdfbf8d261cb7e742b68c999c9595713b
            End
            Begin
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Width =1800
                    ColumnWidth =1740
                    Name ="FieldName"
                    ControlSource ="FieldName"
                    GUID = Begin
                        0x932ef91663c27240b41f37c6c7d8c176
                    End

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3000
                    Width =2820
                    ColumnWidth =2190
                    TabIndex =1
                    Name ="Value"
                    ControlSource ="Value"
                    GUID = Begin
                        0x97c10e0c7d3ea64ba9589484c3fd4899
                    End

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AutoExpand = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1800
                    Width =1200
                    FontWeight =700
                    TabIndex =2
                    ForeColor =255
                    GUID = Begin
                        0x4e705167033786488419cda73f7ea3c4
                    End
                    Name ="cmbCriteria"
                    ControlSource ="Criteria"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;>=;<=;<>;LIKE;NOT LIKE"

                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
            GUID = Begin
                0xa122716974420048a80f6823a27972b9
            End
        End
    End
End
