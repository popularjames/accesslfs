Version =20
VersionRequired =20
Checksum =-2109655494
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11400
    ItemSuffix =24
    Left =1125
    Top =4155
    Right =9525
    Bottom =6120
    DatasheetForeColor =33587200
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf8a3e33cb3a1e340
    End
    RecordSource ="SELECT * FROM GENERAL_Tabs WHERE AccessForm=\"frm_REPORT_Main\"; "
    Caption ="GENERAL_Tabs"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    DatasheetForeColor12 =33587200
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
                0xf15c56c986fee248a6aaff8f386f56f0
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =3915
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0x2454000a532e004785ababb2eef5c761
            End
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =120
                    Height =255
                    ColumnWidth =960
                    ColumnOrder =1
                    Name ="RowID"
                    ControlSource ="RowID"
                    GUID = Begin
                        0xf968be7aab200d4d971ccb5e48742681
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1560
                            Height =255
                            Name ="RowID_Label"
                            Caption ="RowID"
                            GUID = Begin
                                0x555248c64b56f040b5cafa0043d30b13
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =840
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =2
                    TabIndex =2
                    Name ="FormName"
                    ControlSource ="FormName"
                    GUID = Begin
                        0x4cd5f00bc54f6e42821e09a8e6b65513
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =840
                            Width =1560
                            Height =255
                            Name ="FormName_Label"
                            Caption ="FormName"
                            GUID = Begin
                                0xa4b4dfd516ba3c47b86590843875c5f7
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1740
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =3
                    TabIndex =4
                    Name ="AccessForm"
                    ControlSource ="AccessForm"
                    GUID = Begin
                        0xa4d53dfa1bccbe48a7dc062aeb48b0ea
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1740
                            Width =1560
                            Height =255
                            Name ="AccessForm_Label"
                            Caption ="AccessForm"
                            GUID = Begin
                                0xbdc2737ceed36d47a08b206460900577
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =2820
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =0
                    TabIndex =5
                    Name ="FormValue"
                    ControlSource ="FormValue"
                    GUID = Begin
                        0x7a46c1c70e555a4aaf8838202830766d
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2820
                            Width =1560
                            Height =255
                            Name ="FormValue_Label"
                            Caption ="FormValue"
                            GUID = Begin
                                0x027174c922a1aa4aa9a42068a2018928
                            End
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =7260
                    Top =1200
                    Width =3840
                    Height =1110
                    TabIndex =3
                    Name ="ctl_frm_ADMIN_Report_sub3"
                    SourceObject ="Form.frm_ADMIN_Report_sub2"
                    LinkChildFields ="RowID"
                    LinkMasterFields ="RowID"
                    GUID = Begin
                        0x4ef5038db8794144a33d345d64fb8c9a
                    End

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =480
                    Width =3300
                    TabIndex =1
                    Name ="txtTabName"
                    ControlSource ="TabName"
                    GUID = Begin
                        0x9d475f4020e4f04781712196849cb8ee
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =780
                            Height =240
                            Name ="Label23"
                            Caption ="TabName"
                            GUID = Begin
                                0xd965f445100f5e468b408e27e17aa77f
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
                0x5dee1719afc9dd45828066fc07c50fb4
            End
        End
    End
End
