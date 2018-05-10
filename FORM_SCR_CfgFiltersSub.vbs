Version =20
VersionRequired =20
Checksum =232634585
Begin Form
    AutoCenter = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7920
    RowHeight =495
    DatasheetFontHeight =10
    ItemSuffix =4
    Right =16065
    Bottom =13470
    HelpContextId =1126
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x19f209b62f0fe440
    End
    GUID = Begin
        0x1630a8481538934ebb2f96df32367f93
    End
    RecordSource ="SCR_ScreensFilters"
    Caption ="CcaCfgScrScreenFilters"
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
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            GUID = Begin
                0x3aa2e4550e90934dbe2d3c7c05f3decc
            End
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =60
                    Width =3900
                    Height =240
                    HelpContextId =1126
                    Name ="FilterName Label"
                    Caption ="FilterName"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="FilterName_Label"
                    GUID = Begin
                        0x4efdc2d3f617ce4b86f3515431fcbe69
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4020
                    Top =60
                    Width =3840
                    Height =240
                    HelpContextId =1126
                    Name ="FilterSQL Label"
                    Caption ="FilterSQL"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="FilterSQL_Label"
                    GUID = Begin
                        0x575984c3b2fabf428e76535c636e94b5
                    End
                End
            End
        End
        Begin Section
            Height =570
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0xef9e2cb738cd4147877eb6a791a5750e
            End
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =3900
                    Height =450
                    ColumnWidth =1890
                    HelpContextId =1126
                    Name ="FilterName"
                    ControlSource ="FilterName"
                    GUID = Begin
                        0x021ed36a74a0ea4c90fcd4a2de7c4017
                    End

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    Left =4020
                    Top =60
                    Width =3840
                    Height =450
                    ColumnWidth =4215
                    TabIndex =1
                    HelpContextId =1126
                    Name ="FilterSQL"
                    ControlSource ="FilterSQL"
                    GUID = Begin
                        0x9fd96fd55e2bf5439fa768f323d6d53f
                    End

                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
            GUID = Begin
                0x6d577b27eb765f45b86f658e7e898149
            End
        End
    End
End
