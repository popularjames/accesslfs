Version =20
VersionRequired =20
Checksum =-420883455
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3300
    ItemSuffix =4
    Left =1650
    Top =4890
    Right =6360
    Bottom =6105
    DatasheetForeColor =33554687
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa342136cb2a1e340
    End
    GUID = Begin
        0x470c589de4d347469f71770b4bce2d16
    End
    RecordSource ="SELECT GENERAL_Tabs_Linked_ProfileIDs.RowID, GENERAL_Tabs_Linked_ProfileIDs.Prof"
        "ileID FROM GENERAL_Tabs_Linked_ProfileIDs; "
    Caption ="frm_ADMIN_Report_sub2"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    DatasheetForeColor12 =33554687
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
                0x0218c0d8c874eb4f9965a07a7a4f4e39
            End
        End
        Begin Section
            Height =855
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0xd65d860cb07ea641a8f6232226d516a8
            End
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =120
                    Width =900
                    Height =255
                    ColumnWidth =900
                    Name ="RowID"
                    ControlSource ="RowID"
                    GUID = Begin
                        0x256f4760e522c3429531733301606e4e
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
                                0x5943eb202ba69f4eac164339e63c8d4c
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =480
                    Width =1560
                    Height =255
                    ColumnWidth =3525
                    TabIndex =1
                    Name ="ProfileID"
                    ControlSource ="ProfileID"
                    GUID = Begin
                        0x69505f8bca625e4e850d58b69b87d3c2
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =1560
                            Height =255
                            Name ="ProfileID_Label"
                            Caption ="ProfileID"
                            GUID = Begin
                                0x8d589dcc782ec747a91b739d537ce175
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
                0x4b7378a326a11c4499123c13f0fde7d5
            End
        End
    End
End
