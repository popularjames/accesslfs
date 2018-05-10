Version =20
VersionRequired =20
Checksum =556707487
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    ScrollBars =0
    ViewsAllowed =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =10
    ItemSuffix =10
    Left =5295
    Top =2220
    Right =12705
    Bottom =2430
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xff122de06d1ce440
    End
    RecordSource ="SELECT v_R3_Permission_ADMIN_User_Company.UserID, v_R3_Permission_ADMIN_User_Com"
        "pany.ProfileID, v_R3_Permission_ADMIN_User_Company.CompanyName, * FROM v_R3_Perm"
        "ission_ADMIN_User_Company WHERE (((v_R3_Permission_ADMIN_User_Company.UserID)=[c"
        "bUserID]));"
    Caption ="ADMIN_User_Profile"
    DatasheetFontName ="Arial"
    AllowFormView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
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
                0xadc778a4dd491e4987cd614631fd4a5f
            End
        End
        Begin Section
            Height =1575
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0x9f98c76a66dc2d43ac733e5025b06acb
            End
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =120
                    Width =1935
                    Height =255
                    ColumnWidth =2040
                    Name ="UserID"
                    ControlSource ="UserID"
                    FontName ="Tahoma"
                    GUID = Begin
                        0x49cf103cc47feb4cbf5a14981a4a18ba
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1560
                            Height =255
                            Name ="UserID_Label"
                            Caption ="UserID"
                            GUID = Begin
                                0x2dfff24720ecda4c8204268173151464
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
                    ColumnWidth =2190
                    TabIndex =1
                    Name ="ProfileID"
                    ControlSource ="ProfileID"
                    FontName ="Tahoma"
                    GUID = Begin
                        0xcd145a1ce257444f952e69cbc654c74d
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
                                0x8c2be031e9aa3f46a80ae39afed0175c
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1140
                    Top =840
                    TabIndex =2
                    Name ="CompanyName"
                    ControlSource ="CompanyName"
                    GUID = Begin
                        0x7f69d1e11bfc1e4e9cf7d0ba68b98860
                    End

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =840
                            Width =1200
                            Height =240
                            Name ="Label9"
                            Caption ="CompanyName"
                            GUID = Begin
                                0x2efd2888db6d444fa29c8e7f4b91ae8d
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
                0xf5da2d0825b52f43873be2053215678b
            End
        End
    End
End
