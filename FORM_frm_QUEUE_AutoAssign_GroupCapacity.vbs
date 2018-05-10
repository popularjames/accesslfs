Version =20
VersionRequired =20
Checksum =-880488030
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7920
    ItemSuffix =25
    Left =9135
    Top =2445
    Right =17985
    Bottom =5520
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd7a7a680ed0ae440
    End
    RecordSource ="v_QUEUE_AutoAssign_GroupCapacity"
    Caption ="frm_QUEUE_AutoAssign_GroupCapacity"
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
                0xf9de0ca78bf9ce45b252d1948106b431
            End
        End
        Begin Section
            Height =4275
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0x57e5108310cee847a81e1db7a482806e
            End
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =120
                    Width =3480
                    Height =450
                    ColumnWidth =1035
                    ColumnOrder =0
                    Name ="GroupName"
                    ControlSource ="GroupName"
                    GUID = Begin
                        0xd5f822bfc05ac44fa5a3bed29d5d9071
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1200
                            Height =255
                            Name ="GroupName_Label"
                            Caption ="GroupName"
                            GUID = Begin
                                0x905edfd188ab9b4381c9e59e279fb007
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =660
                    Width =900
                    Height =255
                    ColumnWidth =930
                    ColumnOrder =1
                    TabIndex =1
                    Name ="AgeGroup"
                    ControlSource ="AgeGroup"
                    GUID = Begin
                        0xe6022d75a6987b418b98d25efd543187
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =660
                            Width =1200
                            Height =255
                            Name ="AgeGroup_Label"
                            Caption ="AgeGroup"
                            GUID = Begin
                                0x225a937500c93c4e9fcb6cc5ddab5b8c
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =1020
                    Width =900
                    Height =255
                    ColumnWidth =1020
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Productivity"
                    ControlSource ="Productivity"
                    GUID = Begin
                        0xfa57f780e640c045a328a0f2627f9cb3
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1020
                            Width =1200
                            Height =255
                            Name ="Productivity_Label"
                            Caption ="Productivity"
                            GUID = Begin
                                0x9de2bf39446e784b82874cbb21eae41b
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =1380
                    Width =900
                    Height =255
                    ColumnWidth =945
                    ColumnOrder =3
                    TabIndex =3
                    Name ="Incomplete"
                    ControlSource ="Incomplete"
                    GUID = Begin
                        0xe062c28096ce354282083b0caf28de44
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1380
                            Width =1200
                            Height =255
                            Name ="Incomplete_Label"
                            Caption ="Incomplete"
                            GUID = Begin
                                0xee0a62988b4c4a4f95e97da68a635891
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =2460
                    Width =900
                    Height =255
                    ColumnWidth =1575
                    ColumnOrder =6
                    TabIndex =4
                    Name ="CarryOverCapacity"
                    ControlSource ="CarryOverCapacity"
                    GUID = Begin
                        0xa24fdf512ff33548bed83a4d9eb71b7e
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2460
                            Width =1200
                            Height =255
                            Name ="CarryOverCapacity_Label"
                            Caption ="CarryOverCapacity"
                            GUID = Begin
                                0x636bf6b52cdedf4092bd2907b3fb4e0d
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =2820
                    Width =900
                    Height =255
                    ColumnWidth =2130
                    ColumnOrder =7
                    TabIndex =5
                    Name ="CummulCarryOverCapacity"
                    ControlSource ="CummulCarryOverCapacity"
                    GUID = Begin
                        0x54377e44b6789c439a3371b3c4498f69
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2820
                            Width =1200
                            Height =255
                            Name ="CummulCarryOverCapacity_Label"
                            Caption ="CummulCarryOverCapacity"
                            GUID = Begin
                                0xd54dda80eff8114a9cbac72ac4c56c6e
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =3180
                    Width =900
                    Height =255
                    ColumnWidth =1155
                    ColumnOrder =8
                    TabIndex =6
                    Name ="NewCapacity"
                    ControlSource ="NewCapacity"
                    GUID = Begin
                        0x77ac043fea98064291ec350360ca25cf
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3180
                            Width =1200
                            Height =255
                            Name ="NewCapacity_Label"
                            Caption ="NewCapacity"
                            GUID = Begin
                                0x125cc1a415fc354db69b64bb971426b7
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =3540
                    Width =900
                    Height =255
                    ColumnWidth =870
                    ColumnOrder =5
                    TabIndex =7
                    Name ="Assigned"
                    ControlSource ="Assigned"
                    GUID = Begin
                        0xd0edcbb1e808864e858795d96fc659ad
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3540
                            Width =1200
                            Height =255
                            Name ="Assigned_Label"
                            Caption ="Assigned"
                            GUID = Begin
                                0x426d780b947ca84d866c40844be60653
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =3900
                    Width =900
                    Height =255
                    ColumnWidth =1455
                    ColumnOrder =9
                    TabIndex =8
                    Name ="LeftOverCapacity"
                    ControlSource ="LeftOverCapacity"
                    GUID = Begin
                        0xecd8f79e7b7dbb488be2b37337cabbd2
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3900
                            Width =1200
                            Height =255
                            Name ="LeftOverCapacity_Label"
                            Caption ="LeftOverCapacity"
                            GUID = Begin
                                0x3b712f39b161764d8b30fedc684a5075
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6120
                    Top =120
                    Width =1740
                    Height =255
                    ColumnWidth =1575
                    ColumnOrder =10
                    TabIndex =9
                    Name ="AgeGroupFillFactor"
                    ControlSource ="AgeGroupFillFactor"
                    GUID = Begin
                        0x599c4c8cf384094fb8b491aedfe8482c
                    End

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4860
                            Top =120
                            Width =1200
                            Height =255
                            Name ="AgeGroupFillFactor_Label"
                            Caption ="AgeGroupFillFactor"
                            GUID = Begin
                                0x120cc13466637343be4897338aaaba91
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1140
                    Top =1800
                    ColumnOrder =4
                    TabIndex =10
                    Name ="TotalAvailable"
                    ControlSource ="TotalAvailable"
                    GUID = Begin
                        0xdf62960b708cba49b55e50e829204809
                    End

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =720
                            Top =1800
                            Width =1125
                            Height =240
                            Name ="Label24"
                            Caption ="TotalAvailable:"
                            GUID = Begin
                                0x740599a79b08aa47894d759e2b437338
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
                0x21ee9098e02bb742a5c3a4f387c46f32
            End
        End
    End
End
