Version =20
VersionRequired =20
PublishOption =1
Checksum =-861425153
Begin Form
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9240
    DatasheetFontHeight =11
    ItemSuffix =15
    Left =585
    Top =3195
    Right =12810
    Bottom =6735
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x69fa9d0c5575e440
    End
    GUID = Begin
        0x60516980487887419f279e8875394496
    End
    RecordSource ="SELECT AuditClm_RelatedImage_Assign_Worktable.UserID, AuditClm_RelatedImage_Assi"
        "gn_Worktable.SelInd, AuditClm_RelatedImage_Assign_Worktable.ADRInstanceID, Audit"
        "Clm_RelatedImage_Assign_Worktable.CnlyClaimNum, AuditClm_RelatedImage_Assign_Wor"
        "ktable.ICN, AuditClm_RelatedImage_Assign_Worktable.PatFirstName, AuditClm_Relate"
        "dImage_Assign_Worktable.PatLastName, AuditClm_RelatedImage_Assign_Worktable.ClmF"
        "romDt, AuditClm_RelatedImage_Assign_Worktable.ClmThruDt, AuditClm_RelatedImage_A"
        "ssign_Worktable.CnlyProvId, AuditClm_RelatedImage_Assign_Worktable.PatDOB, Audit"
        "Clm_RelatedImage_Assign_Worktable.ClmStatus, AuditClm_RelatedImage_Assign_Workta"
        "ble.ClmStatusDesc\015\012FROM AuditClm_RelatedImage_Assign_Worktable\015\012WHER"
        "E 1 = 2;"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    NoSaveCTIWhenDisabled =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =7560
            Name ="Detail"
            GUID = Begin
                0x466d532a6fff1144967d591df9414d2f
            End
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =1320
                    Height =315
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="UserID"
                    ControlSource ="UserID"
                    GUID = Begin
                        0xd44a0c2427a31b4ca1faea50926e45bb
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =1320
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =1635
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1320
                            Width =690
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label0"
                            Caption ="UserID"
                            GUID = Begin
                                0x79c38f1716a1154fa11745e8f68ee925
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =1320
                            LayoutCachedWidth =810
                            LayoutCachedHeight =1635
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =1740
                    Height =315
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ADRInstanceID"
                    ControlSource ="ADRInstanceID"
                    GUID = Begin
                        0xc037056ca4806e4c930c8328b45f28e2
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =1740
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =2055
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1740
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label2"
                            Caption ="ADRInstanceID"
                            GUID = Begin
                                0x2ddea9fb3365c94dbb3f7371b8bcc3fe
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =1740
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =2055
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =2160
                    Height =315
                    ColumnOrder =2
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="CnlyClaimNum"
                    ControlSource ="CnlyClaimNum"
                    GUID = Begin
                        0x02699b8ef14b61418a78ec9c0373fb80
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =2160
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =2475
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2160
                            Width =1455
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label3"
                            Caption ="CnlyClaimNum"
                            GUID = Begin
                                0x80032abd0e5bd748aa300c24dfd98834
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =2160
                            LayoutCachedWidth =1575
                            LayoutCachedHeight =2475
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =2580
                    Height =315
                    ColumnWidth =1920
                    ColumnOrder =4
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ICN"
                    ControlSource ="ICN"
                    GUID = Begin
                        0x1954ba0065a3d746bb3b8a22db3ab680
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =2580
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =2895
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2580
                            Width =405
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label4"
                            Caption ="ICN"
                            GUID = Begin
                                0x501049b44700374f917157eaeee782b8
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =2580
                            LayoutCachedWidth =525
                            LayoutCachedHeight =2895
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =3000
                    Height =315
                    ColumnWidth =1680
                    ColumnOrder =7
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="PatFirstName"
                    ControlSource ="PatFirstName"
                    GUID = Begin
                        0x09fd72e14d40af428c6e49e2f26807e0
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =3000
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =3315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =3000
                            Width =1335
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label5"
                            Caption ="PatFirstName"
                            GUID = Begin
                                0x01902c4204b47a4ba612f0f18df387b7
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =3000
                            LayoutCachedWidth =1455
                            LayoutCachedHeight =3315
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =3420
                    Height =315
                    ColumnWidth =1635
                    ColumnOrder =6
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="PatLastName"
                    ControlSource ="PatLastName"
                    GUID = Begin
                        0x459d6856a5bfa44aa4927870be8ffa61
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =3420
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =3735
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =3420
                            Width =1290
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label6"
                            Caption ="PatLastName"
                            GUID = Begin
                                0xa24e36d8936415428c45c19b19d8fc7f
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =3420
                            LayoutCachedWidth =1410
                            LayoutCachedHeight =3735
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =3840
                    Height =315
                    ColumnOrder =8
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ClmFromDt"
                    ControlSource ="ClmFromDt"
                    GUID = Begin
                        0xd76c4a4cbc512c4fb510eb8223cee593
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =3840
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =4155
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =3840
                            Width =1125
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label7"
                            Caption ="ClmFromDt"
                            GUID = Begin
                                0x38cb405fcb10204384937a5b20385d8b
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =3840
                            LayoutCachedWidth =1245
                            LayoutCachedHeight =4155
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =4260
                    Height =315
                    ColumnOrder =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ClmThruDt"
                    ControlSource ="ClmThruDt"
                    GUID = Begin
                        0x601fad8fed2d0f4c9d8a72b04df9d46a
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =4260
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =4575
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =4260
                            Width =1065
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label8"
                            Caption ="ClmThruDt"
                            GUID = Begin
                                0xb04afdf626ab7247b19e599d9e08cfbb
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =4260
                            LayoutCachedWidth =1185
                            LayoutCachedHeight =4575
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =4680
                    Height =315
                    ColumnOrder =10
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="CnlyProvId"
                    ControlSource ="CnlyProvId"
                    GUID = Begin
                        0x609e711d32ccbb4fa174e37887daf91b
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =4680
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =4995
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =4680
                            Width =1080
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label9"
                            Caption ="CnlyProvId"
                            GUID = Begin
                                0x7022b1d13e30a14db0017c71210d74cc
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =4680
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =4995
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =240
                    Top =390
                    Width =1560
                    Height =360
                    ColumnWidth =705
                    ColumnOrder =3
                    TabIndex =9
                    BorderColor =10921638
                    Name ="SelInd"
                    ControlSource ="SelInd"
                    GUID = Begin
                        0x4eb82f5e7d14ed499677f8f7ff7dc94e
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =390
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =750
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =470
                            Top =360
                            Width =375
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label11"
                            Caption ="Sel"
                            GUID = Begin
                                0x2b9e978729c10243b7203a42e7bffa11
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =470
                            LayoutCachedTop =360
                            LayoutCachedWidth =845
                            LayoutCachedHeight =675
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5940
                    Top =1500
                    Height =315
                    ColumnOrder =5
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="PatDOB"
                    ControlSource ="PatDOB"
                    GUID = Begin
                        0x5260d6d89d9a1a43bf3a8abc4a04549f
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =5940
                    LayoutCachedTop =1500
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =1815
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4140
                            Top =1500
                            Width =780
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label12"
                            Caption ="PatDOB"
                            GUID = Begin
                                0x83faa9f107de8a4999a8e9ead0ea1d3d
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =4140
                            LayoutCachedTop =1500
                            LayoutCachedWidth =4920
                            LayoutCachedHeight =1815
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6000
                    Top =1980
                    Height =315
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ClmStatus"
                    ControlSource ="ClmStatus"
                    GUID = Begin
                        0xebaa2f7cd35cc4429201dc025baf3ea5
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =6000
                    LayoutCachedTop =1980
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =2295
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4200
                            Top =1980
                            Width =1005
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label13"
                            Caption ="ClmStatus"
                            GUID = Begin
                                0x24968bc7e6f62a4dacbe61ab4e91b3d2
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =4200
                            LayoutCachedTop =1980
                            LayoutCachedWidth =5205
                            LayoutCachedHeight =2295
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6000
                    Top =2520
                    Height =315
                    ColumnWidth =4950
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ClmStatusDesc"
                    ControlSource ="ClmStatusDesc"
                    GUID = Begin
                        0x63eef0b36aa5ff43a19c738a21d8efc7
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =6000
                    LayoutCachedTop =2520
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =2835
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4200
                            Top =2520
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label14"
                            Caption ="ClmStatusDesc"
                            GUID = Begin
                                0x6b61c10cf7f2a44285d9fa9d21d9ccfa
                            End
                            GridlineColor =10921638
                            LayoutCachedLeft =4200
                            LayoutCachedTop =2520
                            LayoutCachedWidth =5640
                            LayoutCachedHeight =2835
                        End
                    End
                End
            End
        End
    End
End
