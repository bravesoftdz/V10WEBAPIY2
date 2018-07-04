inherited FMulFactEff: TFMulFactEff
  Left = 94
  Top = 147
  Width = 732
  Height = 484
  Caption = 'Ecriture Comptable'
  PixelsPerInch = 96
  TextHeight = 13
  inherited iCritGlyphModified: THImage
    Left = 235
  end
  inherited iCritGlyph: THImage
    Left = 203
  end
  inherited Pages: THPageControl2
    Width = 724
    Height = 117
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      inherited Bevel1: TBevel
        Width = 716
        Height = 89
      end
      object HLabel4: THLabel
        Left = 255
        Top = 66
        Width = 37
        Height = 13
        Caption = '&G'#233'n'#233'ral'
        FocusControl = E_GENERAL
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TE_AUXILIAIRE: THLabel
        Left = 8
        Top = 64
        Width = 41
        Height = 13
        Caption = '&Auxiliaire'
        FocusControl = E_AUXILIAIRE
      end
      object TE_EXERCICE: THLabel
        Left = 10
        Top = 10
        Width = 41
        Height = 13
        Caption = 'E&xercice'
        FocusControl = E_EXERCICE
      end
      object TE_DATECOMPTABLE: THLabel
        Left = 255
        Top = 10
        Width = 103
        Height = 13
        Caption = '&Dates comptables  du'
        FocusControl = E_DATECOMPTABLE
      end
      object TE_DATECOMPTABLE2: THLabel
        Left = 468
        Top = 10
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object TE_DATEECHEANCE2: THLabel
        Left = 468
        Top = 38
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATEECHEANCE_
      end
      object TE_DATEECHEANCE: THLabel
        Left = 255
        Top = 38
        Width = 72
        Height = 13
        Caption = '&Ech'#233'ances du '
        FocusControl = E_DATEECHEANCE
      end
      object TE_JOURNAL: THLabel
        Left = 10
        Top = 38
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = E_JOURNAL
      end
      object E_GENERAL: THCritMaskEdit
        Left = 385
        Top = 62
        Width = 181
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 17
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 7
        TagDispatch = 0
        ElipsisButton = True
      end
      object E_AUXILIAIRE: THCritMaskEdit
        Left = 58
        Top = 62
        Width = 180
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 6
        TagDispatch = 0
        DataType = 'TZTTOUS'
        ElipsisButton = True
      end
      object E_EXERCICE: THValComboBox
        Left = 58
        Top = 6
        Width = 181
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 385
        Top = 6
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 1
        Text = '01/01/1900'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = od1900
        ElipsisButton = True
        ControlerDate = True
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 489
        Top = 6
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 2
        Text = '31/12/2099'
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        DefaultDate = od2099
        ElipsisButton = True
        ControlerDate = True
      end
      object E_DATEECHEANCE_: THCritMaskEdit
        Left = 489
        Top = 34
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '31/12/2099'
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        DefaultDate = od2099
        ElipsisButton = True
        ControlerDate = True
      end
      object E_DATEECHEANCE: THCritMaskEdit
        Left = 385
        Top = 34
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '01/01/1900'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = od1900
        ElipsisButton = True
        ControlerDate = True
      end
      object E_JOURNAL: THValComboBox
        Left = 58
        Top = 34
        Width = 180
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        DataType = 'TTJOURNAUX'
      end
      object XX_WHEREAN: TEdit
        Left = 101
        Top = 9
        Width = 13
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 8
        Text = 'E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H"'
        Visible = False
      end
      object E_NUMECHE: THCritMaskEdit
        Left = 120
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 9
        Text = '1'
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_QUALIFPIECE: THCritMaskEdit
        Left = 152
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 10
        Text = 'N'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_ECHE: THCritMaskEdit
        Left = 170
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 11
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_ETATLETTRAGE: THCritMaskEdit
        Left = 188
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 12
        Text = 'PL'
        Visible = False
        TagDispatch = 0
        Operateur = Inferieur
      end
      object E_TRESOLETTRE: THCritMaskEdit
        Left = 204
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 13
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object XX_WHERENTP: TEdit
        Left = 85
        Top = 9
        Width = 13
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 14
        Text = 
          'E_NATUREPIECE="FC" OR E_NATUREPIECE="AC" OR E_NATUREPIECE="OD"  ' +
          'OR E_NATUREPIECE="OC"'
        Visible = False
      end
      object XX_WHEREMONTANT: TEdit
        Left = 69
        Top = 9
        Width = 13
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 15
        Text = 'E_DEBIT+E_CREDIT-E_COUVERTURE>0'
        Visible = False
      end
      object XX_WHERESEL: TEdit
        Left = 53
        Top = 9
        Width = 13
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 16
        Visible = False
      end
    end
    inherited PComplement: THTabSheet
      inherited Bevel2: TBevel
        Width = 716
        Height = 89
      end
      object TE_NUMEROPIECE: THLabel
        Left = 8
        Top = 37
        Width = 42
        Height = 13
        Caption = '&Pi'#232'ce de'
        FocusControl = E_NUMEROPIECE
      end
      object HLabel2: THLabel
        Left = 170
        Top = 35
        Width = 6
        Height = 13
        Caption = #224
      end
      object TE_ETABLISSEMENT: THLabel
        Left = 8
        Top = 10
        Width = 65
        Height = 13
        Caption = '&Etablissement'
        FocusControl = E_ETABLISSEMENT
      end
      object Label14: TLabel
        Left = 296
        Top = 37
        Width = 88
        Height = 13
        Caption = '&Mode de paiement'
        FocusControl = E_MODEPAIE
      end
      object TE_DEVISE: THLabel
        Left = 8
        Top = 66
        Width = 33
        Height = 13
        Caption = '&Devise'
        Enabled = False
        FocusControl = E_DEVISE
      end
      object TE_NATUREPIECE: THLabel
        Left = 296
        Top = 10
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = E_NATUREPIECE
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Left = 82
        Top = 33
        Width = 77
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_NUMEROPIECE_: THCritMaskEdit
        Left = 184
        Top = 33
        Width = 77
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_ETABLISSEMENT: THValComboBox
        Left = 82
        Top = 6
        Width = 181
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
      object E_MODEPAIE: THValComboBox
        Tag = -9980
        Left = 396
        Top = 33
        Width = 171
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        Vide = True
        DataType = 'TTMODEPAIE'
      end
      object E_DEVISE: THValComboBox
        Tag = 1
        Left = 82
        Top = 62
        Width = 181
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        Enabled = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        DataType = 'TTDEVISE'
      end
      object E_NATUREPIECE: THValComboBox
        Left = 396
        Top = 6
        Width = 171
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATPIECEVENTE'
      end
    end
    inherited PAvance: THTabSheet
      inherited Bevel4: TBevel
        Width = 716
        Height = 89
      end
      inherited bEffaceAvance: TToolbarButton97
        Left = 577
      end
      inherited Z_C1: THValComboBox
        Left = 6
        Width = 151
      end
      inherited Z_C2: THValComboBox
        Left = 6
        Width = 151
      end
      inherited Z_C3: THValComboBox
        Left = 6
        Width = 151
      end
      inherited ZO3: THValComboBox
        Left = 167
        Width = 151
      end
      inherited ZO2: THValComboBox
        Left = 167
        Width = 151
      end
      inherited ZO1: THValComboBox
        Left = 167
        Width = 151
      end
      inherited ZV1: THEdit
        Left = 331
      end
      inherited ZV2: THEdit
        Left = 331
      end
      inherited ZV3: THEdit
        Left = 331
      end
      inherited ZG2: THCombobox
        Left = 531
      end
      inherited ZG1: THCombobox
        Left = 531
      end
    end
    inherited PSQL: THTabSheet
      inherited Bevel3: TBevel
        Width = 716
        Height = 89
      end
      inherited Z_SQL: THSQLMemo
        Width = 716
        Height = 89
      end
    end
  end
  inherited Dock971: TDock97
    Top = 117
    Width = 724
    inherited PFiltres: TToolWindow97
      ClientWidth = 724
      ClientAreaWidth = 724
      inherited BCherche: TToolbarButton97
        Left = 508
      end
      inherited lpresentation: THLabel
        Left = 541
      end
      inherited FFiltres: THValComboBox
        Width = 433
      end
      inherited cbPresentations: THValComboBox
        Left = 612
      end
    end
  end
  inherited FListe: THDBGrid
    Top = 158
    Width = 695
    Height = 237
    DragCursor = crMultiDrag
    OnDragDrop = FListeDragDrop
    OnDragOver = FListeDragOver
    OnEndDrag = FListeEndDrag
    OnMouseMove = FListeMouseMove
    OnFlipSelection = FListeFlipSelection
    MultiSelection = True
    MultiFieds = 
      'E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_NUMEROPIECE;E_QUALIFPIECE' +
      ';E_NUMLIGNE;E_NUMECHE;'
    SortEnabled = True
  end
  inherited Panel2: THPanel
    Left = 425
    inherited PListe: THPanel
      Left = 516
    end
  end
  inherited Dock: TDock97
    Top = 417
    Width = 724
    Height = 40
    inherited PanelBouton: TToolWindow97
      ClientHeight = 36
      ClientWidth = 724
      ClientAreaHeight = 36
      ClientAreaWidth = 724
      inherited Binsert: TToolbarButton97 [0]
        Left = 456
      end
      inherited BReduire: TToolbarButton97 [1]
      end
      inherited BAgrandir: TToolbarButton97 [2]
      end
      inherited bSelectAll: TToolbarButton97 [3]
        Visible = True
      end
      inherited BRechercher: TToolbarButton97 [4]
      end
      inherited BParamListe: TToolbarButton97 [5]
      end
      inherited bExport: TToolbarButton97 [6]
      end
      inherited BImprimer: TToolbarButton97 [7]
        Left = 597
      end
      inherited BOuvrir: TToolbarButton97 [8]
        Left = 629
        Hint = 'Lancer le traitement'
        Caption = 'Traitement'
      end
      inherited BAnnuler: TToolbarButton97 [9]
        Left = 661
      end
      inherited BAide: TToolbarButton97 [10]
        Left = 693
      end
      inherited BBlocNote: TToolbarButton97
        Left = 305
        Visible = True
      end
      object bVoirTotSel: TToolbarButton97
        Left = 164
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Voir/masquer le total de la liste des mouvements s'#233'lectionn'#233's'
        AllowAllUp = True
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Caption = 'S'#233'lectionner'
        Flat = False
        Layout = blGlyphTop
        OnClick = bVoirTotSelClick
        GlobalIndexImage = 'Z0661_S16G1'
      end
    end
  end
  object FTotSel: TToolbar97 [7]
    Left = 24
    Top = 241
    Caption = 'Outils de s'#233'lection'
    CloseButton = False
    DockPos = 0
    TabOrder = 5
    Visible = False
    object FPied: TPanel
      Left = 0
      Top = 0
      Width = 530
      Height = 36
      Align = alBottom
      BevelInner = bvRaised
      BevelOuter = bvNone
      TabOrder = 0
      object FSession: TLabel
        Left = 3
        Top = 1
        Width = 155
        Height = 13
        Alignment = taCenter
        Caption = 'Mouvements s'#233'lectionn'#233's :'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
      end
      object Label9: TLabel
        Left = 199
        Top = 1
        Width = 17
        Height = 13
        Alignment = taCenter
        Caption = 'Nb'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
      end
      object Label10: TLabel
        Left = 281
        Top = 1
        Width = 31
        Height = 13
        Caption = 'D'#233'bit'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
      end
      object Label11: TLabel
        Left = 370
        Top = 1
        Width = 34
        Height = 13
        Caption = 'Cr'#233'dit'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
      end
      object Label12: TLabel
        Left = 482
        Top = 1
        Width = 33
        Height = 13
        Caption = 'Solde'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
      end
      object Label13: TLabel
        Left = 20
        Top = 19
        Width = 125
        Height = 15
        AutoSize = False
        Caption = 'Totaux'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
        WordWrap = True
      end
      object ISigneEuro1: TImage
        Left = 423
        Top = 3
        Width = 16
        Height = 16
        AutoSize = True
        Picture.Data = {
          07544269746D6170F6000000424DF60000000000000076000000280000001000
          0000100000000100040000000000800000000000000000000000100000001000
          0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
          C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
          FF00FFFFFFFFFFFFFFFFFFFFFFF44444444FFFFFF44444444444FFFF4444FFFF
          FFF4FFFF444FFFFFFFFFFFFF44FFFFFFFFFFF44444444444FFFFFF4444444444
          4FFFFFF44FFFFFFFFFFFF444444444444FFFFF444444444444FFFFFF44FFFFFF
          FFF4FFFF444FFFFFFF44FFFFF444FFFFF444FFFFFF4444444444FFFFFFF44444
          4FF4}
        Stretch = True
        Transparent = True
        Visible = False
      end
      object FNbP: THNumEdit
        Tag = -1
        Left = 158
        Top = 19
        Width = 60
        Height = 15
        TabStop = False
        AutoSize = False
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Masks.PositiveMask = '#,##0'
        Debit = False
        ParentCtl3D = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 0
        UseRounding = True
        Validate = False
      end
      object FDebitPA: THNumEdit
        Left = 221
        Top = 19
        Width = 91
        Height = 15
        TabStop = False
        AutoSize = False
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        ParentCtl3D = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 1
        UseRounding = True
        Validate = False
      end
      object FCreditPA: THNumEdit
        Left = 312
        Top = 19
        Width = 91
        Height = 15
        TabStop = False
        AutoSize = False
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        ParentCtl3D = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 2
        UseRounding = True
        Validate = False
      end
      object FSoldeA: THNumEdit
        Left = 407
        Top = 19
        Width = 110
        Height = 15
        TabStop = False
        AutoSize = False
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Masks.PositiveMask = '#,##0.00'
        Masks.ZeroMask = '0.00'
        Debit = False
        ParentCtl3D = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 3
        UseRounding = True
        Validate = False
      end
    end
  end
  inherited PCumul: THPanel
    Top = 395
    Width = 724
    TabOrder = 9
  end
  inherited TBlocNote: TToolWindow97
    Left = 209
    Top = 97
  end
  object WFliste2: TToolbar97 [10]
    Left = 8
    Top = 216
    Caption = 'Outils de s'#233'lection'
    CloseButton = False
    DockPos = 0
    DockRow = 1
    TabOrder = 8
    object Panel1: TPanel
      Left = 0
      Top = 0
      Width = 678
      Height = 100
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 0
      object Label3: TLabel
        Left = 0
        Top = 0
        Width = 105
        Height = 100
        Align = alLeft
        AutoSize = False
        Caption = 
          'Pour choisir les '#233'ch'#233'ances, s'#233'lectionner-les dans la liste des m' +
          'ouvements '#224' s'#233'lectionner et faites les "glisser" ici -->'
        WordWrap = True
      end
      object FListe2: THGrid
        Left = 105
        Top = 0
        Width = 573
        Height = 100
        Align = alClient
        ColCount = 9
        DefaultRowHeight = 18
        FixedCols = 0
        RowCount = 10
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
        TabOrder = 0
        OnDragDrop = Fliste2DragDrop
        OnDragOver = Fliste2DragOver
        OnEndDrag = FListe2EndDrag
        OnMouseMove = FListe2MouseMove
        SortedCol = -1
        Titres.Strings = (
          'E_JOURNAL'
          'E_GENERAL'
          'E_AUXILIAIRE'
          'E_DATECOMPTABLE'
          'E_DATEECHEANCE'
          'E_NUMEROPIECE'
          'E_DEBIT'
          'E_CREDIT'
          'IND')
        Couleur = False
        MultiSelect = True
        TitleBold = False
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = 13224395
        ColWidths = (
          43
          75
          84
          79
          73
          57
          62
          71
          72)
      end
    end
  end
  inherited PanVBar: THPanel
    Left = 695
    Top = 158
    Height = 237
    inherited BLineDown: TToolbarButton97
      Top = 82
    end
    inherited BPageDown: TToolbarButton97
      Top = 39
    end
  end
  inherited SQ: TDataSource
    Left = 104
  end
  inherited Q: THQuery
    Left = 56
    Top = 353
  end
  inherited FindDialog: THFindDialog
    Left = 148
  end
  inherited POPF: THPopupMenu
    Left = 156
    Top = 352
  end
  inherited HMTrad: THSystemMenu
    Left = 192
  end
  inherited SD: THSaveDialog
    Left = 280
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Validation impossible : Vous n'#39'avez pas s'#233'lectionnn'#233 +
        ' les comptes de l'#39#233'criture en cours de saisie;W;O;O;O;'
      'Liste des mouvements du compte '#224' s'#233'lectionner'
      
        '2;?caption?;Validation impossible : Le montant global est cr'#233'dit' +
        'eur;W;O;O;O;'
      
        '3;?Caption?;Confirmez-vous l'#39'affectation des '#233'ch'#233'ances s'#233'lection' +
        'n'#233'es?;Q;YN;Y;N;'
      
        '4;?Caption?;Confirmez-vous la suppression des '#233'ch'#233'ances s'#233'lectio' +
        'nn'#233'es du lot ?;Q;YN;Y;N;'
      'Liste des mouvements s'#233'lectionn'#233's :'
      'D'#233'bit'
      'Cr'#233'dit'
      'Solde'
      
        '9;?Caption?;Confirmez-vous les mouvements s'#233'lectionn'#233's ?;Q;YN;Y;' +
        'N;'
      
        '10;?Caption?;Un ou plusieurs mouvements n'#39'ont pas '#233't'#233' pris en co' +
        'mpte car ils sont d'#233'ja s'#233'lectionn'#233's sur une autre ligne.;W;O;O;O' +
        ';'
      '')
    Left = 87
    Top = 357
  end
end
