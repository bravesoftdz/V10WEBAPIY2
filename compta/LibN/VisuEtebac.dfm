object FVisuEtebac: TFVisuEtebac
  Left = 351
  Top = 176
  Width = 621
  Height = 404
  Caption = 'Contr'#244'le d'#39'exhaustivit'#233
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 613
    Height = 116
    ActivePage = PCritere
    Align = alTop
    HotTrack = True
    TabOrder = 0
    TabWidth = 100
    object PCritere: TTabSheet
      Caption = 'Standards'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 605
        Height = 88
        Align = alClient
      end
      object HLabel1: THLabel
        Left = 15
        Top = 10
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = EdtJournal
      end
      object HLabel2: THLabel
        Left = 15
        Top = 38
        Width = 14
        Height = 13
        Caption = '&Du'
        FocusControl = edtMulDateDu
      end
      object HLabel3: THLabel
        Left = 240
        Top = 40
        Width = 13
        Height = 13
        Caption = '&Au'
        FocusControl = EdtMulDateAu
      end
      object HLabel4: THLabel
        Left = 15
        Top = 65
        Width = 50
        Height = 13
        Caption = '&R'#233'f'#233'rence'
        FocusControl = edtReference
      end
      object HLabel5: THLabel
        Left = 240
        Top = 65
        Width = 46
        Height = 13
        Caption = '&Amplitude'
        FocusControl = EdtAmplitude
      end
      object EdtJournal: THCritMaskEdit
        Left = 105
        Top = 10
        Width = 121
        Height = 21
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TTJOURNAUX'
        ElipsisButton = True
      end
      object edtMulDateDu: THCritMaskEdit
        Left = 105
        Top = 35
        Width = 121
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 1
        Text = '  /  /    '
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = odDate
        ElipsisButton = True
        ControlerDate = True
      end
      object EdtMulDateAu: THCritMaskEdit
        Left = 300
        Top = 35
        Width = 121
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 2
        Text = '  /  /    '
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = odDate
        ElipsisButton = True
        ControlerDate = True
      end
      object edtReference: THCritMaskEdit
        Left = 105
        Top = 60
        Width = 121
        Height = 21
        TabOrder = 3
        TagDispatch = 0
      end
      object EdtAmplitude: THCritMaskEdit
        Left = 300
        Top = 60
        Width = 121
        Height = 21
        TabOrder = 4
        TagDispatch = 0
      end
    end
    object TabSheet1: TTabSheet
      Caption = 'Compl'#233'ments'
      ImageIndex = 1
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 605
        Height = 88
        Align = alClient
      end
      object HLabel6: THLabel
        Left = 15
        Top = 15
        Width = 27
        Height = 13
        Caption = '&Table'
        FocusControl = CmbChoixTable
      end
      object HLabel7: THLabel
        Left = 15
        Top = 45
        Width = 38
        Height = 13
        Caption = '&Champs'
        FocusControl = CmbChamps
      end
      object CmbChoixTable: THValComboBox
        Left = 90
        Top = 10
        Width = 146
        Height = 21
        ItemHeight = 13
        TabOrder = 0
        OnChange = CmbChoixTableChange
        Items.Strings = (
          'Ecriture'
          'Saisie de tr'#233'sorerie')
        TagDispatch = 0
        Values.Strings = (
          'ECRITURE'
          'CRELBQE')
      end
      object CmbChamps: THValComboBox
        Left = 90
        Top = 40
        Width = 146
        Height = 21
        ItemHeight = 13
        TabOrder = 1
        OnChange = CmbChampsChange
        Items.Strings = (
          'R'#233'f'#233'rence interne'
          'R'#233'f'#233'rence externe')
        TagDispatch = 0
        Values.Strings = (
          '_REFINTERNE'
          '_REFEXTERNE')
      end
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 116
    Width = 613
    Height = 37
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 33
      ClientWidth = 604
      Caption = 'Filtres'
      ClientAreaHeight = 33
      ClientAreaWidth = 604
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        604
        33)
      object BFiltre: TToolbarButton97
        Left = 6
        Top = 7
        Width = 56
        Height = 21
        Hint = 'Menu filtre'
        DropdownArrow = True
        DropdownMenu = POPF
        Caption = '&Filtres'
        Flat = False
        Layout = blGlyphRight
      end
      object BCherche: TToolbarButton97
        Left = 569
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        Anchors = [akTop, akRight]
        Flat = False
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object FFiltres: THValComboBox
        Left = 68
        Top = 7
        Width = 493
        Height = 21
        Style = csDropDownList
        Anchors = [akLeft, akTop, akRight]
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
      end
    end
  end
  object Dock972: TDock97
    Left = 0
    Top = 341
    Width = 613
    Height = 36
    Position = dpBottom
    object PanelBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 604
      Caption = 'Barre outils multicrit'#232're'
      ClientAreaHeight = 32
      ClientAreaWidth = 604
      DefaultDock = Dock972
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        604
        32)
      object BReduire: TToolbarButton97
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'R'#233'duire la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'R'#233'duire'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Visible = False
        OnClick = BReduireClick
        GlobalIndexImage = 'Z0910_S16G1'
      end
      object BAgrandir: TToolbarButton97
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Agrandir la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Agrandir'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        ParentFont = False
        OnClick = BAgrandirClick
        GlobalIndexImage = 'Z0270_S16G1'
      end
      object bSelectAll: TToolbarButton97
        Left = 132
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Tout s'#233'lectionner'
        AllowAllUp = True
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Caption = 'S'#233'lectionner'
        Flat = False
        Layout = blGlyphTop
        Visible = False
        GlobalIndexImage = 'Z0205_S16G1'
      end
      object BRechercher: TToolbarButton97
        Left = 36
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Chercher'
        Flat = False
        Layout = blGlyphTop
        Visible = False
        GlobalIndexImage = 'Z0077_S16G1'
      end
      object BParamListe: TToolbarButton97
        Left = 68
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Param'#233'trer la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Liste'
        Flat = False
        Layout = blGlyphTop
        Visible = False
        GlobalIndexImage = 'Z0025_S16G1'
      end
      object bExport: TToolbarButton97
        Left = 100
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Exporter la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Exporter'
        Enabled = False
        Flat = False
        Layout = blGlyphTop
        Visible = False
        GlobalIndexImage = 'Z0724_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 476
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Imprimer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        ParentFont = False
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BOuvrir: TToolbarButton97
        Left = 508
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Ouvrir'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Ouvrir'
        Flat = False
        Layout = blGlyphTop
        NumGlyphs = 2
        Visible = False
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BAnnuler: TToolbarButton97
        Left = 540
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Anchors = [akTop, akRight]
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        Layout = blGlyphTop
        ModalResult = 2
        OnClick = BAnnulerClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 572
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        Layout = blGlyphTop
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object Binsert: TToolbarButton97
        Left = 164
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Nouveau'
        DisplayMode = dmGlyphOnly
        Caption = 'Nouveau'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Visible = False
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BBlocNote: TToolbarButton97
        Left = 444
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Bloc-Note'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        GroupIndex = 2
        DisplayMode = dmGlyphOnly
        Caption = 'Bloc-note'
        Flat = False
        Visible = False
        GlobalIndexImage = 'Z0029_S16G1'
      end
    end
  end
  object FListe: THGrid
    Left = 0
    Top = 153
    Width = 613
    Height = 188
    Align = alClient
    ColCount = 3
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSelect]
    TabOrder = 3
    SortedCol = -1
    Titres.Strings = (
      'Num'#233'ro'
      'Commentaire'
      'Date trouv'#233'e')
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = 13224395
    ColWidths = (
      121
      348
      130)
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 454
    Top = 49
  end
  object POPF: TPopupMenu
    OwnerDraw = True
    Left = 488
    Top = 49
  end
end
