object FEtudes: TFEtudes
  Left = 217
  Top = 226
  Width = 863
  Height = 486
  Caption = 'documents de l'#39'appel d'#39'offre'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 31
    Top = 44
    Width = 59
    Height = 13
    Caption = 'Appel d'#39'offre'
  end
  object WinTV: TToolWindow97
    Left = 60
    Top = 108
    ClientHeight = 293
    ClientWidth = 269
    Caption = 'D'#233'finition de l'#39'appel d'#39'offre'
    ClientAreaHeight = 293
    ClientAreaWidth = 269
    TabOrder = 4
    Visible = False
    OnClose = WinTVClose
    object TV: TTreeView
      Left = 0
      Top = 0
      Width = 269
      Height = 293
      Align = alClient
      Color = clInfoBk
      Indent = 19
      TabOrder = 0
      OnClick = TVClick
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 66
    Width = 847
    Height = 310
    Align = alClient
    TabOrder = 0
    object GS: THGrid
      Tag = 1
      Left = 1
      Top = 1
      Width = 845
      Height = 308
      Align = alClient
      ColCount = 6
      DefaultRowHeight = 18
      RowCount = 50
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goTabs]
      TabOrder = 0
      OnDblClick = GSDblClick
      OnEnter = GSEnter
      SortedCol = -1
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = GSRowEnter
      OnCellEnter = GSCellEnter
      OnCellExit = GSCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      OnElipsisClick = GSElipsisClick
      ColWidths = (
        64
        64
        64
        64
        64
        64)
      RowHeights = (
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18)
    end
  end
  object DockBottom: TDock97
    Left = 0
    Top = 416
    Width = 847
    Height = 31
    BackgroundTransparent = True
    Position = dpBottom
    object LblErreur: TLabel
      Left = 260
      Top = 11
      Width = 14
      Height = 13
      Caption = 'XX'
    end
    object Outils97: TToolbar97
      Left = 0
      Top = 0
      Caption = 'Actions'
      CloseButton = False
      DefaultDock = DockBottom
      DockPos = 0
      TabOrder = 0
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 0
        Top = 0
        Width = 40
        Height = 27
        Hint = 'Menu zoom'
        Caption = 'Zoom'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
      object BInfos: TToolbarButton97
        Tag = 1
        Left = 40
        Top = 0
        Width = 40
        Height = 27
        Hint = 'Actions compl'#233'mentaires'
        Caption = 'Actions'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z0105_S16G2'
        IsControl = True
      end
      object BActionsLignes: TToolbarButton97
        Tag = 1
        Left = 80
        Top = 0
        Width = 40
        Height = 27
        Hint = 'Actions lignes'
        Caption = 'Actions'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z0543_S16G2'
        IsControl = True
      end
      object BDelete: TToolbarButton97
        Left = 202
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Supprimer la pi'#232'ce'
        Caption = 'Supprimer'
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BDeleteClick
        GlobalIndexImage = 'Z0005_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 258
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Caption = 'Imprimer'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BNewligne: TToolbarButton97
        Left = 165
        Top = 0
        Width = 37
        Height = 27
        Hint = 'Nouvel article'
        Caption = 'Nouveau document'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopCreat
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BArborescence: TToolbarButton97
        Tag = 1
        Left = 230
        Top = 0
        Width = 28
        Height = 27
        Hint = 'structure de l'#39'appel d'#39'offre'
        Caption = 'Arborescence Paragraphes'
        AllowAllUp = True
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BArborescenceClick
        GlobalIndexImage = 'Z0189_S16G0'
        IsControl = True
      end
      object bMail: TToolbarButton97
        Left = 120
        Top = 0
        Width = 45
        Height = 27
        Hint = 'Envoi mail intervenant'
        Caption = 'Envoi Mail'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopMail
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z2023_S16G1'
      end
    end
    object Valide97: TToolbar97
      Left = 749
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DefaultDock = DockBottom
      DockableTo = [dpTop, dpBottom, dpRight]
      DockPos = 749
      TabOrder = 1
      object BValider: TToolbarButton97
        Left = 0
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Enregistrer la pi'#232'ce'
        Caption = 'Enregistrer'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          BE060000424DBE06000000000000360400002800000024000000120000000100
          0800000000008802000000000000000000000001000000010000000000000000
          80000080000000808000800000008000800080800000C0C0C000C0DCC000F0C8
          A400000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          000000000000000000000000000000000000F0FBFF00A4A0A000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00030303030303
          0303030303030303030303030303030303030303030303030303030303030303
          03030303030303030303030303030303030303030303FF030303030303030303
          03030303030303040403030303030303030303030303030303F8F8FF03030303
          03030303030303030303040202040303030303030303030303030303F80303F8
          FF030303030303030303030303040202020204030303030303030303030303F8
          03030303F8FF0303030303030303030304020202020202040303030303030303
          0303F8030303030303F8FF030303030303030304020202FA0202020204030303
          0303030303F8FF0303F8FF030303F8FF03030303030303020202FA03FA020202
          040303030303030303F8FF03F803F8FF0303F8FF03030303030303FA02FA0303
          03FA0202020403030303030303F8FFF8030303F8FF0303F8FF03030303030303
          FA0303030303FA0202020403030303030303F80303030303F8FF0303F8FF0303
          0303030303030303030303FA0202020403030303030303030303030303F8FF03
          03F8FF03030303030303030303030303FA020202040303030303030303030303
          0303F8FF0303F8FF03030303030303030303030303FA02020204030303030303
          03030303030303F8FF0303F8FF03030303030303030303030303FA0202020403
          030303030303030303030303F8FF0303F8FF03030303030303030303030303FA
          0202040303030303030303030303030303F8FF03F8FF03030303030303030303
          03030303FA0202030303030303030303030303030303F8FFF803030303030303
          030303030303030303FA0303030303030303030303030303030303F803030303
          0303030303030303030303030303030303030303030303030303030303030303
          0303}
        GlyphMask.Data = {00000000}
        Layout = blGlyphTop
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 28
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Annuler'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAbandonClick
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 56
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Aide'
        HelpContext = 119000017
        Caption = 'Aide'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
  end
  object PENTETE: TPanel
    Left = 0
    Top = 0
    Width = 847
    Height = 66
    Align = alTop
    TabOrder = 1
    object LAFF_TIERS: TLabel
      Left = 28
      Top = 20
      Width = 26
      Height = 13
      Caption = 'Client'
    end
    object LAFF_AFFAIRE: TLabel
      Left = 28
      Top = 44
      Width = 59
      Height = 13
      Caption = 'Appel d'#39'offre'
    end
    object LAFF_NOMCLI: TLabel
      Left = 345
      Top = 20
      Width = 72
      Height = 13
      Caption = 'LAFF_NOMCLI'
    end
    object AFF_LIBELLE: TLabel
      Left = 345
      Top = 44
      Width = 67
      Height = 13
      Caption = 'AFF_LIBELLE'
    end
    object TAFF_REGOUPFACT: THLabel
      Left = 28
      Top = 92
      Width = 82
      Height = 13
      Caption = 'Regroup sur fact.'
      Visible = False
    end
    object TAFF_GENERAUTO: THLabel
      Left = 407
      Top = 91
      Width = 95
      Height = 13
      Caption = 'Mode de facturation'
      Visible = False
    end
    object BRechAffaire: TToolbarButton97
      Left = 320
      Top = 65
      Width = 23
      Height = 22
      Hint = 'Rechercher client/affaire'
      Enabled = False
      Opaque = False
      Visible = False
      OnClick = BRechAffaireClick
      GlobalIndexImage = 'Z0002_S16G1'
    end
    object bnewaff: TToolbarButton97
      Left = 88
      Top = 65
      Width = 23
      Height = 22
      Hint = 'Rechercher client/affaire'
      Opaque = False
      Visible = False
      OnClick = bnewaffClick
      GlobalIndexImage = 'Z0053_S16G1'
    end
    object LAFFAIRE: TLabel
      Left = 28
      Top = 70
      Width = 30
      Height = 13
      Caption = 'Affaire'
    end
    object AFF_TIERS: THCritMaskEdit
      Left = 113
      Top = 16
      Width = 109
      Height = 21
      Enabled = False
      TabOrder = 0
      Text = 'AFF_TIERS'
      TagDispatch = 0
    end
    object AFF_AFFAIRE1: THCritMaskEdit
      Left = 113
      Top = 40
      Width = 41
      Height = 21
      Enabled = False
      TabOrder = 1
      Text = 'AFF_AFFAIRE1'
      TagDispatch = 0
    end
    object AFF_AFFAIRE2: THCritMaskEdit
      Left = 159
      Top = 40
      Width = 57
      Height = 21
      Enabled = False
      TabOrder = 2
      Text = 'AFF_AFFAIRE2'
      TagDispatch = 0
    end
    object AFF_AFFAIRE3: THCritMaskEdit
      Left = 220
      Top = 40
      Width = 77
      Height = 21
      Enabled = False
      TabOrder = 3
      Text = 'AFF_AFFAIRE3'
      TagDispatch = 0
    end
    object AFF_AFFAIRE: THCritMaskEdit
      Left = 732
      Top = 12
      Width = 109
      Height = 21
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 4
      Text = 'AFF_AFFAIRE'
      Visible = False
      TagDispatch = 0
    end
    object AFF_AVENANT: THCritMaskEdit
      Left = 300
      Top = 40
      Width = 22
      Height = 21
      Color = clYellow
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 5
      Text = 'AFF_AVENANT'
      TagDispatch = 0
    end
    object REGROUPFACTBIS: THValComboBox
      Left = 113
      Top = 88
      Width = 153
      Height = 21
      ItemHeight = 13
      TabOrder = 6
      Visible = False
      OnChange = REGROUPFACTBISChange
      TagDispatch = 0
    end
    object AFF_GENERAUTO: THValComboBox
      Left = 510
      Top = 88
      Width = 125
      Height = 21
      ItemHeight = 13
      TabOrder = 7
      Visible = False
      OnChange = AFF_GENERAUTOChange
      TagDispatch = 0
    end
    object AFF_AFFAIRE0: THCritMaskEdit
      Left = 324
      Top = 40
      Width = 17
      Height = 21
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 8
      Text = 'AFF_AFFAIRE'
      Visible = False
      TagDispatch = 0
    end
    object NAFF_AFFAIRE1: THCritMaskEdit
      Left = 113
      Top = 66
      Width = 41
      Height = 21
      Enabled = False
      TabOrder = 9
      Text = 'NAFF_AFFAIRE1'
      TagDispatch = 0
    end
    object NAFF_AFFAIRE2: THCritMaskEdit
      Left = 159
      Top = 66
      Width = 57
      Height = 21
      Enabled = False
      TabOrder = 10
      Text = 'NAFF_AFFAIRE2'
      TagDispatch = 0
    end
    object NAFF_AFFAIRE3: THCritMaskEdit
      Left = 220
      Top = 66
      Width = 77
      Height = 21
      Enabled = False
      TabOrder = 11
      Text = 'NAFF_AFFAIRE3'
      TagDispatch = 0
    end
    object NAFF_AVENANT: THCritMaskEdit
      Left = 300
      Top = 66
      Width = 22
      Height = 21
      Color = clYellow
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 12
      Text = 'NAFF_AVENANT'
      TagDispatch = 0
    end
    object NAFF_AFFAIRE: THCritMaskEdit
      Left = 732
      Top = 32
      Width = 109
      Height = 21
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 13
      Text = 'AFF_AFFAIRE'
      Visible = False
      TagDispatch = 0
    end
    object NAFF_AFFAIRE0: THCritMaskEdit
      Left = 344
      Top = 68
      Width = 17
      Height = 21
      Color = clYellow
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 14
      Text = 'AFF_AFFAIRE'
      Visible = False
      TagDispatch = 0
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 376
    Width = 847
    Height = 40
    Align = alBottom
    TabOrder = 2
    object Etat: TLabel
      Left = 20
      Top = 12
      Width = 14
      Height = 13
      Caption = 'XX'
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 472
    Top = 21
  end
  object PopCreat: TPopupMenu
    Left = 212
    Top = 125
  end
  object PopExcel: TPopupMenu
    OnPopup = PopExcelPopup
    Left = 252
    Top = 213
    object ModifXls: TMenuItem
      Caption = 'Consultation/Modification du document Excel'
      OnClick = ModifXlsClick
    end
    object DefStructXls: TMenuItem
      Caption = 'Description de la structure du document Excel'
      OnClick = DefStructXlsClick
    end
    object RecupDonneeXls: TMenuItem
      Caption = 'Int'#233'gration des donn'#233'es dans l'#39'Etude PGI'
      Visible = False
      OnClick = RecupDonneeXlsClick
    end
    object MajDonneesXLS: TMenuItem
      Caption = 'Mise '#224' Jour du Bordereau de Prix'
      OnClick = MajDonneesXLSClick
    end
    object EntreeValDoc: TMenuItem
      Caption = 'Chiffrage dans l'#39'Etude PGI'
      Visible = False
      OnClick = EntreeValDocClick
    end
    object RenVoieBordereauXls: TMenuItem
      Caption = 'Renvoi du chiffrage dans le bordereau'
      Visible = False
      OnClick = RenVoieBordereauXlsClick
    end
    object RenvoieDonneeXls: TMenuItem
      Caption = 'Renvoi du chiffrage dans Excel'
      Visible = False
      OnClick = RenvoieDonneeXlsClick
    end
    object EditionXls: TMenuItem
      Caption = 'Edition du document Excel'
      OnClick = EditionXlsClick
    end
    object N2C: TMenuItem
      Caption = '-'
    end
    object SelectDocE: TMenuItem
      Caption = 'S'#233'lection du document'
      Visible = False
      OnClick = SelectDocEClick
    end
  end
  object PopDocPgi: TPopupMenu
    Left = 316
    Top = 132
    object Modificationdudocument1: TMenuItem
      Caption = 'Modification de l'#39#233'tude'
      OnClick = Modificationdudocument1Click
    end
    object EditionDocPGI: TMenuItem
      Caption = 'Edition de l'#39#233'tude'
      OnClick = EditionDocPGIClick
    end
    object SelectDocP: TMenuItem
      Caption = 'S'#233'lection de l'#39#233'tude'
      OnClick = SelectDocPClick
    end
  end
  object DOPen: TOpenDialog
    Title = 'Fichier source'
    Left = 412
    Top = 125
  end
  object DSave: TSaveDialog
    Options = [ofHideReadOnly, ofNoChangeDir, ofEnableSizing]
    Title = 'Fichier Destination'
    Left = 412
    Top = 193
  end
  object PopMail: TPopupMenu
    OnPopup = PopExcelPopup
    Left = 496
    Top = 237
    object EnvoiLotCotraitance: TMenuItem
      Caption = 'Envoi du lot au cotraitant'
      OnClick = EnvoiLotCotraitanceClick
    end
    object mni10: TMenuItem
      Caption = 'Envoi de l'#39'offre au client'
      OnClick = EnvoiLotGlobalClick
    end
  end
end
