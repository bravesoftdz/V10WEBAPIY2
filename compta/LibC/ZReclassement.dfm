object FReclassement: TFReclassement
  Left = 271
  Top = 105
  Width = 694
  Height = 478
  Caption = 'Reclassement des donn'#233'es'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 216
    Top = 36
    Height = 372
  end
  object DockBottom: TDock97
    Left = 0
    Top = 408
    Width = 678
    Height = 31
    Position = dpBottom
    object Valide97: TToolbar97
      Left = 585
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DockPos = 585
      TabOrder = 0
      object BValide: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Valider la saisie'
        Caption = 'Valider'
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
        Spacing = -1
        Visible = False
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 28
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Fermer'
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
        Spacing = -1
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
        Caption = 'Aide'
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
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
  end
  object CoolBar1: TCoolBar
    Left = 0
    Top = 0
    Width = 678
    Height = 36
    Bands = <
      item
        Control = CoolBar2
        ImageIndex = -1
        MinHeight = 31
        Width = 674
      end>
    object CoolBar2: TCoolBar
      Left = 9
      Top = 0
      Width = 661
      Height = 31
      Bands = <>
      object BTree: TToolbarButton97
        Left = 2
        Top = 0
        Width = 32
        Height = 26
        DropdownArrow = True
        DropdownMenu = POPB
        GlobalIndexImage = 'M0004_S16G1'
      end
      object BChoix: TToolbarButton97
        Left = 34
        Top = 0
        Width = 32
        Height = 26
        DropdownArrow = True
        DropdownMenu = POPC
        GlobalIndexImage = 'M0079_S16G1'
      end
      object BAff: TToolbarButton97
        Left = 66
        Top = 0
        Width = 32
        Height = 26
        DropdownArrow = True
        DropdownMenu = POPAFF
        GlobalIndexImage = 'Z0335_S16G1'
      end
      object BZoom: TToolbarButton97
        Left = 98
        Top = 0
        Width = 32
        Height = 26
        DropdownArrow = True
        DropdownMenu = POPZoom
        GlobalIndexImage = 'Z0061_S16G1'
      end
    end
  end
  object TW: TTreeView
    Left = 0
    Top = 36
    Width = 216
    Height = 372
    Align = alLeft
    Images = IM16
    Indent = 19
    PopupMenu = POPC
    ReadOnly = True
    TabOrder = 2
    OnChange = TWChange
    OnDragDrop = TWDragDrop
    OnDragOver = TWDragOver
    OnExpanding = TWExpanding
  end
  object LW: TListView
    Left = 219
    Top = 36
    Width = 459
    Height = 372
    Align = alClient
    Columns = <>
    LargeImages = IM32
    MultiSelect = True
    ReadOnly = True
    PopupMenu = POPC
    SmallImages = IM16
    TabOrder = 3
    ViewStyle = vsList
    OnDblClick = LWDblClick
  end
  object FParam: TToolWindow97
    Left = 242
    Top = 200
    ClientHeight = 131
    ClientWidth = 249
    Caption = 'Param'#232'tres'
    CloseButton = False
    ClientAreaHeight = 131
    ClientAreaWidth = 249
    Resizable = False
    TabOrder = 4
    Visible = False
    object L1: TLabel
      Left = 10
      Top = 14
      Width = 12
      Height = 13
      Caption = 'L1'
    end
    object L2: TLabel
      Left = 10
      Top = 45
      Width = 12
      Height = 13
      Caption = 'L2'
    end
    object L4: THLabel
      Left = 12
      Top = 72
      Width = 39
      Height = 13
      Caption = 'Auxilaire'
    end
    object E1: THCritMaskEdit
      Left = 137
      Top = 9
      Width = 103
      Height = 21
      TabOrder = 0
      TagDispatch = 0
      ElipsisButton = True
    end
    object E2: THCritMaskEdit
      Left = 137
      Top = 40
      Width = 103
      Height = 21
      TabOrder = 1
      OnChange = E2Change
      TagDispatch = 0
      ElipsisButton = True
    end
    object Panel1: TPanel
      Left = 0
      Top = 99
      Width = 249
      Height = 32
      Align = alBottom
      TabOrder = 3
      DesignSize = (
        249
        32)
      object BValider: TToolbarButton97
        Left = -15
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
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
        ModalResult = 1
        NumGlyphs = 2
        ParentFont = False
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 12
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Cancel = True
        Flat = False
        Glyph.Data = {
          CE070000424DCE07000000000000360000002800000024000000120000000100
          1800000000009807000000000000000000000000000000000000008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284008284008284008284
          0082840082840082840082840082840082848482848482840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284FFFFFF008284008284008284008284008284008284
          0082840082840082840082840082840082840082840082840082840000FF0000
          840000848482840082840082840082840082840082840000FF84828400828400
          8284008284008284008284008284008284008284848284848284FFFFFF008284
          008284008284008284008284008284FFFFFF0082840082840082840082840082
          840082840082840000FF00008400008400008484828400828400828400828400
          00FF000084000084848284008284008284008284008284008284008284848284
          FFFFFF008284848284FFFFFF008284008284008284FFFFFF848284848284FFFF
          FF0082840082840082840082840082840082840000FF00008400008400008400
          00848482840082840000FF000084000084000084000084848284008284008284
          008284008284008284848284FFFFFF008284008284848284FFFFFF008284FFFF
          FF848284008284008284848284FFFFFF00828400828400828400828400828400
          82840000FF000084000084000084000084848284000084000084000084000084
          000084848284008284008284008284008284008284848284FFFFFF0082840082
          84008284848284FFFFFF848284008284008284008284008284848284FFFFFF00
          82840082840082840082840082840082840000FF000084000084000084000084
          0000840000840000840000848482840082840082840082840082840082840082
          84008284848284FFFFFF00828400828400828484828400828400828400828400
          8284FFFFFF848284008284008284008284008284008284008284008284008284
          0000FF0000840000840000840000840000840000848482840082840082840082
          84008284008284008284008284008284008284848284FFFFFF00828400828400
          8284008284008284008284FFFFFF848284008284008284008284008284008284
          0082840082840082840082840082840000840000840000840000840000848482
          8400828400828400828400828400828400828400828400828400828400828400
          8284848284FFFFFF008284008284008284008284008284848284008284008284
          0082840082840082840082840082840082840082840082840082840000FF0000
          8400008400008400008484828400828400828400828400828400828400828400
          8284008284008284008284008284008284848284FFFFFF008284008284008284
          8482840082840082840082840082840082840082840082840082840082840082
          840082840000FF00008400008400008400008400008484828400828400828400
          8284008284008284008284008284008284008284008284008284008284848284
          008284008284008284008284848284FFFFFF0082840082840082840082840082
          840082840082840082840082840000FF00008400008400008484828400008400
          0084000084848284008284008284008284008284008284008284008284008284
          008284008284848284008284008284008284008284008284848284FFFFFF0082
          840082840082840082840082840082840082840082840000FF00008400008400
          00848482840082840000FF000084000084000084848284008284008284008284
          008284008284008284008284008284848284008284008284008284848284FFFF
          FF008284008284848284FFFFFF00828400828400828400828400828400828400
          82840000FF0000840000848482840082840082840082840000FF000084000084
          000084848284008284008284008284008284008284008284848284FFFFFF0082
          84008284848284008284848284FFFFFF008284008284848284FFFFFF00828400
          82840082840082840082840082840082840000FF000084008284008284008284
          0082840082840000FF0000840000840000840082840082840082840082840082
          84008284848284FFFFFFFFFFFF848284008284008284008284848284FFFFFF00
          8284008284848284FFFFFF008284008284008284008284008284008284008284
          0082840082840082840082840082840082840082840000FF0000840000FF0082
          8400828400828400828400828400828400828484828484828400828400828400
          8284008284008284848284FFFFFFFFFFFFFFFFFF848284008284008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284848284848284848284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284008284008284008284
          008284008284008284008284008284008284}
        GlyphMask.Data = {00000000}
        ModalResult = 2
        NumGlyphs = 2
        OnClick = BFermeClick
      end
    end
    object HE: THValComboBox
      Left = 137
      Top = 40
      Width = 103
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 4
      TagDispatch = 0
    end
    object CBDuplic: TCheckBox
      Left = 56
      Top = 0
      Width = 145
      Height = 17
      Alignment = taLeftJustify
      Caption = 'Duplication'
      TabOrder = 5
      Visible = False
    end
    object HV: THValComboBox
      Left = 137
      Top = 9
      Width = 103
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 6
      TagDispatch = 0
      Vide = True
      DataType = 'TTETABLISSEMENT'
    end
    object E4: THCritMaskEdit
      Left = 136
      Top = 68
      Width = 103
      Height = 21
      TabOrder = 2
      OnExit = E4Exit
      TagDispatch = 0
      ElipsisButton = True
    end
  end
  object FRech: TToolWindow97
    Left = 301
    Top = 33
    ClientHeight = 129
    ClientWidth = 305
    Caption = 'Rechercher...'
    ClientAreaHeight = 129
    ClientAreaWidth = 305
    TabOrder = 5
    Visible = False
    object Label1: TLabel
      Left = 8
      Top = 16
      Width = 51
      Height = 13
      Caption = 'rechercher'
    end
    object E3: THCritMaskEdit
      Left = 65
      Top = 13
      Width = 128
      Height = 21
      TabOrder = 0
      TagDispatch = 0
      ElipsisButton = True
    end
    object BSuivantRech: TButton
      Left = 216
      Top = 8
      Width = 75
      Height = 25
      Caption = 'Suivant'
      Default = True
      TabOrder = 1
      OnClick = BSuivantRechClick
    end
    object BAnnulerRech: TButton
      Left = 216
      Top = 40
      Width = 75
      Height = 25
      Caption = 'Annuler'
      TabOrder = 2
      OnClick = BAnnulerRechClick
    end
    object RadioGroup1: TRadioGroup
      Left = 8
      Top = 48
      Width = 185
      Height = 65
      Caption = 'Par'
      ItemIndex = 0
      Items.Strings = (
        'Journal'
        'Folio')
      TabOrder = 3
    end
  end
  object IM32: THImageList
    GlobalIndexImages.Strings = (
      'Z0757_S32G1'
      'Z1841_S32G1'
      'Z0650_S32G1'
      'Z0650_S32G1'
      'Z1528_S32G1'
      'Z1642_S32G1'
      'Z2099_S32G1'
      'Z0650_S32G1')
    Height = 32
    Width = 32
    Left = 65
    Top = 92
  end
  object POPB: TPopupMenu
    Left = 22
    Top = 44
    object Parexercice: TMenuItem
      Caption = 'Par p'#233'riode'
      Checked = True
      OnClick = ParexerciceClick
    end
    object Parjournal: TMenuItem
      Caption = 'Par journal'
      OnClick = ParjournalClick
    end
  end
  object POPC: TPopupMenu
    OnPopup = POPCPopup
    Left = 144
    Top = 84
    object MMOuv: TMenuItem
      Caption = 'Ouvrir...'
      Default = True
      OnClick = MMOuvClick
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object MMJal: TMenuItem
      Caption = 'Changement de Journal'
      OnClick = MMJalClick
    end
    object MMGen: TMenuItem
      Caption = 'Changement de compte'
      OnClick = MMGenClick
    end
    object MMAux: TMenuItem
      Caption = 'Changement par un auxiliaire'
      OnClick = MMAuxClick
    end
    object MMFolio: TMenuItem
      Caption = 'Changement de folio'
      OnClick = MMFolioClick
    end
    object MMPer: TMenuItem
      Caption = 'Changement de P'#233'riode'
      OnClick = MMPerClick
    end
    object MMEta: TMenuItem
      Caption = 'Changement d'#39#233'tablissement'
      OnClick = MMEtaClick
    end
    object MMAna: TMenuItem
      Caption = 'Contr'#244'le et r'#233'paration'
      Enabled = False
      Visible = False
      OnClick = MMAnaClick
    end
    object MMContr: TMenuItem
      Caption = 'Contr'#244'le'
      Enabled = False
      Visible = False
      OnClick = MMContrClick
    end
    object MMRecup: TMenuItem
      Caption = 'R'#233'cup'#233'ration...'
      Enabled = False
      Visible = False
      OnClick = MMRecupClick
    end
    object MMsauv: TMenuItem
      Caption = 'Sauvegarde...'
      Enabled = False
      Visible = False
      OnClick = MMsauvClick
    end
    object MMInsert: TMenuItem
      Caption = 'Insertion...'
      Enabled = False
      Visible = False
      OnClick = MMInsertClick
    end
  end
  object HMTrad: THSystemMenu
    Separator = True
    Traduction = False
    Left = 24
    Top = 96
  end
  object POPAFF: TPopupMenu
    Left = 104
    Top = 48
    object MMCodeExo: TMenuItem
      Caption = 'Affichage du libelle de l'#39'exercice'
      Checked = True
      OnClick = MMCodeExoClick
    end
    object MMAFFJAL: TMenuItem
      Caption = 'Affichage du libelle du journal'
      Checked = True
      OnClick = MMAFFJALClick
    end
    object N2: TMenuItem
      Caption = '-'
    end
    object MMGrand: TMenuItem
      Caption = 'Grandes ic'#244'nes'
      OnClick = MMGrandClick
    end
    object MMPetit: TMenuItem
      Caption = 'Petites ic'#244'nes'
      Checked = True
      OnClick = MMPetitClick
    end
    object N3: TMenuItem
      Caption = '-'
    end
  end
  object POPZoom: TPopupMenu
    Left = 104
    Top = 88
    object MMZoomJal: TMenuItem
      Caption = 'Voir le journal'
      OnClick = MMZoomJalClick
    end
  end
  object IM16: THImageList
    GlobalIndexImages.Strings = (
      'Z1052_S16G1'
      'Z0930_S16G1'
      'Z0930_S16G1'
      'Z1005_S16G1'
      'Z0021_S16G1'
      'Z0859_S16G1'
      'Z0369_S16G1'
      'Z0053_S16G1'
      'Z0016_S16G1'
      'Z0447_S16G1')
    Left = 60
    Top = 48
  end
end
