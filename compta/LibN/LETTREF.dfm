object FLET: TFLET
  Left = 476
  Top = 158
  Width = 409
  Height = 382
  Caption = 'Lettrage sur code de regroupement'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 401
    Height = 279
    ActivePage = PStandards
    Align = alClient
    TabOrder = 1
    object PStandards: TTabSheet
      Caption = 'Standards'
      object TCPTEDEBUT: THLabel
        Left = 10
        Top = 49
        Width = 56
        Height = 13
        Caption = 'Comptes &de'
        FocusControl = CPTEDEBUT
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TCPTEFIN: THLabel
        Left = 256
        Top = 49
        Width = 6
        Height = 13
        Caption = '&'#224
        FocusControl = CPTEFIN
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TFExercice: THLabel
        Left = 10
        Top = 152
        Width = 88
        Height = 13
        AutoSize = False
        Caption = 'E&xercice'
        FocusControl = FExercice
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TFDateCpta1: THLabel
        Left = 10
        Top = 180
        Width = 124
        Height = 13
        AutoSize = False
        Caption = '&Date comptable du'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TFDateCpta2: TLabel
        Left = 252
        Top = 180
        Width = 15
        Height = 13
        AutoSize = False
        Caption = 'au'
      end
      object TFNomChamp: TLabel
        Left = 10
        Top = 101
        Width = 108
        Height = 13
        Caption = '&Zone de regroupement'
      end
      object TFCodeLettre1: TLabel
        Left = 10
        Top = 127
        Width = 78
        Height = 13
        Caption = 'Code &lettrage de'
        FocusControl = FCodeLettre1
      end
      object TFCodeLettre2: THLabel
        Left = 256
        Top = 127
        Width = 6
        Height = 13
        Caption = '&'#224
        FocusControl = CPTEFIN
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TCollDebut: THLabel
        Left = 10
        Top = 75
        Width = 57
        Height = 13
        Caption = 'Collectifs &de'
        FocusControl = CollDebut
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TCollFin: THLabel
        Left = 256
        Top = 75
        Width = 6
        Height = 13
        Caption = '&'#224
        FocusControl = CollFin
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Hlabel2: THLabel
        Left = 236
        Top = 17
        Width = 35
        Height = 13
        Caption = '&Nature '
        FocusControl = FNatCpt
      end
      object CPTEDEBUT: THCpteEdit
        Left = 140
        Top = 45
        Width = 93
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        ZoomTable = tzTiers
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object CPTEFIN: THCpteEdit
        Left = 273
        Top = 45
        Width = 93
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        ZoomTable = tzTiers
        Vide = False
        Bourre = True
        okLocate = False
        SynJoker = False
      end
      object FExercice: THValComboBox
        Left = 139
        Top = 148
        Width = 230
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 9
        OnChange = FExerciceChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object FCodeLettre1: TEdit
        Left = 139
        Top = 123
        Width = 93
        Height = 21
        TabOrder = 7
      end
      object FTypeCompte: TRadioGroup
        Left = 12
        Top = -1
        Width = 219
        Height = 41
        Caption = 'Type de compte '#224' lettrer'
        Columns = 2
        ItemIndex = 1
        Items.Strings = (
          'G'#233'n'#233'raux'
          'Auxiliaires')
        TabOrder = 0
        OnClick = FTypeCompteClick
      end
      object FCodeLettre2: TEdit
        Left = 273
        Top = 123
        Width = 93
        Height = 21
        TabOrder = 8
      end
      object FNomChamp: THValComboBox
        Left = 139
        Top = 97
        Width = 230
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 6
        OnChange = FExerciceChange
        Items.Strings = (
          'E_REFINTERNE'
          'E_REFEXTERNE'
          'E_REFLIBRE'
          'E_AFFAIRE'
          'E_LIBELLE'
          'E_LIBRETEXTE0'
          'E_LIBRETEXTE1'
          'E_LIBRETEXTE2'
          'E_LIBRETEXTE3'
          'E_LIBRETEXTE4'
          'E_LIBRETEXTE5'
          'E_LIBRETEXTE6'
          'E_LIBRETEXTE7'
          'E_LIBRETEXTE8'
          'E_LIBRETEXTE9'
          'E_REFRELEVE'
          'E_REFLETTRAGE'
          'E_NOMLOT')
        TagDispatch = 0
        Values.Strings = (
          'E_REFINTERNE'
          'E_REFEXTERNE'
          'E_REFLIBRE'
          'E_AFFAIRE'
          'E_LIBELLE'
          'E_LIBRETEXTE0'
          'E_LIBRETEXTE1'
          'E_LIBRETEXTE2'
          'E_LIBRETEXTE3'
          'E_LIBRETEXTE4'
          'E_LIBRETEXTE5'
          'E_LIBRETEXTE6'
          'E_LIBRETEXTE7'
          'E_LIBRETEXTE8'
          'E_LIBRETEXTE9'
          'E_REFRELEVE'
          'E_REFLETTRAGE'
          'E_NOMLOT')
      end
      object FTraitePartiel: TCheckBox
        Left = 10
        Top = 204
        Width = 141
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Traiter le lettrage &partiel'
        Checked = True
        State = cbChecked
        TabOrder = 10
        OnClick = FTraitePartielClick
      end
      object FMontant: TCheckBox
        Left = 10
        Top = 230
        Width = 141
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Lettrage sur &montant'
        TabOrder = 12
        OnClick = FMontantClick
      end
      object FMontantNeg: TCheckBox
        Left = 195
        Top = 230
        Width = 190
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Traiter les montants &n'#233'gatifs'
        Enabled = False
        TabOrder = 13
        OnClick = FMontantClick
      end
      object CollDebut: THCpteEdit
        Left = 140
        Top = 71
        Width = 93
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        ZoomTable = tzGCollectif
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object CollFin: THCpteEdit
        Left = 273
        Top = 71
        Width = 93
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        ZoomTable = tzGCollectif
        Vide = False
        Bourre = True
        okLocate = False
        SynJoker = False
      end
      object FNatCpt: THValComboBox
        Left = 273
        Top = 13
        Width = 93
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 1
        OnChange = FNatCptChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATTIERSCPTA'
      end
      object CBOkDEbCre: TCheckBox
        Left = 195
        Top = 204
        Width = 190
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Lettrage si d'#233'bit et cr'#233'dit renseign'#233
        TabOrder = 11
        OnClick = FMontantClick
      end
      object FDateCpta: THCritMaskEdit
        Left = 139
        Top = 176
        Width = 93
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 14
        Text = '  /  /    '
        OnKeyPress = FDateCptaKeyPress
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = od1900
        ControlerDate = True
      end
      object FDateCpta_: THCritMaskEdit
        Left = 273
        Top = 176
        Width = 92
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 15
        Text = '  /  /    '
        OnKeyPress = FDateCptaKeyPress
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = od2099
        ControlerDate = True
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Compl'#233'ments'
      ImageIndex = 2
      object F1: TLabel
        Left = 64
        Top = 85
        Width = 18
        Height = 23
        Caption = '['
        Enabled = False
        Font.Charset = SYMBOL_CHARSET
        Font.Color = clMaroon
        Font.Height = -19
        Font.Name = 'Wingdings 3'
        Font.Style = []
        ParentFont = False
      end
      object TChampMAJ: TLabel
        Left = 84
        Top = 90
        Width = 74
        Height = 13
        Caption = 'Champ '#224' utiliser'
        Enabled = False
      end
      object TValMAJ: TLabel
        Left = 84
        Top = 116
        Width = 107
        Height = 13
        Caption = 'Valeur pour mise '#224' jour'
        Enabled = False
        FocusControl = FValMAJ
      end
      object F2: TLabel
        Left = 64
        Top = 111
        Width = 18
        Height = 23
        Caption = '['
        Enabled = False
        Font.Charset = SYMBOL_CHARSET
        Font.Color = clMaroon
        Font.Height = -19
        Font.Name = 'Wingdings 3'
        Font.Style = []
        ParentFont = False
      end
      object FVol: TCheckBox
        Left = 41
        Top = 30
        Width = 169
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Traiter les gros &volumes'
        TabOrder = 0
      end
      object FMAJ: TCheckBox
        Left = 41
        Top = 57
        Width = 169
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Marquer les '#233'critures trait'#233'es'
        TabOrder = 1
        OnClick = FMAJClick
      end
      object FChampMAJ: THValComboBox
        Left = 198
        Top = 86
        Width = 117
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        Enabled = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 2
        OnChange = FExerciceChange
        Items.Strings = (
          'E_REFINTERNE'
          'E_REFEXTERNE'
          'E_REFLIBRE'
          'E_AFFAIRE'
          'E_LIBELLE'
          'E_LIBRETEXTE0'
          'E_LIBRETEXTE1'
          'E_LIBRETEXTE2'
          'E_LIBRETEXTE3'
          'E_LIBRETEXTE4'
          'E_LIBRETEXTE5'
          'E_LIBRETEXTE6'
          'E_LIBRETEXTE7'
          'E_LIBRETEXTE8'
          'E_LIBRETEXTE9'
          'E_LIBREBOOL0'
          'E_LIBREBOOL1'
          'E_TABLE0'
          'E_TABLE1'
          'E_TABLE2'
          'E_TABLE3')
        TagDispatch = 0
        Values.Strings = (
          'E_REFINTERNE'
          'E_REFEXTERNE'
          'E_REFLIBRE'
          'E_AFFAIRE'
          'E_LIBELLE'
          'E_LIBRETEXTE0'
          'E_LIBRETEXTE1'
          'E_LIBRETEXTE2'
          'E_LIBRETEXTE3'
          'E_LIBRETEXTE4'
          'E_LIBRETEXTE5'
          'E_LIBRETEXTE6'
          'E_LIBRETEXTE7'
          'E_LIBRETEXTE8'
          'E_LIBRETEXTE9'
          'E_LIBREBOOL0'
          'E_LIBREBOOL1'
          'E_TABLE0'
          'E_TABLE1'
          'E_TABLE2'
          'E_TABLE3')
      end
      object FValMAJ: TEdit
        Left = 198
        Top = 112
        Width = 117
        Height = 21
        Enabled = False
        TabOrder = 3
      end
      object FGroupLibres: TGroupBox
        Left = 1
        Top = 168
        Width = 392
        Height = 67
        Caption = 'Tables libres '
        TabOrder = 4
        object TFLibre1: THLabel
          Left = 8
          Top = 32
          Width = 31
          Height = 13
          AutoSize = False
          Caption = '&De'
          FocusControl = FLibre1
        end
        object TFLibre2: THLabel
          Left = 214
          Top = 32
          Width = 6
          Height = 13
          Caption = #224
          FocusControl = FLibre2
        end
        object FLibre1: TEdit
          Left = 50
          Top = 28
          Width = 154
          Height = 21
          AutoSize = False
          CharCase = ecUpperCase
          MaxLength = 17
          ReadOnly = True
          TabOrder = 0
          OnDblClick = FLibre1DblClick
        end
        object FLibre2: TEdit
          Left = 230
          Top = 28
          Width = 154
          Height = 21
          AutoSize = False
          CharCase = ecUpperCase
          MaxLength = 17
          ReadOnly = True
          TabOrder = 1
          OnDblClick = FLibre1DblClick
        end
      end
      object BOptim: TCheckBox
        Left = 41
        Top = 145
        Width = 169
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Optimisation lettrage'
        TabOrder = 5
        Visible = False
        OnClick = FMAJClick
      end
      object FDecoupeGen: TEdit
        Left = 264
        Top = 16
        Width = 121
        Height = 21
        TabOrder = 6
        Visible = False
      end
      object FDecoupeAux: TEdit
        Left = 264
        Top = 48
        Width = 121
        Height = 21
        TabOrder = 7
        Visible = False
      end
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 279
    Width = 401
    Height = 76
    AllowDrag = False
    Position = dpBottom
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 34
      ClientWidth = 401
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 34
      ClientAreaWidth = 401
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BFiltre: TToolbarButton97
        Left = 6
        Top = 7
        Width = 61
        Height = 21
        Hint = 'Menu filtre'
        DropdownArrow = True
        DropdownMenu = POPF
        Caption = '&Filtres'
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object FFiltres: THValComboBox
        Left = 72
        Top = 7
        Width = 295
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = FFiltresChange
        TagDispatch = 0
      end
    end
    object HPB: TToolWindow97
      Left = 0
      Top = 38
      ClientHeight = 34
      ClientWidth = 401
      Caption = 'Actions'
      ClientAreaHeight = 34
      ClientAreaWidth = 401
      DockPos = 0
      DockRow = 1
      FullSize = True
      TabOrder = 1
      object BValider: TToolbarButton97
        Left = 304
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Valider'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
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
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BValiderCEGID: TToolbarButton97
        Tag = 1
        Left = 304
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Arr'#234'ter l'#39'exportation'
        DisplayMode = dmGlyphOnly
        Caption = 'Arr'#234'ter l'#39'exportation'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BValiderCEGIDClick
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 336
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 368
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object TTravail: TLabel
        Left = 16
        Top = 11
        Width = 48
        Height = 13
        Caption = 'TTravail'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
        Visible = False
      end
      object BDelet: TToolbarButton97
        Tag = 1
        Left = 264
        Top = 4
        Width = 28
        Height = 27
        Hint = 'D'#233'lettrage'
        DisplayMode = dmGlyphOnly
        Caption = 'D'#233'lettrage'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BDeletClick
        GlobalIndexImage = 'Z1902_S16G1'
        IsControl = True
      end
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 193
    Top = 269
  end
  object POPF: TPopupMenu
    OnPopup = POPFPopup
    Left = 237
    Top = 269
    object BCreerFiltre: TMenuItem
      Caption = '&Cr'#233'er un filtre'
    end
    object BSaveFiltre: TMenuItem
      Caption = '&Enregistrer le filtre'
    end
    object BDelFiltre: TMenuItem
      Caption = '&Supprimer le filtre'
    end
    object BRenFiltre: TMenuItem
      Caption = '&Renommer le filtre'
    end
    object BNouvRech: TMenuItem
      Caption = '&Nouvelle recherche'
    end
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Lettrage sur code de regroupement;Confirmez-vous le traitement' +
        ' ?;Q;YNC;Y;C;'
      
        '1;Lettrage sur code de regroupement;Le traitement est termin'#233';W;' +
        'O;O;O;'
      '2;D'#233'-lettrage ;Confirmez-vous le traitement ?;Q;YNC;Y;C;')
    Left = 313
    Top = 269
  end
  object Sauve: TSaveDialog
    Filter = 
      'Fichiers textes (*.txt).|*.txt|Fichiers ASCII (*.asc).|*.asc|Fic' +
      'hiers SAARI (*.pn*).|*.pn*|Fichiers EDIFICAS (*.edf).|*.edf|Tous' +
      ' fichiers (*.*).|*.*'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofShowHelp, ofExtensionDifferent, ofPathMustExist, ofNoLongNames]
    Title = 'Exportation des mouvements'
    Left = 269
    Top = 269
  end
  object SauveIni: TSaveDialog
    Filter = 'Fichiers INI (*.ini).|*.ini'
    FilterIndex = 0
    Options = [ofHideReadOnly, ofShowHelp, ofExtensionDifferent, ofPathMustExist, ofNoLongNames]
    Title = 'Exportation des mouvements'
    Left = 344
    Top = 269
  end
  object Timer1: TTimer
    Enabled = False
    OnTimer = Timer1Timer
    Left = 53
    Top = 300
  end
end
