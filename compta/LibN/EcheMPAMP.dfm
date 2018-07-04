object FEcheMPAMP: TFEcheMPAMP
  Left = 306
  Top = 154
  HelpContext = 7505770
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Modifications des conditions de r'#233'glement'
  ClientHeight = 388
  ClientWidth = 343
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PanelBouton: TPanel
    Left = 0
    Top = 353
    Width = 343
    Height = 35
    Align = alBottom
    Caption = ' '
    TabOrder = 2
    object BValider: THBitBtn
      Left = 240
      Top = 4
      Width = 29
      Height = 27
      Hint = 'Valider la modification'
      Default = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BValiderClick
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
      Margin = 3
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BFerme: THBitBtn
      Left = 272
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ModalResult = 2
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
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
      NumGlyphs = 2
    end
    object BAide: THBitBtn
      Left = 304
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Aide'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BAideClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z1117_S16G1'
      IsControl = True
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 343
    Height = 115
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 8
      Top = 31
      Width = 34
      Height = 13
      Caption = 'Journal'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 162
      Top = 31
      Width = 23
      Height = 13
      Caption = 'Date'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label3: TLabel
      Left = 8
      Top = 51
      Width = 32
      Height = 13
      Caption = 'Nature'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label4: TLabel
      Left = 162
      Top = 51
      Width = 37
      Height = 13
      Caption = 'Num'#233'ro'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object E_JOURNAL: TLabel
      Left = 55
      Top = 31
      Width = 106
      Height = 13
      AutoSize = False
      Caption = 'E_JOURNAL'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object E_NATUREPIECE: TLabel
      Left = 55
      Top = 51
      Width = 106
      Height = 13
      AutoSize = False
      Caption = 'E_NATUREPIECE'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object E_DATECOMPTABLE: TLabel
      Left = 216
      Top = 31
      Width = 109
      Height = 13
      AutoSize = False
      Caption = 'E_DATECOMPTABLE'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object E_NUMEROPIECE: TLabel
      Left = 216
      Top = 51
      Width = 109
      Height = 13
      AutoSize = False
      Caption = 'E_NUMEROPIECE'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label5: TLabel
      Left = 162
      Top = 72
      Width = 49
      Height = 13
      Caption = 'Ech'#233'ance'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object E_DATEECHEANCE: TLabel
      Left = 217
      Top = 72
      Width = 109
      Height = 13
      AutoSize = False
      Caption = 'E_DATEECHEANCE'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label7: TLabel
      Left = 7
      Top = 72
      Width = 39
      Height = 13
      Caption = 'Montant'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object E_MONTANT: TLabel
      Left = 55
      Top = 72
      Width = 106
      Height = 13
      AutoSize = False
      Caption = 'E_MONTANT'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label9: TLabel
      Left = 8
      Top = 92
      Width = 44
      Height = 13
      Caption = 'Paiement'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object E_MODEPAIE: TLabel
      Left = 55
      Top = 93
      Width = 106
      Height = 13
      AutoSize = False
      Caption = ' '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object MP_CODEACCEPT: TLabel
      Left = 162
      Top = 93
      Width = 94
      Height = 13
      Caption = 'MP_CODEACCEPT'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object E_AUXILIAIRE: TLabel
      Left = 8
      Top = 7
      Width = 106
      Height = 13
      AutoSize = False
      Caption = 'E_AUXILIAIRE'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object T_LIBELLE: TLabel
      Left = 119
      Top = 7
      Width = 210
      Height = 13
      AutoSize = False
      Caption = 'T_LIBELLE'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 115
    Width = 343
    Height = 238
    Align = alClient
    TabOrder = 1
    object HLabel4: THLabel
      Left = 7
      Top = 41
      Width = 88
      Height = 13
      Caption = '&Mode de paiement'
      FocusControl = FModepaie2
    end
    object HLabel5: THLabel
      Left = 7
      Top = 96
      Width = 82
      Height = 13
      Caption = '&Date d'#39#233'ch'#233'ance'
      FocusControl = FDateEche2
    end
    object FTitreSup: TLabel
      Left = 10
      Top = 15
      Width = 203
      Height = 13
      Caption = 'Caract'#233'ristiques modifiables de l'#39#233'ch'#233'ance '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TMP_CODEACCEPT2: THLabel
      Left = 7
      Top = 69
      Width = 84
      Height = 13
      Caption = 'Code acceptation'
      FocusControl = MP_CODEACCEPT2
    end
    object ZoomRib: TToolbarButton97
      Left = 319
      Top = 120
      Width = 16
      Height = 21
      Hint = 'Choisir le RIB'
      Caption = '...'
      ParentShowHint = False
      ShowHint = True
      OnClick = ZoomRibClick
    end
    object Label6: TLabel
      Left = 6
      Top = 124
      Width = 54
      Height = 13
      Caption = '&RIB / IBAN'
      FocusControl = E_RIB
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TE_REFEXTERNE: THLabel
      Left = 5
      Top = 151
      Width = 67
      Height = 13
      Caption = 'R'#233'f'#233'rence &tir'#233
      FocusControl = FRefTire
    end
    object TE_REFLIBRE: THLabel
      Left = 5
      Top = 180
      Width = 76
      Height = 13
      Caption = 'R'#233'f'#233'rence &tireur'
      FocusControl = FRefTireur
    end
    object TE_NUMTRAITECHQ: THLabel
      Left = 5
      Top = 209
      Width = 85
      Height = 13
      Caption = '&N'#176' traite / ch'#232'que'
      FocusControl = E_NUMTRAITECHQ
    end
    object FDateEche2: TMaskEdit
      Left = 111
      Top = 92
      Width = 202
      Height = 21
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 1
      Text = '  /  /    '
    end
    object FModepaie2: THValComboBox
      Left = 111
      Top = 37
      Width = 202
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = FModepaie2Change
      TagDispatch = 0
      DataType = 'TTMODEPAIE'
    end
    object FRefTire: TEdit
      Left = 111
      Top = 147
      Width = 202
      Height = 21
      TabOrder = 3
    end
    object E_RIB: TEdit
      Left = 111
      Top = 120
      Width = 202
      Height = 21
      TabOrder = 2
      OnDblClick = E_RIBDblClick
      OnKeyDown = E_RIBKeyDown
    end
    object FRefTireur: TEdit
      Left = 111
      Top = 176
      Width = 202
      Height = 21
      TabOrder = 4
    end
    object E_NUMTRAITECHQ: TEdit
      Left = 111
      Top = 205
      Width = 202
      Height = 21
      TabOrder = 5
    end
    object MP_CODEACCEPT2: THValComboBox
      Left = 111
      Top = 65
      Width = 202
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 6
      TagDispatch = 0
      Vide = True
      VideString = 'Non renseign'#233
      DataType = 'TTLETTRECHANGE'
    end
    object E_RIB2: TEdit
      Left = 159
      Top = 120
      Width = 42
      Height = 21
      TabOrder = 7
      Visible = False
      OnDblClick = E_RIBDblClick
      OnKeyDown = E_RIBKeyDown
    end
    object E_RIB3: TEdit
      Left = 207
      Top = 120
      Width = 78
      Height = 21
      TabOrder = 8
      Visible = False
      OnDblClick = E_RIBDblClick
      OnKeyDown = E_RIBKeyDown
    end
    object E_RIB4: TEdit
      Left = 291
      Top = 120
      Width = 22
      Height = 21
      TabOrder = 9
      Visible = False
      OnDblClick = E_RIBDblClick
      OnKeyDown = E_RIBKeyDown
    end
  end
  object HME: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?Caption?;Vous devez renseigner une date valide.;W;O;O;O;'
      
        '1;?Caption?;La date d'#39#233'ch'#233'ance doit respecter la plage de saisie' +
        ' autoris'#233'e.;W;O;O;O;'
      'Non renseign'#233'.'
      ' ')
    Left = 263
    Top = 97
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 88
    Top = 90
  end
end
