object FBudcomp: TFBudcomp
  Left = 247
  Top = 154
  Width = 467
  Height = 306
  HelpContext = 15217150
  BorderIcons = [biSystemMenu]
  Caption = 'Saisie des informations compl'#233'mentaires'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 244
    Width = 459
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    BevelOuter = bvNone
    TabOrder = 1
    object BValide: THBitBtn
      Left = 361
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Valider la saisie'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BValideClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      GlobalIndexImage = 'Z0127_S16G1'
      IsControl = True
    end
    object BFerme: THBitBtn
      Left = 393
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BFermeClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0021_S16G1'
      IsControl = True
    end
    object BAide: THBitBtn
      Left = 425
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
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 459
    Height = 244
    ActivePage = TS1
    Align = alClient
    TabOrder = 0
    object TS1: TTabSheet
      Caption = 'Compl'#233'ments'
      object GInt: TGroupBox
        Left = 0
        Top = 0
        Width = 451
        Height = 49
        Align = alTop
        Caption = ' Informations internes '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object TBE_REFINTERNE: THLabel
          Left = 7
          Top = 20
          Width = 50
          Height = 13
          Caption = '&R'#233'f'#233'rence'
          FocusControl = BE_REFINTERNE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBE_LIBELLE: THLabel
          Left = 200
          Top = 20
          Width = 30
          Height = 13
          Caption = '&Libell'#233
          FocusControl = BE_LIBELLE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BE_REFINTERNE: TEdit
          Tag = 1
          Left = 64
          Top = 18
          Width = 117
          Height = 20
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentFont = False
          TabOrder = 0
          OnChange = BE_REFINTERNEChange
        end
        object BE_LIBELLE: TEdit
          Tag = 2
          Left = 244
          Top = 18
          Width = 200
          Height = 20
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentFont = False
          TabOrder = 1
          OnChange = BE_REFINTERNEChange
        end
      end
      object GQte: TGroupBox
        Left = 0
        Top = 49
        Width = 451
        Height = 72
        Align = alTop
        Caption = ' Saisie des quantit'#233's '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TBE_QTE1: THLabel
          Left = 7
          Top = 21
          Width = 49
          Height = 13
          Caption = 'Quantit'#233' &1'
          Color = clBtnFace
          FocusControl = BE_QTE1
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TBE_QTE2: THLabel
          Left = 7
          Top = 48
          Width = 49
          Height = 13
          Caption = 'Quantit'#233' &2'
          Color = clBtnFace
          FocusControl = BE_QTE2
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TBE_QUALIFQTE1: THLabel
          Left = 143
          Top = 21
          Width = 54
          Height = 13
          Caption = '&Qualif qt'#233' 1'
          Color = clBtnFace
          FocusControl = BE_QUALIFQTE1
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TBE_QUALIFQTE2: THLabel
          Left = 143
          Top = 48
          Width = 54
          Height = 13
          Caption = 'Q&ualif qt'#233' 2'
          Color = clBtnFace
          FocusControl = BE_QUALIFQTE2
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TBE_AFFAIRE: THLabel
          Left = 347
          Top = 21
          Width = 57
          Height = 13
          Caption = 'Code &affaire'
          Color = clBtnFace
          FocusControl = BE_AFFAIRE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object BE_QTE1: THNumEdit
          Tag = 7
          Left = 64
          Top = 17
          Width = 67
          Height = 20
          AutoSize = False
          Decimals = 2
          Digits = 12
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Masks.PositiveMask = '#,##0.000'
          Debit = False
          ParentFont = False
          TabOrder = 0
          UseRounding = True
          Validate = False
          OnChange = BE_REFINTERNEChange
        end
        object BE_QTE2: THNumEdit
          Tag = 8
          Left = 64
          Top = 44
          Width = 67
          Height = 20
          AutoSize = False
          Decimals = 2
          Digits = 12
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Masks.PositiveMask = '#,##0.000'
          Debit = False
          ParentFont = False
          TabOrder = 2
          UseRounding = True
          Validate = False
          OnChange = BE_REFINTERNEChange
        end
        object BE_QUALIFQTE1: THValComboBox
          Tag = 7
          Left = 204
          Top = 17
          Width = 104
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 1
          OnChange = BE_REFINTERNEChange
          TagDispatch = 0
          DataType = 'TTQUALUNITMESURE'
        end
        object BE_QUALIFQTE2: THValComboBox
          Tag = 8
          Left = 204
          Top = 44
          Width = 104
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 3
          OnChange = BE_REFINTERNEChange
          TagDispatch = 0
          DataType = 'TTQUALUNITMESURE'
        end
        object BE_AFFAIRE: TEdit
          Tag = 1
          Left = 319
          Top = 46
          Width = 125
          Height = 20
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 17
          ParentFont = False
          TabOrder = 4
          OnChange = BE_REFINTERNEChange
        end
      end
      object GBloc: TGroupBox
        Left = 0
        Top = 121
        Width = 451
        Height = 95
        Align = alClient
        Caption = ' Bloc-Notes '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object BE_BLOCNOTE: THRichEditOLE
          Tag = 9
          Left = 8
          Top = 14
          Width = 437
          Height = 73
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnChange = BE_REFINTERNEChange
          Margins.Top = 0
          Margins.Bottom = 0
          Margins.Left = 0
          Margins.Right = 0
          ContainerName = 'Document'
          ObjectMenuPrefix = '&Object'
          LinesRTF.Strings = (
            
              '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fswiss Ar' +
              'ial;}}'
            '{\colortbl ;\red0\green0\blue0;}'
            
              '{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1\pard\cf1\f0\' +
              'fs16 '
            '\par }')
        end
      end
    end
    object TS2: TTabSheet
      Caption = 'Informations libres'
      object TBE_LIBRETEXTE1: THLabel
        Left = 2
        Top = 20
        Width = 58
        Height = 13
        Caption = 'Texte libre 1'
        FocusControl = BE_LIBRETEXTE1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBE_LIBRETEXTE2: THLabel
        Left = 2
        Top = 44
        Width = 58
        Height = 13
        Caption = 'Texte libre 2'
        FocusControl = BE_LIBRETEXTE2
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBE_LIBRETEXTE3: THLabel
        Left = 2
        Top = 68
        Width = 58
        Height = 13
        Caption = 'Texte libre 3'
        FocusControl = BE_LIBRETEXTE3
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBE_LIBRETEXTE4: THLabel
        Left = 2
        Top = 92
        Width = 58
        Height = 13
        Caption = 'Texte libre 4'
        FocusControl = BE_LIBRETEXTE4
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBE_LIBRETEXTE5: THLabel
        Left = 2
        Top = 116
        Width = 58
        Height = 13
        Caption = 'Texte libre 5'
        FocusControl = BE_LIBRETEXTE5
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBE_TABLE0: THLabel
        Left = 2
        Top = 160
        Width = 58
        Height = 13
        Caption = '&Table libre 1'
        Color = clBtnFace
        FocusControl = BE_TABLE0
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object TBE_TABLE1: THLabel
        Left = 236
        Top = 160
        Width = 58
        Height = 13
        Caption = 'Ta&ble libre 2'
        Color = clBtnFace
        FocusControl = BE_TABLE1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object TBE_TABLE2: THLabel
        Left = 2
        Top = 184
        Width = 58
        Height = 13
        Caption = 'T&able libre 3'
        Color = clBtnFace
        FocusControl = BE_TABLE2
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object TBE_TABLE3: THLabel
        Left = 236
        Top = 184
        Width = 58
        Height = 13
        Caption = 'Tabl&e libre 4'
        Color = clBtnFace
        FocusControl = BE_TABLE3
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object BE_LIBRETEXTE1: TEdit
        Tag = 2
        Left = 94
        Top = 16
        Width = 343
        Height = 20
        AutoSize = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentFont = False
        TabOrder = 0
        OnChange = BE_REFINTERNEChange
      end
      object BE_LIBRETEXTE2: TEdit
        Tag = 2
        Left = 94
        Top = 40
        Width = 343
        Height = 20
        AutoSize = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentFont = False
        TabOrder = 1
        OnChange = BE_REFINTERNEChange
      end
      object BE_LIBRETEXTE3: TEdit
        Tag = 2
        Left = 94
        Top = 64
        Width = 343
        Height = 20
        AutoSize = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentFont = False
        TabOrder = 2
        OnChange = BE_REFINTERNEChange
      end
      object BE_LIBRETEXTE4: TEdit
        Tag = 2
        Left = 94
        Top = 88
        Width = 343
        Height = 20
        AutoSize = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentFont = False
        TabOrder = 3
        OnChange = BE_REFINTERNEChange
      end
      object BE_LIBRETEXTE5: TEdit
        Tag = 2
        Left = 94
        Top = 112
        Width = 343
        Height = 20
        AutoSize = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentFont = False
        TabOrder = 4
        OnChange = BE_REFINTERNEChange
      end
      object BE_TABLE0: THCpteEdit
        Left = 94
        Top = 156
        Width = 129
        Height = 21
        MaxLength = 17
        TabOrder = 5
        OnChange = BE_REFINTERNEChange
        ZoomTable = tzNatBud0
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object BE_TABLE1: THCpteEdit
        Left = 308
        Top = 156
        Width = 129
        Height = 21
        MaxLength = 17
        TabOrder = 6
        OnChange = BE_REFINTERNEChange
        ZoomTable = tzNatBud1
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object BE_TABLE2: THCpteEdit
        Left = 94
        Top = 180
        Width = 129
        Height = 21
        MaxLength = 17
        TabOrder = 7
        OnChange = BE_REFINTERNEChange
        ZoomTable = tzNatBud2
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object BE_TABLE3: THCpteEdit
        Left = 308
        Top = 180
        Width = 129
        Height = 21
        MaxLength = 17
        TabOrder = 8
        OnChange = BE_REFINTERNEChange
        ZoomTable = tzNatBud3
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
    end
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Saisie des informations compl'#233'mentaires'
      'Modification des informations compl'#233'mentaires'
      'Consultation des informations compl'#233'mentaires'
      
        '3;Informations compl'#233'mentaires;Voulez-vous enregistrer les modif' +
        'ications ?;Q;YNC;Y;Y;')
    Left = 60
    Top = 248
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 24
    Top = 248
  end
end
