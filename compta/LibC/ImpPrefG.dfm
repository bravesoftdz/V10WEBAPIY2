object FImporPRef: TFImporPRef
  Left = 363
  Top = 228
  HelpContext = 7109080
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Plan de r'#233'f'#233'rence n'#176' : '
  ClientHeight = 360
  ClientWidth = 323
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
  object Pappli: TPanel
    Left = 0
    Top = 0
    Width = 323
    Height = 325
    Align = alClient
    Enabled = False
    TabOrder = 0
    object Label1: TLabel
      Left = 39
      Top = 8
      Width = 245
      Height = 13
      Caption = 'La validation exportera les caract'#233'ristiques'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object HGBOptionImpression: TGroupBox
      Left = 5
      Top = 201
      Width = 314
      Height = 55
      Caption = ' Options d'#39'impression du grand livre '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      object PR_SOLDEPROGRESSIF: TCheckBox
        Left = 6
        Top = 17
        Width = 109
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Solde progressif'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
      end
      object PR_TOTAUXMENSUELS: TCheckBox
        Left = 188
        Top = 17
        Width = 109
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Totaux mensuels'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
      end
      object PR_SAUTPAGE: TCheckBox
        Left = 6
        Top = 36
        Width = 109
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Saut de page'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
      end
      object PR_CENTRALISABLE: TCheckBox
        Left = 188
        Top = 36
        Width = 109
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Centralisable'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
      end
    end
    object GBDescrip: TGroupBox
      Left = 4
      Top = 27
      Width = 314
      Height = 98
      Caption = ' Descriptif '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      object TPR_COMPTE: THLabel
        Left = 5
        Top = 18
        Width = 36
        Height = 13
        Caption = 'Compte'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object tPR_ABREGE: THLabel
        Left = 164
        Top = 18
        Width = 34
        Height = 13
        Caption = 'Abr'#233'g'#233
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TPR_LIBELLE: THLabel
        Left = 5
        Top = 47
        Width = 30
        Height = 13
        Caption = 'Libell'#233
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TPR_SENS: THLabel
        Left = 5
        Top = 75
        Width = 24
        Height = 13
        Caption = 'Sens'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TPR_NATUREGENE: THLabel
        Left = 164
        Top = 75
        Width = 32
        Height = 13
        Caption = 'Nature'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object PR_COMPTE: TEdit
        Left = 44
        Top = 14
        Width = 106
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
      end
      object PR_ABREGE: TEdit
        Left = 202
        Top = 14
        Width = 106
        Height = 21
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
      end
      object PR_LIBELLE: TEdit
        Left = 44
        Top = 43
        Width = 264
        Height = 21
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
      end
      object PR_SENS: THValComboBox
        Left = 44
        Top = 71
        Width = 106
        Height = 21
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 3
        TagDispatch = 0
        DataType = 'TTSENS'
      end
      object PR_NATUREGENE: THValComboBox
        Left = 202
        Top = 71
        Width = 106
        Height = 21
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 4
        TagDispatch = 0
        DataType = 'TTNATGENE'
      end
    end
    object HGBOptionaxes: TGroupBox
      Left = 5
      Top = 163
      Width = 314
      Height = 38
      Caption = ' Ventilation '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      object PR_VENTILABLE1: TCheckBox
        Left = 8
        Top = 14
        Width = 48
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Axe 1'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
      end
      object PR_VENTILABLE2: TCheckBox
        Left = 68
        Top = 14
        Width = 48
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Axe 2'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
      end
      object PR_VENTILABLE3: TCheckBox
        Left = 129
        Top = 14
        Width = 48
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Axe 3'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
      end
      object PR_VENTILABLE4: TCheckBox
        Left = 189
        Top = 14
        Width = 48
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Axe 4'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
      end
      object PR_VENTILABLE5: TCheckBox
        Left = 249
        Top = 14
        Width = 48
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Axe 5'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 4
      end
    end
    object GBCarac: TGroupBox
      Left = 4
      Top = 125
      Width = 314
      Height = 38
      Caption = ' Caract'#233'ristiques '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 4
      object PR_COLLECTIF: TCheckBox
        Left = 8
        Top = 16
        Width = 65
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Collectif'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
      end
      object PR_POINTABLE: TCheckBox
        Left = 116
        Top = 16
        Width = 65
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Pointable'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
      end
      object PR_LETTRABLE: TCheckBox
        Left = 224
        Top = 16
        Width = 73
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Lettrable'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
      end
    end
    object TG_BLOCNOTE: TGroupBox
      Left = 4
      Top = 256
      Width = 314
      Height = 63
      Caption = ' Bloc-Notes '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      object PR_BLOCNOTE: THRichEditOLE
        Left = 8
        Top = 13
        Width = 298
        Height = 42
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 0
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
            'fs16 PR_BLOCNOTE'
          '\par }')
      end
      object CbMess: TComboBox
        Left = 36
        Top = 36
        Width = 74
        Height = 21
        Style = csDropDownList
        Color = clYellow
        ItemHeight = 13
        TabOrder = 1
        Visible = False
        Items.Strings = (
          'Compte:'
          'du plan de r'#233'f'#233'rence n'#176)
      end
    end
  end
  object PBouton: TPanel
    Left = 0
    Top = 325
    Width = 323
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 1
    object BImprimer: TBitBtn
      Left = 192
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Imprimer'
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      Glyph.Data = {
        4E010000424D4E01000000000000760000002800000012000000120000000100
        040000000000D800000000000000000000001000000010000000000000000000
        BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00DDDDDDDDDDDD
        DDDDDD000000DDD00000000000DDDD000000DD0777777777070DDD000000D000
        000000000070DD000000D0777777FFF77000DD000000D077777799977070DD00
        0000D0000000000000770D000000D0777777777707070D000000DD0000000000
        70700D000000DDD0FFFFFFFF07070D000000DDDD0FCCCCCF0000DD000000DDDD
        0FFFFFFFF0DDDD000000DDDDD0FCCCCCF0DDDD000000DDDDD0FFFFFFFF0DDD00
        0000DDDDDD000000000DDD000000DDDDDDDDDDDDDDDDDD000000DDDDDDDDDDDD
        DDDDDD000000DDDDDDDDDDDDDDDDDD000000}
      Margin = 2
    end
    object BValider: THBitBtn
      Left = 224
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Exporter'
      Default = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ModalResult = 1
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0192_S16G1'
      IsControl = True
    end
    object HelpBtn: THBitBtn
      Left = 288
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
      TabOrder = 3
      OnClick = HelpBtnClick
      Glyph.Data = {
        BE060000424DBE06000000000000360400002800000024000000120000000100
        0800000000008802000000000000000000000001000000010000000000000000
        80000080000000808000800000008000800080800000C0C0C000C0DCC000F0CA
        A600000000000000000000000000000000000000000000000000000000000000
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
        03030606030303030303030303030303030303FFFF0303030303030303030303
        0303030303060404060303030303030303030303030303F8F8FF030303030303
        030303030303030303FE06060403030303030303030303030303F8FF03F8FF03
        0303030303030303030303030303FE060603030303030303030303030303F8FF
        FFF8FF0303030303030303030303030303030303030303030303030303030303
        030303F8F8030303030303030303030303030303030304040603030303030303
        0303030303030303FFFF03030303030303030303030303030306060604030303
        0303030303030303030303F8F8F8FF0303030303030303030303030303FE0606
        0403030303030303030303030303F8FF03F8FF03030303030303030303030303
        03FE06060604030303030303030303030303F8FF03F8FF030303030303030303
        030303030303FE060606040303030303030303030303F8FF0303F8FF03030303
        0303030303030303030303FE060606040303030303030303030303F8FF0303F8
        FF030303030303030303030404030303FE060606040303030303030303FF0303
        F8FF0303F8FF030303030303030306060604030303FE06060403030303030303
        F8F8FF0303F8FF0303F8FF03030303030303FE06060604040406060604030303
        030303F8FF03F8FFFFFFF80303F8FF0303030303030303FE0606060606060606
        06030303030303F8FF0303F8F8F8030303F8FF030303030303030303FEFE0606
        060606060303030303030303F8FFFF030303030303F803030303030303030303
        0303FEFEFEFEFE03030303030303030303F8F8FFFFFFFFFFF803030303030303
        0303030303030303030303030303030303030303030303F8F8F8F8F803030303
        0303}
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BFerme: THBitBtn
      Left = 256
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ModalResult = 2
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      GlobalIndexImage = 'Z1770_S16G1'
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 200
    Top = 276
  end
end
