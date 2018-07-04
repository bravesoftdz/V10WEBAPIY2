object GCFModifZS: TGCFModifZS
  Left = 445
  Top = 193
  HelpContext = 110000141
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Modifications de zones en s'#233'rie'
  ClientHeight = 249
  ClientWidth = 622
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  ShowHint = True
  OnCloseQuery = FormCloseQuery
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 622
    Height = 214
    Align = alClient
    TabOrder = 0
    object BAjouter: TToolbarButton97
      Left = 211
      Top = 88
      Width = 50
      Height = 25
      Hint = 'Modifier selon ...'
      Layout = blGlyphRight
      NumGlyphs = 2
      ParentShowHint = False
      ShowHint = True
      OnClick = BAjouterClick
      GlobalIndexImage = 'Z0056_S16G2'
    end
    object BEnlever: TToolbarButton97
      Left = 211
      Top = 120
      Width = 50
      Height = 25
      Hint = 'Annuler modification'
      NumGlyphs = 2
      ParentShowHint = False
      ShowHint = True
      OnClick = BEnleverClick
      GlobalIndexImage = 'Z0077_S16G2'
    end
    object FZModifiable: TGroupBox
      Left = 4
      Top = 2
      Width = 200
      Height = 206
      Caption = '&Zones Modifiables'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      object FListe: TStringGrid
        Left = 5
        Top = 16
        Width = 190
        Height = 184
        ColCount = 3
        DefaultColWidth = 185
        DefaultRowHeight = 18
        FixedCols = 0
        RowCount = 1
        FixedRows = 0
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Options = [goDrawFocusSelected]
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 0
        OnDblClick = FListeDblClick
      end
    end
    object FModif: TGroupBox
      Left = 265
      Top = 2
      Width = 348
      Height = 206
      Caption = '&Modifications'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      object FBox: TScrollBox
        Left = 5
        Top = 16
        Width = 338
        Height = 184
        VertScrollBar.Range = 500
        AutoScroll = False
        Color = clWindow
        ParentColor = False
        TabOrder = 0
        object Panels_On: TPanel
          Left = 0
          Top = 0
          Width = 317
          Height = 30
          Align = alTop
          Caption = 'Panels_On'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          Visible = False
          OnClick = Panels_OnEnter
          OnEnter = Panels_OnEnter
        end
      end
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 214
    Width = 622
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 1
    object BOK: THBitBtn
      Left = 524
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Modifier fiche par fiche'
      ModalResult = 1
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      Visible = False
      OnClick = BOKClick
      Layout = blGlyphTop
      GlobalIndexImage = 'Z0127_S16G1'
    end
    object BAnnuler: THBitBtn
      Left = 556
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      ModalResult = 2
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BAide: THBitBtn
      Left = 588
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Aide'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BAideClick
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
      NumGlyphs = 2
    end
    object BValider: THBitBtn
      Tag = 1
      Left = 492
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Modifier toutes les fiches'
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
      TabOrder = 3
      OnClick = BValiderClick
      Layout = blGlyphTop
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0184_S16G1'
      IsControl = True
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Modifications en s'#233'rie des comptes g'#233'n'#233'raux'
      'Modifications en s'#233'rie des comptes auxiliaires'
      'Modifications en s'#233'rie des comptes budg'#233'taires'
      'Modifications en s'#233'rie des sections analytiques'
      'Modifications en s'#233'rie des articles'
      'Modifications en s'#233'rie des sections budg'#233'taires'
      'Modifications en s'#233'rie des budgets')
    Left = 362
    Top = 68
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 32
    Top = 12
  end
end
