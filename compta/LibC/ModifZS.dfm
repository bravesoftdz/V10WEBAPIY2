object FModifZS: TFModifZS
  Left = 336
  Top = 81
  HelpContext = 600
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
    Height = 209
    Align = alClient
    TabOrder = 0
    object BAjouter: TToolbarButton97
      Left = 211
      Top = 82
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
      Top = 114
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
        TabStop = True
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
  object Dock: TDock97
    Left = 0
    Top = 209
    Width = 622
    Height = 40
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 36
      ClientWidth = 622
      Caption = 'Actions'
      ClientAreaHeight = 36
      ClientAreaWidth = 622
      DockPos = 0
      FullSize = True
      ParentShowHint = False
      ShowHint = False
      TabOrder = 0
      object BValider: TToolbarButton97
        Left = 525
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Modification : modifier fiche '#224' fiche'
        Caption = 'Valider'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ModalResult = 1
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object BAnnuler: TToolbarButton97
        Left = 556
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAnnulerClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 588
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        Caption = 'Aide'
        DisplayMode = dmGlyphOnly
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
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BOK: TToolbarButton97
        Left = 494
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Modification en rafale : modifier toutes les fiches'
        Caption = 'Valider'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ModalResult = 1
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
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
      'Modifications en s'#233'rie des budgets'
      'Modifications en s'#233'rie des '#233'critures')
    Left = 362
    Top = 68
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 104
    Top = 60
  end
end
