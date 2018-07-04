object StrutDlg: TStrutDlg
  Left = 343
  Top = 198
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = 'R'#233'sultat'
  ClientHeight = 293
  ClientWidth = 585
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label2: TLabel
    Left = 8
    Top = 80
    Width = 128
    Height = 13
    Caption = 'Nouvelles valeurs trouv'#233'es'
  end
  object Compteur: TLabel
    Left = 8
    Top = 260
    Width = 3
    Height = 13
  end
  object StringGrid2: TStringGrid
    Left = 0
    Top = 0
    Width = 585
    Height = 250
    Align = alClient
    ColCount = 1
    DefaultColWidth = 100
    DefaultRowHeight = 16
    FixedCols = 0
    RowCount = 1
    FixedRows = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goEditing, goThumbTracking]
    TabOrder = 0
    OnEnter = StringGrid1Enter
    OnKeyDown = StringGrid2KeyDown
    OnSelectCell = SelectCell
    OnTopLeftChanged = StringGrid2TopLeftChanged
  end
  object Dock971: TDock97
    Left = 0
    Top = 250
    Width = 585
    Height = 43
    AllowDrag = False
    Position = dpBottom
    object PBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 39
      ClientWidth = 585
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 39
      ClientAreaWidth = 585
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object btnOK: TToolbarButton97
        Left = 514
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        AllowAllUp = True
        Default = True
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ModalResult = 1
        NumGlyphs = 2
        ParentFont = False
        Spacing = -1
        OnClick = btnOKClick
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object btnFermer: TToolbarButton97
        Left = 546
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        AllowAllUp = True
        Cancel = True
        Flat = False
        ModalResult = 2
        OnClick = btnFermerClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object bDefaire: TToolbarButton97
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Annuler les modifications'
        AllowAllUp = True
        DisplayMode = dmGlyphOnly
        Caption = 'Annuler'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Spacing = -1
        Visible = False
        GlobalIndexImage = 'Z0075_S16G1'
        IsControl = True
      end
      object Binsert: TToolbarButton97
        Left = 36
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Nouveau'
        AllowAllUp = True
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Visible = False
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 482
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object HelpBtn: TToolbarButton97
        Left = 150
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        AllowAllUp = True
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Spacing = -1
        Visible = False
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
  end
  object Timer1: TTimer
    Enabled = False
    OnTimer = Timer1Timer
    Left = 176
    Top = 136
  end
end
