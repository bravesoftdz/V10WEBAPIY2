object FMAJBDDBTP: TFMAJBDDBTP
  Left = 359
  Top = 309
  Width = 589
  Height = 458
  BorderIcons = [biSystemMenu]
  Caption = 'mise '#224' jour BDD -> LSE BUSINESS'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object LCBDEST: TLabel
    Left = 24
    Top = 28
    Width = 111
    Height = 13
    Caption = 'Base de r'#233'f'#233'rence BTP'
  end
  object BLANCETRAIT: TToolbarButton97
    Left = 197
    Top = 72
    Width = 177
    Height = 41
    Caption = 'Lancer le traitement'
    OnClick = BLANCETRAITClick
    GlobalIndexImage = 'Z0011_S16G1'
  end
  object CURVUE: THLabel
    Left = 0
    Top = 372
    Width = 573
    Height = 13
    Align = alBottom
    Alignment = taCenter
    Caption = 'CURVUE'
    Visible = False
  end
  object Dock971: TDock97
    Left = 0
    Top = 385
    Width = 573
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object PBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 573
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 31
      ClientAreaWidth = 573
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        573
        31)
      object BValider: TToolbarButton97
        Left = 477
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Default = True
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ModalResult = 1
        ParentFont = False
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0127_S16G1'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 509
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Cancel = True
        Flat = False
        ModalResult = 2
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object HelpBtn: TToolbarButton97
        Left = 541
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Spacing = -1
        OnClick = HelpBtnClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
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
        OnClick = bDefaireClick
        GlobalIndexImage = 'M0080_S16G1'
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
        OnClick = BinsertClick
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BDelete: TToolbarButton97
        Left = 68
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Supprimer'
        AllowAllUp = True
        Flat = False
        Visible = False
        OnClick = BDeleteClick
        GlobalIndexImage = 'Z0005_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 445
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Anchors = [akTop, akRight]
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Visible = False
        GlobalIndexImage = 'Z0369_S16G1'
      end
    end
  end
  object BCSRC: TComboBox
    Left = 148
    Top = 24
    Width = 145
    Height = 21
    ItemHeight = 13
    TabOrder = 1
  end
  object Mresult: TMemo
    Left = 0
    Top = 128
    Width = 573
    Height = 244
    Align = alBottom
    Lines.Strings = (
      'Mresult')
    ScrollBars = ssVertical
    TabOrder = 2
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 508
    Top = 8
  end
end
