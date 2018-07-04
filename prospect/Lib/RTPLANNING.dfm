object RT_Planning: TRT_Planning
  Left = 422
  Top = 249
  Width = 648
  Height = 435
  Caption = 'Planning'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panneau: TPanel
    Left = 0
    Top = 364
    Width = 632
    Height = 32
    Align = alBottom
    Alignment = taRightJustify
    TabOrder = 0
    DesignSize = (
      632
      32)
    object BCalendrier: TToolbarButton97
      Left = 65
      Top = 3
      Width = 30
      Height = 26
      Hint = 'Calendrier'
      Opaque = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BCalendrierClick
      GlobalIndexImage = 'Z0665_S16G1'
    end
    object BQuitter: TToolbarButton97
      Left = 580
      Top = 3
      Width = 28
      Height = 26
      Hint = 'Quitter'
      Alignment = taRightJustify
      Anchors = [akTop, akRight]
      Layout = blGlyphRight
      Opaque = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BQuitterClick
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BPageSuiv: TToolbarButton97
      Left = 34
      Top = 3
      Width = 30
      Height = 26
      Hint = 'Suivant'
      Opaque = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BPageSuivClick
      GlobalIndexImage = 'Z0440_S16G1'
    end
    object BPagePrec: TToolbarButton97
      Left = 3
      Top = 3
      Width = 30
      Height = 26
      Hint = 'Pr'#233'c'#233'dent'
      Opaque = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BPagePrecClick
      GlobalIndexImage = 'Z0384_S16G1'
    end
    object Bimprimer: TToolbarButton97
      Left = 549
      Top = 3
      Width = 30
      Height = 26
      Hint = 'Imprimer'
      Anchors = [akTop, akRight]
      Layout = blGlyphRight
      Opaque = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BimprimerClick
      GlobalIndexImage = 'Z0369_S16G1'
    end
    object TRessource: TLabel
      Left = 364
      Top = 8
      Width = 54
      Height = 13
      Caption = 'Intervenant'
    end
    object BAide: TToolbarButton97
      Left = 610
      Top = 3
      Width = 28
      Height = 26
      Hint = 'Aide'
      Anchors = [akTop, akRight]
      DisplayMode = dmGlyphOnly
      Caption = 'Aide'
      Layout = blGlyphTop
      Opaque = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BAideClick
      GlobalIndexImage = 'Z1117_S16G1'
    end
    object BRecharger: TToolbarButton97
      Left = 518
      Top = 3
      Width = 30
      Height = 26
      Hint = 'Rafra'#238'chir'
      Anchors = [akTop, akRight]
      Opaque = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BRechargerClick
      GlobalIndexImage = 'Z1028_S16G1'
    end
    object BEnvoieOutlook: TToolbarButton97
      Left = 96
      Top = 3
      Width = 30
      Height = 26
      Hint = 'Envoie vers Outlook'
      ParentShowHint = False
      ShowHint = True
      Visible = False
      OnClick = BEnvoieOutlookClick
      GlobalIndexImage = 'Z0223_S16G1'
    end
    object DateEdit: THCritMaskEdit
      Left = 156
      Top = 5
      Width = 77
      Height = 21
      Color = clWhite
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      TabOrder = 0
      TagDispatch = 0
    end
    object RESSOURCE: THCritMaskEdit
      Left = 431
      Top = 5
      Width = 77
      Height = 21
      CharCase = ecUpperCase
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      OnExit = RESSOURCEExit
      TagDispatch = 0
      ElipsisButton = True
      OnElipsisClick = RESSOURCEElipsisClick
    end
  end
  object PPlanning: TPanel
    Left = 0
    Top = 0
    Width = 632
    Height = 364
    Align = alClient
    TabOrder = 1
  end
end
