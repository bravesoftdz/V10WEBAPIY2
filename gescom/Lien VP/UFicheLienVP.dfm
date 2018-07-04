object FFicheTestVP: TFFicheTestVP
  Left = 466
  Top = 361
  Width = 302
  Height = 245
  Caption = 'Mise '#224' jour de VisualProjet'
  Color = clBtnFace
  DefaultMonitor = dmPrimary
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  DesignSize = (
    286
    206)
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 24
    Top = 29
    Width = 148
    Height = 13
    Caption = 'Mise '#224' jour des horaires soci'#233't'#233
  end
  object Label2: TLabel
    Left = 24
    Top = 54
    Width = 107
    Height = 13
    Caption = 'Mise '#224' jour des m'#233'tiers'
  end
  object Label3: TLabel
    Left = 24
    Top = 80
    Width = 125
    Height = 13
    Caption = 'Mise '#224' jour des ressources'
  end
  object ImgHrs: THImage
    Left = 243
    Top = 27
    Width = 16
    Height = 16
    Anchors = [akTop, akBottom]
    AutoSize = True
    Center = True
    Enabled = False
    Transparent = True
    Visible = False
    GlobalIndexImage = 'M0003_S16G1'
  end
  object ImgMetiers: THImage
    Left = 243
    Top = 52
    Width = 16
    Height = 16
    Anchors = [akTop, akBottom]
    AutoSize = True
    Center = True
    Enabled = False
    Transparent = True
    Visible = False
    GlobalIndexImage = 'M0003_S16G1'
  end
  object ImgRessources: THImage
    Left = 243
    Top = 78
    Width = 16
    Height = 16
    Anchors = [akTop, akBottom]
    AutoSize = True
    Center = True
    Enabled = False
    Transparent = True
    Visible = False
    GlobalIndexImage = 'M0003_S16G1'
  end
  object Label4: TLabel
    Left = 24
    Top = 107
    Width = 109
    Height = 13
    Caption = 'Mise '#224' jour des cong'#233's'
  end
  object ImgConges: THImage
    Left = 243
    Top = 105
    Width = 16
    Height = 16
    Anchors = [akTop, akBottom]
    AutoSize = True
    Center = True
    Enabled = False
    Transparent = True
    Visible = False
    GlobalIndexImage = 'M0003_S16G1'
  end
  object Label5: TLabel
    Left = 24
    Top = 132
    Width = 115
    Height = 13
    Caption = 'Mise '#224' jour Interventions'
  end
  object ImgInterv: THImage
    Left = 243
    Top = 130
    Width = 16
    Height = 16
    Anchors = [akTop, akBottom]
    AutoSize = True
    Center = True
    Enabled = False
    Transparent = True
    Visible = False
    GlobalIndexImage = 'M0003_S16G1'
  end
  object PBAS: THPanel
    Left = 0
    Top = 172
    Width = 286
    Height = 34
    Align = alBottom
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    DesignSize = (
      286
      34)
    object BValider: TToolbarButton97
      Left = 220
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
      ParentFont = False
      Spacing = -1
      OnClick = BValiderClick
      GlobalIndexImage = 'Z0127_S16G1'
      IsControl = True
    end
    object BFerme: TToolbarButton97
      Left = 252
      Top = 2
      Width = 28
      Height = 27
      Hint = 'Fermer'
      AllowAllUp = True
      Anchors = [akTop, akRight]
      Cancel = True
      Flat = False
      ModalResult = 2
      GlobalIndexImage = 'Z0021_S16G1'
    end
  end
end
