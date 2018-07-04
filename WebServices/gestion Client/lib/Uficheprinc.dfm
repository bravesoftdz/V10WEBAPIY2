object FPrinc: TFPrinc
  Left = 443
  Top = 262
  Width = 481
  Height = 504
  Caption = 'Connexion'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object LUSER: THLabel
    Left = 56
    Top = 51
    Width = 46
    Height = 13
    Caption = 'Utilisateur'
  end
  object LPASS: THLabel
    Left = 56
    Top = 75
    Width = 64
    Height = 13
    Caption = 'Mot de passe'
  end
  object PBAS: THPanel
    Left = 0
    Top = 425
    Width = 465
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    DesignSize = (
      465
      41)
    object BQUIT: TToolbarButton97
      Left = 426
      Top = 8
      Width = 27
      Height = 27
      Anchors = [akRight, akBottom]
      OnClick = BQUITClick
      GlobalIndexImage = 'Z0021_S24G1'
    end
    object BVALIDE: TToolbarButton97
      Left = 394
      Top = 8
      Width = 27
      Height = 27
      Anchors = [akRight, akBottom]
      OnClick = BVALIDEClick
      GlobalIndexImage = 'Z0127_S24G1'
    end
  end
  object USER: TEdit97
    Left = 151
    Top = 48
    Width = 229
    Height = 19
    TabOrder = 1
    Text = 'CEGID'
  end
  object PASSW: TEdit97
    Left = 151
    Top = 72
    Width = 229
    Height = 19
    TabOrder = 2
    Text = '123456'
  end
  object MTEXT: THMemo
    Left = 0
    Top = 112
    Width = 465
    Height = 313
    Align = alBottom
    Lines.Strings = (
      'MTEXT')
    TabOrder = 3
  end
  object HmTrad: THSystemMenu
    Separator = True
    Traduction = False
    Left = 424
    Top = 8
  end
end
