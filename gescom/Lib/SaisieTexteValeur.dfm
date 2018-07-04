object FSaisieTexteValeur: TFSaisieTexteValeur
  Left = 193
  Top = 222
  Width = 594
  Height = 106
  Caption = 'Saisie'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poScreenCenter
  DesignSize = (
    586
    72)
  PixelsPerInch = 96
  TextHeight = 13
  object Libelle: THLabel
    Left = 3
    Top = 11
    Width = 3
    Height = 13
  end
  object bOk: TToolbarButton97
    Left = 549
    Top = 41
    Width = 27
    Height = 26
    Flat = False
    OnClick = bOkClick
    GlobalIndexImage = 'Z0127_S16G1'
  end
  object bAnnuler: TToolbarButton97
    Left = 517
    Top = 41
    Width = 27
    Height = 26
    Flat = False
    OnClick = bAnnulerClick
    GlobalIndexImage = 'Z0021_S16G1'
  end
  object Saisie: TEdit
    Left = 144
    Top = 7
    Width = 433
    Height = 21
    TabOrder = 0
  end
  object GroupBox1: TGroupBox
    Left = -9
    Top = 30
    Width = 596
    Height = 7
    Anchors = [akLeft, akRight, akBottom]
    TabOrder = 1
  end
end
