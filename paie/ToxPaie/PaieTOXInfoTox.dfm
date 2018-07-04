object FicheInfoTox: TFicheInfoTox
  Left = 526
  Top = 289
  Width = 231
  Height = 47
  BorderIcons = [biSystemMenu]
  Caption = 'Fichier :'
  Color = clInfoBk
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object LLibelle: TLabel
    Left = 4
    Top = 0
    Width = 96
    Height = 16
    Caption = 'Taille en octets :'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = []
    ParentFont = False
  end
  object LTaille: TLabel
    Left = 108
    Top = 0
    Width = 82
    Height = 16
    Caption = '0,000,000,000'
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlue
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
  end
end
