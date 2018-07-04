object FMPEscompte: TFMPEscompte
  Left = 219
  Top = 200
  Width = 422
  Height = 242
  Caption = 'Escompte'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 12
    Top = 57
    Width = 15
    Height = 13
    Caption = 'HT'
  end
  object Label2: TLabel
    Left = 22
    Top = 85
    Width = 21
    Height = 13
    Caption = 'TVA'
  end
  object Label3: TLabel
    Left = 5
    Top = 113
    Width = 48
    Height = 13
    Caption = 'Taux TVA'
  end
  object Label4: TLabel
    Left = 6
    Top = 141
    Width = 48
    Height = 13
    Caption = 'Taux ESC'
  end
  object ESCOMPTABLE: TCheckBox
    Left = 56
    Top = 24
    Width = 149
    Height = 17
    Caption = 'ESCOMPTABLE'
    TabOrder = 0
  end
  object COMPTEHT: THCritMaskEdit
    Left = 60
    Top = 52
    Width = 121
    Height = 21
    TabOrder = 1
    Text = 'COMPTEHT'
    TagDispatch = 0
    DataType = 'TZGENERAL'
  end
  object COMPTETVA: THCritMaskEdit
    Left = 60
    Top = 81
    Width = 121
    Height = 21
    TabOrder = 2
    Text = 'COMPTEHT'
    TagDispatch = 0
    DataType = 'TZGENERAL'
  end
  object TAUXTVA: THCritMaskEdit
    Left = 60
    Top = 108
    Width = 121
    Height = 21
    TabOrder = 3
    Text = 'TAUXTVA'
    TagDispatch = 0
    OpeType = otReel
  end
  object E_NUMEROPIECE: THCritMaskEdit
    Left = 240
    Top = 84
    Width = 121
    Height = 21
    TabOrder = 4
    Text = 'E_NUMEROPIECE'
    TagDispatch = 0
    OpeType = otReel
  end
  object Button1: TButton
    Left = 244
    Top = 144
    Width = 75
    Height = 25
    Caption = 'OK'
    TabOrder = 5
    OnClick = Button1Click
  end
  object TAUXESC: THCritMaskEdit
    Left = 60
    Top = 136
    Width = 121
    Height = 21
    TabOrder = 6
    Text = 'TAUXESC'
    TagDispatch = 0
    OpeType = otReel
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 244
    Top = 20
  end
end
