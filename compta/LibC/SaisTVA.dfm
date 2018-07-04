object FSaisTVA: TFSaisTVA
  Left = 267
  Top = 272
  HelpContext = 7244350
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Contr'#244'le de TVA '#224' la ligne'
  ClientHeight = 274
  ClientWidth = 403
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HLabel15: THLabel
    Left = 79
    Top = 3
    Width = 49
    Height = 13
    Caption = 'Comptes'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentColor = False
    ParentFont = False
  end
  object HLabel14: THLabel
    Left = 203
    Top = 3
    Width = 32
    Height = 13
    Caption = 'Ligne'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentColor = False
    ParentFont = False
  end
  object HLabel11: THLabel
    Left = 291
    Top = 3
    Width = 70
    Height = 13
    Caption = 'Cumul pi'#232'ce'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentColor = False
    ParentFont = False
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 16
    Width = 385
    Height = 77
    Caption = 'Saisi'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentColor = False
    ParentFont = False
    TabOrder = 0
    object HTPF1: THLabel
      Left = 8
      Top = 55
      Width = 24
      Height = 13
      Caption = 'TPF'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
    end
    object HLabel12: THLabel
      Left = 8
      Top = 36
      Width = 25
      Height = 13
      Caption = 'TVA'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
    end
    object HLabel1: THLabel
      Left = 8
      Top = 19
      Width = 26
      Height = 13
      Caption = 'H.T.'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
    end
    object SGP: THNumEdit
      Tag = 1
      Left = 267
      Top = 16
      Width = 108
      Height = 19
      Color = clWhite
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      NumericType = ntDC
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 0
      UseRounding = True
      Validate = False
    end
    object SAP: THNumEdit
      Tag = 1
      Left = 267
      Top = 34
      Width = 108
      Height = 19
      Color = clWhite
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      NumericType = ntDC
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 1
      UseRounding = True
      Validate = False
    end
    object SFP: THNumEdit
      Tag = 1
      Left = 267
      Top = 52
      Width = 108
      Height = 19
      Color = clWhite
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      NumericType = ntDC
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 2
      UseRounding = True
      Validate = False
    end
    object SFL: THNumEdit
      Tag = 1
      Left = 160
      Top = 52
      Width = 108
      Height = 19
      Color = clWhite
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      NumericType = ntDC
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 3
      UseRounding = True
      Validate = False
    end
    object SAL: THNumEdit
      Tag = 1
      Left = 160
      Top = 34
      Width = 108
      Height = 19
      Color = clWhite
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      NumericType = ntDC
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 4
      UseRounding = True
      Validate = False
    end
    object SGL: THNumEdit
      Tag = 1
      Left = 160
      Top = 16
      Width = 108
      Height = 19
      Color = clWhite
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      NumericType = ntDC
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 5
      UseRounding = True
      Validate = False
    end
    object SGC: TEdit
      Left = 40
      Top = 16
      Width = 121
      Height = 19
      Color = clWhite
      Ctl3D = False
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 6
    end
    object SAC: TEdit
      Left = 40
      Top = 34
      Width = 121
      Height = 19
      Color = clWhite
      Ctl3D = False
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 7
    end
    object SFC: TEdit
      Left = 40
      Top = 52
      Width = 121
      Height = 19
      Color = clWhite
      Ctl3D = False
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 8
    end
  end
  object GroupBox2: TGroupBox
    Left = 8
    Top = 100
    Width = 385
    Height = 65
    Caption = 'Calcul'#233
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    object HTPF2: THLabel
      Left = 12
      Top = 41
      Width = 24
      Height = 13
      Caption = 'TPF'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object HLabel7: THLabel
      Left = 12
      Top = 22
      Width = 25
      Height = 13
      Caption = 'TVA'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object CAP: THNumEdit
      Tag = 1
      Left = 267
      Top = 19
      Width = 108
      Height = 19
      Color = clWhite
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 0
      UseRounding = True
      Validate = False
    end
    object CFP: THNumEdit
      Tag = 1
      Left = 267
      Top = 37
      Width = 108
      Height = 19
      Color = clWhite
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 1
      UseRounding = True
      Validate = False
    end
    object CFL: THNumEdit
      Tag = 1
      Left = 160
      Top = 37
      Width = 108
      Height = 19
      Color = clWhite
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 2
      UseRounding = True
      Validate = False
    end
    object CAL: THNumEdit
      Tag = 1
      Left = 160
      Top = 19
      Width = 108
      Height = 19
      Color = clWhite
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 3
      UseRounding = True
      Validate = False
    end
    object CAC: TEdit
      Left = 40
      Top = 19
      Width = 121
      Height = 19
      Color = clWhite
      Ctl3D = False
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 4
    end
    object CFC: TEdit
      Left = 40
      Top = 37
      Width = 121
      Height = 19
      Color = clWhite
      Ctl3D = False
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 5
    end
  end
  object GroupBox3: TGroupBox
    Left = 8
    Top = 172
    Width = 385
    Height = 65
    Caption = 'Ecarts'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentColor = False
    ParentFont = False
    TabOrder = 2
    object HTPF3: THLabel
      Left = 8
      Top = 41
      Width = 24
      Height = 13
      Caption = 'TPF'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGreen
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object HLabel17: THLabel
      Left = 8
      Top = 22
      Width = 25
      Height = 13
      Caption = 'TVA'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGreen
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object EAP: THNumEdit
      Tag = 3
      Left = 267
      Top = 19
      Width = 108
      Height = 19
      Color = clWhite
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 0
      UseRounding = True
      Validate = False
    end
    object EFP: THNumEdit
      Tag = 3
      Left = 267
      Top = 37
      Width = 108
      Height = 19
      Color = clWhite
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 1
      UseRounding = True
      Validate = False
    end
    object EFL: THNumEdit
      Tag = 3
      Left = 160
      Top = 37
      Width = 108
      Height = 19
      Color = clWhite
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 2
      UseRounding = True
      Validate = False
    end
    object EAL: THNumEdit
      Tag = 3
      Left = 160
      Top = 19
      Width = 108
      Height = 19
      Color = clWhite
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 3
      UseRounding = True
      Validate = False
    end
    object EAC: TEdit
      Left = 40
      Top = 19
      Width = 121
      Height = 19
      Color = clWhite
      Ctl3D = False
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 4
    end
    object EFC: TEdit
      Left = 40
      Top = 37
      Width = 121
      Height = 19
      Color = clWhite
      Ctl3D = False
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 5
    end
  end
  object BValide: THBitBtn
    Left = 330
    Top = 243
    Width = 28
    Height = 27
    Hint = 'OK'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 3
    OnClick = BValideClick
    Margin = 2
    Spacing = -1
    GlobalIndexImage = 'Z0127_S16G1'
    IsControl = True
  end
  object BAide: THBitBtn
    Left = 362
    Top = 243
    Width = 28
    Height = 27
    Hint = 'Aide'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 4
    OnClick = BAideClick
    Margin = 2
    Spacing = -1
    GlobalIndexImage = 'Z1117_S16G1'
    IsControl = True
  end
  object Panel1: TPanel
    Left = 12
    Top = 248
    Width = 309
    Height = 21
    Color = clInfoBk
    Enabled = False
    TabOrder = 5
    object HINFO: TLabel
      Left = 4
      Top = 4
      Width = 41
      Height = 13
      Caption = 'Ligne N'#176
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 72
    Top = 8
  end
end
