object ZFEch: TZFEch
  Left = 474
  Top = 268
  BorderStyle = bsDialog
  Caption = 'Saisie mono-'#233'ch'#233'ance'
  ClientHeight = 110
  ClientWidth = 238
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HLabel1: THLabel
    Left = 4
    Top = 12
    Width = 88
    Height = 13
    Caption = '&Mode de paiement'
    FocusControl = FMODE
  end
  object HLabel2: THLabel
    Left = 4
    Top = 40
    Width = 82
    Height = 13
    Caption = '&Date d'#39#233'ch'#233'ance'
    FocusControl = FEch
  end
  object FMODE: THValComboBox
    Left = 96
    Top = 8
    Width = 137
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 0
    TagDispatch = 0
    DataType = 'TTMODEPAIE'
  end
  object FEch: TMaskEdit
    Left = 96
    Top = 36
    Width = 137
    Height = 21
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    TabOrder = 1
    Text = '01/01/1900'
    OnKeyPress = FEchKeyPress
  end
  object BValider: THBitBtn
    Left = 140
    Top = 73
    Width = 28
    Height = 27
    Hint = 'Valider'
    Default = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    OnClick = BValiderClick
    Margin = 2
    NumGlyphs = 2
    Spacing = -1
    GlobalIndexImage = 'Z0127_S16G1'
    IsControl = True
  end
  object BAnnuler: THBitBtn
    Left = 172
    Top = 73
    Width = 28
    Height = 27
    Hint = 'Fermer'
    Cancel = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ModalResult = 2
    ParentFont = False
    TabOrder = 3
    Margin = 2
    Spacing = -1
    GlobalIndexImage = 'Z0021_S16G1'
    IsControl = True
  end
  object Baide: THBitBtn
    Left = 204
    Top = 73
    Width = 28
    Height = 27
    Hint = 'Aide'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 4
    GlobalIndexImage = 'Z1117_S16G1'
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?Caption?;Saisie mono-'#233'ch'#233'ance;Vous devez renseigner un mode d' +
        'e paiement.;E;O;O;O;'
      
        '1;?Caption?;Saisie mono-'#233'ch'#233'nace;Vous devez valider les informat' +
        'ions.;E;O;O;O;'
      '2;?Caption?;Date d'#39#233'ch'#233'ance incorrecte.;E;O;O;O;')
    Left = 8
    Top = 64
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 48
    Top = 64
  end
end
