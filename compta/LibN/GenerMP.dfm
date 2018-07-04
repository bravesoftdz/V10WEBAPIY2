object FGenereMPPlus: TFGenereMPPlus
  Left = 325
  Top = 222
  HelpContext = 999999213
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = 'Param'#232'tres de g'#233'n'#233'ration'
  ClientHeight = 181
  ClientWidth = 239
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  ShowHint = True
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HDateReg: THLabel
    Left = 5
    Top = 12
    Width = 116
    Height = 13
    Caption = '&Date de comptabilisation'
    FocusControl = DATEGENERATION
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
  end
  object HLabel4: THLabel
    Left = 5
    Top = 78
    Width = 78
    Height = 13
    Caption = 'N'#176' de &bordereau'
    FocusControl = NUMENCADECA
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object H_MODEGENERE: TLabel
    Left = 5
    Top = 111
    Width = 88
    Height = 13
    Caption = '&Mode de paiement'
    FocusControl = MODEGENERE
  end
  object DATEGENERATION: THCritMaskEdit
    Left = 142
    Top = 8
    Width = 96
    Height = 21
    Ctl3D = True
    EditMask = '!99/99/0000;1;_'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    MaxLength = 10
    ParentCtl3D = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 0
    Text = '  /  /    '
    TagDispatch = 0
    Operateur = Superieur
    OpeType = otDate
    ElipsisButton = True
    ControlerDate = True
  end
  object CChoixDate: TCheckBox
    Left = 5
    Top = 45
    Width = 145
    Height = 17
    Caption = '&Forcer date d'#39#233'ch'#233'ance'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    OnClick = CChoixDateClick
  end
  object DATEECHEANCE: THCritMaskEdit
    Left = 142
    Top = 43
    Width = 96
    Height = 21
    Ctl3D = True
    Enabled = False
    EditMask = '!99/99/0000;1;_'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    MaxLength = 10
    ParentCtl3D = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 2
    Text = '  /  /    '
    TagDispatch = 0
    Operateur = Superieur
    OpeType = otDate
    ElipsisButton = True
    ControlerDate = True
  end
  object NUMENCADECA: THCritMaskEdit
    Left = 142
    Top = 74
    Width = 96
    Height = 21
    CharCase = ecUpperCase
    Ctl3D = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    MaxLength = 17
    ParentCtl3D = False
    ParentFont = False
    TabOrder = 3
    TagDispatch = 0
    DataType = 'CPNUMENCADECA'
    ElipsisButton = True
  end
  object MODEGENERE: THValComboBox
    Left = 142
    Top = 107
    Width = 96
    Height = 21
    Style = csDropDownList
    Ctl3D = False
    ItemHeight = 13
    ParentCtl3D = False
    TabOrder = 4
    OnChange = MODEGENEREChange
    TagDispatch = 0
    Vide = True
    VideString = 'Inchang'#233
    DataType = 'TTMODEPAIE'
  end
  object CCVisuEcr: TCheckBox
    Left = 5
    Top = 134
    Width = 188
    Height = 17
    Caption = '&Visualiser les '#233'critures g'#233'n'#233'r'#233'es'
    Checked = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    State = cbChecked
    TabOrder = 8
    Visible = False
    OnClick = CChoixDateClick
  end
  object CCCumulEcr: TCheckBox
    Left = 5
    Top = 152
    Width = 184
    Height = 17
    Caption = '&Regrouper les '#233'critures g'#233'n'#233'r'#233'es'
    Checked = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    State = cbChecked
    TabOrder = 9
    Visible = False
    OnClick = CChoixDateClick
  end
  object BValider: THBitBtn
    Left = 140
    Top = 153
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
    TabOrder = 5
    OnClick = BValiderClick
    Glyph.Data = {
      BE060000424DBE06000000000000360400002800000024000000120000000100
      0800000000008802000000000000000000000001000000010000000000000000
      80000080000000808000800000008000800080800000C0C0C000C0DCC000F0C8
      A400000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000F0FBFF00A4A0A000808080000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00030303030303
      0303030303030303030303030303030303030303030303030303030303030303
      03030303030303030303030303030303030303030303FF030303030303030303
      03030303030303040403030303030303030303030303030303F8F8FF03030303
      03030303030303030303040202040303030303030303030303030303F80303F8
      FF030303030303030303030303040202020204030303030303030303030303F8
      03030303F8FF0303030303030303030304020202020202040303030303030303
      0303F8030303030303F8FF030303030303030304020202FA0202020204030303
      0303030303F8FF0303F8FF030303F8FF03030303030303020202FA03FA020202
      040303030303030303F8FF03F803F8FF0303F8FF03030303030303FA02FA0303
      03FA0202020403030303030303F8FFF8030303F8FF0303F8FF03030303030303
      FA0303030303FA0202020403030303030303F80303030303F8FF0303F8FF0303
      0303030303030303030303FA0202020403030303030303030303030303F8FF03
      03F8FF03030303030303030303030303FA020202040303030303030303030303
      0303F8FF0303F8FF03030303030303030303030303FA02020204030303030303
      03030303030303F8FF0303F8FF03030303030303030303030303FA0202020403
      030303030303030303030303F8FF0303F8FF03030303030303030303030303FA
      0202040303030303030303030303030303F8FF03F8FF03030303030303030303
      03030303FA0202030303030303030303030303030303F8FFF803030303030303
      030303030303030303FA0303030303030303030303030303030303F803030303
      0303030303030303030303030303030303030303030303030303030303030303
      0303}
    Margin = 2
    NumGlyphs = 2
    Spacing = -1
    IsControl = True
  end
  object BAnnuler: THBitBtn
    Left = 172
    Top = 153
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
    TabOrder = 6
    OnClick = BAnnulerClick
    Margin = 2
    Spacing = -1
    GlobalIndexImage = 'Z0021_S16G1'
    IsControl = True
  end
  object Baide: THBitBtn
    Left = 204
    Top = 153
    Width = 28
    Height = 27
    Hint = 'Aide'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 7
    OnClick = BaideClick
    GlobalIndexImage = 'Z1117_S16G1'
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Saisie mono-'#233'ch'#233'ance;Vous devez renseigner un mode de paiement' +
        ';W;O;O;O;'
      
        '1;Saisie mono-'#233'ch'#233'ance;Vous devez valider les informations;W;O;O' +
        ';O;'
      
        '2;Saisie mono-'#233'ch'#233'ance;La date d'#39#233'ch'#233'ance doit respecter la plag' +
        'e de saisie autoris'#233'e;W;O;O;O;'
      '3;?caption?;Vous devez renseigner une date valide.;W;O;O;O;'
      
        '4;?caption?;La date de r'#232'glement que vous avez renseign'#233'e est su' +
        'r un exercice non ouvert.;W;O;O;O;'
      
        '5;?caption?;La date de r'#232'glement que vous avez renseign'#233'e est su' +
        'r un exercice non ouvert.;W;O;O;O;'
      
        '6;?caption?;La date de r'#232'glement que vous avez renseign'#233'e est an' +
        't'#233'rieure '#224' la cl'#244'ture provisoire.;W;O;O;O;'
      
        '7;?caption?;La date de r'#232'glement que vous avez renseign'#233'e est an' +
        't'#233'rieure '#224' la cl'#244'ture d'#233'finitive.;W;O;O;O;'
      
        '8;?caption?;La date que vous avez renseign'#233'e n'#39'est pas valide;W;' +
        'O;O;O;'
      
        '9;?caption?;La date que vous avez renseign'#233'e n'#39'est pas dans un e' +
        'xercice ouvert;W;O;O;O;'
      
        '10;?caption?;La date que vous avez renseign'#233'e est ant'#233'rieure '#224' u' +
        'ne cl'#244'ture;W;O;O;O;'
      
        '11;?caption?;La date que vous avez renseign'#233'e est ant'#233'rieure '#224' u' +
        'ne cl'#244'ture;W;O;O;O;'
      
        '12;?caption?;La date que vous avez renseign'#233'e est en dehors des ' +
        'limites autoris'#233'es;W;O;O;O;'
      
        '13;?caption?;La formule saisie est trop longue. Tout ne sera pas' +
        ' retenu,vous devez la recomposer.;W;O;O;O;')
    Left = 100
    Top = 71
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 100
    Top = 103
  end
end
