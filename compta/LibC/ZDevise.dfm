object ZFDevise: TZFDevise
  Left = 465
  Top = 216
  HelpContext = 7242100
  BorderStyle = bsDialog
  Caption = 'Saisie autres monnaies'
  ClientHeight = 300
  ClientWidth = 429
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
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 9
    Top = 9
    Width = 411
    Height = 260
    Align = alClient
    TabOrder = 0
    object HLabel2: THLabel
      Left = 16
      Top = 26
      Width = 23
      Height = 13
      Caption = 'Date'
      FocusControl = FDate
    end
    object HLDEVISE: THLabel
      Left = 16
      Top = 66
      Width = 33
      Height = 13
      Caption = 'De&vise'
      FocusControl = FDevise
    end
    object HLTAUX: THLabel
      Left = 244
      Top = 26
      Width = 24
      Height = 13
      Caption = 'Taux'
      FocusControl = FTaux
    end
    object HLDEBIT: THLabel
      Left = 16
      Top = 110
      Width = 25
      Height = 13
      Caption = '&D'#233'bit'
      FocusControl = DEB
    end
    object HLMONTANTD: THLabel
      Left = 244
      Top = 110
      Width = 39
      Height = 13
      Caption = 'Montant'
      FocusControl = FPivotDebit
    end
    object HLCREDIT: THLabel
      Left = 16
      Top = 152
      Width = 27
      Height = 13
      Caption = '&Cr'#233'dit'
      FocusControl = CRE
    end
    object HLMONTANTC: THLabel
      Left = 244
      Top = 153
      Width = 39
      Height = 13
      Caption = 'Montant'
      FocusControl = FPivotCredit
    end
    object SepBevel: TBevel
      Left = -4
      Top = 183
      Width = 450
      Height = 9
      Shape = bsTopLine
    end
    object HLTDEBIT: THLabel
      Left = 16
      Top = 203
      Width = 25
      Height = 13
      Caption = 'D'#233'bit'
    end
    object HLTCREDIT: THLabel
      Left = 16
      Top = 226
      Width = 27
      Height = 13
      Caption = 'Cr'#233'dit'
    end
    object HLTSOLDE: THLabel
      Left = 244
      Top = 212
      Width = 27
      Height = 13
      Caption = 'Solde'
    end
    object ISigneEuro: TImage
      Left = 387
      Top = 24
      Width = 16
      Height = 16
      AutoSize = True
      Picture.Data = {
        07544269746D6170F6000000424DF60000000000000076000000280000001000
        0000100000000100040000000000800000000000000000000000100000001000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00FFFFFFFFFFFFFFFFFFFFFFF44444444FFFFFF44444444444FFFF4444FFFF
        FFF4FFFF444FFFFFFFFFFFFF44FFFFFFFFFFF44444444444FFFFFF4444444444
        4FFFFFF44FFFFFFFFFFFF444444444444FFFFF444444444444FFFFFF44FFFFFF
        FFF4FFFF444FFFFFFF44FFFFF444FFFFF444FFFFFF4444444444FFFFFFF44444
        4FF4}
      Stretch = True
      Transparent = True
    end
    object FDate: TMaskEdit
      Left = 68
      Top = 22
      Width = 90
      Height = 21
      Color = clBtnFace
      Enabled = False
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 3
      Text = '01/01/1900'
    end
    object FDevise: THValComboBox
      Left = 68
      Top = 62
      Width = 137
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = FDeviseChange
      TagDispatch = 0
      DataType = 'TTDEVISEETAT'
    end
    object FTaux: THNumEdit
      Left = 296
      Top = 22
      Width = 90
      Height = 21
      Color = clBtnFace
      Decimals = 2
      Digits = 12
      Enabled = False
      Masks.PositiveMask = '#,#####0.00000'
      Debit = False
      TabOrder = 4
      UseRounding = True
      Validate = False
    end
    object FPivotDebit: THNumEdit
      Left = 296
      Top = 106
      Width = 90
      Height = 21
      Color = clBtnFace
      Decimals = 2
      Digits = 12
      Enabled = False
      Masks.PositiveMask = '#,##0.00'
      Debit = True
      NumericType = ntDC
      TabOrder = 1
      UseRounding = True
      Validate = False
    end
    object FPivotCredit: THNumEdit
      Left = 296
      Top = 149
      Width = 90
      Height = 21
      Color = clBtnFace
      Decimals = 2
      Digits = 12
      Enabled = False
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      NumericType = ntDC
      TabOrder = 2
      UseRounding = True
      Validate = False
    end
    object FTotalSolde: THNumEdit
      Left = 296
      Top = 207
      Width = 90
      Height = 21
      Color = clBtnFace
      Decimals = 2
      Digits = 12
      Enabled = False
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      NumericType = ntDC
      TabOrder = 5
      UseRounding = True
      Validate = False
    end
    object FTotalDebit: THNumEdit
      Left = 68
      Top = 198
      Width = 90
      Height = 21
      Color = clBtnFace
      Decimals = 2
      Digits = 12
      Enabled = False
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      TabOrder = 6
      UseRounding = True
      Validate = False
    end
    object FTotalCredit: THNumEdit
      Left = 68
      Top = 220
      Width = 90
      Height = 21
      Color = clBtnFace
      Decimals = 2
      Digits = 12
      Enabled = False
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      TabOrder = 7
      UseRounding = True
      Validate = False
    end
    object DEB: TEdit
      Left = 68
      Top = 106
      Width = 90
      Height = 21
      TabOrder = 8
      Text = '0,00'
      OnChange = DEBChange
    end
    object CRE: TEdit
      Left = 68
      Top = 149
      Width = 90
      Height = 21
      TabOrder = 9
      Text = '0,00'
      OnChange = CREChange
    end
  end
  object DockTop: TDock97
    Left = 0
    Top = 0
    Width = 429
    Height = 9
  end
  object DockRight: TDock97
    Left = 420
    Top = 9
    Width = 9
    Height = 260
    Position = dpRight
  end
  object DockBottom: TDock97
    Left = 0
    Top = 269
    Width = 429
    Height = 31
    Position = dpBottom
    object Outils: TToolbar97
      Left = 0
      Top = 0
      Caption = 'Outils'
      DockPos = 0
      TabOrder = 1
      object BChancel: TToolbarButton97
        Left = 0
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Chancellerie'
        DisplayMode = dmGlyphOnly
        Caption = 'Chancel.'
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BChancelClick
        GlobalIndexImage = 'Z0653_S16G1'
      end
      object BSolde: TToolbarButton97
        Tag = 1
        Left = 28
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Calcul du solde'
        DisplayMode = dmGlyphOnly
        Caption = 'Solde'
        NumGlyphs = 2
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BSoldeClick
        GlobalIndexImage = 'Z0051_S16G2'
      end
      object BSaisTaux: TToolbarButton97
        Left = 56
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Saisie du taux volatil'
        DisplayMode = dmGlyphOnly
        Caption = 'Taux'
        NumGlyphs = 2
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BSaisTauxClick
        GlobalIndexImage = 'Z2233_S16G2'
      end
    end
    object Valide97: TToolbar97
      Left = 332
      Top = 0
      Caption = 'Valide97'
      DockPos = 332
      TabOrder = 0
      object BValide: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Valider la saisie'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValideClick
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 28
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Fermer'
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        ModalResult = 7
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 56
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
  end
  object DockLeft: TDock97
    Left = 0
    Top = 9
    Width = 9
    Height = 260
    Position = dpLeft
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?caption?;Vous devez renseigner une devise;W;O;O;O;'
      '1;?caption?;Vous devez renseigner un montant;W;O;O;O;'
      '2;?caption?;Vous devez valider les informations;W;O;O;O;'
      
        '3;?caption?;Vous ne pouvez pas saisir des montants n'#233'gatifs;W;O;' +
        'O;O;'
      
        '4;?caption?;Le montant que vous avez saisi est en dehors de la f' +
        'ourchette autoris'#233'e;W;O;O;O;'
      
        '5;?caption?;ATTENTION. La parit'#233' est incorrecte. Voulez-vous la ' +
        'renseigner ?;Q;YN;Y;N;O;'
      
        '6;?caption?;ATTENTION : Le taux en cours est de 1. Voulez-vous s' +
        'aisir ce taux dans la table de chancellerie;Q;YN;Y;N;O;'
      '7;?caption?;ATTENTION : Le taux volatile est perdu;W;O;O;O;'
      
        '8;?caption?;La parit'#233' fixe est mal renseign'#233'e. Voulez-vous la sa' +
        'isir ?;Q;YN;Y;N;O;'
      
        '9;?caption?;Voulez-vous saisir ce taux dans la table de chancell' +
        'erie ?;Q;YN;Y;N;O;')
    Left = 171
    Top = 124
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 210
    Top = 125
  end
end
