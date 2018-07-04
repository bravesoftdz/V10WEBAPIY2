inherited FAssistGenereCAB: TFAssistGenereCAB
  Left = 307
  Top = 197
  Caption = 'Assistant de g'#233'n'#233'ration des codes '#224' barres'
  ClientWidth = 529
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Caption = 'Etape 1/1'
  end
  inherited lAide: THLabel
    Top = 264
  end
  object HLabel1: THLabel [2]
    Left = 200
    Top = 220
    Width = 40
    Height = 13
    Caption = 'HLabel1'
  end
  inherited bFin: TToolbarButton97
    Caption = '&Ex'#233'cuter'
  end
  inherited P: THPageControl2
    Height = 236
  end
  inherited cControls: THListBox
    Top = 252
  end
  object GBcab: TGroupBox [13]
    Left = 208
    Top = 96
    Width = 293
    Height = 97
    Caption = ' Type de g'#233'n'#233'ration '
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clNavy
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentColor = False
    ParentFont = False
    TabOrder = 5
    object rbMajBlanc: TRadioButton
      Left = 16
      Top = 22
      Width = 257
      Height = 17
      Caption = 'g'#233'n'#233'ration des codes '#224' barres non renseign'#233's'
      Checked = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      TabStop = True
    end
    object rbMajTous: TRadioButton
      Left = 16
      Top = 44
      Width = 257
      Height = 17
      Caption = 'g'#233'n'#233'ration de tous les codes '#224' barres'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
    end
    object rbEfface: TRadioButton
      Left = 16
      Top = 66
      Width = 257
      Height = 17
      Caption = 'suppression des codes '#224' barres'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
    end
  end
  object Panel1: TPanel [14]
    Left = 188
    Top = 24
    Width = 329
    Height = 45
    Caption = 'G'#233'n'#233'ration des codes '#224' barres'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 6
  end
end
