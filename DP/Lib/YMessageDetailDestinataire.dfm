inherited FMsgDetailDest: TFMsgDetailDest
  Left = 279
  Top = 178
  VertScrollBar.Range = 0
  BorderStyle = bsDialog
  Caption = 'D'#233'tail des destinataires'
  ClientHeight = 269
  ClientWidth = 573
  KeyPreview = False
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 232
    Width = 573
    Height = 37
    inherited PBouton: TToolWindow97
      ClientHeight = 33
      ClientWidth = 573
      ClientAreaHeight = 33
      ClientAreaWidth = 573
      inherited BValider: TToolbarButton97
        Left = 477
        Visible = False
      end
      inherited BFerme: TToolbarButton97
        Left = 541
      end
      inherited HelpBtn: TToolbarButton97
        Left = 509
        Visible = False
      end
      inherited BImprimer: TToolbarButton97
        Left = 445
      end
    end
  end
  object Panel1: TPanel [1]
    Left = 0
    Top = 0
    Width = 573
    Height = 232
    Align = alClient
    TabOrder = 1
    object HLabel3: THLabel
      Left = 11
      Top = 16
      Width = 7
      Height = 13
      Caption = 'A'
    end
    object HLabel2: THLabel
      Left = 8
      Top = 130
      Width = 13
      Height = 13
      Caption = 'Cc'
    end
    object HRDestinataireA: THRichEditOLE
      Left = 32
      Top = 16
      Width = 529
      Height = 97
      TabStop = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      ScrollBars = ssVertical
      TabOrder = 0
      WantReturns = False
      Margins.Top = 0
      Margins.Bottom = 0
      Margins.Left = 0
      Margins.Right = 0
      ContainerName = 'Document'
      ObjectMenuPrefix = '&Object'
      LinesRTF.Strings = (
        '{\rtf1\ansi\deff0{\fonttbl{\f0\fnil Arial;}}'
        
          '{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1\pard\lang103' +
          '6\f0\fs16 HRDestinataireA'
        '\par }')
    end
    object HRDestinataireCc: THRichEditOLE
      Left = 32
      Top = 128
      Width = 529
      Height = 97
      TabStop = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      ScrollBars = ssVertical
      TabOrder = 1
      WantReturns = False
      Margins.Top = 0
      Margins.Bottom = 0
      Margins.Left = 0
      Margins.Right = 0
      ContainerName = 'Document'
      ObjectMenuPrefix = '&Object'
      LinesRTF.Strings = (
        '{\rtf1\ansi\deff0{\fonttbl{\f0\fnil Arial;}}'
        
          '{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1\pard\lang103' +
          '6\f0\fs16 HRDestinataireCc'
        '\par }')
    end
  end
end
