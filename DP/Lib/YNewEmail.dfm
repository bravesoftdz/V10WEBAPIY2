inherited FNewEmail: TFNewEmail
  Left = 665
  Top = 344
  Height = 151
  Caption = 'Liste de distribution'
  OldCreateOrder = True
  ShowHint = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 82
    inherited PBouton: TToolWindow97
      ParentShowHint = False
      ShowHint = True
      inherited BValider: TToolbarButton97
        ModalResult = 0
        ParentShowHint = False
        ShowHint = True
      end
      inherited BFerme: TToolbarButton97
        ParentShowHint = False
        ShowHint = True
      end
      inherited Binsert: TToolbarButton97
        Hint = 'Nouveau email'
        ParentShowHint = False
        ShowHint = True
      end
    end
  end
  object Panel1: TPanel [1]
    Left = 0
    Top = 0
    Width = 292
    Height = 82
    Align = alClient
    TabOrder = 1
    object HLabel1: THLabel
      Left = 16
      Top = 20
      Width = 25
      Height = 13
      Caption = 'Email'
    end
    object HLabel2: THLabel
      Left = 16
      Top = 52
      Width = 22
      Height = 13
      Caption = 'Nom'
    end
    object Email: THCritMaskEdit
      Left = 64
      Top = 16
      Width = 225
      Height = 21
      MaxLength = 255
      TabOrder = 0
      OnKeyPress = EmailKeyPress
      TagDispatch = 0
    end
    object Nom: THCritMaskEdit
      Left = 64
      Top = 48
      Width = 225
      Height = 21
      MaxLength = 255
      TabOrder = 1
      TagDispatch = 0
    end
  end
  inherited HMTrad: THSystemMenu
    Top = 152
  end
end
