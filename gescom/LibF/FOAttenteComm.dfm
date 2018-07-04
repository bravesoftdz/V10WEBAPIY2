inherited FFOAttenteComm: TFFOAttenteComm
  Left = 321
  Top = 212
  Width = 441
  Height = 229
  Cursor = crHourGlass
  Caption = 'Attente de communication'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  object LAttention: TFlashingLabel [0]
    Left = 4
    Top = 84
    Width = 137
    Height = 37
    Alignment = taCenter
    Caption = 'Attention'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -32
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    Visible = False
    Flashing = True
  end
  object LTexte: THLabel [1]
    Left = 4
    Top = 16
    Width = 422
    Height = 58
    Alignment = taCenter
    Caption = 'Vous êtes en attente de communication avec le site central...'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -24
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    WordWrap = True
  end
  object LTexte1: THLabel [2]
    Left = 4
    Top = 128
    Width = 395
    Height = 16
    Alignment = taCenter
    Caption = 
      'Si vous quittez cet écran les communications seront interrompues' +
      '.'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsItalic]
    ParentFont = False
    WordWrap = True
  end
  inherited Dock971: TDock97
    Top = 167
    Width = 433
    inherited PBouton: TToolWindow97
      ClientWidth = 429
      ClientAreaWidth = 429
      inherited BValider: TToolbarButton97
        Visible = False
      end
      inherited BFerme: TToolbarButton97
        Visible = False
      end
    end
  end
  inherited HMTrad: THSystemMenu
    Left = 384
    Top = 160
  end
end
