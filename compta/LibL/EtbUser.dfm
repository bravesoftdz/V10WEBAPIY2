object FEtbUser: TFEtbUser
  Left = 564
  Top = 111
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  ClientHeight = 161
  ClientWidth = 192
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -10
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  ShowHint = True
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object LBBanque: TListBox
    Left = 0
    Top = 0
    Width = 192
    Height = 128
    Align = alClient
    ItemHeight = 13
    TabOrder = 0
    OnDblClick = LBBanqueDblClick
  end
  object Panel1: TPanel
    Left = 0
    Top = 128
    Width = 192
    Height = 33
    Align = alBottom
    TabOrder = 1
    object BAnnuler: THBitBtn
      Left = 124
      Top = 3
      Width = 27
      Height = 28
      Hint = 'Fermer'
      Cancel = True
      ModalResult = 2
      TabOrder = 0
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BValider: THBitBtn
      Left = 91
      Top = 3
      Width = 28
      Height = 28
      Hint = 'Valider'
      Default = True
      ModalResult = 1
      TabOrder = 1
      NumGlyphs = 2
      GlobalIndexImage = 'Z0003_S16G2'
    end
    object BAide: THBitBtn
      Left = 156
      Top = 3
      Width = 28
      Height = 28
      Hint = 'Aide'
      TabOrder = 2
      GlobalIndexImage = 'Z1117_S16G1'
    end
  end
  object HMsgBox1: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -13
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Choix de la carte d'#39'appel'
      'Choix du destinataire'
      'Le fichier journal est introuvable'
      'Aucun destinataire'
      'Aucune carte'
      'Veuillez utiliser le Manager de communications Etebac'
      'La fonction W32L341 est introuvable'
      'Le fichier W32L341.DLL est introuvable'
      'Le fichier W32L341DBG.DLL est introuvable')
    Left = 15
    Top = 12
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 76
    Top = 11
  end
end
