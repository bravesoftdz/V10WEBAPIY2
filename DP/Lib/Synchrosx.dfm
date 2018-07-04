inherited FSynchro: TFSynchro
  Left = 373
  Top = 226
  Caption = 'Synchronisation Sx'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  object TDATEECR: THLabel [0]
    Left = 26
    Top = 97
    Width = 113
    Height = 13
    Caption = 'Date d'#39'arrêté comptable'
  end
  object TRANSFERTVERS: TRadioGroup [2]
    Left = 26
    Top = 24
    Width = 272
    Height = 33
    Columns = 2
    ItemIndex = 0
    Items.Strings = (
      'Envoyer'
      'Recevoir')
    TabOrder = 1
  end
  object DATEECR: THCritMaskEdit [3]
    Left = 160
    Top = 93
    Width = 96
    Height = 21
    EditMask = '!00/00/0000;1;_'
    MaxLength = 10
    TabOrder = 2
    Text = '  /  /    '
    TagDispatch = 0
    Operateur = Egal
    OpeType = otDate
    ElipsisButton = True
    ControlerDate = True
  end
end
