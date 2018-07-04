inherited fEchgImp_Parametres: TfEchgImp_Parametres
  Left = 387
  Top = 267
  Width = 364
  Height = 334
  Caption = 'fEchgImp_Parametres'
  OldCreateOrder = True
  ShowHint = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 258
    Width = 347
    Height = 43
    inherited PBouton: TToolWindow97
      ClientHeight = 39
      ClientWidth = 347
      ClientAreaHeight = 39
      ClientAreaWidth = 347
      inherited BValider: TToolbarButton97
        Left = 251
      end
      inherited BFerme: TToolbarButton97
        Left = 283
      end
      inherited HelpBtn: TToolbarButton97
        Left = 315
      end
      inherited BImprimer: TToolbarButton97
        Left = 219
      end
    end
  end
  object GroupBox2: TGroupBox [1]
    Left = 9
    Top = 14
    Width = 336
    Height = 195
    Caption = ' Valeurs par d'#233'faut  (si absentes dans fichier)  '
    TabOrder = 1
    object labDate: THLabel
      Left = 28
      Top = 34
      Width = 23
      Height = 13
      Caption = 'Date'
    end
    object labLibelle: THLabel
      Left = 28
      Top = 59
      Width = 30
      Height = 13
      Caption = 'Libell'#233
    end
    object libJournal: THLabel
      Left = 28
      Top = 106
      Width = 34
      Height = 13
      Caption = 'Journal'
    end
    object libCompte: THLabel
      Left = 28
      Top = 133
      Width = 36
      Height = 13
      Caption = 'Compte'
    end
    object LibLeurre: THLabel
      Left = 36
      Top = 82
      Width = 3
      Height = 13
      Caption = ' '
      Visible = False
    end
    object HLabel1: THLabel
      Left = 28
      Top = 162
      Width = 65
      Height = 13
      Caption = 'Etablissement'
    end
    object HValCB_Libelle: TEdit
      Left = 104
      Top = 55
      Width = 212
      Height = 21
      Hint = 
        'Libell'#233' par d'#233'faut. Est utilis'#233' lorsque absent dans le fichier '#224 +
        ' importer.'
      TabOrder = 0
    end
    object CB_AvecDateImport: TCheckBox
      Left = 104
      Top = 78
      Width = 149
      Height = 17
      Hint = 'Permet d'#39'ajouter la date syst'#232'me au libell'#233' par d'#233'faut.'
      Caption = '+  date syst'#232'me'
      TabOrder = 1
    end
    object Hcme_Journal: THCritMaskEdit
      Left = 104
      Top = 102
      Width = 212
      Height = 21
      Hint = 
        'Journal par d'#233'faut. Est utilis'#233' lorsque absent dans le fichier '#224 +
        ' importer.'
      TabOrder = 2
      TagDispatch = 0
      DataType = 'CPJOURNALSYNCHRO'
      ElipsisButton = True
    end
    object Hcme_Compte: THCritMaskEdit
      Left = 104
      Top = 129
      Width = 212
      Height = 21
      Hint = 
        'Compte par d'#233'faut. Est utilis'#233' lorsque absent dans le fichier '#224' ' +
        'importer.'
      TabOrder = 3
      TagDispatch = 0
      DataType = 'TZGENERAL'
      ElipsisButton = True
      ControlerDate = True
    end
    object Hcme_Etablissement: THCritMaskEdit
      Left = 104
      Top = 158
      Width = 212
      Height = 21
      Hint = 'Etablissement'
      TabOrder = 4
      TagDispatch = 0
      DataType = 'TTETABLISSEMENT'
      ElipsisButton = True
      ControlerDate = True
    end
  end
  object GroupBox3: TGroupBox [2]
    Left = 9
    Top = 215
    Width = 338
    Height = 43
    TabOrder = 2
    object cbRAZPrealable: TCheckBox
      Left = 40
      Top = 16
      Width = 237
      Height = 17
      Hint = 
        'Si coch'#233'e, un dialogue proposera, avant l'#39'int'#233'gration, de choisi' +
        'r les '#233'critures '#224' supprimer __'
      Alignment = taLeftJustify
      Caption = 'Effacement pr'#233'alable'
      Checked = True
      State = cbChecked
      TabOrder = 0
    end
  end
  object cmeDate: THCritMaskEdit [3]
    Left = 115
    Top = 42
    Width = 209
    Height = 21
    Hint = 
      'Date par d'#233'faut. Est utilis'#233'e lorsque absente dans le fichier '#224' ' +
      'importer.'
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    ParentShowHint = False
    ShowHint = True
    TabOrder = 3
    Text = '  /  /    '
    TagDispatch = 0
    OpeType = otDate
    ElipsisButton = True
    ElipsisAutoHide = True
    Libelle = LibLeurre
    ControlerDate = True
  end
  inherited HMTrad: THSystemMenu
    Left = 212
    Top = 424
  end
end
