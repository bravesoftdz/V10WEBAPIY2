inherited FSynchro: TFSynchro
  Left = 188
  Top = 148
  Width = 390
  Height = 218
  Caption = 'Synchronisation des tables libres'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  object btn_Synchro: TToolbarButton97 [0]
    Left = 180
    Top = 92
    Width = 21
    Height = 22
    ImageIndex = 1
    Images = iml_Liste
    OnClick = btn_SynchroClick
  end
  inherited Dock971: TDock97
    Top = 156
    Width = 382
    inherited PBouton: TToolWindow97
      ClientWidth = 382
      ClientAreaWidth = 382
      inherited BValider: TToolbarButton97
        Left = 281
      end
      inherited BFerme: TToolbarButton97
        Left = 313
      end
      inherited HelpBtn: TToolbarButton97
        Left = 345
      end
      inherited BImprimer: TToolbarButton97
        Left = 249
      end
    end
  end
  object GroupBox1: TGroupBox [2]
    Left = 4
    Top = 52
    Width = 165
    Height = 73
    Caption = 'Gestion commerciale '
    TabOrder = 1
    object cbo_GCClient: THValComboBox
      Left = 12
      Top = 40
      Width = 137
      Height = 21
      Style = csDropDownList
      DropDownCount = 4
      ItemHeight = 13
      TabOrder = 0
      TagDispatch = 0
      Values.Strings = (
        '1'
        '2'
        '3'
        '4'
        '5'
        '6'
        '7'
        '8'
        '9'
        'A')
    end
    object rbt_Client: TRadioButton
      Left = 16
      Top = 20
      Width = 49
      Height = 17
      Caption = 'Client'
      Checked = True
      TabOrder = 1
      TabStop = True
      OnClick = rbt_ClientClick
    end
    object rbt_Fournisseur: TRadioButton
      Left = 68
      Top = 20
      Width = 77
      Height = 17
      Caption = 'Fournisseur'
      TabOrder = 2
      OnClick = rbt_FournisseurClick
    end
    object cbo_GCFournisseur: THValComboBox
      Left = 12
      Top = 40
      Width = 137
      Height = 21
      ItemHeight = 13
      TabOrder = 3
      Visible = False
      TagDispatch = 0
      Values.Strings = (
        '1'
        '2'
        '3')
    end
  end
  object GroupBox2: TGroupBox [3]
    Left = 212
    Top = 52
    Width = 165
    Height = 73
    Caption = 'Comptabilit'#233
    TabOrder = 2
    object cbo_CPTA: THValComboBox
      Left = 12
      Top = 40
      Width = 137
      Height = 21
      Style = csDropDownList
      DropDownCount = 4
      ItemHeight = 13
      TabOrder = 0
      TagDispatch = 0
      Values.Strings = (
        '0'
        '1'
        '2'
        '3'
        '4'
        '5'
        '6'
        '7'
        '8'
        '9')
    end
  end
  object TobViewer1: TTobViewer [4]
    Left = 192
    Top = 156
    Width = 0
    Height = 0
    RowCount = 0
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Arial'
    Font.Style = []
  end
  inherited HMTrad: THSystemMenu
    Left = 60
    Top = 0
  end
  object iml_Liste: THImageList
    GlobalIndexImages.Strings = (
      'Z0407_S16G1'
      'Z0497_S16G1')
    ImageType = itMask
    Left = 128
    Top = 260
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Vous devez s'#233'lectionner une table de destination.;W;' +
        'O;O;O;'
      '1;?caption?;Vous devez s'#233'lectionner une table source.;W;O;O;O;'
      
        '2;?caption?;Si la table libre en comptabilit'#233' contient des infor' +
        'mations auxquelles des tiers font ref'#233'rence, ce traitement entra' +
        'inera une perte d'#39'informations.#10#13 D'#233'sirez-vous continuer ?;E' +
        ';YN;Y;N;'
      
        '3;?caption?;Si la table libre en gestion commerciale contient de' +
        's informations auxquelles des tiers font ref'#233'rence, ce traitemen' +
        't entrainera une perte d'#39'informations.#10#13 D'#233'sirez-vous contin' +
        'uer ?;E;YN;Y;N;'
      
        '4;?caption?;Souhaitez-vous mettre '#224' jour les pi'#232'ces en gestion c' +
        'ommerciale ?;Q;YN;Y;N;'
      
        '5;?caption?;Impossible de mettre '#224' jour la structure de la table' +
        ' libre.;W;O;O;O;'
      
        '6;?caption?;Impossible de mettre '#224' jour les donn'#233'es de la table ' +
        'libre.;W;O;O;O;'
      
        '7;?caption?;Impossible de mettre '#224' jour tous les tiers en gestio' +
        'n commercial.;W;O;O;O;'
      ''
      '')
    Left = 308
    Top = 12
  end
end
