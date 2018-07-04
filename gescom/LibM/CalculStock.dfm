inherited FCalculStock: TFCalculStock
  Left = 237
  Top = 134
  Width = 413
  Height = 530
  Caption = 'Recalcul des stocks'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  object TGP_DEPOT: THLabel [0]
    Left = 16
    Top = 48
    Width = 29
    Height = 13
    Caption = 'Dépôt'
  end
  object TDateDepart: THLabel [1]
    Left = 259
    Top = 127
    Width = 53
    Height = 13
    Caption = 'à partir du :'
  end
  object TOk: TLabel [2]
    Left = 10
    Top = 422
    Width = 371
    Height = 29
    Caption = 'Traitement terminé avec succès'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clGreen
    Font.Height = -24
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    Visible = False
  end
  object HLabel1: THLabel [3]
    Left = 344
    Top = 112
    Width = 55
    Height = 13
    Caption = 'Initialise à 0'
  end
  inherited Dock971: TDock97
    Top = 468
    Width = 405
    inherited PBouton: TToolWindow97
      ClientWidth = 401
      ClientAreaWidth = 401
      inherited BValider: TToolbarButton97
        Left = 298
      end
      inherited BFerme: TToolbarButton97
        Left = 330
      end
      inherited HelpBtn: TToolbarButton97
        Left = 362
      end
      inherited BImprimer: TToolbarButton97
        Left = 266
      end
    end
  end
  object TChkStockNull: TCheckBox [5]
    Left = 16
    Top = 25
    Width = 169
    Height = 17
    Caption = 'Vérification des fiches stocks'
    Checked = True
    State = cbChecked
    TabOrder = 1
  end
  object TChkDepotLigne: TCheckBox [6]
    Left = 16
    Top = 6
    Width = 225
    Height = 17
    Caption = 'Vérification des dépôts sur les documents'
    TabOrder = 0
  end
  object TChkPhysique: TCheckBox [7]
    Left = 16
    Top = 125
    Width = 169
    Height = 17
    Caption = 'Mise à jour du stock physique'
    Checked = True
    State = cbChecked
    TabOrder = 6
    OnClick = TChkPhysiqueClick
  end
  object TChkVenteFFO: TCheckBox [8]
    Left = 40
    Top = 143
    Width = 129
    Height = 17
    Caption = 'Ventes Front-Office'
    Checked = True
    Enabled = False
    State = cbChecked
    TabOrder = 8
  end
  object TChkEntreesSorties: TCheckBox [9]
    Left = 40
    Top = 167
    Width = 185
    Height = 17
    Caption = 'Entrées / Sorties exceptionnelles'
    Checked = True
    Enabled = False
    State = cbChecked
    TabOrder = 11
  end
  object TChkEcartINV: TCheckBox [10]
    Left = 40
    Top = 191
    Width = 113
    Height = 17
    Caption = 'Ecart d'#39'inventaire'
    Checked = True
    Enabled = False
    State = cbChecked
    TabOrder = 14
  end
  object TChkTransfert: TCheckBox [11]
    Left = 40
    Top = 215
    Width = 73
    Height = 17
    Caption = 'Transfert'
    Checked = True
    Enabled = False
    State = cbChecked
    TabOrder = 17
  end
  object TChkLivreFou: TCheckBox [12]
    Left = 40
    Top = 239
    Width = 129
    Height = 17
    Caption = 'Livraison fournisseur'
    Checked = True
    Enabled = False
    State = cbChecked
    TabOrder = 20
  end
  object TChkLivreClient: TCheckBox [13]
    Left = 40
    Top = 263
    Width = 97
    Height = 17
    Caption = 'Livraison client'
    Checked = True
    Enabled = False
    State = cbChecked
    TabOrder = 23
  end
  object TChkPrepaCli: TCheckBox [14]
    Left = 40
    Top = 287
    Width = 121
    Height = 17
    Caption = 'Préparation livraison'
    Checked = True
    Enabled = False
    State = cbChecked
    TabOrder = 26
    Visible = False
  end
  object TChkReserveFou: TCheckBox [15]
    Left = 40
    Top = 311
    Width = 153
    Height = 17
    Caption = 'Commande fournisseur'
    Checked = True
    Enabled = False
    State = cbChecked
    TabOrder = 29
    Visible = False
  end
  object TchkReserveCli: TCheckBox [16]
    Left = 40
    Top = 335
    Width = 113
    Height = 17
    Caption = 'Commande client'
    Checked = True
    Enabled = False
    State = cbChecked
    TabOrder = 32
    Visible = False
  end
  object TChkClotureStock: TCheckBox [17]
    Left = 16
    Top = 70
    Width = 217
    Height = 17
    Caption = 'A partir de la dernière clôture de stock'
    Checked = True
    State = cbChecked
    TabOrder = 3
    OnClick = TChkClotureStockClick
  end
  object DtVenteFFO: THCritMaskEdit [18]
    Left = 259
    Top = 141
    Width = 100
    Height = 21
    Enabled = False
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    TabOrder = 9
    Text = '  /  /    '
    TagDispatch = 0
    OpeType = otDate
    DefaultDate = od1900
    ElipsisButton = True
    ControlerDate = True
  end
  object DtEntreesSorties: THCritMaskEdit [19]
    Left = 259
    Top = 165
    Width = 100
    Height = 21
    Enabled = False
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    TabOrder = 12
    Text = '  /  /    '
    TagDispatch = 0
    OpeType = otDate
    DefaultDate = od1900
    ElipsisButton = True
    ControlerDate = True
  end
  object DtEcartINV: THCritMaskEdit [20]
    Left = 259
    Top = 189
    Width = 100
    Height = 21
    Enabled = False
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    TabOrder = 15
    Text = '  /  /    '
    TagDispatch = 0
    OpeType = otDate
    DefaultDate = od1900
    ElipsisButton = True
    ControlerDate = True
  end
  object DtTransfert: THCritMaskEdit [21]
    Left = 259
    Top = 213
    Width = 100
    Height = 21
    Enabled = False
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    TabOrder = 18
    Text = '  /  /    '
    TagDispatch = 0
    OpeType = otDate
    DefaultDate = od1900
    ElipsisButton = True
    ControlerDate = True
  end
  object DtLivreFou: THCritMaskEdit [22]
    Left = 259
    Top = 237
    Width = 100
    Height = 21
    Enabled = False
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    TabOrder = 21
    Text = '  /  /    '
    TagDispatch = 0
    OpeType = otDate
    DefaultDate = od1900
    ElipsisButton = True
    ControlerDate = True
  end
  object DtLivreClient: THCritMaskEdit [23]
    Left = 259
    Top = 261
    Width = 100
    Height = 21
    Enabled = False
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    TabOrder = 24
    Text = '  /  /    '
    TagDispatch = 0
    OpeType = otDate
    DefaultDate = od1900
    ElipsisButton = True
    ControlerDate = True
  end
  object DtPrepaCli: THCritMaskEdit [24]
    Left = 259
    Top = 285
    Width = 100
    Height = 21
    Enabled = False
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    TabOrder = 27
    Text = '  /  /    '
    Visible = False
    TagDispatch = 0
    OpeType = otDate
    DefaultDate = od1900
    ElipsisButton = True
    ControlerDate = True
  end
  object DtReserveFou: THCritMaskEdit [25]
    Left = 259
    Top = 309
    Width = 100
    Height = 21
    Enabled = False
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    TabOrder = 30
    Text = '  /  /    '
    Visible = False
    TagDispatch = 0
    OpeType = otDate
    DefaultDate = od1900
    ElipsisButton = True
    ControlerDate = True
  end
  object DtReserveCli: THCritMaskEdit [26]
    Left = 259
    Top = 333
    Width = 100
    Height = 21
    Enabled = False
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    TabOrder = 33
    Text = '  /  /    '
    Visible = False
    TagDispatch = 0
    OpeType = otDate
    DefaultDate = od1900
    ElipsisButton = True
    ControlerDate = True
  end
  object GP_DEPOT: THMultiValComboBox [27]
    Left = 59
    Top = 44
    Width = 156
    Height = 21
    TabOrder = 2
    Plus = 'GDE_SURSITE="X"'
    Abrege = False
    DataType = 'GCDEPOT'
    Complete = False
    OuInclusif = False
  end
  object TChkClotureDate: TCheckBox [28]
    Left = 16
    Top = 380
    Width = 145
    Height = 17
    Caption = 'Clotûre du stock à date '
    TabOrder = 36
  end
  object DtClotureDate: THCritMaskEdit [29]
    Left = 259
    Top = 378
    Width = 100
    Height = 21
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    TabOrder = 37
    Text = '  /  /    '
    TagDispatch = 0
    OpeType = otDate
    DefaultDate = od1900
    ElipsisButton = True
    ControlerDate = True
  end
  object TChkEgal: TCheckBox [30]
    Left = 40
    Top = 90
    Width = 305
    Height = 17
    Caption = 'Prise en compte des mouvements de la journée de clôture'
    Checked = True
    State = cbChecked
    TabOrder = 5
  end
  object TChkEgalDate: TCheckBox [31]
    Left = 40
    Top = 360
    Width = 313
    Height = 17
    Caption = 'Prise en compte des mouvements de la date choisie'
    Checked = True
    Enabled = False
    State = cbChecked
    TabOrder = 35
  end
  object TChkTrace: TCheckBox [32]
    Left = 16
    Top = 403
    Width = 249
    Height = 17
    Caption = 'Trace dans fichier LogRecalculStock.log'
    Checked = True
    State = cbChecked
    TabOrder = 38
  end
  object DtLastDateCloture: THCritMaskEdit [33]
    Left = 259
    Top = 68
    Width = 100
    Height = 21
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    TabOrder = 4
    Text = '  /  /    '
    TagDispatch = 0
    OpeType = otDate
    DefaultDate = od1900
    ElipsisButton = True
    ControlerDate = True
  end
  object TChkVenteFFO_0: TCheckBox [34]
    Left = 368
    Top = 143
    Width = 17
    Height = 17
    Enabled = False
    TabOrder = 10
  end
  object TChkEntreesSorties_0: TCheckBox [35]
    Left = 368
    Top = 167
    Width = 17
    Height = 17
    Enabled = False
    TabOrder = 13
  end
  object TChkEcartINV_0: TCheckBox [36]
    Left = 368
    Top = 191
    Width = 17
    Height = 17
    Enabled = False
    TabOrder = 16
  end
  object TChkTransfert_0: TCheckBox [37]
    Left = 368
    Top = 215
    Width = 17
    Height = 17
    Enabled = False
    TabOrder = 19
  end
  object TChkLivreFou_0: TCheckBox [38]
    Left = 368
    Top = 239
    Width = 17
    Height = 17
    Enabled = False
    TabOrder = 22
  end
  object TChkLivreClient_0: TCheckBox [39]
    Left = 368
    Top = 263
    Width = 17
    Height = 17
    Enabled = False
    TabOrder = 25
  end
  object TChkPrepaCli_0: TCheckBox [40]
    Left = 368
    Top = 287
    Width = 17
    Height = 17
    Enabled = False
    TabOrder = 28
    Visible = False
  end
  object TChkReserveFou_0: TCheckBox [41]
    Left = 368
    Top = 311
    Width = 17
    Height = 17
    Enabled = False
    TabOrder = 31
    Visible = False
  end
  object TchkReserveCli_0: TCheckBox [42]
    Left = 368
    Top = 335
    Width = 17
    Height = 17
    Enabled = False
    TabOrder = 34
    Visible = False
  end
  object TChkPhysique_0: TCheckBox [43]
    Left = 368
    Top = 125
    Width = 17
    Height = 17
    Enabled = False
    TabOrder = 7
  end
  inherited HMTrad: THSystemMenu
    Left = 323
    Top = 7
  end
end
