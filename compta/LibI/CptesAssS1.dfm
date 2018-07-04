inherited FComptesAssocies: TFComptesAssocies
  Left = 376
  Top = 185
  HelpContext = 2210000
  BorderStyle = bsDialog
  Caption = 'Comptes associ'#233's :'
  ClientHeight = 368
  ClientWidth = 572
  PixelsPerInch = 96
  TextHeight = 13
  inherited FListe: THDBGrid
    Left = 344
    Width = 228
    Height = 330
    Columns = <
      item
        Expanded = False
        Title.Alignment = taCenter
        Title.Caption = 'Code'
        Width = 70
        Visible = True
      end
      item
        Expanded = False
        Title.Alignment = taCenter
        Title.Caption = 'Libell'#233
        Width = 125
        Visible = True
      end>
  end
  inherited Pappli: THPanel
    Width = 344
    Height = 330
    object HPanel1: THPanel
      Left = 0
      Top = 0
      Width = 344
      Height = 34
      Align = alTop
      FullRepaint = False
      TabOrder = 0
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object HLabel1: THLabel
        Left = 18
        Top = 11
        Width = 114
        Height = 13
        Caption = 'Compte d'#39'&immobilisation '
        FocusControl = IC_COMPTEIMMO
      end
      object IC_COMPTEIMMO: THDBEdit
        Left = 192
        Top = 6
        Width = 121
        Height = 21
        DataField = 'IC_COMPTEIMMO'
        DataSource = STa
        TabOrder = 0
        OnEnter = OnCompteEnter
        OnExit = OnCompteExit
        ElipsisButton = True
        ElipsisAutoHide = True
        OnElipsisClick = OnEllipsisCompteImmo
        Obligatory = False
      end
    end
    object HPanel2: THPanel
      Left = 0
      Top = 34
      Width = 344
      Height = 80
      Align = alTop
      FullRepaint = False
      TabOrder = 1
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object HLabel2: THLabel
        Left = 18
        Top = 10
        Width = 117
        Height = 13
        Caption = 'Compte d'#39'&amortissement '
        FocusControl = IC_COMPTEAMORT
      end
      object HLabel4: THLabel
        Left = 18
        Top = 34
        Width = 95
        Height = 13
        Caption = 'Compte de &dotation '
        FocusControl = IC_COMPTEDOTATION
      end
      object HLabel11: THLabel
        Left = 18
        Top = 58
        Width = 149
        Height = 13
        Caption = 'Compte de &reprise d'#39'exploitation'
        FocusControl = IC_REPEXPLOIT
      end
      object IC_COMPTEAMORT: THDBEdit
        Left = 191
        Top = 6
        Width = 121
        Height = 21
        DataField = 'IC_COMPTEAMORT'
        DataSource = STa
        TabOrder = 0
        OnEnter = OnCompteEnter
        OnExit = OnCompteExit
        ElipsisButton = True
        ElipsisAutoHide = True
        OnElipsisClick = OnEllipsisCompteImmo
        Obligatory = False
      end
      object IC_COMPTEDOTATION: THDBEdit
        Left = 191
        Top = 30
        Width = 121
        Height = 21
        DataField = 'IC_COMPTEDOTATION'
        DataSource = STa
        TabOrder = 1
        OnEnter = OnCompteEnter
        OnExit = OnCompteExit
        ElipsisButton = True
        ElipsisAutoHide = True
        OnElipsisClick = OnEllipsisCompteImmo
        Obligatory = False
      end
      object IC_REPEXPLOIT: THDBEdit
        Left = 191
        Top = 54
        Width = 121
        Height = 21
        DataField = 'IC_REPEXPLOIT'
        DataSource = STa
        TabOrder = 2
        OnEnter = OnCompteEnter
        OnExit = OnCompteExit
        ElipsisButton = True
        ElipsisAutoHide = True
        OnElipsisClick = OnEllipsisCompteImmo
        Obligatory = False
      end
    end
    object HPanel3: THPanel
      Left = 0
      Top = 114
      Width = 344
      Height = 55
      Align = alTop
      FullRepaint = False
      TabOrder = 2
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object HLabel8: THLabel
        Left = 18
        Top = 9
        Width = 163
        Height = 13
        Caption = 'Compte de dotation &exceptionnelle'
        FocusControl = IC_DOTATIONEXC
      end
      object HLabel12: THLabel
        Left = 18
        Top = 33
        Width = 156
        Height = 13
        Caption = 'Compte de reprise e&xceptionnelle'
        FocusControl = IC_REPEXCEP
      end
      object IC_DOTATIONEXC: THDBEdit
        Left = 192
        Top = 6
        Width = 121
        Height = 21
        DataField = 'IC_DOTATIONEXC'
        DataSource = STa
        TabOrder = 0
        OnEnter = OnCompteEnter
        OnExit = OnCompteExit
        ElipsisButton = True
        ElipsisAutoHide = True
        OnElipsisClick = OnEllipsisCompteImmo
        Obligatory = False
      end
      object IC_REPEXCEP: THDBEdit
        Left = 192
        Top = 29
        Width = 121
        Height = 21
        DataField = 'IC_REPEXCEP'
        DataSource = STa
        TabOrder = 1
        OnEnter = OnCompteEnter
        OnExit = OnCompteExit
        ElipsisButton = True
        ElipsisAutoHide = True
        OnElipsisClick = OnEllipsisCompteImmo
        Obligatory = False
      end
    end
    object HPanel4: THPanel
      Left = 0
      Top = 169
      Width = 344
      Height = 79
      Align = alTop
      FullRepaint = False
      TabOrder = 3
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object HLabel9: THLabel
        Left = 18
        Top = 9
        Width = 162
        Height = 13
        Caption = 'Compte de valeur &nette comptable'
        FocusControl = IC_VACEDEE
      end
      object HLabel10: THLabel
        Left = 18
        Top = 33
        Width = 141
        Height = 13
        Caption = 'Compte d'#39'amortissement &c'#233'd'#233
        FocusControl = IC_AMORTCEDE
      end
      object HLabel3: THLabel
        Left = 18
        Top = 57
        Width = 158
        Height = 13
        Caption = 'Compte de valeur d'#39'&origine c'#233'd'#233'e'
        FocusControl = IC_VOACEDE
      end
      object IC_VACEDEE: THDBEdit
        Left = 191
        Top = 5
        Width = 121
        Height = 21
        DataField = 'IC_VACEDEE'
        DataSource = STa
        TabOrder = 0
        OnEnter = OnCompteEnter
        OnExit = OnCompteExit
        ElipsisButton = True
        ElipsisAutoHide = True
        OnElipsisClick = OnEllipsisCompteImmo
        Obligatory = False
      end
      object IC_AMORTCEDE: THDBEdit
        Left = 192
        Top = 29
        Width = 121
        Height = 21
        DataField = 'IC_AMORTCEDE'
        DataSource = STa
        TabOrder = 1
        OnEnter = OnCompteEnter
        OnExit = OnCompteExit
        ElipsisButton = True
        ElipsisAutoHide = True
        OnElipsisClick = OnEllipsisCompteImmo
        Obligatory = False
      end
      object IC_VOACEDE: THDBEdit
        Left = 192
        Top = 53
        Width = 121
        Height = 21
        DataField = 'IC_VOACEDE'
        DataSource = STa
        TabOrder = 2
        OnEnter = OnCompteEnter
        OnExit = OnCompteExit
        ElipsisButton = True
        ElipsisAutoHide = True
        OnElipsisClick = OnEllipsisCompteImmo
        Obligatory = False
      end
    end
    object HPanel5: THPanel
      Left = 0
      Top = 248
      Width = 344
      Height = 82
      Align = alClient
      FullRepaint = False
      TabOrder = 4
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object HLabel6: THLabel
        Left = 18
        Top = 33
        Width = 152
        Height = 13
        Caption = 'Compte de &provision d'#233'rogatoire'
        FocusControl = IC_PROVISDEROG
      end
      object HLabel7: THLabel
        Left = 18
        Top = 58
        Width = 144
        Height = 13
        Caption = 'Compte de reprise d'#233'roga&toire '
        FocusControl = IC_REPRISEDEROG
      end
      object HLabel5: THLabel
        Left = 18
        Top = 9
        Width = 95
        Height = 13
        Caption = 'Compte d'#233'ro&gatoire '
        FocusControl = IC_COMPTEDEROG
      end
      object IC_COMPTEDEROG: THDBEdit
        Left = 192
        Top = 5
        Width = 121
        Height = 21
        DataField = 'IC_COMPTEDEROG'
        DataSource = STa
        TabOrder = 0
        OnEnter = OnCompteEnter
        OnExit = OnCompteExit
        ElipsisButton = True
        ElipsisAutoHide = True
        OnElipsisClick = OnEllipsisCompteImmo
        Obligatory = False
      end
      object IC_PROVISDEROG: THDBEdit
        Left = 192
        Top = 29
        Width = 121
        Height = 21
        DataField = 'IC_PROVISDEROG'
        DataSource = STa
        TabOrder = 1
        OnEnter = OnCompteEnter
        OnExit = OnCompteExit
        ElipsisButton = True
        ElipsisAutoHide = True
        OnElipsisClick = OnEllipsisCompteImmo
        Obligatory = False
      end
      object IC_REPRISEDEROG: THDBEdit
        Left = 192
        Top = 54
        Width = 121
        Height = 21
        DataField = 'IC_REPRISEDEROG'
        DataSource = STa
        TabOrder = 2
        OnEnter = OnCompteEnter
        OnExit = OnCompteExit
        ElipsisButton = True
        ElipsisAutoHide = True
        OnElipsisClick = OnEllipsisCompteImmo
        Obligatory = False
      end
    end
  end
  inherited DBNav: TDBNavigator
    Hints.Strings = ()
  end
  inherited Dock971: TDock97
    Top = 330
    Width = 572
    inherited PBouton: TToolWindow97
      ClientWidth = 572
      ClientAreaWidth = 572
      inherited BFirst: TToolbarButton97
        Left = 364
      end
      inherited BPrev: TToolbarButton97
        Left = 408
      end
      inherited BNext: TToolbarButton97
        Left = 452
      end
      inherited BLast: TToolbarButton97
        Left = 496
      end
      inherited FAutoSave: THCheckbox
        Left = 157
      end
    end
  end
  inherited Ta: THTable
    TableName = 'IMMOCPTE'
    Left = 356
    Top = 49
    object TaIC_COMPTEIMMO: TStringField
      FieldName = 'IC_COMPTEIMMO'
      Size = 17
    end
    object TaIC_COMPTEAMORT: TStringField
      FieldName = 'IC_COMPTEAMORT'
      Size = 17
    end
    object TaIC_COMPTEDOTATION: TStringField
      FieldName = 'IC_COMPTEDOTATION'
      Size = 17
    end
    object TaIC_COMPTEDEROG: TStringField
      FieldName = 'IC_COMPTEDEROG'
      Size = 17
    end
    object TaIC_REPRISEDEROG: TStringField
      FieldName = 'IC_REPRISEDEROG'
      Size = 17
    end
    object TaIC_DOTATIONEXC: TStringField
      FieldName = 'IC_DOTATIONEXC'
      Size = 17
    end
    object TaIC_VACEDEE: TStringField
      FieldName = 'IC_VACEDEE'
      Size = 17
    end
    object TaIC_AMORTCEDE: TStringField
      FieldName = 'IC_AMORTCEDE'
      Size = 17
    end
    object TaIC_VOACEDE: TStringField
      FieldName = 'IC_VOACEDE'
      Size = 17
    end
    object TaIC_PROVISDEROG: TStringField
      FieldName = 'IC_PROVISDEROG'
      Size = 17
    end
    object TaIC_REPEXPLOIT: TStringField
      FieldName = 'IC_REPEXPLOIT'
      Size = 17
    end
    object TaIC_REPEXCEP: TStringField
      FieldName = 'IC_REPEXCEP'
      Size = 17
    end
  end
  inherited STa: TDataSource
    Left = 389
    Top = 49
  end
  inherited HM: THMsgBox
    Mess.Strings = (
      
        '0;?caption?;Voulez-vous enregistrer les modifications ?;Q;YNC;Y;' +
        'C;'
      
        '1;?caption?;Confirmez-vous la suppression de l'#39'enregistrement ?;' +
        'Q;YNC;N;C;'
      '2;?caption?;Vous devez renseigner un code.;W;O;O;O;'
      '3;?caption?;Vous devez renseigner un libell'#233'.;W;O;O;O;'
      
        '4;?caption?;Le code que vous avez saisi existe d'#233'j'#224'. Vous devez ' +
        'le modifier.;W;O;O;O;'
      'L'#39'enregistrement est inaccessible'
      
        '6;?caption?;Un compte n'#39'a pas '#233't'#233' cr'#233#233'. Cr'#233'er tous les comptes a' +
        'ssoci'#233's saisis avant de valider l'#39'enregistrement.;W;O;O;O;'
      'Compte : '
      
        '8;?Caption?;Voulez-vous mettre '#224' jour les fiches d'#233'j'#224' cr'#233#233'es ?;Q' +
        ';YN;Y;N;'
      'fiche a '#233't'#233' mise '#224' jour.'
      'fiches ont '#233't'#233' mises '#224' jour.'
      '11;?caption?;;W;O;O;O;'
      'Comptes associ'#233's'
      '13;?caption?;Un compte est incorrect.;E;O;O;O;')
    Left = 487
    Top = 49
  end
  inherited HM2: THMsgBox
    Left = 454
    Top = 49
  end
  inherited HMTrad: THSystemMenu
    MaxInsideSize = False
    ResizeDBGrid = False
    Left = 520
    Top = 49
  end
  object TGene: TTable
    DatabaseName = 'SOC'
    TableName = 'GENERAUX'
    Left = 422
    Top = 49
  end
end
