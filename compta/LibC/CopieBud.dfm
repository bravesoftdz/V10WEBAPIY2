inherited FCopieBud: TFCopieBud
  Left = 135
  Top = 296
  HelpContext = 15270000
  VertScrollBar.Range = 0
  BorderStyle = bsSingle
  Caption = 'Recopie de budget'
  ClientHeight = 311
  ClientWidth = 592
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 276
    Width = 592
    inherited PBouton: TToolWindow97
      ClientWidth = 592
      ClientAreaWidth = 592
      inherited BValider: TToolbarButton97
        Left = 496
      end
      inherited BFerme: TToolbarButton97
        Left = 528
      end
      inherited HelpBtn: TToolbarButton97
        Left = 560
      end
      inherited BImprimer: TToolbarButton97
        Left = 464
      end
    end
  end
  object Psource: TPanel [1]
    Left = 0
    Top = 0
    Width = 296
    Height = 276
    Align = alLeft
    TabOrder = 1
    object Bevel1: TBevel
      Left = 166
      Top = 87
      Width = 106
      Height = 21
    end
    object TBuds: TLabel
      Left = 20
      Top = 36
      Width = 34
      Height = 13
      Caption = '&Source'
      FocusControl = BudS
    end
    object TNatS: TLabel
      Left = 20
      Top = 64
      Width = 32
      Height = 13
      Caption = '&Nature'
    end
    object NbE: TLabel
      Left = 20
      Top = 91
      Width = 125
      Height = 13
      Caption = 'Nombre de pi'#232'ce trouv'#233'es'
    end
    object NbEcr: TLabel
      Left = 179
      Top = 91
      Width = 75
      Height = 13
      Alignment = taCenter
      AutoSize = False
      Caption = 'NBEcr'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label1: TLabel
      Left = 124
      Top = 7
      Width = 101
      Height = 16
      Caption = 'Budget source'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object BudS: THValComboBox
      Left = 92
      Top = 32
      Width = 180
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = BudSChange
      TagDispatch = 0
      DataType = 'TTBUDJAL'
    end
    object GbCS: TGroupBox
      Left = 9
      Top = 108
      Width = 277
      Height = 79
      Caption = ' Exercice '
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      object TExoDebS: TLabel
        Left = 11
        Top = 22
        Width = 29
        Height = 13
        Caption = 'D'#233'but'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TExoFinS: TLabel
        Left = 11
        Top = 51
        Width = 14
        Height = 13
        Caption = 'Fin'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object ExoDebS: THValComboBox
        Left = 86
        Top = 18
        Width = 185
        Height = 21
        Style = csSimple
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TTEXERCICEBUDGET'
      end
      object ExoFinS: THValComboBox
        Left = 86
        Top = 47
        Width = 185
        Height = 21
        Style = csSimple
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 1
        TagDispatch = 0
        DataType = 'TTEXERCICEBUDGET'
      end
      object GenS: TEdit
        Left = 212
        Top = 8
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        Visible = False
      end
      object SecS: TEdit
        Left = 188
        Top = 8
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        Visible = False
      end
      object SattS: TEdit
        Left = 140
        Top = 8
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        Visible = False
      end
      object GattS: TEdit
        Left = 164
        Top = 8
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        Visible = False
      end
      object ShS: TEdit
        Left = 116
        Top = 8
        Width = 20
        Height = 21
        Color = clYellow
        TabOrder = 6
        Visible = False
      end
      object AxS: THValComboBox
        Left = 236
        Top = 8
        Width = 20
        Height = 21
        Style = csSimple
        Color = clYellow
        ItemHeight = 13
        TabOrder = 7
        Visible = False
        TagDispatch = 0
        DataType = 'TTAXE'
      end
    end
    object NatS: THValComboBox
      Left = 92
      Top = 60
      Width = 180
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 2
      OnChange = NatSChange
      TagDispatch = 0
      DataType = 'TTNATECRBUD'
    end
    object GbPer: TGroupBox
      Left = 9
      Top = 190
      Width = 277
      Height = 79
      Caption = ' P'#233'riode '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      object TPerdebS: TLabel
        Left = 11
        Top = 22
        Width = 29
        Height = 13
        Caption = 'D'#233'&but'
        FocusControl = PerDebS
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TPerFinS: TLabel
        Left = 11
        Top = 51
        Width = 14
        Height = 13
        Caption = '&Fin'
        FocusControl = PerFinS
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object PerDebS: THValComboBox
        Left = 86
        Top = 18
        Width = 185
        Height = 21
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 0
        TagDispatch = 0
      end
      object PerFinS: THValComboBox
        Left = 86
        Top = 47
        Width = 185
        Height = 21
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 1
        TagDispatch = 0
      end
      object CbS: THValComboBox
        Left = 42
        Top = 36
        Width = 34
        Height = 21
        Style = csDropDownList
        Color = clYellow
        ItemHeight = 13
        TabOrder = 2
        Visible = False
        TagDispatch = 0
      end
    end
  end
  object PDestination: TPanel [2]
    Left = 296
    Top = 0
    Width = 296
    Height = 276
    Align = alClient
    TabOrder = 2
    object TBudD: TLabel
      Left = 20
      Top = 36
      Width = 53
      Height = 13
      Caption = '&Destination'
    end
    object TNatD: TLabel
      Left = 20
      Top = 64
      Width = 32
      Height = 13
      Caption = 'N&ature'
    end
    object TCoef: TLabel
      Left = 20
      Top = 91
      Width = 205
      Height = 13
      Caption = '&Coefficient '#224' appliquer pour la copie (+ ou -)'
    end
    object LPct: TLabel
      Left = 275
      Top = 91
      Width = 8
      Height = 13
      Caption = '%'
    end
    object Label2: TLabel
      Left = 120
      Top = 7
      Width = 130
      Height = 16
      Caption = 'Budget destination'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Coef: TEdit
      Left = 232
      Top = 87
      Width = 40
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 3
    end
    object BudD: THValComboBox
      Left = 92
      Top = 32
      Width = 191
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = BudSChange
      TagDispatch = 0
      DataType = 'TTBUDJAL'
    end
    object GroupBox2: TGroupBox
      Left = 9
      Top = 108
      Width = 277
      Height = 79
      Caption = ' Exercice '
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      object TExoDebD: TLabel
        Left = 11
        Top = 22
        Width = 29
        Height = 13
        Caption = 'D'#233'but'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TExoFinD: TLabel
        Left = 11
        Top = 51
        Width = 14
        Height = 13
        Caption = 'Fin'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object ExoDebD: THValComboBox
        Left = 86
        Top = 18
        Width = 185
        Height = 21
        Style = csSimple
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 7
        TagDispatch = 0
        DataType = 'TTEXERCICEBUDGET'
      end
      object ExoFinD: THValComboBox
        Left = 86
        Top = 47
        Width = 185
        Height = 21
        Style = csSimple
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TTEXERCICEBUDGET'
      end
      object GenD: TEdit
        Left = 184
        Top = 12
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        Visible = False
      end
      object SecD: TEdit
        Left = 160
        Top = 12
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        Visible = False
      end
      object GattD: TEdit
        Left = 136
        Top = 12
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        Visible = False
      end
      object SattD: TEdit
        Left = 112
        Top = 12
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        Visible = False
      end
      object ShD: TEdit
        Left = 87
        Top = 12
        Width = 20
        Height = 21
        Color = clYellow
        TabOrder = 5
        Visible = False
      end
      object AxD: THValComboBox
        Left = 208
        Top = 12
        Width = 20
        Height = 21
        Style = csSimple
        Color = clYellow
        ItemHeight = 13
        TabOrder = 6
        Visible = False
        TagDispatch = 0
        DataType = 'TTAXE'
      end
    end
    object NatD: THValComboBox
      Left = 92
      Top = 60
      Width = 191
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 2
      TagDispatch = 0
      DataType = 'TTNATECRBUD'
    end
    object GroupBox1: TGroupBox
      Left = 9
      Top = 190
      Width = 277
      Height = 79
      Caption = ' P'#233'riode '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 4
      object TPerdebF: TLabel
        Left = 11
        Top = 22
        Width = 29
        Height = 13
        Caption = 'D'#233'b&ut'
        FocusControl = PerDebD
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TPerFinF: TLabel
        Left = 11
        Top = 51
        Width = 14
        Height = 13
        Caption = 'F&in'
        FocusControl = PerFinD
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object PerDebD: THValComboBox
        Left = 86
        Top = 18
        Width = 185
        Height = 21
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 0
        TagDispatch = 0
      end
      object PerFinD: THValComboBox
        Left = 86
        Top = 47
        Width = 185
        Height = 21
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 1
        TagDispatch = 0
      end
      object CbD: THValComboBox
        Left = 42
        Top = 36
        Width = 34
        Height = 21
        Style = csDropDownList
        Color = clYellow
        ItemHeight = 13
        TabOrder = 2
        Visible = False
        TagDispatch = 0
      end
    end
  end
  inherited HMTrad: THSystemMenu
    Left = 280
    Top = 152
  end
  object QCopi: TQuery
    DatabaseName = 'SOC'
    RequestLive = True
    Left = 281
    Top = 96
  end
  object QNbEcr: TQuery
    DatabaseName = 'SOC'
    SQL.Strings = (
      
        'Select Distinct BE_NUMEROPIECE From BUDECR Where BE_BUDJAL=:BudJ' +
        ' '
      'And BE_NATUREBUD=:BudN')
    Left = 281
    Top = 48
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'BudJ'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'BudN'
        ParamType = ptUnknown
      end>
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Copie de budget;D'#233'sirez vous copier le budget s'#233'lectionn'#233'?;Q;Y' +
        'N;N;N;'
      
        '1;Copie de budget;Le journal source comporte aucune ligne d'#39#233'cri' +
        'ture!;W;O;O;O;'
      '2;Copie de budget;Aucun journal a '#233't'#233' s'#233'lectionn'#233';W;O;O;O;'
      
        '3;Copie de budget;Aucune nature de budget a '#233't'#233' s'#233'lectionn'#233'e;W;O' +
        ';O;O;'
      
        '4;Copie de budget;Le nombre de p'#233'riode du journal de destination' +
        ' est diff'#233'rent du journal source;W;O;O;O;'
      
        '5;Copie de budget;Les comptes g'#233'n'#233'raux budg'#233'taires sont diff'#233'ren' +
        'ts. Vous devez choisir un autre journal de destination;W;O;O;O;'
      
        '6;Copie de budget;Les sections budg'#233'taires sont diff'#233'rentes. Vou' +
        's devez choisir un autre journal de destination;W;O;O;O;'
      
        '7;Copie de budget;L'#39'axe du journal de destination est diff'#233'rent ' +
        'de l'#39'axe du journal source! Vous devez choisir un autre journal ' +
        'de destination;W;O;O;O;'
      'Insertion des enregistrements en cours'
      'Nombre de pi'#232'ce trouv'#233'e'
      'Nombre de pi'#232'ces trouv'#233'es'
      
        '11;Copie de budget;Le compte g'#233'n'#233'ral d'#39'attente du journal de des' +
        'tination est diff'#233'rent du compte g'#233'n'#233'ral d'#39'attente du journal so' +
        'urce! Vous devez choisir un autre journal de destination;W;O;O;O' +
        ';'
      
        '12;Copie de budget;La section budg'#233'taire d'#39'attente du journal de' +
        ' destination est diff'#233'rent de la section budg'#233'taire d'#39'attente du' +
        ' journal source! Vous devez choisir un autre journal de destinat' +
        'ion;W;O;O;O;'
      
        '13;Copie de budget;Le coefficient '#224' appliquer pour la copie doit' +
        ' '#234'tre une valeur num'#233'rique;W;O;O;O;'
      
        '14;Copie de budget;Incoh'#233'rence dans le choix des p'#233'riodes du bud' +
        'get source;W;O;O;O;'
      
        '15;Copie de budget;Incoh'#233'rence dans le choix des p'#233'riodes du bud' +
        'get destination;W;O;O;O;'
      
        '16;Copie de budget;Le nombre de p'#233'riode du budget destination es' +
        't diff'#233'rente du nombre de p'#233'riode du budget source;W;O;O;O;'
      'ATTENTION : Recopie non effectu'#233'e !'
      
        '18;Copie de budget;La recopie de budget s'#39'est correctement effec' +
        'tu'#233'e !;E;O;O;O;')
    Left = 281
    Top = 2
  end
end
