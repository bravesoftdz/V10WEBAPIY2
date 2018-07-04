object FVideSect: TFVideSect
  Left = 306
  Top = 189
  HelpContext = 7379000
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Transferts inter-sections'
  ClientHeight = 431
  ClientWidth = 377
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock: TDock97
    Left = 0
    Top = 395
    Width = 377
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 377
      Caption = 'Barre d'#39'outils'
      ClientAreaHeight = 32
      ClientAreaWidth = 377
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object FlashVide: TFlashingLabel
        Left = 23
        Top = 10
        Width = 192
        Height = 13
        Caption = '<< Traitement de g'#233'n'#233'ration en cours >>'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Visible = False
      end
      object BGenere: TToolbarButton97
        Tag = 1
        Left = 247
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Liste des pi'#232'ces analytiques g'#233'n'#233'r'#233'es'
        DisplayMode = dmGlyphOnly
        Caption = 'G'#233'n'#233'rer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 0
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BGenereClick
        GlobalIndexImage = 'Z1826_S16G1'
        IsControl = True
      end
      object BValide: TToolbarButton97
        Tag = 1
        Left = 279
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Lancer le transfert inter-sections'
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValideClick
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 311
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Fermer'
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAbandonClick
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 343
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
  end
  object HPanel2: THPanel
    Left = 0
    Top = 354
    Width = 377
    Height = 41
    Align = alBottom
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object BFiltre: TToolbarButton97
      Left = 8
      Top = 8
      Width = 65
      Height = 22
      Hint = 'Menu filtre'
      DropdownArrow = True
      DropdownMenu = POPF
      Caption = '&Filtre'
      Flat = False
    end
    object FFiltres: THValComboBox
      Left = 88
      Top = 8
      Width = 241
      Height = 21
      ItemHeight = 13
      TabOrder = 0
      Text = 'FFiltres'
      TagDispatch = 0
    end
  end
  object HPanel1: THPanel
    Left = 0
    Top = 0
    Width = 377
    Height = 354
    Align = alClient
    Caption = 'HPanel1'
    FullRepaint = False
    TabOrder = 2
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object Pages: TPageControl
      Left = 1
      Top = 1
      Width = 375
      Height = 352
      ActivePage = TabSheet1
      Align = alClient
      TabHeight = 1
      TabOrder = 0
      TabWidth = 373
      object TabSheet1: TTabSheet
        object GCRIT: TGroupBox
          Left = 8
          Top = 0
          Width = 359
          Height = 186
          Caption = ' Crit'#232'res de s'#233'lection '
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clActiveCaption
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 0
          object HLabel3: THLabel
            Left = 12
            Top = 18
            Width = 103
            Height = 13
            Caption = '&Montants '#224' r'#233'partir du'
            Color = clBtnFace
            FocusControl = Y_DATECOMPTABLE
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object HLabel6: THLabel
            Left = 224
            Top = 18
            Width = 12
            Height = 13
            Caption = 'au'
            FocusControl = Y_DATECOMPTABLE_
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object TY_AXE: THLabel
            Left = 12
            Top = 79
            Width = 18
            Height = 13
            Caption = '&Axe'
            Color = clBtnFace
            FocusControl = Y_AXE
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object TS_CLEREPARTITION: THLabel
            Left = 12
            Top = 105
            Width = 82
            Height = 13
            Caption = 'Cl'#233'  de &r'#233'partition'
            Color = clBtnFace
            FocusControl = S_CLEREPARTITION
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object HLabel1: THLabel
            Left = 12
            Top = 132
            Width = 87
            Height = 13
            Caption = '&Calcul de la cl'#233' du'
            Color = clBtnFace
            FocusControl = Y_DATECOMPTABLE1
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object HLabel2: THLabel
            Left = 224
            Top = 132
            Width = 12
            Height = 13
            Caption = 'au'
            FocusControl = Y_DATECOMPTABLE1_
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object TGENEEMET: THLabel
            Left = 38
            Top = 43
            Width = 77
            Height = 13
            Caption = 'sur les &g'#233'n'#233'raux'
            Color = clBtnFace
            FocusControl = GENEEMET
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object TGENERECEPT: THLabel
            Left = 38
            Top = 159
            Width = 77
            Height = 13
            Caption = 'sur les &g'#233'n'#233'raux'
            Color = clBtnFace
            FocusControl = GENERECEPT
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object Y_DATECOMPTABLE: THCritMaskEdit
            Left = 124
            Top = 14
            Width = 93
            Height = 21
            Ctl3D = True
            EditMask = '!99/99/0000;1;_'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            MaxLength = 10
            ParentCtl3D = False
            ParentFont = False
            TabOrder = 0
            Text = '  /  /    '
            OnKeyPress = Y_DATECOMPTABLEKeyPress
            TagDispatch = 0
            Operateur = Superieur
            OpeType = otDate
            ControlerDate = True
          end
          object Y_DATECOMPTABLE_: THCritMaskEdit
            Left = 248
            Top = 14
            Width = 93
            Height = 21
            Ctl3D = True
            EditMask = '!99/99/0000;1;_'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            MaxLength = 10
            ParentCtl3D = False
            ParentFont = False
            TabOrder = 1
            Text = '  /  /    '
            OnKeyPress = Y_DATECOMPTABLEKeyPress
            TagDispatch = 0
            Operateur = Inferieur
            OpeType = otDate
            ControlerDate = True
          end
          object Y_AXE: THValComboBox
            Left = 124
            Top = 75
            Width = 218
            Height = 21
            Style = csDropDownList
            Ctl3D = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ItemHeight = 13
            ParentCtl3D = False
            ParentFont = False
            TabOrder = 3
            OnChange = Y_AXEChange
            TagDispatch = 0
            DataType = 'TTAXE'
          end
          object S_CLEREPARTITION: THValComboBox
            Left = 124
            Top = 101
            Width = 218
            Height = 21
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ItemHeight = 13
            ParentFont = False
            TabOrder = 4
            OnChange = S_CLEREPARTITIONChange
            TagDispatch = 0
            DataType = 'TTCLEREPART1'
          end
          object Y_DATECOMPTABLE1: THCritMaskEdit
            Left = 124
            Top = 128
            Width = 93
            Height = 21
            Ctl3D = True
            EditMask = '!99/99/0000;1;_'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            MaxLength = 10
            ParentCtl3D = False
            ParentFont = False
            TabOrder = 5
            Text = '  /  /    '
            OnKeyPress = Y_DATECOMPTABLEKeyPress
            TagDispatch = 0
            Operateur = Superieur
            OpeType = otDate
            ControlerDate = True
          end
          object Y_DATECOMPTABLE1_: THCritMaskEdit
            Left = 248
            Top = 128
            Width = 93
            Height = 21
            Ctl3D = True
            EditMask = '!99/99/0000;1;_'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            MaxLength = 10
            ParentCtl3D = False
            ParentFont = False
            TabOrder = 6
            Text = '  /  /    '
            OnKeyPress = Y_DATECOMPTABLEKeyPress
            TagDispatch = 0
            Operateur = Inferieur
            OpeType = otDate
            ControlerDate = True
          end
          object GENEEMET: TEdit
            Left = 124
            Top = 39
            Width = 217
            Height = 21
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            MaxLength = 35
            ParentFont = False
            TabOrder = 2
          end
          object GENERECEPT: TEdit
            Left = 124
            Top = 155
            Width = 217
            Height = 21
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            MaxLength = 35
            ParentFont = False
            TabOrder = 7
          end
        end
        object GPARAM: TGroupBox
          Left = 8
          Top = 190
          Width = 359
          Height = 123
          Caption = ' Param'#232'tres de g'#233'n'#233'ration '
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clActiveCaption
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 1
          object H_JOURNALE: THLabel
            Left = 8
            Top = 18
            Width = 34
            Height = 13
            Caption = '&Journal'
            FocusControl = Y_JOURNAL
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object HLabel9: THLabel
            Left = 8
            Top = 45
            Width = 101
            Height = 13
            Caption = '&Date comptabilisation'
            FocusControl = DATEGENERATION
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object HLabel7: THLabel
            Left = 8
            Top = 71
            Width = 50
            Height = 13
            Caption = 'R'#233'&f'#233'rence'
            FocusControl = Y_REFINTERNE
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object HLabel10: THLabel
            Left = 8
            Top = 98
            Width = 30
            Height = 13
            Caption = '&Libell'#233
            FocusControl = Y_LIBELLE
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object Y_JOURNAL: THValComboBox
            Tag = 1
            Left = 124
            Top = 14
            Width = 218
            Height = 21
            Style = csDropDownList
            Ctl3D = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ItemHeight = 13
            ParentCtl3D = False
            ParentFont = False
            TabOrder = 0
            OnChange = Y_JOURNALChange
            TagDispatch = 0
            DataType = 'TTJALANALYTIQUE'
          end
          object DATEGENERATION: THCritMaskEdit
            Left = 124
            Top = 41
            Width = 69
            Height = 21
            Ctl3D = True
            EditMask = '!99/99/0000;1;_'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            MaxLength = 10
            ParentCtl3D = False
            ParentFont = False
            TabOrder = 1
            Text = '  /  /    '
            OnExit = DATEGENERATIONExit
            TagDispatch = 0
            Operateur = Superieur
            OpeType = otDate
            ControlerDate = True
          end
          object Y_REFINTERNE: TEdit
            Left = 124
            Top = 67
            Width = 217
            Height = 21
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            MaxLength = 35
            ParentFont = False
            TabOrder = 2
          end
          object Y_LIBELLE: TEdit
            Left = 124
            Top = 94
            Width = 218
            Height = 21
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            MaxLength = 35
            ParentFont = False
            TabOrder = 3
          end
          object VERIFQTE: TCheckBox
            Left = 204
            Top = 45
            Width = 137
            Height = 17
            Alignment = taLeftJustify
            Caption = '&V'#233'rification des quantit'#233's'
            Checked = True
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
            State = cbChecked
            TabOrder = 4
          end
        end
      end
    end
  end
  object HMCle: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -12
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Transfert inter-sections;Vous devez renseigner une date valide' +
        ';W;O;O;O;'
      
        '1;Transfert inter-sections;La date que vous avez renseign'#233'e est ' +
        'sur un exercice non ouvert;W;O;O;O;'
      
        '2;Transfert inter-sections;La date que vous avez renseign'#233'e est ' +
        'sur un exercice non ouvert;W;O;O;O;'
      
        '3;Transfert inter-sections;La date que vous avez renseign'#233'e est ' +
        'ant'#233'rieure '#224' la cl'#244'ture provisoire;W;O;O;O;'
      
        '4;Transfert inter-sections;La date que vous avez renseign'#233'e est ' +
        'ant'#233'rieure '#224' la cl'#244'ture d'#233'finitive;W;O;O;O;'
      
        '5;Transfert inter-sections;Vous devez renseigner une cl'#233' de r'#233'pa' +
        'rtition;W;O;O;O;'
      
        '6;Transfert inter-sections;Vous devez renseigner un journal;W;O;' +
        'O;O;'
      
        '7;Transfert inter-sections;Vous devez renseigner un axe analytiq' +
        'ue;W;O;O;O;'
      
        '8;Transfert inter-sections;Il n'#39'existe aucun mouvement analytiqu' +
        'e pour ces crit'#232'res;E;O;O;O;'
      
        '9;Transfert inter-sections;Il n'#39'existe aucune section de transfe' +
        'rt pour cette cl'#233' de r'#233'partition;E;O;O;O;'
      
        '10;Transfert inter-sections;La somme des mouvements sur les sect' +
        'ions receptrices est nulle;E;O;O;O;'
      
        '11;Transfert inter-sections;Confirmez-vous la g'#233'n'#233'ration du tran' +
        'sfert et des pi'#232'ces associ'#233'es ?;Q;YN;N;N;'
      
        '12;Transfert inter-sections;Voulez-vous voir la liste des pi'#232'ces' +
        ' g'#233'n'#233'r'#233'es?;Q;YN;Y;Y;'
      
        '13;Transfert inter-sections;Certains qualifiants quantit'#233's des m' +
        'ouvements sont incoh'#233'rents. Voulez-vous lancer le traitement sur' +
        ' les qualifiants correspondant '#224' la cl'#233' de r'#233'partition ?;Q;YN;Y;' +
        'Y;'
      'ATTENTION! G'#233'n'#233'ration non effectu'#233'e'
      
        '15;Transfert inter-sections;Vous devez renseigner un journal qui' +
        ' porte sur l'#39'axe choisi.;W;O;O;O;'
      'ATTENTION! G'#233'n'#233'ration effectu'#233'e mais non m'#233'moris'#233'e.'
      
        '17;Transfert inter-sections;Un transfert a d'#233'j'#224' '#233't'#233' effectu'#233' sur' +
        ' tout ou partie de la p'#233'riode choisie. Confirmez-vous l'#39'op'#233'ratio' +
        'n ?;Q;YN;Y;Y;'
      
        '18;Transfert inter-sections;Transfert impossible : vous devez d'#233 +
        'finir un compte g'#233'n'#233'ral d'#39'attente de l'#39'axe;W;O;O;O;'
      
        '19;Transfert inter-sections;Vous devez renseigner un journal qui' +
        ' poss'#232'de un facturier.;W;O;O;O;'
      
        '20;Transfert inter-sections;La d'#233'finition des crit'#232'res sur compt' +
        'es g'#233'n'#233'raux est incorrecte.;W;O;O;O;'
      '21;')
    Left = 308
    Top = 80
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 224
    Top = 80
  end
  object POPF: TPopupMenu
    Left = 344
    Top = 347
  end
end
