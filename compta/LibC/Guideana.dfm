object FGuideAna: TFGuideAna
  Left = 166
  Top = 180
  Width = 544
  Height = 300
  HelpContext = 1430100
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Ventilation Analytique du Guide'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object DockTop: TDock97
    Left = 0
    Top = 0
    Width = 536
    Height = 9
  end
  object DockRight: TDock97
    Left = 527
    Top = 9
    Width = 9
    Height = 236
    Position = dpRight
  end
  object DockLeft: TDock97
    Left = 0
    Top = 9
    Width = 9
    Height = 236
    Position = dpLeft
  end
  object DockBottom: TDock97
    Left = 0
    Top = 245
    Width = 536
    Height = 28
    Position = dpBottom
    object Valide97: TToolbar97
      Left = 442
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DockPos = 442
      TabOrder = 0
      object BValider: TToolbarButton97
        Left = 0
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Valider'
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
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0127_S16G1'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 27
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Fermer'
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 54
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Aide'
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
    object PPied: TToolbar97
      Left = 0
      Top = 0
      Caption = 'Outils'
      CloseButton = False
      DockPos = 0
      TabOrder = 1
      object BInsLigne: TToolbarButton97
        Left = 0
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Ins'#233'rer ligne'
        ParentShowHint = False
        ShowHint = True
        OnClick = BInsLigneClick
        GlobalIndexImage = 'Z0074_S16G1'
      end
      object BDelLigne: TToolbarButton97
        Left = 27
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Supprimer ligne'
        ParentShowHint = False
        ShowHint = True
        OnClick = BDelLigneClick
        GlobalIndexImage = 'Z0072_S16G1'
      end
      object BZoom: TToolbarButton97
        Tag = 1
        Left = 54
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Zoom section'
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
        OnClick = BZoomClick
        GlobalIndexImage = 'Z0016_S16G1'
        IsControl = True
      end
    end
  end
  object PFEN: TPanel
    Left = 9
    Top = 9
    Width = 518
    Height = 236
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object PEntete: TPanel
      Left = 0
      Top = 0
      Width = 518
      Height = 29
      Align = alTop
      BevelOuter = bvNone
      Caption = ' '
      TabOrder = 0
      object Bevel2: TBevel
        Left = -3
        Top = 47
        Width = 546
        Height = 1
      end
      object PAxe: TTabControl
        Left = 0
        Top = 2
        Width = 541
        Height = 27
        TabOrder = 0
        Tabs.Strings = (
          'Axe n'#176'1'
          'Axe n'#176'2'
          'Axe n'#176'3'
          'Axe n'#176'4'
          'Axe n'#176'5')
        TabIndex = 0
        OnChange = PAxeChange
      end
    end
    object FListe: THGrid
      Left = 0
      Top = 29
      Width = 518
      Height = 207
      Align = alClient
      ColCount = 9
      DefaultColWidth = 20
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goEditing, goAlwaysShowEditor]
      TabOrder = 1
      OnDblClick = FListeDblClick
      OnSetEditText = FListeSetEditText
      SortedCol = -1
      Titres.Strings = (
        'N'#176';C'
        '*;C'
        'Section'
        '*;C'
        'Pourcentage'
        '*;C'
        'Qt'#233' n'#176'1'
        '*;C'
        'Qt'#233' n'#176'2')
      Couleur = False
      MultiSelect = False
      TitleBold = True
      TitleCenter = True
      OnCellExit = FListeCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
      ColWidths = (
        20
        17
        95
        17
        99
        19
        113
        20
        107)
    end
    object HS: THCpteEdit
      Left = 288
      Top = 2
      Width = 17
      Height = 21
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      Text = '!!!'
      Visible = False
      ZoomTable = tzSection
      Vide = False
      Bourre = True
      okLocate = False
      SynJoker = False
    end
  end
  object Messages: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Ventilation analytique du guide; Ce compte n'#39'est pas ventilabl' +
        'e sur cet axe;W;O;O;O;'
      
        '1;Ventilation analytique du guide; Section analytique ferm'#233'e : v' +
        'ous ne pouvez pas saisir d'#39#233'critures sur ce compte;W;O;O;O;'
      
        '2;Ventilation analytique du guide; Section analytique inexistant' +
        'e : vous devez imp'#233'rativement lui associer un arr'#234't en saisie.;W' +
        ';O;O;O;'
      
        '3;Ventilation analytique du guide; Section analytique non rensei' +
        'gn'#233'e : vous devez param'#233'trer un arr'#234't en saisie ou bien renseign' +
        'er la section.;W;O;O;O;'
      
        '4;Ventilation analytique du guide; Pourcentage incorrect.;W;O;O;' +
        'O;'
      
        '5;Ventilation analytique du guide; Vous devez indiquer une secti' +
        'on analytique existante;W;O;O;O;'
      '6;'
      '7;'
      '8;'
      '9;'
      '10;'
      '11;'
      
        '12;Ventilation analytique du guide; Erreur de syntaxe dans la fo' +
        'rmule.;W;O;O;O;'
      
        '13;Ventilation analytique du guide; Erreur dans la formule : r'#233'f' +
        #233'rence incorrecte '#224' une ligne.;W;O;O;O;'
      
        '14;Ventilation analytique du guide;Cette ligne est vide !;W;O;O;' +
        'O;'
      
        '15;Ventilation analytique du guide; Erreur dans la formule : r'#233'f' +
        #233'rence '#224' un champ inexistant.;W;O;O;O;'
      
        '16;Ventilation analytique du guide; Ventilation incorrecte : la ' +
        'somme des ventilations doit '#234'tre '#233'gale '#224' 100 %.;W;O;O;O;'
      'Enregistrement non valid'#233' !')
    Left = 48
    Top = 113
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 104
    Top = 112
  end
end
