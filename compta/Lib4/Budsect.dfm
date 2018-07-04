object Fbudsect: TFbudsect
  Left = 351
  Top = 107
  HelpContext = 15133000
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Section budg'#233'taire : '
  ClientHeight = 217
  ClientWidth = 516
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  Scaled = False
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object BPages: TPageControl
    Left = 0
    Top = 0
    Width = 516
    Height = 182
    ActivePage = ZL
    Align = alClient
    TabOrder = 0
    object PCaract: TTabSheet
      HelpContext = 15133100
      Caption = 'Caract'#233'ristiques'
      object TBS_BUDSECT: THLabel
        Left = 5
        Top = 16
        Width = 36
        Height = 13
        Caption = 'S&ection'
        FocusControl = BS_BUDSECT
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBS_LIBELLE: THLabel
        Left = 5
        Top = 43
        Width = 30
        Height = 13
        Caption = '&Libell'#233
        FocusControl = BS_LIBELLE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBS_SENS: THLabel
        Left = 5
        Top = 99
        Width = 24
        Height = 13
        Caption = '&Sens'
        FocusControl = BS_SENS
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBS_ABREGE: THLabel
        Left = 306
        Top = 43
        Width = 34
        Height = 13
        Caption = '&Abr'#233'g'#233
        FocusControl = BS_ABREGE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBS_SIGNE: TLabel
        Left = 306
        Top = 70
        Width = 27
        Height = 13
        Caption = 'S&igne'
        Color = clYellow
        FocusControl = BS_SIGNE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        Visible = False
      end
      object TBS_PARENT: TLabel
        Left = 328
        Top = 99
        Width = 67
        Height = 13
        Caption = 'Budget &parent'
        Color = clYellow
        FocusControl = BS_PARENT
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        Visible = False
      end
      object TBS_REPORTDISPO: TLabel
        Left = 402
        Top = 99
        Width = 97
        Height = 13
        Caption = '&Report du disponible'
        Color = clYellow
        FocusControl = BS_REPORTDISPO
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        Visible = False
      end
      object TBS_FORMULE: TLabel
        Left = 257
        Top = 99
        Width = 37
        Height = 13
        Caption = '&Formule'
        Color = clYellow
        FocusControl = BS_FORMULE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        Visible = False
      end
      object TBS_AXE: TLabel
        Left = 5
        Top = 70
        Width = 18
        Height = 13
        Caption = 'A&xe'
        FocusControl = BS_AXE
      end
      object TBS_RUB: THLabel
        Left = 226
        Top = 16
        Width = 114
        Height = 13
        Caption = 'Code &rubrique '#224' g'#233'n'#233'rer'
        FocusControl = BS_RUB
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object BS_BUDSECT: TDBEdit
        Left = 55
        Top = 12
        Width = 150
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        DataField = 'BS_BUDSECT'
        DataSource = SBudsect
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
      end
      object BS_LIBELLE: TDBEdit
        Left = 55
        Top = 39
        Width = 235
        Height = 21
        Ctl3D = True
        DataField = 'BS_LIBELLE'
        DataSource = SBudsect
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
      end
      object BS_SENS: THDBValComboBox
        Left = 55
        Top = 95
        Width = 193
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 0
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 6
        TagDispatch = 0
        DataType = 'TTSENS'
        DataField = 'BS_SENS'
        DataSource = SBudsect
      end
      object BS_ABREGE: TDBEdit
        Left = 354
        Top = 39
        Width = 150
        Height = 21
        Ctl3D = True
        DataField = 'BS_ABREGE'
        DataSource = SBudsect
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
      end
      object BS_SIGNE: THDBValComboBox
        Left = 354
        Top = 67
        Width = 150
        Height = 21
        Style = csDropDownList
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 0
        ParentFont = False
        TabOrder = 5
        Visible = False
        TagDispatch = 0
        DataType = 'TTRUBSIGNE'
        DataField = 'BS_SIGNE'
        DataSource = SBudsect
      end
      object BS_REPORTDISPO: THDBValComboBox
        Left = 375
        Top = 96
        Width = 26
        Height = 21
        Style = csDropDownList
        Color = clYellow
        ItemHeight = 0
        TabOrder = 12
        Visible = False
        TagDispatch = 0
        DataField = 'BS_REPORTDISPO'
        DataSource = SBudsect
      end
      object BS_PARENT: TDBEdit
        Left = 440
        Top = 96
        Width = 21
        Height = 21
        Color = clYellow
        DataField = 'BS_PARENT'
        DataSource = SBudsect
        TabOrder = 11
        Visible = False
      end
      object BS_ATTENTE: TDBCheckBox
        Left = 163
        Top = 126
        Width = 156
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Sec&tion budg'#233'taire d'#39'attente'
        DataField = 'BS_ATTENTE'
        DataSource = SBudsect
        TabOrder = 8
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
      object BS_HT: TDBCheckBox
        Left = 4
        Top = 126
        Width = 108
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Budget &hors taxes'
        DataField = 'BS_HT'
        DataSource = SBudsect
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 7
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
      object BS_FORMULE: TDBEdit
        Left = 303
        Top = 99
        Width = 14
        Height = 21
        Color = clYellow
        DataField = 'BS_FORMULE'
        DataSource = SBudsect
        TabOrder = 10
        Visible = False
      end
      object BS_CONFIDENTIEL: TDBCheckBox
        Left = 388
        Top = 126
        Width = 108
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Confidentiel'
        DataField = 'BS_CONFIDENTIEL'
        DataSource = SBudsect
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 9
        ValueChecked = '1'
        ValueUnchecked = '0'
      end
      object BS_AXE: THDBValComboBox
        Left = 55
        Top = 67
        Width = 193
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 4
        OnChange = BS_AXEChange
        TagDispatch = 0
        DataType = 'TTAXE'
        DataField = 'BS_AXE'
        DataSource = SBudsect
      end
      object BS_RUB: TDBEdit
        Left = 354
        Top = 12
        Width = 150
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        DataField = 'BS_RUB'
        DataSource = SBudsect
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
      end
    end
    object PComplement: TTabSheet
      HelpContext = 15133200
      Caption = 'Compl'#233'ments'
      object FListe: THGrid
        Left = 5
        Top = 5
        Width = 498
        Height = 147
        ColCount = 2
        DefaultRowHeight = 18
        FixedCols = 0
        RowCount = 2
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
        TabOrder = 0
        OnKeyDown = FListeKeyDown
        OnKeyPress = FListeKeyPress
        SortedCol = -1
        Titres.Strings = (
          'Sections analytiques'
          'Sections analytiques exclues')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnCellExit = FListeCellExit
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = clSilver
        ColWidths = (
          238
          238)
      end
    end
    object PInfo: TTabSheet
      HelpContext = 15133300
      Caption = 'Informations'
      object HGBDates: TGroupBox
        Left = 2
        Top = 2
        Width = 503
        Height = 54
        Caption = ' Dates '
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object TBS_DATECREATION: THLabel
          Left = 62
          Top = 0
          Width = 39
          Height = 13
          Caption = ' Cr'#233#233' le '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBS_DATEMODIF: THLabel
          Left = 165
          Top = 0
          Width = 51
          Height = 13
          Caption = ' Modifi'#233' le '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBS_DATEOUVERTURE: THLabel
          Left = 276
          Top = 0
          Width = 49
          Height = 13
          Caption = ' Ouvert le '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBS_DATEFERMETURE: THLabel
          Left = 387
          Top = 0
          Width = 46
          Height = 13
          Caption = ' Ferm'#233' le '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BS_FERME: TDBCheckBox
          Left = 116
          Top = 0
          Width = 41
          Height = 13
          Caption = 'Ferm'#233
          Color = clYellow
          DataField = 'BS_FERME'
          DataSource = SBudsect
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 0
          ValueChecked = 'X'
          ValueUnchecked = '-'
          Visible = False
        end
        object BS_DATECREATION: TDBEdit
          Left = 54
          Top = 22
          Width = 68
          Height = 19
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'BS_DATECREATION'
          DataSource = SBudsect
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          ReadOnly = True
          TabOrder = 1
        end
        object BS_DATEMODIF: TDBEdit
          Left = 157
          Top = 22
          Width = 68
          Height = 19
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'BS_DATEMODIF'
          DataSource = SBudsect
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          ReadOnly = True
          TabOrder = 2
        end
        object BS_DATEOUVERTURE: TDBEdit
          Left = 268
          Top = 22
          Width = 68
          Height = 19
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'BS_DATEOUVERTURE'
          DataSource = SBudsect
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          ReadOnly = True
          TabOrder = 3
        end
        object BS_DATEFERMETURE: TDBEdit
          Left = 379
          Top = 22
          Width = 68
          Height = 19
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = False
          DataField = 'BS_DATEFERMETURE'
          DataSource = SBudsect
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          ReadOnly = True
          TabOrder = 4
        end
      end
      object TBS_BLOCNOTE: TGroupBox
        Left = 2
        Top = 56
        Width = 503
        Height = 96
        Caption = ' Bloc Note '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object BS_BLOCNOTE: THDBRichEditOLE
          Left = 7
          Top = 13
          Width = 488
          Height = 76
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          MaxLength = 262140
          ParentCtl3D = False
          ParentFont = False
          ScrollBars = ssVertical
          TabOrder = 0
          Margins.Top = 0
          Margins.Bottom = 0
          Margins.Left = 0
          Margins.Right = 0
          ContainerName = 'Document'
          ObjectMenuPrefix = '&Object'
          LinesRTF.Strings = (
            
              '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fswiss Ar' +
              'ial;}}'
            '{\colortbl ;\red0\green0\blue0;}'
            
              '{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1\pard\cf1\f0\' +
              'fs16 BS_BLOCNOTE'
            '\par }')
          DataField = 'BS_BLOCNOTE'
          DataSource = SBudsect
        end
      end
    end
    object ZL: TTabSheet
      Caption = 'Tables libres'
      object TBS_TABLE0: THLabel
        Left = 3
        Top = 18
        Width = 89
        Height = 13
        Caption = '&Premi'#232're table libre'
        FocusControl = BS_TABLE0
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBS_TABLE1: THLabel
        Left = 3
        Top = 46
        Width = 95
        Height = 13
        Caption = '&Deuxi'#232'me table libre'
        FocusControl = BS_TABLE1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBS_TABLE2: THLabel
        Left = 3
        Top = 75
        Width = 93
        Height = 13
        Caption = '&Troisi'#232'me table libre'
        FocusControl = BS_TABLE2
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBS_TABLE3: THLabel
        Left = 3
        Top = 103
        Width = 96
        Height = 13
        Caption = '&Quatri'#232'me table libre'
        FocusControl = BS_TABLE3
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBS_TABLE4: THLabel
        Left = 3
        Top = 131
        Width = 97
        Height = 13
        Caption = '&Cinqui'#232'me table libre'
        FocusControl = BS_TABLE4
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBS_TABLE5: THLabel
        Left = 258
        Top = 18
        Width = 84
        Height = 13
        Caption = '&Sixi'#232'me table libre'
        FocusControl = BS_TABLE5
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBS_TABLE6: THLabel
        Left = 258
        Top = 46
        Width = 92
        Height = 13
        Caption = 'S&epti'#232'me table libre'
        FocusControl = BS_TABLE6
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBS_TABLE7: THLabel
        Left = 258
        Top = 75
        Width = 89
        Height = 13
        Caption = '&Huiti'#232'me table libre'
        FocusControl = BS_TABLE7
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBS_TABLE8: THLabel
        Left = 258
        Top = 103
        Width = 96
        Height = 13
        Caption = '&Neuvi'#232'me table libre'
        FocusControl = BS_TABLE8
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBS_TABLE9: THLabel
        Left = 258
        Top = 131
        Width = 85
        Height = 13
        Caption = 'Di&xi'#232'me table libre'
        FocusControl = BS_TABLE9
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object BS_TABLE0: THDBCpteEdit
        Left = 103
        Top = 14
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 0
        ZoomTable = tzNatBudS0
        Vide = True
        Bourre = False
        okLocate = True
        SynJoker = False
        DataField = 'BS_TABLE0'
        DataSource = SBudsect
      end
      object BS_TABLE1: THDBCpteEdit
        Left = 103
        Top = 42
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 1
        ZoomTable = tzNatBudS1
        Vide = True
        Bourre = False
        okLocate = True
        SynJoker = False
        DataField = 'BS_TABLE1'
        DataSource = SBudsect
      end
      object BS_TABLE2: THDBCpteEdit
        Left = 103
        Top = 71
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 2
        ZoomTable = tzNatBudS2
        Vide = True
        Bourre = False
        okLocate = True
        SynJoker = False
        DataField = 'BS_TABLE2'
        DataSource = SBudsect
      end
      object BS_TABLE3: THDBCpteEdit
        Left = 103
        Top = 99
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 3
        ZoomTable = tzNatBudS3
        Vide = True
        Bourre = False
        okLocate = True
        SynJoker = False
        DataField = 'BS_TABLE3'
        DataSource = SBudsect
      end
      object BS_TABLE4: THDBCpteEdit
        Left = 103
        Top = 127
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 4
        ZoomTable = tzNatBudS4
        Vide = True
        Bourre = False
        okLocate = True
        SynJoker = False
        DataField = 'BS_TABLE4'
        DataSource = SBudsect
      end
      object BS_TABLE5: THDBCpteEdit
        Left = 356
        Top = 14
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 5
        ZoomTable = tzNatBudS5
        Vide = True
        Bourre = False
        okLocate = True
        SynJoker = False
        DataField = 'BS_TABLE5'
        DataSource = SBudsect
      end
      object BS_TABLE6: THDBCpteEdit
        Left = 356
        Top = 42
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 6
        ZoomTable = tzNatBudS6
        Vide = True
        Bourre = False
        okLocate = True
        SynJoker = False
        DataField = 'BS_TABLE6'
        DataSource = SBudsect
      end
      object BS_TABLE7: THDBCpteEdit
        Left = 356
        Top = 71
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 7
        ZoomTable = tzNatBudS7
        Vide = True
        Bourre = False
        okLocate = True
        SynJoker = False
        DataField = 'BS_TABLE7'
        DataSource = SBudsect
      end
      object BS_TABLE8: THDBCpteEdit
        Left = 356
        Top = 99
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 8
        ZoomTable = tzNatBudS8
        Vide = True
        Bourre = False
        okLocate = True
        SynJoker = False
        DataField = 'BS_TABLE8'
        DataSource = SBudsect
      end
      object BS_TABLE9: THDBCpteEdit
        Left = 356
        Top = 127
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 9
        ZoomTable = tzNatBudS9
        Vide = True
        Bourre = False
        okLocate = True
        SynJoker = False
        DataField = 'BS_TABLE9'
        DataSource = SBudsect
      end
    end
  end
  object HPB: TPanel
    Left = 0
    Top = 182
    Width = 516
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    Ctl3D = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentCtl3D = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = False
    TabOrder = 1
    object FAutoSave: TCheckBox
      Left = 248
      Top = 8
      Width = 34
      Height = 17
      Caption = 'Auto'
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      TabOrder = 0
      Visible = False
    end
    object BAide: THBitBtn
      Left = 484
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Aide'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 10
      OnClick = BAideClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z1117_S16G1'
      IsControl = True
    end
    object BAnnuler: THBitBtn
      Left = 120
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Annuler la derni'#232're action'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
      OnClick = BAnnulerClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0075_S16G1'
      IsControl = True
    end
    object BValider: THBitBtn
      Left = 420
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Valider'
      Default = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 8
      OnClick = BValiderClick
      Glyph.Data = {
        BE060000424DBE06000000000000360400002800000024000000120000000100
        0800000000008802000000000000000000000001000000010000000000000000
        80000080000000808000800000008000800080800000C0C0C000C0DCC000F0C8
        A400000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        000000000000000000000000000000000000F0FBFF00A4A0A000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00030303030303
        0303030303030303030303030303030303030303030303030303030303030303
        03030303030303030303030303030303030303030303FF030303030303030303
        03030303030303040403030303030303030303030303030303F8F8FF03030303
        03030303030303030303040202040303030303030303030303030303F80303F8
        FF030303030303030303030303040202020204030303030303030303030303F8
        03030303F8FF0303030303030303030304020202020202040303030303030303
        0303F8030303030303F8FF030303030303030304020202FA0202020204030303
        0303030303F8FF0303F8FF030303F8FF03030303030303020202FA03FA020202
        040303030303030303F8FF03F803F8FF0303F8FF03030303030303FA02FA0303
        03FA0202020403030303030303F8FFF8030303F8FF0303F8FF03030303030303
        FA0303030303FA0202020403030303030303F80303030303F8FF0303F8FF0303
        0303030303030303030303FA0202020403030303030303030303030303F8FF03
        03F8FF03030303030303030303030303FA020202040303030303030303030303
        0303F8FF0303F8FF03030303030303030303030303FA02020204030303030303
        03030303030303F8FF0303F8FF03030303030303030303030303FA0202020403
        030303030303030303030303F8FF0303F8FF03030303030303030303030303FA
        0202040303030303030303030303030303F8FF03F8FF03030303030303030303
        03030303FA0202030303030303030303030303030303F8FFF803030303030303
        030303030303030303FA0303030303030303030303030303030303F803030303
        0303030303030303030303030303030303030303030303030303030303030303
        0303}
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BImprimer: THBitBtn
      Left = 388
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Imprimer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 7
      OnClick = BImprimerClick
      Margin = 2
      GlobalIndexImage = 'Z0369_S16G1'
    end
    object BInsert: THBitBtn
      Left = 152
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Nouveau'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      OnClick = BInsertClick
      GlobalIndexImage = 'Z0053_S16G1'
    end
    object BFerme: THBitBtn
      Left = 452
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 9
      OnClick = BFermeClick
      GlobalIndexImage = 'Z1770_S16G1'
    end
    object BFirst: THBitBtn
      Left = 4
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Premier'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BFirstClick
      GlobalIndexImage = 'Z0301_S16G1'
    end
    object BPrev: THBitBtn
      Left = 32
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Pr'#233'c'#233'dent'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BPrevClick
      GlobalIndexImage = 'Z0017_S16G1'
    end
    object BNext: THBitBtn
      Left = 60
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Suivant'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      OnClick = BNextClick
      GlobalIndexImage = 'Z0031_S16G1'
    end
    object BLast: THBitBtn
      Left = 88
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Dernier'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      OnClick = BLastClick
      GlobalIndexImage = 'Z0264_S16G1'
    end
  end
  object DBNav: TDBNavigator
    Left = 404
    Top = 0
    Width = 80
    Height = 18
    DataSource = SBudsect
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
    ConfirmDelete = False
    TabOrder = 2
    Visible = False
  end
  object SBudsect: TDataSource
    DataSet = QBudsect
    OnDataChange = SBudsectDataChange
    Left = 200
    Top = 196
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Sections budg'#233'taires;Voulez-vous enregistrer les modifications' +
        '?;Q;YNC;Y;C;'
      
        '1;Sections budg'#233'taires;Confirmez-vous la suppression de l'#39'enregi' +
        'strement?;Q;YNC;N;C;'
      '2;Sections budg'#233'taires;Vous devez renseigner le code !;W;O;O;O;'
      
        '3;Sections budg'#233'taires;Vous devez renseigner le libell'#233' !;W;O;O;' +
        'O;'
      
        '4;Sections budg'#233'taires;Le code que vous avez saisi existe d'#233'ja. ' +
        'Vous devez le modifier.;W;O;O;O;'
      'L'#39'enregistrement est inaccessible'
      
        '6;Sections budg'#233'taires;Cette section budg'#233'taire comporte des '#233'cr' +
        'itures budg'#233'taires, vous ne pouvez pas modifier l'#39'axe ;W;O;O;O;'
      'L'#39'acc'#232's '#224' ce compte  est interdit.'
      
        '8;Sections budg'#233'taires;ATTENTION : Votre saisie des comptes est ' +
        'trop longue, elle sera tronqu'#233'e.;E;O;O;O;'
      
        '9;Sections budg'#233'taires;Le longueur du code que vous avez saisi e' +
        'st inf'#233'rieure '#224' la longueur d'#233'finit pour cette soci'#233't'#233'.;W;O;O;O;'
      
        '10;Sections budg'#233'taires;Le longueur du code que vous avez saisi ' +
        'est sup'#233'rieure '#224' la longueur d'#233'finit pour cette soci'#233't'#233'.;W;O;O;O' +
        ';'
      
        '11;Sections budg'#233'taires;Vous devez renseigner le code rubrique !' +
        ';W;O;O;O;'
      
        '12;Sections budg'#233'taires;ATTENTION! Vous n'#39'avez pas associ'#233' de se' +
        'ctions analytiques '#224' votre section budg'#233'taire.D'#233'sirez vous conti' +
        'nuer?;Q;YNC;N;N;'
      
        '13;Sections budg'#233'taires;ATTENTION! Le code rubrique que vous ave' +
        'z renseign'#233' existe d'#233'j'#224'. La g'#233'n'#233'ration automatique des rubriques' +
        ' ne sera pas compl'#232'te.D'#233'sirez vous continuer?;Q;YNC;N;N;'
      
        '14;Sections budg'#233'taires;ATTENTION. Il y a des codes de tables li' +
        'bres qui n'#39'existent pas. D'#233'sirez vous continuer?;Q;YNC;N;C;')
    Left = 352
    Top = 65520
  end
  object QBudsect: TQuery
    AutoCalcFields = False
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    SQL.Strings = (
      'Select * From BUDSECT')
    dataBaseName = 'SOC'
    RequestLive = True
    UpdateMode = upWhereChanged
    Left = 348
    Top = 184
    object QBudsectBS_BUDSECT: TStringField
      FieldName = 'BS_BUDSECT'
      Size = 17
    end
    object QBudsectBS_LIBELLE: TStringField
      FieldName = 'BS_LIBELLE'
      Size = 35
    end
    object QBudsectBS_ABREGE: TStringField
      FieldName = 'BS_ABREGE'
      Size = 17
    end
    object QBudsectBS_SECTIONRUB: TStringField
      FieldName = 'BS_SECTIONRUB'
      Size = 250
    end
    object QBudsectBS_EXCLURUB: TStringField
      FieldName = 'BS_EXCLURUB'
      Size = 250
    end
    object QBudsectBS_SIGNE: TStringField
      FieldName = 'BS_SIGNE'
      Size = 3
    end
    object QBudsectBS_REPORTDISPO: TStringField
      FieldName = 'BS_REPORTDISPO'
      Size = 3
    end
    object QBudsectBS_PARENT: TStringField
      FieldName = 'BS_PARENT'
      Size = 17
    end
    object QBudsectBS_ATTENTE: TStringField
      FieldName = 'BS_ATTENTE'
      Size = 1
    end
    object QBudsectBS_FORMULE: TStringField
      FieldName = 'BS_FORMULE'
      Size = 200
    end
    object QBudsectBS_DATECREATION: TDateTimeField
      FieldName = 'BS_DATECREATION'
    end
    object QBudsectBS_DATEMODIF: TDateTimeField
      FieldName = 'BS_DATEMODIF'
    end
    object QBudsectBS_DATEOUVERTURE: TDateTimeField
      FieldName = 'BS_DATEOUVERTURE'
    end
    object QBudsectBS_DATEFERMETURE: TDateTimeField
      FieldName = 'BS_DATEFERMETURE'
    end
    object QBudsectBS_FERME: TStringField
      FieldName = 'BS_FERME'
      Size = 1
    end
    object QBudsectBS_SENS: TStringField
      FieldName = 'BS_SENS'
      Size = 3
    end
    object QBudsectBS_BLOCNOTE: TMemoField
      FieldName = 'BS_BLOCNOTE'
      BlobType = ftMemo
      Size = 1
    end
    object QBudsectBS_CONFIDENTIEL: TStringField
      FieldName = 'BS_CONFIDENTIEL'
      Size = 1
    end
    object QBudsectBS_HT: TStringField
      FieldName = 'BS_HT'
      Size = 1
    end
    object QBudsectBS_UTILISATEUR: TStringField
      FieldName = 'BS_UTILISATEUR'
      Size = 3
    end
    object QBudsectBS_SOCIETE: TStringField
      FieldName = 'BS_SOCIETE'
      Size = 3
    end
    object QBudsectBS_BLOQUANT: TStringField
      FieldName = 'BS_BLOQUANT'
      Size = 3
    end
    object QBudsectBS_CREERPAR: TStringField
      FieldName = 'BS_CREERPAR'
      Size = 3
    end
    object QBudsectBS_EXPORTE: TStringField
      FieldName = 'BS_EXPORTE'
      Size = 3
    end
    object QBudsectBS_TABLE0: TStringField
      FieldName = 'BS_TABLE0'
      Size = 17
    end
    object QBudsectBS_TABLE1: TStringField
      FieldName = 'BS_TABLE1'
      Size = 17
    end
    object QBudsectBS_TABLE2: TStringField
      FieldName = 'BS_TABLE2'
      Size = 17
    end
    object QBudsectBS_TABLE3: TStringField
      FieldName = 'BS_TABLE3'
      Size = 17
    end
    object QBudsectBS_TABLE4: TStringField
      FieldName = 'BS_TABLE4'
      Size = 17
    end
    object QBudsectBS_TABLE5: TStringField
      FieldName = 'BS_TABLE5'
      Size = 17
    end
    object QBudsectBS_TABLE6: TStringField
      FieldName = 'BS_TABLE6'
      Size = 17
    end
    object QBudsectBS_TABLE7: TStringField
      FieldName = 'BS_TABLE7'
      Size = 17
    end
    object QBudsectBS_TABLE8: TStringField
      FieldName = 'BS_TABLE8'
      Size = 17
    end
    object QBudsectBS_TABLE9: TStringField
      FieldName = 'BS_TABLE9'
      Size = 17
    end
    object QBudsectBS_AXE: TStringField
      FieldName = 'BS_AXE'
      Size = 3
    end
    object QBudsectBS_RUB: TStringField
      DisplayWidth = 6
      FieldName = 'BS_RUB'
      Size = 6
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    OnBeforeTraduc = HMTradBeforeTraduc
    Traduction = True
    Left = 384
    Top = 65520
  end
end
