object EtatPCL: TEtatPCL
  Left = 174
  Top = 111
  Width = 600
  Height = 399
  Caption = 'Etats libres'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 592
    Height = 300
    ActivePage = Standards
    Align = alClient
    ParentShowHint = False
    ShowHint = False
    TabOrder = 0
    object Standards: TTabSheet
      Caption = 'Standards'
      object LFile: THLabel
        Left = 5
        Top = 253
        Width = 45
        Height = 13
        Caption = 'Maquette'
        FocusControl = FMaq
        Contour = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Tag2 = 0
        Tag3 = 0
        TagExport = 0
      end
      object RechFile: TToolbarButton97
        Left = 267
        Top = 248
        Width = 17
        Height = 23
        Hint = 'Parcourir'
        Caption = '...'
        ParentShowHint = False
        ShowHint = True
        OnClick = RechFileClick
      end
      object GB1: TGroupBox
        Tag = 1
        Left = 3
        Top = 4
        Width = 578
        Height = 56
        TabOrder = 0
        object HLabel4: THLabel
          Left = 81
          Top = 26
          Width = 41
          Height = 13
          AutoSize = False
          Caption = '&Exercice'
          FocusControl = Exo1
          Contour = False
          Ombre = False
          OmbreDecalX = 0
          OmbreDecalY = 0
          OmbreColor = clBlack
          Tag2 = 0
          Tag3 = 0
          TagExport = 0
        end
        object HLabel6: THLabel
          Left = 231
          Top = 26
          Width = 43
          Height = 13
          Caption = '&Dates du'
          FocusControl = FD11
          Contour = False
          Ombre = False
          OmbreDecalX = 0
          OmbreDecalY = 0
          OmbreColor = clBlack
          Tag2 = 0
          Tag3 = 0
          TagExport = 0
        end
        object Label7: TLabel
          Left = 350
          Top = 26
          Width = 12
          Height = 13
          Caption = 'au'
          FocusControl = FD12
        end
        object Formule1: TEdit
          Left = 83
          Top = 22
          Width = 488
          Height = 21
          TabOrder = 6
          Visible = False
          OnEnter = Formule1Enter
          OnExit = Formule1Exit
        end
        object Exo1: THValComboBox
          Left = 125
          Top = 22
          Width = 100
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 13
          ParentCtl3D = False
          Sorted = False
          TabOrder = 2
          OnChange = Exo1Change
          Exhaustif = exOui
          Abrege = False
          Vide = False
          HideInGrid = False
          DataType = 'TTEXERCICE'
          DataTypeParametrable = False
          ComboWidth = 0
          DisableTab = False
        end
        object FD11: TMaskEdit
          Left = 280
          Top = 22
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 3
          Text = '  /  /    '
          OnKeyPress = FD11KeyPress
        end
        object FD12: TMaskEdit
          Left = 368
          Top = 22
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 4
          Text = '  /  /    '
          OnKeyPress = FD11KeyPress
        end
        object OkCol1: TCheckBox
          Tag = 1
          Left = 7
          Top = -1
          Width = 69
          Height = 17
          Caption = 'Colonne 1'
          Checked = True
          State = cbChecked
          TabOrder = 0
          OnClick = OkCol1Click
        end
        object Sit1: TCheckBox
          Left = 439
          Top = 24
          Width = 137
          Height = 17
          Caption = 'Sur balance de situation'
          TabOrder = 5
        end
        object Type1: THValComboBox
          Left = 6
          Top = 22
          Width = 70
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 13
          ParentCtl3D = False
          Sorted = False
          TabOrder = 1
          OnChange = Type1Change
          Items.Strings = (
            'Rubrique'
            'Formule')
          Exhaustif = exOui
          Abrege = False
          Vide = False
          HideInGrid = False
          DataTypeParametrable = False
          Values.Strings = (
            'RUB'
            'FOR')
          ComboWidth = 0
          DisableTab = False
        end
        object FValVar1: TSpinEdit
          Tag = 1
          Left = 128
          Top = 8
          Width = 41
          Height = 22
          Color = clYellow
          MaxValue = 0
          MinValue = 0
          TabOrder = 7
          Value = 0
          Visible = False
        end
      end
      object FAvecPourcent: TCheckBox
        Left = 307
        Top = 251
        Width = 110
        Height = 17
        Caption = 'Avec pourcentage'
        TabOrder = 5
      end
      object FAvecDetail: TCheckBox
        Left = 443
        Top = 251
        Width = 81
        Height = 17
        Caption = 'Voir le détail'
        TabOrder = 6
      end
      object FMaq: TEdit
        Left = 64
        Top = 249
        Width = 204
        Height = 21
        TabOrder = 4
      end
      object GB2: TGroupBox
        Tag = 2
        Left = 3
        Top = 65
        Width = 578
        Height = 56
        TabOrder = 1
        object HLabel1: THLabel
          Left = 81
          Top = 26
          Width = 56
          Height = 13
          AutoSize = False
          Caption = '&Exercice'
          FocusControl = Exo2
          Contour = False
          Ombre = False
          OmbreDecalX = 0
          OmbreDecalY = 0
          OmbreColor = clBlack
          Tag2 = 0
          Tag3 = 0
          TagExport = 0
        end
        object HLabel2: THLabel
          Left = 231
          Top = 26
          Width = 43
          Height = 13
          Caption = '&Dates du'
          FocusControl = FD21
          Contour = False
          Ombre = False
          OmbreDecalX = 0
          OmbreDecalY = 0
          OmbreColor = clBlack
          Tag2 = 0
          Tag3 = 0
          TagExport = 0
        end
        object Label1: TLabel
          Left = 350
          Top = 26
          Width = 12
          Height = 13
          Caption = 'au'
          FocusControl = FD22
        end
        object Formule2: TEdit
          Left = 83
          Top = 22
          Width = 488
          Height = 21
          TabOrder = 6
          Visible = False
          OnEnter = Formule1Enter
          OnExit = Formule1Exit
        end
        object Exo2: THValComboBox
          Left = 125
          Top = 22
          Width = 100
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 13
          ParentCtl3D = False
          Sorted = False
          TabOrder = 2
          OnChange = Exo1Change
          Exhaustif = exOui
          Abrege = False
          Vide = False
          HideInGrid = False
          DataType = 'TTEXERCICE'
          DataTypeParametrable = False
          ComboWidth = 0
          DisableTab = False
        end
        object FD21: TMaskEdit
          Left = 280
          Top = 22
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 3
          Text = '  /  /    '
          OnKeyPress = FD11KeyPress
        end
        object FD22: TMaskEdit
          Left = 368
          Top = 22
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 4
          Text = '  /  /    '
          OnKeyPress = FD11KeyPress
        end
        object OkCol2: TCheckBox
          Tag = 2
          Left = 7
          Top = -1
          Width = 69
          Height = 17
          Caption = 'Colonne 2'
          TabOrder = 0
          OnClick = OkCol1Click
        end
        object Sit2: TCheckBox
          Left = 439
          Top = 24
          Width = 136
          Height = 17
          Caption = 'Sur balance de situation'
          TabOrder = 5
        end
        object Type2: THValComboBox
          Left = 6
          Top = 22
          Width = 70
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 13
          ParentCtl3D = False
          Sorted = False
          TabOrder = 1
          OnChange = Type1Change
          Items.Strings = (
            'Rubrique'
            'Formule')
          Exhaustif = exOui
          Abrege = False
          Vide = False
          HideInGrid = False
          DataTypeParametrable = False
          Values.Strings = (
            'RUB'
            'FOR')
          ComboWidth = 0
          DisableTab = False
        end
        object FValVar2: TSpinEdit
          Tag = 2
          Left = 128
          Top = 8
          Width = 41
          Height = 22
          Color = clYellow
          MaxValue = 0
          MinValue = 0
          TabOrder = 7
          Value = 0
          Visible = False
        end
      end
      object GB3: TGroupBox
        Tag = 3
        Left = 3
        Top = 125
        Width = 578
        Height = 56
        TabOrder = 2
        object HLabel3: THLabel
          Left = 81
          Top = 26
          Width = 56
          Height = 13
          AutoSize = False
          Caption = '&Exercice'
          FocusControl = Exo3
          Contour = False
          Ombre = False
          OmbreDecalX = 0
          OmbreDecalY = 0
          OmbreColor = clBlack
          Tag2 = 0
          Tag3 = 0
          TagExport = 0
        end
        object HLabel9: THLabel
          Left = 231
          Top = 26
          Width = 43
          Height = 13
          Caption = '&Dates du'
          FocusControl = FD31
          Contour = False
          Ombre = False
          OmbreDecalX = 0
          OmbreDecalY = 0
          OmbreColor = clBlack
          Tag2 = 0
          Tag3 = 0
          TagExport = 0
        end
        object Label2: TLabel
          Left = 350
          Top = 26
          Width = 12
          Height = 13
          Caption = 'au'
          FocusControl = FD32
        end
        object Formule3: TEdit
          Left = 83
          Top = 22
          Width = 488
          Height = 21
          TabOrder = 6
          Visible = False
          OnEnter = Formule1Enter
          OnExit = Formule1Exit
        end
        object Exo3: THValComboBox
          Left = 125
          Top = 22
          Width = 100
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 13
          ParentCtl3D = False
          Sorted = False
          TabOrder = 2
          OnChange = Exo1Change
          Exhaustif = exOui
          Abrege = False
          Vide = False
          HideInGrid = False
          DataType = 'TTEXERCICE'
          DataTypeParametrable = False
          ComboWidth = 0
          DisableTab = False
        end
        object FD31: TMaskEdit
          Left = 280
          Top = 22
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 3
          Text = '  /  /    '
          OnKeyPress = FD11KeyPress
        end
        object FD32: TMaskEdit
          Left = 368
          Top = 22
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 4
          Text = '  /  /    '
          OnKeyPress = FD11KeyPress
        end
        object OkCol3: TCheckBox
          Tag = 3
          Left = 7
          Top = -1
          Width = 69
          Height = 17
          Caption = 'Colonne 3'
          TabOrder = 0
          OnClick = OkCol1Click
        end
        object Sit3: TCheckBox
          Left = 439
          Top = 24
          Width = 136
          Height = 17
          Caption = 'Sur balance de situation'
          TabOrder = 5
        end
        object Type3: THValComboBox
          Left = 6
          Top = 22
          Width = 70
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 13
          ParentCtl3D = False
          Sorted = False
          TabOrder = 1
          OnChange = Type1Change
          Items.Strings = (
            'Rubrique'
            'Formule')
          Exhaustif = exOui
          Abrege = False
          Vide = False
          HideInGrid = False
          DataTypeParametrable = False
          Values.Strings = (
            'RUB'
            'FOR')
          ComboWidth = 0
          DisableTab = False
        end
        object FValVar3: TSpinEdit
          Tag = 3
          Left = 128
          Top = 8
          Width = 41
          Height = 22
          Color = clYellow
          MaxValue = 0
          MinValue = 0
          TabOrder = 7
          Value = 0
          Visible = False
        end
      end
      object GB4: TGroupBox
        Tag = 4
        Left = 3
        Top = 185
        Width = 578
        Height = 56
        TabOrder = 3
        object HLabel10: THLabel
          Left = 81
          Top = 26
          Width = 56
          Height = 13
          AutoSize = False
          Caption = '&Exercice'
          FocusControl = Exo4
          Contour = False
          Ombre = False
          OmbreDecalX = 0
          OmbreDecalY = 0
          OmbreColor = clBlack
          Tag2 = 0
          Tag3 = 0
          TagExport = 0
        end
        object HLabel11: THLabel
          Left = 231
          Top = 26
          Width = 43
          Height = 13
          Caption = '&Dates du'
          FocusControl = FD41
          Contour = False
          Ombre = False
          OmbreDecalX = 0
          OmbreDecalY = 0
          OmbreColor = clBlack
          Tag2 = 0
          Tag3 = 0
          TagExport = 0
        end
        object Label3: TLabel
          Left = 350
          Top = 26
          Width = 12
          Height = 13
          Caption = 'au'
          FocusControl = FD42
        end
        object Formule4: TEdit
          Left = 83
          Top = 22
          Width = 488
          Height = 21
          TabOrder = 6
          Visible = False
          OnEnter = Formule1Enter
          OnExit = Formule1Exit
        end
        object Exo4: THValComboBox
          Left = 125
          Top = 22
          Width = 100
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 13
          ParentCtl3D = False
          Sorted = False
          TabOrder = 2
          OnChange = Exo1Change
          Exhaustif = exOui
          Abrege = False
          Vide = False
          HideInGrid = False
          DataType = 'TTEXERCICE'
          DataTypeParametrable = False
          ComboWidth = 0
          DisableTab = False
        end
        object FD41: TMaskEdit
          Left = 280
          Top = 22
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 3
          Text = '  /  /    '
          OnKeyPress = FD11KeyPress
        end
        object FD42: TMaskEdit
          Left = 368
          Top = 22
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 4
          Text = '  /  /    '
          OnKeyPress = FD11KeyPress
        end
        object OkCol4: TCheckBox
          Tag = 4
          Left = 7
          Top = -1
          Width = 69
          Height = 17
          Caption = 'Colonne 4'
          TabOrder = 0
          OnClick = OkCol1Click
        end
        object Sit4: TCheckBox
          Left = 439
          Top = 24
          Width = 136
          Height = 17
          Caption = 'Sur balance de situation'
          TabOrder = 5
        end
        object Type4: THValComboBox
          Left = 6
          Top = 22
          Width = 70
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 13
          ParentCtl3D = False
          Sorted = False
          TabOrder = 1
          OnChange = Type1Change
          Items.Strings = (
            'Rubrique'
            'Formule')
          Exhaustif = exOui
          Abrege = False
          Vide = False
          HideInGrid = False
          DataTypeParametrable = False
          Values.Strings = (
            'RUB'
            'FOR')
          ComboWidth = 0
          DisableTab = False
        end
        object FValVar4: TSpinEdit
          Tag = 4
          Left = 128
          Top = 8
          Width = 41
          Height = 22
          Color = clYellow
          MaxValue = 0
          MinValue = 0
          TabOrder = 7
          Value = 0
          Visible = False
        end
      end
    end
    object Complements: TTabSheet
      Caption = 'Compléments'
      ImageIndex = 1
      object HLabel7: THLabel
        Left = 9
        Top = 68
        Width = 111
        Height = 13
        AutoSize = False
        Caption = '&Etablissements'
        FocusControl = FEtab
        Contour = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Tag2 = 0
        Tag3 = 0
        TagExport = 0
      end
      object HLabel8: THLabel
        Left = 9
        Top = 122
        Width = 111
        Height = 13
        AutoSize = False
        Caption = '&Devises'
        FocusControl = FDevises
        Contour = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Tag2 = 0
        Tag3 = 0
        TagExport = 0
      end
      object TFRESOL: THLabel
        Left = 9
        Top = 15
        Width = 59
        Height = 13
        Caption = 'Rés&olution...'
        FocusControl = FRESOL
        Contour = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Tag2 = 0
        Tag3 = 0
        TagExport = 0
      end
      object HLabel5: THLabel
        Left = 367
        Top = 15
        Width = 73
        Height = 13
        Caption = 'Format &montant'
        FocusControl = FFormat
        Contour = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Tag2 = 0
        Tag3 = 0
        TagExport = 0
      end
      object FDevises: THValComboBox
        Left = 149
        Top = 118
        Width = 164
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        Sorted = False
        TabOrder = 3
        Exhaustif = exOui
        Abrege = False
        Vide = True
        HideInGrid = False
        DataType = 'TTDEVISE'
        DataTypeParametrable = False
        ComboWidth = 0
        DisableTab = False
      end
      object FEtab: THValComboBox
        Left = 149
        Top = 64
        Width = 163
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        Sorted = False
        TabOrder = 2
        Exhaustif = exOui
        Abrege = False
        Vide = True
        HideInGrid = False
        DataType = 'TTETABLISSEMENT'
        DataTypeParametrable = False
        ComboWidth = 0
        DisableTab = False
      end
      object FAffiche: TRadioGroup
        Left = 328
        Top = 107
        Width = 200
        Height = 43
        Caption = ' Affichage en... '
        Columns = 3
        ItemIndex = 0
        Items.Strings = (
          '&Franc'
          '&Devise'
          '&Euro')
        TabOrder = 4
      end
      object FAvec: TGroupBox
        Left = 9
        Top = 176
        Width = 304
        Height = 56
        Caption = 'Intégrer les écritures de ... '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        object FSimu: TCheckBox
          Left = 24
          Top = 24
          Width = 81
          Height = 17
          Caption = 'Simulation'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object FSitu: TCheckBox
          Left = 120
          Top = 24
          Width = 81
          Height = 17
          Caption = 'Situation'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object FRevi: TCheckBox
          Left = 216
          Top = 24
          Width = 81
          Height = 17
          Caption = 'Révision'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
      end
      object FVoirNom: TCheckBox
        Left = 11
        Top = 243
        Width = 134
        Height = 17
        Caption = 'Voir le nom des cellules'
        TabOrder = 6
      end
      object FRESOL: THValComboBox
        Left = 149
        Top = 11
        Width = 163
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        Sorted = False
        TabOrder = 0
        Items.Strings = (
          'Avec décimales'
          'Sans décimales'
          'en K(ilo)'
          'en M(éga)')
        Exhaustif = exOui
        Abrege = False
        Vide = False
        HideInGrid = False
        DataTypeParametrable = False
        Values.Strings = (
          'C'
          'F'
          'K'
          'M')
        ComboWidth = 0
        DisableTab = False
      end
      object FFormat: TEdit
        Left = 453
        Top = 11
        Width = 120
        Height = 21
        TabOrder = 1
      end
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 300
    Width = 592
    Height = 72
    AllowDrag = False
    Position = dpBottom
    object PanelFiltre: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 588
      Caption = 'Filtres'
      ClientAreaHeight = 32
      ClientAreaWidth = 588
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BFiltre: TToolbarButton97
        Left = 6
        Top = 6
        Width = 56
        Height = 21
        Hint = 'Menu filtre'
        DropdownMenu = POPF
        Caption = '&Filtres'
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object FFiltres: TComboBox
        Left = 72
        Top = 6
        Width = 465
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = FFiltresChange
      end
    end
    object HPB: TToolWindow97
      Left = 0
      Top = 36
      ClientHeight = 32
      ClientWidth = 588
      Caption = 'Barre d'#39'outils'
      ClientAreaHeight = 32
      ClientAreaWidth = 588
      DockPos = 0
      DockRow = 1
      FullSize = True
      TabOrder = 1
      object BParamListe: TToolbarButton97
        Left = 6
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Paramétrer liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Paramétrer liste'
        Flat = False
        Glyph.Data = {
          F6000000424DF600000000000000760000002800000010000000100000000100
          0400000000008000000000000000000000001000000010000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          777777777777077777777000000060000007770EEEE0660EE0777770EE0660EE
          077777770EE00EE07777777770EEEE0777777777060EE0607777777706600660
          77777770666600607777777066660B0077777777006660507777777777000005
          0777777777777770507777777777777705077777777777777007}
        ParentShowHint = False
        ShowHint = True
        OnClick = BParamListeClick
      end
      object BValider: TToolbarButton97
        Left = 494
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Lancer l'#39'état'
        DisplayMode = dmGlyphOnly
        Caption = 'OK'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
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
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 526
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Annuler'
        Flat = False
        Glyph.Data = {
          CE070000424DCE07000000000000360000002800000024000000120000000100
          1800000000009807000000000000000000000000000000000000008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284008284008284008284
          0082840082840082840082840082840082848482848482840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284FFFFFF008284008284008284008284008284008284
          0082840082840082840082840082840082840082840082840082840000FF0000
          840000848482840082840082840082840082840082840000FF84828400828400
          8284008284008284008284008284008284008284848284848284FFFFFF008284
          008284008284008284008284008284FFFFFF0082840082840082840082840082
          840082840082840000FF00008400008400008484828400828400828400828400
          00FF000084000084848284008284008284008284008284008284008284848284
          FFFFFF008284848284FFFFFF008284008284008284FFFFFF848284848284FFFF
          FF0082840082840082840082840082840082840000FF00008400008400008400
          00848482840082840000FF000084000084000084000084848284008284008284
          008284008284008284848284FFFFFF008284008284848284FFFFFF008284FFFF
          FF848284008284008284848284FFFFFF00828400828400828400828400828400
          82840000FF000084000084000084000084848284000084000084000084000084
          000084848284008284008284008284008284008284848284FFFFFF0082840082
          84008284848284FFFFFF848284008284008284008284008284848284FFFFFF00
          82840082840082840082840082840082840000FF000084000084000084000084
          0000840000840000840000848482840082840082840082840082840082840082
          84008284848284FFFFFF00828400828400828484828400828400828400828400
          8284FFFFFF848284008284008284008284008284008284008284008284008284
          0000FF0000840000840000840000840000840000848482840082840082840082
          84008284008284008284008284008284008284848284FFFFFF00828400828400
          8284008284008284008284FFFFFF848284008284008284008284008284008284
          0082840082840082840082840082840000840000840000840000840000848482
          8400828400828400828400828400828400828400828400828400828400828400
          8284848284FFFFFF008284008284008284008284008284848284008284008284
          0082840082840082840082840082840082840082840082840082840000FF0000
          8400008400008400008484828400828400828400828400828400828400828400
          8284008284008284008284008284008284848284FFFFFF008284008284008284
          8482840082840082840082840082840082840082840082840082840082840082
          840082840000FF00008400008400008400008400008484828400828400828400
          8284008284008284008284008284008284008284008284008284008284848284
          008284008284008284008284848284FFFFFF0082840082840082840082840082
          840082840082840082840082840000FF00008400008400008484828400008400
          0084000084848284008284008284008284008284008284008284008284008284
          008284008284848284008284008284008284008284008284848284FFFFFF0082
          840082840082840082840082840082840082840082840000FF00008400008400
          00848482840082840000FF000084000084000084848284008284008284008284
          008284008284008284008284008284848284008284008284008284848284FFFF
          FF008284008284848284FFFFFF00828400828400828400828400828400828400
          82840000FF0000840000848482840082840082840082840000FF000084000084
          000084848284008284008284008284008284008284008284848284FFFFFF0082
          84008284848284008284848284FFFFFF008284008284848284FFFFFF00828400
          82840082840082840082840082840082840000FF000084008284008284008284
          0082840082840000FF0000840000840000840082840082840082840082840082
          84008284848284FFFFFFFFFFFF848284008284008284008284848284FFFFFF00
          8284008284848284FFFFFF008284008284008284008284008284008284008284
          0082840082840082840082840082840082840082840000FF0000840000FF0082
          8400828400828400828400828400828400828484828484828400828400828400
          8284008284008284848284FFFFFFFFFFFFFFFFFF848284008284008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284848284848284848284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284008284008284008284
          008284008284008284008284008284008284}
        ModalResult = 2
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
      end
      object BAide: TToolbarButton97
        Left = 558
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
        Glyph.Data = {
          CE070000424DCE07000000000000360000002800000024000000120000000100
          1800000000009807000000000000000000000000000000000000008080008080
          0080800080800080800080800080800080808080008080000080800080800080
          8000808000808000808000808000808000808000808000808000808000808000
          8080008080F8FCF8F8FCF8008080008080008080008080008080008080008080
          0080800080800080800080800080800080800080800080800080808080008000
          0080000080800000808000808000808000808000808000808000808000808000
          8080008080008080008080008080008080808080808080F8FCF8008080008080
          0080800080800080800080800080800080800080800080800080800080800080
          80008080008080F8FC0080800080800080000000808000808000808000808000
          8080008080008080008080008080008080008080008080008080808080F8FCF8
          008080808080F8FCF80080800080800080800080800080800080800080800080
          80008080008080008080008080008080008080008080F8FC0080800080800000
          8080008080008080008080008080008080008080008080008080008080008080
          008080008080808080F8FCF8F8FCF8808080F8FCF80080800080800080800080
          8000808000808000808000808000808000808000808000808000808000808000
          8080008080008080008080008080008080008080008080008080008080008080
          0080800080800080800080800080800080800080808080808080800080800080
          8000808000808000808000808000808000808000808000808000808000808000
          8080008080008080008080008080800000800000808000008080008080008080
          0080800080800080800080800080800080800080800080800080800080800080
          80008080F8FCF8F8FCF800808000808000808000808000808000808000808000
          8080008080008080008080008080008080008080008080808000808000808000
          8000000080800080800080800080800080800080800080800080800080800080
          80008080008080008080008080808080808080808080F8FCF800808000808000
          8080008080008080008080008080008080008080008080008080008080008080
          008080F8FC008080008080008000000080800080800080800080800080800080
          80008080008080008080008080008080008080008080808080F8FCF800808080
          8080F8FCF8008080008080008080008080008080008080008080008080008080
          008080008080008080008080008080F8FC008080008080008080008000000080
          8000808000808000808000808000808000808000808000808000808000808000
          8080808080F8FCF8008080808080F8FCF8008080008080008080008080008080
          008080008080008080008080008080008080008080008080008080008080F8FC
          0080800080800080800080000000808000808000808000808000808000808000
          8080008080008080008080008080808080F8FCF8008080008080808080F8FCF8
          0080800080800080800080800080800080800080800080800080800080800080
          80008080008080008080008080F8FC0080800080800080800080000000808000
          8080008080008080008080008080008080008080008080008080008080808080
          F8FCF8008080008080808080F8FCF80080800080800080800080800080800080
          80008080008080008080008080800000800000008080008080008080F8FC0080
          8000808000808000800000008080008080008080008080008080008080008080
          008080F8FCF8008080008080808080F8FCF8008080008080808080F8FCF80080
          8000808000808000808000808000808000808000808080800080800080800080
          0000008080008080008080F8FC00808000808000800000008080008080008080
          008080008080008080008080808080808080F8FCF8008080008080808080F8FC
          F8008080008080808080F8FCF800808000808000808000808000808000808000
          8080F8FC00808000808000808000800000800000800000808000808000808000
          800000008080008080008080008080008080008080808080F8FCF80080808080
          80F8FCF8F8FCF8F8FCF8808080008080008080808080F8FCF800808000808000
          8080008080008080008080008080008080F8FC00808000808000808000808000
          8080008080008080008080008080000080800080800080800080800080800080
          80808080F8FCF800808000808080808080808080808000808000808000808080
          8080F8FCF8008080008080008080008080008080008080008080008080008080
          F8FC00F8FC008080008080008080008080008080008080000080800080800080
          80008080008080008080008080008080808080F8FCF8F8FCF800808000808000
          8080008080008080008080808080008080008080008080008080008080008080
          008080008080008080008080008080008080F8FC00F8FC00F8FC00F8FC00F8FC
          0000808000808000808000808000808000808000808000808000808000808080
          8080808080F8FCF8F8FCF8F8FCF8F8FCF8F8FCF8808080008080008080008080
          0080800080800080800080800080800080800080800080800080800080800080
          8000808000808000808000808000808000808000808000808000808000808000
          8080008080008080008080008080008080808080808080808080808080808080
          008080008080008080008080008080008080}
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        IsControl = True
      end
    end
  end
  object FVariation: TRadioGroup
    Left = 232
    Top = 256
    Width = 185
    Height = 105
    Caption = 'Variation'
    Items.Strings = (
      'Simple'
      'Sur 12 mois'
      'Ramené au plus petit'
      'Ramené au plus grand')
    TabOrder = 2
    Visible = False
    OnClick = FVariationClick
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Répartition des échéances;Erreur de répartition : le montant n' +
        #39'est pas réparti complètement.;W;O;O;O;'
      
        '1;Répartition des échéances;Le montant est trop élevé pour ce mo' +
        'de de paiement.Voulez-vous passer au mode de remplacement ?;Q;YN' +
        ';Y;N;'
      
        '2;Répartition des échéances;Vous ne pouvez pas renseigner des mo' +
        'ntants négatifs.;W;O;O;O;'
      
        '3;Répartition des échéances;Les dates d'#39'échéance doivent respect' +
        'er la plage de saisie autorisée;W;O;O;O;')
    Left = 496
    Top = 240
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 400
    Top = 240
  end
  object POPF: TPopupMenu
    Left = 448
    Top = 240
    object BCreerFiltre: TMenuItem
      Caption = '&Créer un filtre'
      OnClick = BCreerFiltreClick
    end
    object BSaveFiltre: TMenuItem
      Caption = '&Enregistrer le filtre'
      OnClick = BSaveFiltreClick
    end
    object BDelFiltre: TMenuItem
      Caption = '&Supprimer le filtre'
      OnClick = BDelFiltreClick
    end
    object BRenFiltre: TMenuItem
      Caption = '&Renommer le filtre'
      OnClick = BRenFiltreClick
    end
    object BNouvRech: TMenuItem
      Caption = '&Nouvelle recherche'
      OnClick = BNouvRechClick
    end
  end
  object Sauve: TSaveDialog
    Filter = 'Fichiers textes (*.txt).|*.txt|Tous fichiers (*.*)|*.*'
    Options = [ofHideReadOnly, ofShowHelp, ofExtensionDifferent, ofPathMustExist, ofNoLongNames]
    Title = 'Exportation des mouvements'
    Left = 352
    Top = 240
  end
end
