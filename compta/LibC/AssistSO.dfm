inherited FAssistSoc: TFAssistSoc
  Left = 236
  Top = 160
  Caption = 'Assistant de param'#233'trage de soci'#233't'#233
  ClientHeight = 313
  ClientWidth = 514
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Left = 52
    Top = 289
    Width = 79
  end
  inherited lAide: THLabel
    Left = 139
    Top = 6
    Width = 35
    Height = 23
    Visible = False
  end
  object ip1: TToolbarButton97 [2]
    Left = 12
    Top = 9
    Width = 29
    Height = 24
    Glyph.Data = {
      4E010000424D4E01000000000000760000002800000012000000120000000100
      040000000000D800000000000000000000001000000010000000000000000000
      80000080000000808000800000008000800080800000C0C0C000808080000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
      7777770000007777777007777777770000007777770880777777770000007777
      7708807777007700000077777770807770880700000077770077077770880700
      0000777088077777770807000000770888807770077077000000708888077708
      8077770000007088807770888077770000007088807708888077770000007088
      0770888807777700000077007770888807777700000077777770888077777700
      0000777777770007777777000000777777777777777777000000777777777777
      777777000000777777777777777777000000}
    GlyphMask.Data = {00000000}
    Visible = False
  end
  object ip2: TToolbarButton97 [3]
    Left = 54
    Top = 10
    Width = 23
    Height = 23
    Glyph.Data = {
      4E010000424D4E01000000000000760000002800000012000000120000000100
      040000000000D800000000000000000000001000000010000000000000000000
      80000080000000808000800000008000800080800000C0C0C000808080000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
      7777770000007777777007777777770000007777770880777777770000007777
      7708807777007700000077777770807770880700000077770077077770880700
      0000777088077777770807000000770888807770077077000000708888077708
      8077770000007088807770888077770000007088807708888077770000007088
      0770888807777700000077007770888807777700000077777770888077177700
      0000777777770007719177000000777777777777797917000000777777777777
      777797000000777777777777777777000000}
    GlyphMask.Data = {00000000}
    Visible = False
  end
  object bStep: TToolbarButton97 [4]
    Left = 5
    Top = 285
    Width = 39
    Height = 23
    DropdownArrow = True
    DropdownMenu = StepMenu
    Glyph.Data = {
      4E010000424D4E01000000000000760000002800000012000000120000000100
      040000000000D800000000000000000000001000000010000000000000000000
      80000080000000808000800000008000800080800000C0C0C000808080000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
      7777770000007777777007777777770000007777770880777777770000007777
      7708807777007700000077777770807770880700000077770077077770880700
      0000777088077777770807000000770888807770077077000000708888077708
      8077770000007088807770888077770000007088807708888077770000007088
      0770888807777700000077007770888807777700000077777770888077777700
      0000777777770007777777000000777777777777777777000000777777777777
      777777000000777777777777777777000000}
    GlyphMask.Data = {00000000}
    Opaque = False
  end
  inherited bPrecedent: TToolbarButton97
    Top = 285
  end
  inherited bSuivant: TToolbarButton97
    Top = 285
  end
  inherited bFin: TToolbarButton97
    Top = 285
  end
  inherited bAnnuler: TToolbarButton97
    Top = 285
  end
  inherited bAide: TToolbarButton97
    Top = 313
  end
  inherited Plan: THPanel
    Left = 173
    Top = 1
    Height = 252
  end
  inherited GroupBox1: TGroupBox
    Top = 271
  end
  inherited P: TPageControl
    Left = 171
    Top = 2
    Width = 338
    Height = 275
    ActivePage = Analytique
    object Welcome: TTabSheet
      Hint = 'Coordonn'#233'es de la soci'#233't'#233
      Caption = 'Welcome'
      object HLabel1: THLabel
        Left = 17
        Top = 29
        Width = 311
        Height = 13
        Caption = 'Bienvenue dans l'#39'assistant de param'#233'trage de soci'#233't'#233'.'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object TSO_LIBELLE: THLabel
        Left = 29
        Top = 121
        Width = 69
        Height = 13
        Caption = 'Raison sociale'
        FocusControl = SO_LIBELLE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel2: THLabel
        Left = 29
        Top = 96
        Width = 62
        Height = 13
        Caption = 'Code soci'#233't'#233
        FocusControl = SO_LIBELLE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel3: THLabel
        Left = 25
        Top = 63
        Width = 237
        Height = 13
        AutoSize = False
        Caption = 'Indiquez le code et la raison sociale de la soci'#233't'#233' :'
      end
      object bSoc1: TToolbarButton97
        Left = 70
        Top = 146
        Width = 195
        Height = 29
        Caption = 'Coordonn'#233'es de la soci'#233't'#233
        Glyph.Data = {
          66010000424D6601000000000000760000002800000014000000140000000100
          040000000000F000000000000000000000001000000010000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          3333333300003333333333333333333300003333333333333333333300003333
          3333333333333333000033800000000000003333000033007A7A7A7A7A7A0333
          0000330F07A7A7A7A7A703330000330A0A7A7A7A7A7A70330000330FA0A7A7A7
          A7A7A0330000330AF07A7A7A7A7A7A030000330FAF000007A7A7A7030000330A
          FAFAFAF0000000330000330FAFAFAFAFAFA033330000330AFAFAFAFAFAF03333
          0000330FAFAF0000000333330000333000003333333333330000333333333333
          3333333300003333333333333333333300003333333333333333333300003333
          33333333333333330000}
        GlyphMask.Data = {00000000}
        Opaque = False
        ParentShowHint = False
        ShowHint = False
        OnClick = bSoc1Click
      end
      object HDossierEuro: THLabel
        Left = 25
        Top = 195
        Width = 292
        Height = 13
        AutoSize = False
        Caption = 'Indiquez si vous souhaitez tenir le dossier en Euro'
      end
      object SO_LIBELLE: TEdit
        Left = 107
        Top = 117
        Width = 205
        Height = 21
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
      end
      object SO_CODE: TEdit
        Left = 107
        Top = 92
        Width = 53
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        MaxLength = 3
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
      end
      object GroupBox2: TGroupBox
        Left = 17
        Top = 44
        Width = 311
        Height = 7
        TabOrder = 2
      end
      object DossierEuro: TCheckBox
        Left = 24
        Top = 216
        Width = 273
        Height = 17
        Caption = 'Euro comme monnaie de tenue de la comptabilit'#233
        TabOrder = 3
      end
    end
    object PlanRef: TTabSheet
      Hint = 'Cr'#233'ation du plan comptable'
      Caption = 'PlanRef'
      object bCreerPlanComptable: TToolbarButton97
        Left = 92
        Top = 220
        Width = 153
        Height = 25
        Caption = '&Cr'#233'er le plan comptable'
        Enabled = False
        Opaque = False
        OnClick = bCreerPlanComptableClick
        GlobalIndexImage = 'Z0104_S16G1'
      end
      object HLabel6: THLabel
        Left = 205
        Top = 34
        Width = 84
        Height = 13
        Caption = 'Plan de r'#233'f'#233'rence'
      end
      object HLabel4: THLabel
        Left = 87
        Top = 5
        Width = 159
        Height = 16
        AutoSize = False
        Caption = 'Cr'#233'ation du plan comptable'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object bSoc2: TToolbarButton97
        Left = 9
        Top = 29
        Width = 153
        Height = 25
        Caption = '&Param'#232'tres comptables...'
        Opaque = False
        OnClick = bSoc2Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object SO_NUMPLANREF: TSpinEdit
        Left = 294
        Top = 30
        Width = 35
        Height = 22
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxValue = 10
        MinValue = 0
        ParentFont = False
        TabOrder = 0
        Value = 0
        OnChange = SO_NUMPLANREFChange
      end
      object GPlanRef: THDBGrid
        Left = 9
        Top = 60
        Width = 320
        Height = 158
        DataSource = SPlanRef
        Options = [dgColLines, dgRowLines, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
        OnDblClick = GPlanRefDblClick
        Row = 0
        MultiSelection = False
        SortEnabled = False
        MyDefaultRowHeight = 0
        Columns = <
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'PR_REPORTDETAIL'
            Title.Alignment = taCenter
            Title.Caption = '*'
            Width = 15
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PR_COMPTE'
            ReadOnly = True
            Title.Caption = 'Compte'
            Width = 73
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PR_LIBELLE'
            ReadOnly = True
            Title.Caption = 'Libell'#233
            Width = 211
            Visible = True
          end>
      end
      object GroupBox3: TGroupBox
        Left = 87
        Top = 18
        Width = 159
        Height = 7
        TabOrder = 2
      end
    end
    object ParamCPT: TTabSheet
      Hint = 'Param'#232'tres comptables'
      Caption = 'ParamCPT'
      object bSoc3: TToolbarButton97
        Left = 76
        Top = 89
        Width = 190
        Height = 25
        Caption = 'Fourchettes de comptes...'
        Margin = 4
        Opaque = False
        OnClick = bSoc3Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bSoc4: TToolbarButton97
        Left = 76
        Top = 120
        Width = 190
        Height = 25
        Caption = 'D'#233'finition des comptes sp'#233'ciaux...'
        Margin = 4
        Opaque = False
        OnClick = bSoc4Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bSoc5: TToolbarButton97
        Left = 76
        Top = 152
        Width = 190
        Height = 25
        Caption = 'Autres param'#232'tres comptables...'
        Margin = 4
        Opaque = False
        OnClick = bSoc5Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object HLabel8: THLabel
        Left = 93
        Top = 47
        Width = 135
        Height = 16
        AutoSize = False
        Caption = 'Param'#232'tres comptables'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object bSoc7: TToolbarButton97
        Left = 76
        Top = 183
        Width = 190
        Height = 25
        Caption = 'Comptes collectifs par d'#233'faut...'
        Margin = 4
        Opaque = False
        OnClick = bSoc7Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object GroupBox4: TGroupBox
        Left = 93
        Top = 60
        Width = 135
        Height = 7
        TabOrder = 0
      end
    end
    object Users: TTabSheet
      Hint = 'Gestion des utilisateurs'
      Caption = 'Users'
      object HLabel5: THLabel
        Left = 88
        Top = 56
        Width = 133
        Height = 13
        Caption = 'Gestion des utilisateurs'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object bUsers: TToolbarButton97
        Left = 67
        Top = 121
        Width = 200
        Height = 25
        Caption = 'Cr'#233'ation des utilisateurs...'
        Margin = 5
        Opaque = False
        OnClick = bUsersClick
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bUserGroups: TToolbarButton97
        Left = 67
        Top = 89
        Width = 200
        Height = 25
        Caption = 'Cr'#233'ation des groupes d'#39'utilisateurs...'
        Margin = 5
        Opaque = False
        OnClick = bUserGroupsClick
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object GroupBox5: TGroupBox
        Left = 88
        Top = 68
        Width = 133
        Height = 7
        TabOrder = 0
      end
    end
    object EtabExo: TTabSheet
      Hint = 'Etablissements, exercices, devises...'
      Caption = 'EtabExo'
      object HLabel7: THLabel
        Left = 59
        Top = 22
        Width = 210
        Height = 13
        Caption = 'Etablissements, exercices, devises...'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object bEtab: TToolbarButton97
        Left = 75
        Top = 60
        Width = 180
        Height = 25
        Caption = 'Cr'#233'ation des '#233'tablissements...'
        Margin = 5
        Opaque = False
        OnClick = bEtabClick
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bExos: TToolbarButton97
        Left = 75
        Top = 93
        Width = 180
        Height = 25
        Caption = 'Cr'#233'ation des exercices...'
        Margin = 5
        Opaque = False
        OnClick = bExosClick
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bDevises: TToolbarButton97
        Left = 75
        Top = 125
        Width = 180
        Height = 25
        Caption = 'Choix des devises...'
        Margin = 5
        Opaque = False
        OnClick = bDevisesClick
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bUnites: TToolbarButton97
        Left = 75
        Top = 190
        Width = 180
        Height = 25
        Caption = 'Unit'#233's de mesure...'
        Margin = 5
        Opaque = False
        OnClick = bUnitesClick
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bSoc6: TToolbarButton97
        Left = 75
        Top = 158
        Width = 180
        Height = 25
        Caption = 'Param'#232'tres soci'#233't'#233' li'#233's...'
        Margin = 5
        Opaque = False
        OnClick = bSoc6Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object GroupBox6: TGroupBox
        Left = 59
        Top = 35
        Width = 210
        Height = 7
        TabOrder = 0
      end
    end
    object Reglements: TTabSheet
      Hint = 'Param'#232'tres des r'#232'glements'
      Caption = 'Reglements'
      object HLabel9: THLabel
        Left = 82
        Top = 55
        Width = 153
        Height = 13
        Caption = 'Param'#232'tres des r'#232'glements'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object bModePaiement: TToolbarButton97
        Left = 86
        Top = 100
        Width = 148
        Height = 25
        Caption = 'Modes de paiement'
        Margin = 5
        Opaque = False
        OnClick = bModePaiementClick
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bModeRegl: TToolbarButton97
        Left = 85
        Top = 132
        Width = 148
        Height = 25
        Caption = 'Conditions de r'#232'glement'
        Margin = 5
        Opaque = False
        OnClick = bModeReglClick
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object GroupBox7: TGroupBox
        Left = 82
        Top = 68
        Width = 153
        Height = 7
        TabOrder = 0
      end
    end
    object Analytique: TTabSheet
      Hint = 'Param'#233'trage de la comptabilit'#233' analytique'
      Caption = 'Analytique'
      object HLabel10: THLabel
        Left = 50
        Top = 10
        Width = 238
        Height = 13
        Caption = 'Param'#233'trage de la comptabilit'#233' analytique'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object bAna1: TToolbarButton97
        Left = 60
        Top = 38
        Width = 220
        Height = 25
        Caption = 'Personnalisation des axes analytiques...'
        Margin = 5
        Opaque = False
        OnClick = bAna1Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bAna2: TToolbarButton97
        Left = 60
        Top = 67
        Width = 220
        Height = 25
        Caption = 'Structure des sections analytiques...'
        Margin = 5
        Opaque = False
        OnClick = bAna2Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bAna3: TToolbarButton97
        Left = 60
        Top = 181
        Width = 220
        Height = 25
        DropdownArrow = True
        DropdownMenu = PopAxes
        Caption = 'Plan de correspondance analytique...'
        Margin = 5
        Opaque = False
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bAna4: TToolbarButton97
        Left = 60
        Top = 95
        Width = 220
        Height = 25
        Caption = 'Sections analytiques...'
        Margin = 5
        Opaque = False
        OnClick = bAna4Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bAna5: TToolbarButton97
        Left = 60
        Top = 124
        Width = 220
        Height = 25
        Caption = 'Ventilations analytiques...'
        Margin = 5
        Opaque = False
        OnClick = bAna5Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bAna6: TToolbarButton97
        Left = 60
        Top = 209
        Width = 220
        Height = 25
        Caption = 'Cl'#233's de r'#233'partitions inter-sections...'
        Margin = 5
        Opaque = False
        OnClick = bAna6Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object nAna7: TToolbarButton97
        Tag = 1
        Left = 60
        Top = 152
        Width = 220
        Height = 25
        DropdownArrow = True
        DropdownMenu = PopAxRupt
        Caption = 'Plan de rupture analytique...'
        Margin = 5
        Opaque = False
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object GroupBox8: TGroupBox
        Left = 50
        Top = 23
        Width = 238
        Height = 7
        TabOrder = 0
      end
    end
    object PlanCpt: TTabSheet
      Hint = 'Param'#233'trage du plan comptable g'#233'n'#233'ral'
      Caption = 'PlanCpt'
      object HLabel11: THLabel
        Left = 53
        Top = 36
        Width = 226
        Height = 13
        Caption = 'Param'#233'trage du plan comptable g'#233'n'#233'ral'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object bCpt1: TToolbarButton97
        Left = 72
        Top = 142
        Width = 192
        Height = 25
        Caption = 'Plan de correspondance g'#233'n'#233'ral...'
        Margin = 5
        Opaque = False
        OnClick = bCpt1Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bCpt2: TToolbarButton97
        Left = 72
        Top = 75
        Width = 192
        Height = 23
        Caption = 'Plan comptable g'#233'n'#233'ral...'
        Margin = 5
        Opaque = False
        OnClick = bCpt2Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bCpt3: TToolbarButton97
        Left = 72
        Top = 109
        Width = 192
        Height = 25
        Caption = 'Plan de rupture g'#233'n'#233'ral...'
        Margin = 5
        Opaque = False
        OnClick = bCpt3Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object GroupBox9: TGroupBox
        Left = 53
        Top = 49
        Width = 226
        Height = 7
        TabOrder = 0
      end
    end
    object TVA: TTabSheet
      Hint = 'Param'#233'trage des taxes'
      Caption = 'TVA'
      object HLabel12: THLabel
        Left = 93
        Top = 33
        Width = 130
        Height = 13
        Caption = 'Param'#233'trage des taxes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object bTaxes1: TToolbarButton97
        Left = 99
        Top = 75
        Width = 120
        Height = 25
        Caption = 'R'#233'gimes fiscaux...'
        Margin = 5
        Opaque = False
        OnClick = bTaxes1Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bTaxes2: TToolbarButton97
        Left = 99
        Top = 112
        Width = 120
        Height = 23
        Caption = 'Codes de TVA...'
        Margin = 5
        Opaque = False
        OnClick = bTaxes2Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bTaxes3: TToolbarButton97
        Left = 99
        Top = 148
        Width = 120
        Height = 25
        Caption = 'Codes de TPF...'
        Margin = 5
        Opaque = False
        OnClick = bTaxes3Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object GroupBox10: TGroupBox
        Left = 93
        Top = 46
        Width = 130
        Height = 7
        TabOrder = 0
      end
    end
    object Tiers: TTabSheet
      Hint = 'Param'#233'trage des tiers'
      Caption = 'Tiers'
      object bTiers1: TToolbarButton97
        Left = 59
        Top = 158
        Width = 222
        Height = 25
        Caption = 'Formes juridiques...'
        Margin = 5
        Opaque = False
        OnClick = bTiers1Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bTiers2: TToolbarButton97
        Left = 59
        Top = 133
        Width = 222
        Height = 25
        Caption = 'Fonctions...'
        Margin = 5
        Opaque = False
        OnClick = bTiers2Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bTiers3: TToolbarButton97
        Left = 59
        Top = 59
        Width = 222
        Height = 25
        Caption = 'Param'#232'tres de relances des traites...'
        Margin = 5
        Opaque = False
        OnClick = bTiers3Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bTiers4: TToolbarButton97
        Left = 59
        Top = 83
        Width = 222
        Height = 25
        Caption = 'Param'#232'tres de relances des r'#232'glements...'
        Margin = 5
        Opaque = False
        OnClick = bTiers4Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bTiers5: TToolbarButton97
        Left = 59
        Top = 108
        Width = 222
        Height = 25
        Caption = 'Plan de rupture auxiliaire...'
        Margin = 5
        Opaque = False
        OnClick = bTiers5Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bTiers6: TToolbarButton97
        Left = 59
        Top = 207
        Width = 222
        Height = 25
        Caption = 'Plan de correspondance auxiliaire...'
        Margin = 5
        Opaque = False
        OnClick = bTiers6Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bTiers9: TToolbarButton97
        Left = 59
        Top = 34
        Width = 222
        Height = 25
        Caption = 'Tiers...'
        Margin = 5
        Opaque = False
        OnClick = bTiers9Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bTiers10: TToolbarButton97
        Left = 59
        Top = 182
        Width = 222
        Height = 25
        Caption = 'Civilit'#233's...'
        Margin = 5
        Opaque = False
        OnClick = bTiers10Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object HLabel13: THLabel
        Left = 100
        Top = 8
        Width = 124
        Height = 13
        Caption = 'Param'#233'trage des tiers'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object GroupBox11: TGroupBox
        Left = 100
        Top = 21
        Width = 124
        Height = 7
        TabOrder = 0
      end
    end
    object Divers: TTabSheet
      Hint = 'Autres param'#232'tres'
      Caption = 'Divers'
      object HLabel14: THLabel
        Left = 111
        Top = 10
        Width = 103
        Height = 13
        Caption = 'Autres param'#232'tres'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object bDivers1: TToolbarButton97
        Left = 81
        Top = 73
        Width = 182
        Height = 25
        Caption = 'Journaux comptables...'
        Margin = 5
        Opaque = False
        OnClick = bDivers1Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bDivers2: TToolbarButton97
        Left = 81
        Top = 40
        Width = 182
        Height = 25
        Caption = 'Compteurs comptables...'
        Margin = 5
        Opaque = False
        OnClick = bDivers2Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bDivers3: TToolbarButton97
        Left = 81
        Top = 107
        Width = 182
        Height = 25
        Caption = 'Libell'#233's automatiques...'
        Margin = 5
        Opaque = False
        OnClick = bDivers3Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bDivers4: TToolbarButton97
        Left = 81
        Top = 140
        Width = 182
        Height = 25
        Caption = 'Sc'#233'narios de saisie comptable...'
        Margin = 5
        Opaque = False
        OnClick = bDivers4Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bDivers5: TToolbarButton97
        Left = 81
        Top = 174
        Width = 182
        Height = 25
        Caption = 'Contrats d'#39'abonnements...'
        Margin = 5
        Opaque = False
        OnClick = bDivers5Click
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object GroupBox12: TGroupBox
        Left = 111
        Top = 23
        Width = 103
        Height = 7
        TabOrder = 0
      end
    end
  end
  object SSociete: TDataSource [15]
    Left = 133
    Top = 353
  end
  object SPlanRef: TDataSource [16]
    DataSet = QPlanRef
    Left = 134
    Top = 211
  end
  object QPlanRef: TQuery [17]
    DatabaseName = 'SOC'
    RequestLive = True
    SQL.Strings = (
      'SELECT * FROM PLANREF WHERE PR_NUMPLAN = 1 ORDER BY PR_COMPTE')
    Left = 85
    Top = 210
  end
  object StepMenu: TPopupMenu [18]
    Left = 101
    Top = 4
    object FinEtape: TMenuItem
      Caption = 'Etape termin'#233'e'
      OnClick = FinEtapeClick
    end
    object N1: TMenuItem
      Caption = '-'
    end
  end
  object PopAxes: TPopupMenu [19]
    Left = 48
    Top = 46
    object Axe1: TMenuItem
      Caption = 'Axe n'#176'1'
      OnClick = AxesClick
    end
    object Axe2: TMenuItem
      Caption = 'Axe n'#176'2'
      OnClick = AxesClick
    end
    object Axe3: TMenuItem
      Caption = 'Axe n'#176'3'
      OnClick = AxesClick
    end
    object Axe4: TMenuItem
      Caption = 'Axe n'#176'4'
      OnClick = AxesClick
    end
    object Axe5: TMenuItem
      Caption = 'Axe n'#176'5'
      OnClick = AxesClick
    end
  end
  inherited Msg: THMsgBox
    Mess.Strings = (
      'Etape'
      '1;?caption?;Voulez-vous quitter l'#39'assistant ?;Q;YN;Y;C;'
      
        '2;?caption?;Vous devez renseigner la longueur des comptes g'#233'n'#233'ra' +
        'ux et les caract'#232'res de bourrage.;W;O;O;O;')
    Left = 10
    Top = 46
  end
  object PopAxRupt: TPopupMenu
    Left = 99
    Top = 46
    object MenuItem1: TMenuItem
      Caption = 'Axe n'#176'1'
      OnClick = AxesRuptClick
    end
    object MenuItem2: TMenuItem
      Caption = 'Axe n'#176'2'
      OnClick = AxesRuptClick
    end
    object MenuItem3: TMenuItem
      Caption = 'Axe n'#176'3'
      OnClick = AxesRuptClick
    end
    object MenuItem4: TMenuItem
      Caption = 'Axe n'#176'4'
      OnClick = AxesRuptClick
    end
    object MenuItem5: TMenuItem
      Caption = 'Axe n'#176'5'
      OnClick = AxesRuptClick
    end
  end
end
