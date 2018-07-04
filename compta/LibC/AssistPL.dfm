inherited FAssistPL: TFAssistPL
  Left = 315
  Top = 199
  Caption = 'Assistant comptabilit'#233
  ClientWidth = 563
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Top = 313
  end
  inherited lAide: THLabel
    Left = 175
    Top = 283
    Width = 337
    Height = 17
  end
  inherited bPrecedent: TToolbarButton97
    Left = 316
  end
  inherited bSuivant: TToolbarButton97
    Left = 399
  end
  inherited bFin: TToolbarButton97
    Left = 483
  end
  inherited bAnnuler: TToolbarButton97
    Left = 233
  end
  inherited bAide: TToolbarButton97
    Left = 150
  end
  inherited GroupBox1: TGroupBox [7]
    Left = 5
    Top = 297
    Width = 601
  end
  inherited cControls: TListBox [8]
  end
  inherited PanelImage: THPanel [9]
    TabOrder = 2
  end
  inherited Plan: THPanel [10]
  end
  inherited P: TPageControl [11]
    Width = 374
    ActivePage = PDebut
    TabOrder = 3
    object PDebut: TTabSheet
      Caption = 'PDebut'
      object HLabel1: THLabel
        Left = 17
        Top = 4
        Width = 220
        Height = 13
        Caption = 'Bienvenue dans l'#39'assistant de cr'#233'ation'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object TSO_LIBELLE: THLabel
        Left = 17
        Top = 98
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
      object HLabel5: THLabel
        Left = 17
        Top = 73
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
      object HLabel6: THLabel
        Left = 17
        Top = 52
        Width = 237
        Height = 13
        AutoSize = False
        Caption = 'Indiquez le code et la raison sociale de la soci'#233't'#233' :'
      end
      object bCoordonnees: TToolbarButton97
        Left = 57
        Top = 123
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
        OnClick = bCoordonneesClick
      end
      object LABEL_CREATION: THLabel
        Left = 16
        Top = 181
        Width = 116
        Height = 13
        Caption = 'Plan de correspondance'
      end
      object HLabel3: THLabel
        Left = 17
        Top = 20
        Width = 121
        Height = 13
        Caption = 'de dossier comptable'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object HLabel9: THLabel
        Left = 16
        Top = 204
        Width = 86
        Height = 13
        Caption = 'Monnaie de tenue'
      end
      object BParamDevise: TToolbarButton97
        Left = 339
        Top = 199
        Width = 23
        Height = 22
        Hint = 'Selection d'#233'taill'#233'e du plan type'
        OnClick = BParamDeviseClick
        GlobalIndexImage = 'Z0008_S16G1'
      end
      object Label3: TLabel
        Left = 16
        Top = 226
        Width = 124
        Height = 13
        Caption = 'Nom du fichier d'#39#233'change '
        Transparent = True
        Visible = False
      end
      object TMODECREAT: THLabel
        Left = 16
        Top = 157
        Width = 83
        Height = 13
        Caption = 'Mode de cr'#233'ation'
      end
      object FICHIER_IMPORT: THCritMaskEdit
        Left = 154
        Top = 177
        Width = 184
        Height = 21
        TabOrder = 4
        Visible = False
        OnClick = FICHIER_IMPORTClick
        TagDispatch = 0
        ElipsisButton = True
        ElipsisAutoHide = True
      end
      object SO_LIBELLE: TEdit
        Left = 98
        Top = 94
        Width = 225
        Height = 21
        Ctl3D = True
        Enabled = False
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
        Left = 98
        Top = 69
        Width = 53
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        Enabled = False
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
        Top = 33
        Width = 311
        Height = 7
        TabOrder = 2
      end
      object SO_NUMPLANREF: THValComboBox
        Left = 154
        Top = 176
        Width = 184
        Height = 21
        ItemHeight = 13
        TabOrder = 3
        TagDispatch = 0
      end
      object MONNAIE_TENUE: THValComboBox
        Left = 154
        Top = 199
        Width = 184
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 5
        TagDispatch = 0
        DataType = 'TTDEVISE'
      end
      object FICHENAME: THCritMaskEdit
        Left = 154
        Top = 222
        Width = 184
        Height = 21
        TabOrder = 6
        Visible = False
        OnChange = MODECREATChange
        TagDispatch = 0
        DataType = 'OPENFILE(*.TRA;*.ZIP;*.TRT;*.*)'
        ElipsisButton = True
      end
      object MODECREAT: THValComboBox
        Left = 154
        Top = 153
        Width = 184
        Height = 21
        ItemHeight = 13
        TabOrder = 7
        OnChange = MODECREATChange
        Items.Strings = (
          'Plan de correspondance'
          'Fichier')
        TagDispatch = 0
        Values.Strings = (
          'PLA'
          'FIC')
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'TabSheet2'
      ImageIndex = 1
      object HLabel4: THLabel
        Left = 79
        Top = 21
        Width = 159
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
      object bParamComptable: TToolbarButton97
        Left = 57
        Top = 83
        Width = 192
        Height = 25
        Alignment = taLeftJustify
        Caption = 'Param'#232'tres &comptables...           '
        Margin = 4
        Opaque = False
        OnClick = bParamComptableClick
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object ToolbarButton971: TToolbarButton97
        Left = 57
        Top = 53
        Width = 192
        Height = 25
        Alignment = taLeftJustify
        DropdownAlways = True
        DropdownArrow = True
        DropdownMenu = PopParamPlan
        Caption = 'Param'#233'trage du &Plan Comptable'
        Margin = 4
        Opaque = False
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bFourchettes: TToolbarButton97
        Left = 57
        Top = 113
        Width = 192
        Height = 25
        Alignment = taLeftJustify
        Caption = 'Fourchettes de comptes...'
        Margin = 4
        Opaque = False
        OnClick = bFourchettesClick
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object GroupBox3: TGroupBox
        Left = 79
        Top = 34
        Width = 138
        Height = 7
        TabOrder = 0
      end
    end
    object TabSheet3: TTabSheet
      Caption = 'TabSheet3'
      ImageIndex = 2
      object HLabel8: THLabel
        Left = 61
        Top = 15
        Width = 180
        Height = 16
        AutoSize = False
        Caption = 'Caract'#233'ristiques de l'#39'entreprise'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object bMonnaie: TToolbarButton97
        Left = 28
        Top = 47
        Width = 190
        Height = 25
        Caption = 'Monnaie'
        ImageIndex = 0
        Margin = 4
        Opaque = False
        OnClick = bMonnaieClick
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bComptaAux: TToolbarButton97
        Left = 28
        Top = 74
        Width = 190
        Height = 25
        Caption = 'Comptabilit'#233' auxiliaire'
        ImageIndex = 1
        Images = ImageList
        Margin = 4
        Opaque = False
        OnClick = bComptaAuxClick
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bDevise: TToolbarButton97
        Left = 28
        Top = 101
        Width = 190
        Height = 25
        Caption = 'Gestion des devises'
        ImageIndex = 0
        Images = ImageList
        Margin = 4
        Opaque = False
        OnClick = OnBoutonActivationClick
        OnMouseDown = OnBoutonActivationMouseDown
        OnMouseEnter = bButton97MouseEnter
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object bPreferences: TToolbarButton97
        Left = 28
        Top = 155
        Width = 190
        Height = 25
        Caption = 'Pr'#233'f'#233'rences'
        Margin = 4
        Opaque = False
        OnClick = bPreferencesClick
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bEtablissement: TToolbarButton97
        Left = 28
        Top = 182
        Width = 190
        Height = 25
        Caption = 'Gestion des '#233'tablissements'
        ImageIndex = 0
        Images = ImageList
        Margin = 4
        Opaque = False
        Visible = False
        OnClick = OnBoutonActivationClick
        OnMouseDown = OnBoutonActivationMouseDown
        OnMouseEnter = bButton97MouseEnter
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bContrat: TToolbarButton97
        Left = 28
        Top = 128
        Width = 190
        Height = 25
        Caption = 'Gestion des contrats'
        ImageIndex = 0
        Images = ImageList
        Margin = 4
        Opaque = False
        OnClick = OnBoutonActivationClick
        OnMouseDown = OnBoutonActivationMouseDown
        OnMouseEnter = bButton97MouseEnter
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object GroupBox4: TGroupBox
        Left = 61
        Top = 28
        Width = 180
        Height = 7
        TabOrder = 0
      end
    end
    object TabSheet4: TTabSheet
      Caption = 'TabSheet4'
      ImageIndex = 3
      PopupMenu = PopParam
      object HLabel2: THLabel
        Left = 64
        Top = 0
        Width = 198
        Height = 13
        Caption = 'P'#233'rim'#232'tre fonctionnel de la mission'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object bAnalytique: TToolbarButton97
        Left = 60
        Top = 22
        Width = 220
        Height = 25
        Caption = 'Comptabilit'#233' analytique'
        ImageIndex = 0
        Images = ImageList
        Margin = 5
        Opaque = False
        OnClick = OnBoutonActivationClick
        OnMouseDown = OnBoutonActivationMouseDown
        OnMouseEnter = bButton97MouseEnter
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bBDSDynamique: TToolbarButton97
        Left = 60
        Top = 212
        Width = 220
        Height = 25
        Caption = 'Balances de situation dynamiques'
        ImageIndex = 0
        Images = ImageList
        Margin = 5
        Opaque = False
        OnClick = OnBoutonActivationClick
        OnMouseDown = OnBoutonActivationMouseDown
        OnMouseEnter = bButton97MouseEnter
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bTVA: TToolbarButton97
        Left = 60
        Top = 49
        Width = 220
        Height = 25
        Caption = 'D'#233'claration de TVA'
        ImageIndex = 0
        Images = ImageList
        Margin = 5
        Opaque = False
        OnClick = OnBoutonActivationClick
        OnMouseDown = OnBoutonActivationMouseDown
        OnMouseEnter = bButton97MouseEnter
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bTresorerie: TToolbarButton97
        Left = 60
        Top = 76
        Width = 220
        Height = 25
        Caption = 'Gestion de tr'#233'sorerie'
        ImageIndex = 0
        Images = ImageList
        Margin = 5
        Opaque = False
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bTiers: TToolbarButton97
        Left = 60
        Top = 103
        Width = 220
        Height = 25
        Caption = 'Gestion des tiers'
        ImageIndex = 0
        Images = ImageList
        Margin = 5
        Opaque = False
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bSuivi: TToolbarButton97
        Left = 60
        Top = 130
        Width = 220
        Height = 25
        Caption = 'Suivi annexe'
        ImageIndex = 0
        Images = ImageList
        Margin = 5
        Opaque = False
        OnClick = OnBoutonActivationClick
        OnMouseDown = OnBoutonActivationMouseDown
        OnMouseEnter = bButton97MouseEnter
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bImmobilisation: TToolbarButton97
        Left = 60
        Top = 157
        Width = 220
        Height = 25
        Caption = 'Immobilisations'
        ImageIndex = 0
        Images = ImageList
        Margin = 5
        Opaque = False
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object bSaisieQte: TToolbarButton97
        Left = 60
        Top = 184
        Width = 220
        Height = 25
        Caption = 'Gestion des quantit'#233's'
        ImageIndex = 0
        Images = ImageList
        Margin = 5
        Opaque = False
        OnClick = OnBoutonActivationClick
        OnMouseDown = OnBoutonActivationMouseDown
        OnMouseEnter = bButton97MouseEnter
        GlobalIndexImage = 'Z0003_S16G1'
      end
      object TGroupBox
        Left = 64
        Top = 12
        Width = 201
        Height = 7
        TabOrder = 0
      end
    end
  end
  inherited Msg: THMsgBox
    Mess.Strings = (
      'Etape'
      '1;Assistant;Voulez-vous quitter l'#39'assistant ?;Q;YN;Y;C;'
      'Plan de correspondance'
      'Fichier d'#39'import')
  end
  object ActionList1: TActionList
    Left = 121
    Top = 130
  end
  object PopParam: TPopupMenu
    Left = 472
    Top = 124
    object Activer: TMenuItem
      Caption = '&Activer'
      OnClick = ActiverClick
    end
    object Desactiver: TMenuItem
      Caption = '&D'#233'sactiver'
      OnClick = DesactiverClick
    end
  end
  object ImageList: THImageList
    GlobalIndexImages.Strings = (
      'Z0294_S16G1'
      'Z0003_S16G1')
    Left = 471
    Top = 74
  end
  object PopParamPlan: TPopupMenu
    Left = 520
    Top = 114
    object PlanComptesGeneraux: TMenuItem
      Caption = 'Param'#233'trage du plan des comptes g'#233'n'#233'raux'
      OnClick = PlanComptesGenerauxClick
    end
    object JournauxComptables: TMenuItem
      Caption = 'Journaux comptables'
      OnClick = JournauxComptablesClick
    end
    object LibelleAuto: TMenuItem
      Caption = 'Libell'#233's automatiques'
      OnClick = LibelleAutoClick
    end
    object GuideSaisie: TMenuItem
      Caption = 'D'#233'finition des guides de saisie'
      OnClick = GuideSaisieClick
    end
  end
  object OpenDialog: TOpenDialog
    Left = 428
    Top = 77
  end
end
