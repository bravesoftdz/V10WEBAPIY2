inherited FMulScript: TFMulScript
  Left = 224
  Top = 181
  Width = 722
  Height = 557
  HelpContext = 1010
  Caption = 'Liste des scripts'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: THPageControl2
    Width = 714
    Height = 113
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      inherited Bevel1: TBevel
        Width = 706
        Height = 85
      end
      object TDOMAINE: TLabel
        Left = 24
        Top = 12
        Width = 42
        Height = 13
        Caption = 'Domaine'
      end
      object Label2: TLabel
        Left = 268
        Top = 12
        Width = 25
        Height = 13
        Caption = 'Code'
        FocusControl = CLE
      end
      object Label3: TLabel
        Left = 24
        Top = 59
        Width = 58
        Height = 13
        Caption = 'Compl'#233'ment'
        FocusControl = CLEPAR
      end
      object Label4: TLabel
        Left = 24
        Top = 35
        Width = 30
        Height = 13
        Caption = 'Libell'#233
        FocusControl = Comment
      end
      object Label1: TLabel
        Left = 470
        Top = 12
        Width = 32
        Height = 13
        Caption = 'Nature'
      end
      object Label5: TLabel
        Left = 470
        Top = 35
        Width = 33
        Height = 13
        Caption = 'Editeur'
        FocusControl = Editeur
      end
      object ToolbarButton971: TToolbarButton97
        Left = 242
        Top = 54
        Width = 28
        Height = 22
        Hint = 'Appliquer crit'#232'res'
        Flat = False
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object CLE: TEdit
        Left = 320
        Top = 8
        Width = 143
        Height = 21
        TabOrder = 0
      end
      object XX_WHERE: TEdit
        Left = 401
        Top = -3
        Width = 17
        Height = 19
        TabStop = False
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 4
        Visible = False
      end
      object CLEPAR: TEdit
        Left = 104
        Top = 55
        Width = 133
        Height = 21
        TabOrder = 3
      end
      object Comment: TEdit
        Left = 104
        Top = 31
        Width = 360
        Height = 21
        TabOrder = 2
      end
      object Editeur: TEdit
        Left = 518
        Top = 31
        Width = 143
        Height = 21
        TabOrder = 5
      end
      object Nature: TComboBox
        Left = 518
        Top = 8
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        Items.Strings = (
          '<Tous>'
          'Dossier'
          'Journal'
          'Balance'
          'Synchronisation')
      end
      object Domaine: THValComboBox
        Left = 105
        Top = 8
        Width = 141
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 6
        TagDispatch = 0
        ComboWidth = 300
      end
    end
    inherited PComplement: THTabSheet
      Caption = 'Compl'#233'ment'
      inherited Bevel2: TBevel
        Width = 706
        Height = 85
      end
    end
    inherited PAvance: THTabSheet
      TabVisible = False
      inherited Bevel4: TBevel
        Width = 706
        Height = 85
      end
      inherited ZV1: THEdit
        Width = 201
      end
      inherited ZV2: THEdit
        Width = 201
      end
      inherited ZV3: THEdit
        Width = 201
      end
      inherited ZG2: THCombobox
        Left = 489
      end
      inherited ZG1: THCombobox
        Left = 489
      end
    end
    inherited PSQL: THTabSheet
      TabVisible = False
      inherited Bevel3: TBevel
        Width = 706
        Height = 85
      end
      inherited Z_SQL: THSQLMemo
        Width = 706
        Height = 85
      end
    end
  end
  inherited Dock971: TDock97
    Top = 113
    Width = 714
    inherited PFiltres: TToolWindow97
      ClientWidth = 714
      ClientAreaWidth = 714
      Visible = False
      inherited BFiltre: TToolbarButton97
        Enabled = False
        Visible = False
      end
      inherited BCherche: TToolbarButton97
        Left = 506
        Top = 8
        Height = 22
        Visible = False
      end
      inherited lpresentation: THLabel
        Left = 465
        Visible = False
      end
      inherited FFiltres: THValComboBox
        Left = 72
        Width = 425
        Enabled = False
        Visible = False
      end
      inherited cbPresentations: THValComboBox
        Left = 144
        Width = 69
        Visible = False
      end
    end
  end
  inherited FListe: THDBGrid
    Top = 154
    Width = 685
    Height = 299
    MultiSelection = True
    MultiFieds = 'CLE;'
  end
  inherited Panel2: THPanel
    inherited PListe: THPanel
      Left = 486
    end
  end
  inherited Dock: TDock97
    Top = 475
    Width = 714
    Height = 48
    inherited PanelBouton: TToolWindow97
      ClientHeight = 44
      ClientWidth = 714
      ClientAreaHeight = 44
      ClientAreaWidth = 714
      inherited BAgrandir: TToolbarButton97
        Left = 8
        Top = 0
      end
      inherited bSelectAll: TToolbarButton97
        Left = 301
        Visible = True
      end
      inherited BRechercher: TToolbarButton97
        Left = 32
        Top = 0
      end
      inherited BParamListe: TToolbarButton97
        Left = 237
        Visible = False
      end
      inherited bExport: TToolbarButton97
        Left = 269
        Enabled = True
      end
      inherited BImprimer: TToolbarButton97
        Left = 577
      end
      inherited BOuvrir: TToolbarButton97
        Left = 609
      end
      inherited BAnnuler: TToolbarButton97
        Left = 641
      end
      inherited BAide: TToolbarButton97
        Tag = 1010
        Left = 673
      end
      inherited Binsert: TToolbarButton97
        Left = 72
        Top = 0
        Visible = True
        OnClick = BinsertClick
      end
      inherited BBlocNote: TToolbarButton97
        Left = 545
      end
      object BDelete: TToolbarButton97
        Left = 104
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Supprimer'
        DisplayMode = dmGlyphOnly
        Caption = 'Supprimer'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BDeleteClick
        GlobalIndexImage = 'Z0005_S16G1'
      end
      object BDelim: TToolbarButton97
        Left = 543
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Acc'#232's au param'#233'trage d'#233'limit'#233
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Flat = False
        Glyph.Data = {
          42010000424D4201000000000000760000002800000011000000110000000100
          040000000000CC00000000000000000000001000000010000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          77777000000077777777777777777000000070000000000007777000000070FF
          FFFFFFFF07777000000070FCCFCCCCCF07777000000070FFFFFFFFFF07777000
          000070FCCFCCCCCF07777000000070FFFFFFFFFF07777000000070FFFFFFF0FF
          07777000000070F00FFF0B0F07770000000070F0F0F0B0F000700000000070FF
          0B0B0F0FBF0000000000700000F0F0FBFBF0000000007777770B0FBFBFB00000
          000077777770FBFBFB0000000000777777770000007000000000777777777777
          777770000000}
        GlyphMask.Data = {00000000}
        ParentShowHint = False
        ShowHint = True
        OnClick = BDelimClick
      end
      object BSauve: TToolbarButton97
        Left = 136
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Sauvegarder'
        DisplayMode = dmGlyphOnly
        Caption = 'Supprimer'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BSauveClick
        GlobalIndexImage = 'Z0027_S16G1'
      end
    end
  end
  inherited PCumul: THPanel
    Top = 453
    Width = 714
  end
  inherited TBlocNote: TToolWindow97
    Left = 417
    Top = 325
  end
  inherited PanVBar: THPanel
    Left = 685
    Top = 154
    Height = 299
  end
  inherited SQ: TDataSource
    AutoEdit = False
    Left = 122
    Top = 236
  end
  inherited Q: THQuery
    SQL.Strings = (
      'SELECT *  from PGZIMPREQ.DB'
      'Where CLEPAR<>"SQL"')
    dataBaseName = 'GLOBAL'
    object QTable1: TStringField
      DisplayLabel = 'Editeur'
      DisplayWidth = 15
      FieldName = 'Table1'
      Size = 30
    end
    object QCLE: TStringField
      DisplayLabel = 'Code'
      DisplayWidth = 15
      FieldName = 'CLE'
      Size = 12
    end
    object QCOMMENT: TStringField
      DisplayLabel = 'Libell'#233
      DisplayWidth = 50
      FieldName = 'COMMENT'
      Size = 60
    end
    object QTable0: TStringField
      DisplayLabel = 'Nature'
      DisplayWidth = 15
      FieldName = 'Table0'
      Size = 30
    end
    object QCLEPAR: TStringField
      DisplayLabel = 'Compl'#233'ment'
      DisplayWidth = 15
      FieldName = 'CLEPAR'
      Size = 10
    end
    object QDATEDEMODIF: TDateField
      DisplayLabel = 'Date modif.'
      DisplayWidth = 15
      FieldName = 'DATEDEMODIF'
    end
    object QDOMAINE: TStringField
      DisplayLabel = 'Domaine'
      FieldName = 'DOMAINE'
      Size = 2
    end
  end
  inherited FindDialog: THFindDialog
    Top = 236
  end
  inherited POPF: THPopupMenu
    Left = 278
    Top = 232
  end
  inherited HMTrad: THSystemMenu
    Left = 175
    Top = 244
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Visualisation des journaux'
      'Modification des journaux'
      
        '2;Journaux;Le zoom est impossible : aucun code n'#39'est renseign'#233';W' +
        ';O;O;O;'
      '3;Journaux;ATTENTION : ce journal est ferm'#233'.;E;O;O;O;'
      'Journaux')
    Left = 177
    Top = 184
  end
  object HFindDialog1: THFindDialog
    OnFind = FindDialogFind
    Left = 196
    Top = 292
  end
end
