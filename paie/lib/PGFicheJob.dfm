object PFicheJob: TPFicheJob
  Left = 416
  Top = 134
  Width = 608
  Height = 530
  BorderIcons = []
  Caption = 'Param'#233'trage d'#39'une t'#226'che planifi'#233'e'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  OnCanResize = FormCanResize
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PanelHaut: THPanel
    Left = 0
    Top = 0
    Width = 600
    Height = 449
    Align = alClient
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object GroupIdentification: THGroupBox
      Left = 0
      Top = 0
      Width = 600
      Height = 105
      Align = alTop
      Caption = '&Identification'
      TabOrder = 0
      DesignSize = (
        600
        105)
      object HLabel1: THLabel
        Left = 8
        Top = 20
        Width = 81
        Height = 13
        AutoSize = False
        Caption = '&Libell'#233
        FocusControl = SKJ_LIBELLE
      end
      object HLabel2: THLabel
        Left = 8
        Top = 76
        Width = 81
        Height = 13
        AutoSize = False
        Caption = 'Libell'#233' auto'
        FocusControl = SKJ_LIBELLEAUTO
      end
      object PanelJob: THPanel
        Left = 498
        Top = 9
        Width = 97
        Height = 33
        Hint = 'N'#176' de Job en cours d'#39#233'dition'
        Anchors = [akTop, akRight]
        BevelOuter = bvSpace
        BevelWidth = 3
        Caption = 'N'#176' 18888'
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -21
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 4
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
      end
      object SKJ_LIBELLE: THCritMaskEdit
        Left = 91
        Top = 16
        Width = 345
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        OnChange = ChangeField
        TagDispatch = 0
      end
      object SKJ_ACTIF: THCheckbox
        Left = 8
        Top = 47
        Width = 97
        Height = 17
        Hint = 'Permet de reprendre ou de suspendre l'#39'activit'#233' d'#39'un job'
        Caption = 'T'#226'che &active'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = ChangeField
      end
      object SKJ_LIBELLEAUTO: THCritMaskEdit
        Left = 91
        Top = 72
        Width = 501
        Height = 21
        Hint = 
          'Ce libell'#233' est g'#233'n'#233'r'#233' en fonction de la planification du job (le' +
          'cture seule)'
        ParentShowHint = False
        ReadOnly = True
        ShowHint = True
        TabOrder = 3
        TagDispatch = 0
      end
      object SKJ_FORCEEXEC: THCheckbox
        Left = 120
        Top = 45
        Width = 107
        Height = 17
        Hint = 
          'Permet de forcer l'#39'ex'#233'cution d'#39'un job sans attendre son heure de' +
          ' lancement'
        Caption = '&Forcer l'#39'ex'#233'cution'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = ChangeField
      end
    end
    object GroupeCaracteristiques: THGroupBox
      Left = 0
      Top = 105
      Width = 600
      Height = 104
      Align = alTop
      Caption = '&Caract'#233'ristiques'
      TabOrder = 1
      object HLabel3: THLabel
        Left = 8
        Top = 20
        Width = 81
        Height = 13
        AutoSize = False
        Caption = '&eMail'
      end
      object HLabel4: THLabel
        Left = 334
        Top = 20
        Width = 147
        Height = 13
        AutoSize = False
        Caption = '&Niveau de Confidentialit'#233
      end
      object HLabel5: THLabel
        Left = 8
        Top = 44
        Width = 81
        Height = 13
        AutoSize = False
        Caption = 'Ex'#233'cutable'
      end
      object HLabel6: THLabel
        Left = 8
        Top = 68
        Width = 81
        Height = 13
        AutoSize = False
        Caption = 'Action'
      end
      object SKJ_EMAIL: THCritMaskEdit
        Left = 91
        Top = 16
        Width = 213
        Height = 21
        Hint = 
          'Adresse de messagerie internet (SMTP) '#224' laquelle sera envoy'#233' l'#39'i' +
          'nformation sur le traitement du job'
        CharCase = ecLowerCase
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnChange = ChangeField
        TagDispatch = 0
      end
      object SKJ_CONFIDENTIEL: THSpinEdit
        Left = 486
        Top = 16
        Width = 57
        Height = 22
        Hint = 'Permet de fixer un niveau de confidentialit'#233' sur le job'
        MaxValue = 10
        MinValue = 0
        TabOrder = 1
        Value = 0
        OnChange = ChangeField
      end
      object SKJ_EXENAME: THCritMaskEdit
        Left = 91
        Top = 40
        Width = 345
        Height = 21
        Hint = 'Nom du programme qui sera ex'#233'cut'#233
        CharCase = ecLowerCase
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnChange = ChangeField
        TagDispatch = 0
      end
      object SKJACTION: THCritMaskEdit
        Left = 91
        Top = 64
        Width = 345
        Height = 21
        Hint = 'Code action qui sera envoy'#233' au programme'
        CharCase = ecUpperCase
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnChange = ChangeField
        TagDispatch = 0
      end
      object SKJ_NOMJOB: THCritMaskEdit
        Left = 440
        Top = 48
        Width = 49
        Height = 21
        Color = clYellow
        TabOrder = 4
        Visible = False
        TagDispatch = 0
      end
      object SKJ_LIBREJOB: THCritMaskEdit
        Left = 496
        Top = 48
        Width = 49
        Height = 21
        Color = clYellow
        TabOrder = 5
        Visible = False
        TagDispatch = 0
      end
      object SKJ_TYPEJOB: THCritMaskEdit
        Left = 448
        Top = 72
        Width = 121
        Height = 21
        Color = clYellow
        TabOrder = 6
        Visible = False
        TagDispatch = 0
      end
    end
    object PanelFrame: THPanel
      Left = 0
      Top = 209
      Width = 600
      Height = 240
      Align = alClient
      FullRepaint = False
      TabOrder = 2
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 449
    Width = 600
    Height = 47
    AllowDrag = False
    Position = dpBottom
    object PanelBas: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 43
      ClientWidth = 600
      ClientAreaHeight = 43
      ClientAreaWidth = 600
      DockPos = 0
      DragHandleStyle = dhNone
      FullSize = True
      TabOrder = 0
      DesignSize = (
        600
        43)
      object Baide: TToolbarButton97
        Left = 571
        Top = 5
        Width = 28
        Height = 28
        Hint = 'Aide (F1)'
        Anchors = [akTop, akRight]
        Flat = False
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object BFerme: TToolbarButton97
        Left = 538
        Top = 5
        Width = 28
        Height = 28
        Hint = 'Ferme la fiche (ESC)'
        Anchors = [akTop, akRight]
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BValide: TToolbarButton97
        Left = 506
        Top = 5
        Width = 28
        Height = 28
        Hint = 'Valide la fiche (F10)'
        Anchors = [akTop, akRight]
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BValideClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object Bsupprime: TToolbarButton97
        Left = 5
        Top = 5
        Width = 28
        Height = 28
        Hint = 'Supprime la fiche (F7)'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BsupprimeClick
        GlobalIndexImage = 'Z0005_S16G1'
      end
      object Binsert: TToolbarButton97
        Left = 37
        Top = 5
        Width = 28
        Height = 28
        Hint = 'Cr'#233'ation d'#39'une nouvelle fiche (F8)'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BinsertClick
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BAgrandir: TToolbarButton97
        Left = 69
        Top = 5
        Width = 28
        Height = 28
        Hint = 'Affichage des caract'#233'ristiques'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BAgrandirClick
        GlobalIndexImage = 'Z0270_S16G1'
      end
      object BReduire: TToolbarButton97
        Left = 101
        Top = 5
        Width = 28
        Height = 28
        Hint = 'Suppression de l'#39'affichage des caract'#233'ristiques'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BReduireClick
        GlobalIndexImage = 'Z0910_S16G1'
      end
      object BPurge: TToolbarButton97
        Left = 277
        Top = 5
        Width = 28
        Height = 28
        Hint = 'Initialisation du job'
        Enabled = False
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BPurgeClick
        GlobalIndexImage = 'Z0338_S16G1'
      end
    end
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    LockedCtrls.Strings = (
      'bFiltre'
      'bDimension')
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 380
    Top = 228
  end
  object Mess: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?Caption?;Confirmez-vous la r'#233'initialisation du job.?;Q;YN;N;N')
    Left = 288
    Top = 233
  end
end
