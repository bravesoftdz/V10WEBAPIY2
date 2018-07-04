object FSaisLett: TFSaisLett
  Left = 230
  Top = 268
  Width = 485
  Height = 200
  HelpContext = 7244500
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Liste des '#233'ch'#233'ances de la pi'#232'ce'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HP6: TPanel
    Left = 0
    Top = 138
    Width = 477
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 0
    object BValider: TToolbarButton97
      Left = 380
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Lancer le lettrage'
      Flat = False
      NumGlyphs = 2
      ParentShowHint = False
      ShowHint = True
      OnClick = BValiderClick
      GlobalIndexImage = 'Z0003_S16G2'
    end
    object BAnnuler: TToolbarButton97
      Left = 412
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Flat = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BAnnulerClick
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BAide: TToolbarButton97
      Left = 444
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Aide'
      Flat = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BAideClick
      GlobalIndexImage = 'Z1117_S16G1'
    end
    object Distinguer: TCheckBox
      Left = 40
      Top = 8
      Width = 225
      Height = 17
      Alignment = taLeftJustify
      Caption = 'Distinguer les Euros'
      Checked = True
      State = cbChecked
      TabOrder = 0
    end
  end
  object GL: TDBGrid
    Left = 0
    Top = 0
    Width = 477
    Height = 138
    Align = alClient
    DataSource = SEche
    Options = [dgTitles, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit]
    ReadOnly = True
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnDblClick = GLDblClick
  end
  object Pages: TPageControl
    Left = 12
    Top = 68
    Width = 333
    Height = 57
    ActivePage = TabSheet1
    TabOrder = 2
    Visible = False
    object TabSheet1: TTabSheet
      Caption = 'Pi'#232'ce'
      object E_ETATLETTRAGE: THCritMaskEdit
        Left = 196
        Top = 4
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        Text = 'RI'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object E_NUMECHE: THCritMaskEdit
        Left = 172
        Top = 4
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        Text = '0'
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_ECHE: THCritMaskEdit
        Left = 148
        Top = 4
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_ETABLISSEMENT: THCritMaskEdit
        Left = 124
        Top = 4
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Left = 100
        Top = 4
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        Visible = False
        TagDispatch = 0
        Operateur = Egal
        OpeType = otReel
      end
      object E_DEVISE: THCritMaskEdit
        Left = 76
        Top = 4
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_NATUREPIECE: THCritMaskEdit
        Left = 48
        Top = 4
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 6
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 24
        Top = 4
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 7
        Visible = False
        TagDispatch = 0
        Operateur = Egal
        OpeType = otDate
        ControlerDate = True
      end
      object E_JOURNAL: THCritMaskEdit
        Left = 0
        Top = 4
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 8
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_TRESOLETTRE: THCritMaskEdit
        Left = 220
        Top = 4
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 9
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object E_QUALIFPIECE: THCritMaskEdit
        Left = 244
        Top = 4
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 10
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_EXERCICE: THCritMaskEdit
        Left = 268
        Top = 4
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 11
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object XX_WHERE: TEdit
        Left = 293
        Top = 9
        Width = 13
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 12
        Visible = False
      end
    end
  end
  object QEche: THQuery
    AutoCalcFields = False
    LockType = ltReadOnly
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    dataBaseName = 'SOC'
    PageCriteres = Pages
    MAJAuto = False
    Distinct = False
    Left = 92
    Top = 12
  end
  object SEche: TDataSource
    DataSet = QEche
    Left = 136
    Top = 12
  end
  object HLS: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Lettrage non autoris'#233' : cette '#233'ch'#233'ance est d'#233'j'#224' tota' +
        'lement lettr'#233'e;W;O;O;O;'
      'Distinguer'
      '2;')
    Left = 180
    Top = 12
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 220
    Top = 12
  end
  object TT: TTimer
    Enabled = False
    OnTimer = TTTimer
    Left = 288
    Top = 12
  end
end
