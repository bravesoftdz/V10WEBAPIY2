object FNouveauLot: TFNouveauLot
  Left = 349
  Top = 241
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Saisie d'#39'un lot'
  ClientHeight = 191
  ClientWidth = 287
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PGeneral: THPanel
    Left = 0
    Top = 0
    Width = 287
    Height = 156
    Align = alClient
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TNumeroLot: THLabel
      Left = 18
      Top = 77
      Width = 66
      Height = 13
      Caption = '&Num'#233'ro de lot'
      FocusControl = NumeroLot
    end
    object TDateLot: THLabel
      Left = 18
      Top = 101
      Width = 90
      Height = 13
      Caption = '&Date de fabrication'
      FocusControl = DateLot
    end
    object TPhysique: THLabel
      Left = 18
      Top = 125
      Width = 105
      Height = 13
      Caption = '&Quantit'#233' r'#233'ceptionn'#233'e'
      FocusControl = Physique
    end
    object TArticle: THLabel
      Left = 18
      Top = 16
      Width = 29
      Height = 13
      Caption = 'Article'
    end
    object TDepot: THLabel
      Left = 18
      Top = 40
      Width = 29
      Height = 13
      Caption = 'D'#233'p'#244't'
    end
    object Article: THLabel
      Left = 61
      Top = 16
      Width = 29
      Height = 13
      Caption = 'Article'
    end
    object Depot: THLabel
      Left = 61
      Top = 40
      Width = 29
      Height = 13
      Caption = 'Depot'
    end
    object NumeroLot: THCritMaskEdit
      Left = 132
      Top = 73
      Width = 113
      Height = 21
      TabOrder = 0
      TagDispatch = 0
    end
    object DateLot: THCritMaskEdit
      Left = 132
      Top = 97
      Width = 68
      Height = 21
      EditMask = '!99/99/9999;1;_'
      MaxLength = 10
      TabOrder = 1
      Text = '  /  /    '
      TagDispatch = 0
      OpeType = otDate
      ControlerDate = True
    end
    object Physique: THCritMaskEdit
      Left = 132
      Top = 121
      Width = 113
      Height = 21
      TabOrder = 2
      OnExit = PhysiqueExit
      TagDispatch = 0
      OpeType = otReel
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 156
    Width = 287
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 287
      Caption = 'ToolWindow971'
      ClientAreaHeight = 31
      ClientAreaWidth = 287
      DockPos = 0
      DragHandleStyle = dhNone
      FullSize = True
      Resizable = False
      TabOrder = 0
      object baide: TToolbarButton97
        Left = 251
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object bfermer: TToolbarButton97
        Left = 221
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Flat = False
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object bvalider: TToolbarButton97
        Left = 190
        Top = 3
        Width = 29
        Height = 27
        Hint = 'Valider'
        Flat = False
        ModalResult = 1
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = bvaliderClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?caption?;Vous n'#39'avez pas renseign'#233' de num'#233'ro de lot;W;O;O;O;'
      
        '1;?caption?;La date de fabrication doit '#234'tre ant'#233'rieur '#224' la date' +
        ' de connexion;W;O;O;O;'
      '2;?caption?;Un num'#233'ro de lot identique existe d'#233'j'#224';W;O;O;O;'
      
        '3;?caption?;Vous n'#39'avez pas valid'#233'.t#13 Voulez-vous continuer?;Q' +
        ';YN;Y;N;')
    Left = 256
    Top = 8
  end
end
