object FSaisieDate: TFSaisieDate
  Left = 325
  Top = 325
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  ClientHeight = 134
  ClientWidth = 527
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnClose = FormClose
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 38
    Top = 27
    Width = 98
    Height = 13
    Caption = 'Date de consultation'
  end
  object PSAISIEDATE: THPanel
    Left = 0
    Top = 0
    Width = 527
    Height = 97
    Align = alClient
    Alignment = taLeftJustify
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object DateDu: TLabel
      Left = 16
      Top = 13
      Width = 98
      Height = 13
      Caption = 'Date de consultation'
      FocusControl = DateDebut
      Transparent = True
    end
    object TParamPlanning: TLabel
      Left = 16
      Top = 42
      Width = 35
      Height = 13
      Caption = 'Mod'#232'le'
      Transparent = True
    end
    object LPARAMPLANNING: TLabel
      Left = 187
      Top = 42
      Width = 181
      Height = 13
      AutoSize = False
    end
    object LModeplanning: TLabel
      Left = 272
      Top = 13
      Width = 83
      Height = 13
      Caption = 'Type de Planning'
      FocusControl = DateDebut
      Transparent = True
    end
    object DateDebut: THCritMaskEdit
      Left = 120
      Top = 8
      Width = 73
      Height = 22
      AutoSelect = False
      AutoSize = False
      EditMask = '!99/99/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentFont = False
      TabOrder = 0
      Text = '  /  /    '
      OnChange = DateDebutChange
      TagDispatch = 0
      OpeType = otDate
      ControlerDate = True
    end
    object ParamPlanning: THCritMaskEdit
      Left = 119
      Top = 37
      Width = 58
      Height = 22
      AutoSelect = False
      AutoSize = False
      CharCase = ecUpperCase
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      OnExit = ParamPlanningExit
      TagDispatch = 0
      Operateur = Egal
      ElipsisButton = True
      OnElipsisClick = ParamPlanningElipsisClick
    end
    object ParamPlanningSauve: THCritMaskEdit
      Left = 435
      Top = 35
      Width = 87
      Height = 22
      AutoSelect = False
      AutoSize = False
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      Visible = False
      TagDispatch = 0
      Operateur = Egal
      ElipsisButton = True
    end
    object MODEPLANNING: THValComboBox
      Left = 368
      Top = 8
      Width = 153
      Height = 21
      ItemHeight = 13
      TabOrder = 3
      Text = 'MODEPLANNING'
      OnChange = MODEPLANNINGChange
      TagDispatch = 0
      DataType = 'BTMODEPLANNING'
    end
    object CBTOUS: THCheckbox
      Left = 119
      Top = 64
      Width = 57
      Height = 17
      Caption = 'Tous'
      TabOrder = 4
      OnClick = CBTOUSClick
    end
  end
  object Dock972: TDock97
    Left = 0
    Top = 97
    Width = 527
    Height = 37
    AllowDrag = False
    Position = dpBottom
    object Toolbar972: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 33
      ClientWidth = 527
      ActivateParent = False
      Caption = 'Outils Planning'
      ClientAreaHeight = 33
      ClientAreaWidth = 527
      DockPos = 0
      DragHandleStyle = dhNone
      FullSize = True
      Resizable = False
      TabOrder = 0
      DesignSize = (
        527
        33)
      object BQUITTER: TToolbarButton97
        Left = 495
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Quitter'
        Alignment = taLeftJustify
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BQUITTERClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BSaisie: TToolbarButton97
        Left = 464
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Valider'
        Alignment = taRightJustify
        Anchors = [akTop, akRight]
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Saisie'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
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
        GlyphMask.Data = {00000000}
        Layout = blGlyphTop
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BSaisieClick
        IsControl = True
      end
    end
  end
  object HSystemMenu1: THSystemMenu
    Separator = True
    Traduction = False
    Left = 16
    Top = 68
  end
end
