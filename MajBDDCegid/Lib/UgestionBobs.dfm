object Form2: TForm2
  Left = 326
  Top = 180
  Width = 870
  Height = 669
  Caption = 'Form2'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object HPanel1: THPanel
    Left = 0
    Top = 0
    Width = 854
    Height = 113
    Align = alTop
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object Label1: TLabel
      Left = 32
      Top = 15
      Width = 67
      Height = 13
      Caption = 'Emplacement '
    end
    object ToolbarButton972: TToolbarButton97
      Left = 533
      Top = 10
      Width = 23
      Height = 22
      OnClick = ToolbarButton972Click
      GlobalIndexImage = 'O0001_S24G1'
    end
    object LanceTrait: TToolbarButton97
      Left = 818
      Top = 51
      Width = 26
      Height = 26
      OnClick = LanceTraitClick
      GlobalIndexImage = 'M0003_S24G1'
    end
    object Label2: TLabel
      Left = 32
      Top = 56
      Width = 31
      Height = 13
      Caption = 'Fichier'
    end
    object BfindFile: TToolbarButton97
      Left = 533
      Top = 51
      Width = 23
      Height = 22
      OnClick = BfindFileClick
      GlobalIndexImage = 'M0026_S16G1'
    end
    object EMPLACE: TEdit97
      Left = 152
      Top = 9
      Width = 377
      Height = 25
      TabOrder = 0
    end
    object FILENAME: TEdit97
      Left = 152
      Top = 50
      Width = 377
      Height = 25
      TabOrder = 1
    end
    object GroupBox1: TGroupBox
      Left = 568
      Top = 2
      Width = 265
      Height = 33
      Caption = 'Type'
      TabOrder = 2
      object RBBOB: TRadioButton
        Left = 8
        Top = 14
        Width = 65
        Height = 17
        Caption = 'BOB'
        Checked = True
        TabOrder = 0
        TabStop = True
      end
      object RBMDL: TRadioButton
        Left = 88
        Top = 14
        Width = 113
        Height = 17
        Caption = 'MDL'
        TabOrder = 1
      end
    end
  end
  object HPanel2: THPanel
    Left = 0
    Top = 113
    Width = 854
    Height = 479
    Align = alClient
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object GS: THGrid
      Left = 0
      Top = 0
      Width = 854
      Height = 479
      Align = alClient
      ColCount = 3
      DefaultRowHeight = 18
      RowCount = 3
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
      TabOrder = 0
      SortedCol = -1
      Couleur = False
      MultiSelect = True
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = clSilver
    end
  end
  object HPanel3: THPanel
    Left = 0
    Top = 592
    Width = 854
    Height = 39
    Align = alBottom
    FullRepaint = False
    TabOrder = 2
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    DesignSize = (
      854
      39)
    object TBSelect: TToolbarButton97
      Left = 44
      Top = 5
      Width = 26
      Height = 26
      Hint = 'Tout S'#233'lectionner'
      ParentShowHint = False
      ShowHint = True
      OnClick = TBSelectClick
      GlobalIndexImage = 'Z0205_S24G1'
    end
    object BSaveMDL: TToolbarButton97
      Left = 787
      Top = 5
      Width = 26
      Height = 26
      Hint = 'G'#233'n'#233'rer le Mod'#232'le'
      ParentShowHint = False
      ShowHint = True
      OnClick = BSaveMDLClick
      GlobalIndexImage = 'O0057_S24G1'
    end
    object ToolbarButton971: TToolbarButton97
      Left = 818
      Top = 5
      Width = 26
      Height = 26
      Anchors = [akRight, akBottom]
      Cancel = True
      OnClick = ToolbarButton971Click
      GlobalIndexImage = 'M0002_S24G1'
    end
    object BGENERBOB: TToolbarButton97
      Left = 752
      Top = 5
      Width = 26
      Height = 26
      OnClick = BGENERBOBClick
      GlobalIndexImage = 'M0006_S16G1'
    end
  end
  object OfindFile: TOpenDialog
    DefaultExt = '*.BOB'
    Filter = 'Fichiers Bob|*.BOB|Fichiers Mod'#232'les|*.MDL'
    Left = 568
    Top = 48
  end
  object HmTrad: THSystemMenu
    ResizeDBGrid = True
    Separator = True
    Traduction = False
    Left = 600
    Top = 48
  end
  object SaveMDL: TSaveDialog
    DefaultExt = '*.MDL'
    Filter = 'Mod'#232'le|*.MDL'
    Left = 576
    Top = 80
  end
  object SauveBOB: TSaveDialog
    DefaultExt = '*.BOB'
    Filter = 'Fichier BOB|*.BOB'
    Left = 616
    Top = 80
  end
end
