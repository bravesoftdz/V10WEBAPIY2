object FSaveRestore: TFSaveRestore
  Left = 282
  Top = 138
  Width = 640
  Height = 480
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 261
    Top = 0
    Height = 446
  end
  object Panel3: TPanel
    Left = 264
    Top = 0
    Width = 368
    Height = 446
    Align = alClient
    TabOrder = 0
    object Panel4: TPanel
      Left = 1
      Top = 1
      Width = 366
      Height = 164
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 0
      object LSave: TLabel
        Left = 8
        Top = 97
        Width = 135
        Height = 13
        Caption = 'Sauvegarder dans le fichier :'
      end
      object LRest: TLabel
        Left = 8
        Top = 97
        Width = 90
        Height = 13
        Caption = 'Fichier '#224' restaurer :'
        Visible = False
      end
      object Label1: TLabel
        Left = 8
        Top = 63
        Width = 124
        Height = 13
        Caption = 'Param'#232'tres '#224' sauvegarder'
      end
      object Ldomaine: TLabel
        Left = 8
        Top = 23
        Width = 42
        Height = 13
        Caption = 'Domaine'
      end
      object EdFicName: TEdit
        Left = 8
        Top = 113
        Width = 281
        Height = 21
        TabOrder = 0
      end
      object BitBtn1: THBitBtn
        Left = 289
        Top = 113
        Width = 23
        Height = 21
        Caption = '...'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -21
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 1
        OnClick = BitBtn1Click
      end
      object CBTables: TComboBox
        Left = 144
        Top = 59
        Width = 167
        Height = 21
        ItemHeight = 13
        TabOrder = 2
        OnChange = CBTablesChange
        Items.Strings = (
          'Domaine'
          'Dictionnaire'
          'Scripts'
          'Correspondance')
      end
      object RGSR: TRadioGroup
        Left = 4
        Top = 136
        Width = 214
        Height = 27
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          '&Sauvegarde'
          '&Restauration')
        TabOrder = 3
        Visible = False
        OnClick = RGSRClick
      end
      object Domaine: THValComboBox
        Left = 144
        Top = 20
        Width = 125
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 4
        TagDispatch = 0
        ComboWidth = 300
      end
    end
    object Panel5: TPanel
      Left = 1
      Top = 165
      Width = 366
      Height = 280
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 1
      object lblTraitement: TLabel
        Left = 146
        Top = 184
        Width = 90
        Height = 13
        Caption = 'Label de traitement'
      end
      object Panel6: TPanel
        Left = 0
        Top = 243
        Width = 366
        Height = 37
        Align = alBottom
        TabOrder = 0
        DesignSize = (
          366
          37)
        object bOuvrir: TToolbarButton97
          Left = 304
          Top = 5
          Width = 28
          Height = 27
          Hint = 'Valider les modifications'
          AllowAllUp = True
          Anchors = [akTop, akRight]
          Flat = False
          ParentShowHint = False
          ShowHint = True
          OnClick = bOuvrirClick
          GlobalIndexImage = 'Z0127_S16G1'
        end
        object BAide: TToolbarButton97
          Left = 334
          Top = 5
          Width = 28
          Height = 27
          Hint = 'Aide'
          Anchors = [akTop, akRight]
          DisplayMode = dmGlyphOnly
          Caption = 'Aide'
          Flat = False
          Layout = blGlyphTop
          ParentShowHint = False
          ShowHint = True
          OnClick = BAideClick
          GlobalIndexImage = 'Z1117_S16G1'
        end
      end
      object Chart: TChart
        Left = 84
        Top = 44
        Width = 197
        Height = 125
        AllowPanning = pmNone
        AllowZoom = False
        BackWall.Brush.Color = clWhite
        BackWall.Brush.Style = bsClear
        BackWall.Color = clSilver
        BackWall.Pen.Visible = False
        Title.Text.Strings = (
          'TChart')
        Title.Visible = False
        AxisVisible = False
        BackColor = clSilver
        Chart3DPercent = 30
        ClipPoints = False
        Frame.Visible = False
        Legend.Visible = False
        View3DOptions.Elevation = 315
        View3DOptions.Orthogonal = False
        View3DOptions.Perspective = 0
        View3DOptions.Rotation = 360
        View3DWalls = False
        BevelOuter = bvNone
        TabOrder = 1
        Visible = False
        object Series1: TPieSeries
          Marks.ArrowLength = 8
          Marks.Style = smsPercent
          Marks.Visible = False
          SeriesColor = clRed
          ShowInLegend = False
          OtherSlice.Text = 'Autre'
          PieValues.DateTime = False
          PieValues.Name = 'Camembert'
          PieValues.Multiplier = 1.000000000000000000
          PieValues.Order = loNone
        end
      end
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 261
    Height = 446
    Align = alLeft
    Caption = 'Panel1'
    TabOrder = 1
    object TV: TTreeView
      Left = 1
      Top = 49
      Width = 259
      Height = 396
      Align = alClient
      HideSelection = False
      Images = ImageList1
      Indent = 19
      StateImages = ImageList1
      TabOrder = 0
      OnChange = TVChange
      OnDblClick = TVDblClick
    end
    object Panel2: TPanel
      Left = 1
      Top = 1
      Width = 259
      Height = 48
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 1
      object Label2: TLabel
        Left = 24
        Top = 8
        Width = 194
        Height = 26
        Caption = 'Selectionner les '#233'l'#233'ments '#224' sauvegarder par double clic'
        WordWrap = True
      end
    end
  end
  object OpenDialog: TOpenDialog
    Filter = 'Fichiers Sauvegarde (*.cix)|*.CIX'
    Left = 65
    Top = 86
  end
  object ImageList1: THImageList
    GlobalIndexImages.Strings = (
      'Z2396_S16G1'
      'Z0007_S16G1'
      'Z0007_S16G1'
      'Z0437_S16G1'
      'Z0033_S16G1'
      'Z0034_S16G1'
      'Z0034_S16G1'
      'Z2395_S16G1'
      'Z0007_S16G1'
      'Z0241_S16G1'
      'Z0145_S16G1'
      'Z0152_S16G1'
      'Z2415_S16G1')
    Left = 107
    Top = 79
  end
end
