inherited FAssistantPieceEpure: TFAssistantPieceEpure
  Left = 142
  Top = 409
  Caption = 'Assistant Epuration des pi'#232'ces'
  ClientHeight = 386
  ClientWidth = 616
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Top = 363
  end
  inherited lAide: THLabel
    Left = 184
    Top = 333
    Width = 428
  end
  inherited bPrecedent: TToolbarButton97
    Left = 370
    Top = 357
  end
  inherited bSuivant: TToolbarButton97
    Left = 453
    Top = 357
  end
  inherited bFin: TToolbarButton97
    Left = 537
    Top = 357
  end
  inherited bAnnuler: TToolbarButton97
    Left = 287
    Top = 357
  end
  inherited bAide: TToolbarButton97
    Left = 168
    Top = 357
  end
  object bNote: TToolbarButton97 [7]
    Tag = 1
    Left = 249
    Top = 355
    Width = 28
    Height = 27
    Hint = 'Informations compl'#233'mentaires'
    AllowAllUp = True
    GroupIndex = 1
    DisplayMode = dmGlyphOnly
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Layout = blGlyphTop
    Opaque = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    Spacing = -1
    OnClick = bNoteClick
    GlobalIndexImage = 'Z0029_S16G1'
    IsControl = True
  end
  inherited Plan: THPanel
    Left = 180
    Top = 29
  end
  inherited GroupBox1: THGroupBox
    Top = 345
    Width = 619
  end
  inherited P: THPageControl2
    Left = 177
    Top = 0
    Width = 433
    Height = 321
    ActivePage = TabGrid
    object TabGrid: TTabSheet
      Caption = 'TabGrid'
      object PGENERAL: THPanel
        Left = 0
        Top = 0
        Width = 425
        Height = 293
        Align = alClient
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object PTITRE: THPanel
          Left = 1
          Top = 1
          Width = 423
          Height = 41
          Align = alTop
          Caption = 'Epuration des pi'#232'ces'
          FullRepaint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clActiveCaption
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 0
          BackGroundEffect = bdFlat
          ColorShadow = clWindowText
          ColorStart = clBtnFace
          TextEffect = tenone
        end
        object GListePieces: THGrid
          Left = 1
          Top = 42
          Width = 423
          Height = 250
          Align = alClient
          DefaultColWidth = 108
          DefaultRowHeight = 18
          FixedCols = 0
          Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goRowSelect]
          TabOrder = 1
          OnDblClick = GListePiecesDblClick
          SortedCol = -1
          Couleur = False
          MultiSelect = False
          TitleBold = False
          TitleCenter = True
          OnRowEnter = GListePiecesRowEnter
          ColCombo = 0
          SortEnabled = True
          SortRowExclude = 0
          TwoColors = False
          AlternateColor = clSilver
          ColWidths = (
            29
            98
            115
            127
            42)
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'TabSheet2'
      ImageIndex = 1
      object TLBRecapitulatif: TListBox
        Left = 0
        Top = 0
        Width = 425
        Height = 293
        Align = alClient
        ItemHeight = 13
        TabOrder = 0
      end
    end
  end
  object TDescriptif: TToolWindow97 [13]
    Left = 46
    Top = 178
    ClientHeight = 93
    ClientWidth = 235
    Caption = 'Informations compl'#233'mentaires'
    ClientAreaHeight = 93
    ClientAreaWidth = 235
    TabOrder = 5
    Visible = False
    OnClose = TDescriptifClose
    object Descriptif: THRichEditOLE
      Left = 0
      Top = 0
      Width = 235
      Height = 93
      Align = alClient
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      Margins.Top = 0
      Margins.Bottom = 0
      Margins.Left = 0
      Margins.Right = 0
      ContainerName = 'Document'
      ObjectMenuPrefix = '&Object'
      LinesRTF.Strings = (
        
          '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fnil Aria' +
          'l;}}'
        
          '{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1\pard\f0\fs16' +
          ' '
        '\par }')
    end
  end
end
