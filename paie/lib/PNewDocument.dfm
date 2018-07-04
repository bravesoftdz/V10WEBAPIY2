inherited FPNewDocument: TFPNewDocument
  Left = 384
  Top = 219
  Width = 504
  Height = 469
  Caption = 'Nouveau document'
  OldCreateOrder = True
  Position = poOwnerFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 392
    Width = 496
    Height = 43
    inherited PBouton: TToolWindow97
      ClientHeight = 39
      ClientWidth = 496
      ClientAreaHeight = 39
      ClientAreaWidth = 496
      inherited BValider: TToolbarButton97
        Left = 400
        ParentShowHint = False
        ShowHint = True
      end
      inherited BFerme: TToolbarButton97
        Left = 432
        ParentShowHint = False
        ShowHint = True
      end
      inherited HelpBtn: TToolbarButton97
        Left = 464
        ParentShowHint = False
        ShowHint = True
      end
      inherited BImprimer: TToolbarButton97
        Left = 368
      end
    end
  end
  object PnlProprietes: THPanel [1]
    Left = 0
    Top = 0
    Width = 496
    Height = 392
    Align = alClient
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object LblFichier: THLabel
      Left = 16
      Top = 69
      Width = 31
      Height = 13
      Caption = 'Fichier'
      FocusControl = Fichier
    end
    object LblDossier: THLabel
      Left = 16
      Top = 28
      Width = 32
      Height = 13
      Caption = 'Salari'#233
    end
    object LblProprietes: THLabel
      Left = 12
      Top = 47
      Width = 135
      Height = 13
      Caption = 'Propri'#233't'#233's du document'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsUnderline]
      ParentFont = False
    end
    object LblLibelleDoc: THLabel
      Left = 16
      Top = 94
      Width = 53
      Height = 13
      Caption = 'Description'
      FocusControl = YDO_LIBELLEDOC
    end
    object LblEmplacement: THLabel
      Left = 12
      Top = 3
      Width = 31
      Height = 13
      Caption = 'Liens'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsUnderline]
      ParentFont = False
    end
    object LblDateDeb: THLabel
      Left = 16
      Top = 119
      Width = 88
      Height = 13
      Caption = 'Date du document'
    end
    object LblNatDocument: THLabel
      Left = 16
      Top = 144
      Width = 32
      Height = 13
      Caption = 'Nature'
    end
    object LblBlocnote: THLabel
      Left = 16
      Top = 301
      Width = 61
      Height = 13
      Caption = 'Commentaire'
      FocusControl = YDO_BLOCNOTE
    end
    object LblMotsCles: THLabel
      Left = 16
      Top = 244
      Width = 45
      Height = 13
      Caption = 'Mots cl'#233's'
    end
    object BDocument: TToolbarButton97
      Left = 456
      Top = 65
      Width = 19
      Height = 21
      Hint = 'Nouvelle s'#233'lection'
      Glyph.Data = {
        42010000424D4201000000000000760000002800000011000000110000000100
        040000000000CC000000CE0E0000D80E00001000000000000000000000000000
        BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
        77777000000070000000000000077000000070FFF0FFFFFF0F077000000070F0
        F0F0FF0F0F077000000070000000000000077000000077709999999077777000
        0000777090000907777770000000777090709007777770000000777090099700
        77777000000077709099070007777000000077709990770BB077700000007770
        9907770BB07770000000777090777770BB0770000000777007777770B0077000
        00007770777777770BB070000000777777777777000770000000777777777777
        777770000000}
      GlyphMask.Data = {00000000}
      Opaque = False
      ShowBorderWhenInactive = True
      Visible = False
      OnClick = BDocumentClick
    end
    object TRTD_TABLELIBREGED1: THLabel
      Left = 16
      Top = 169
      Width = 58
      Height = 13
      Caption = 'Table libre 1'
      FocusControl = RTD_TABLELIBREGED1
    end
    object TRTD_TABLELIBREGED2: THLabel
      Left = 16
      Top = 194
      Width = 58
      Height = 13
      Caption = 'Table libre 2'
      FocusControl = RTD_TABLELIBREGED2
    end
    object TRTD_TABLELIBREGED3: THLabel
      Left = 16
      Top = 221
      Width = 58
      Height = 13
      Caption = 'Table libre 3'
      FocusControl = RTD_TABLELIBREGED3
    end
    object Fichier: THCritMaskEdit
      Left = 127
      Top = 65
      Width = 328
      Height = 21
      TabOrder = 2
      OnChange = OnDocChange
      TagDispatch = 0
      DataType = 'OPENFILE(*.*;*.*)'
      ElipsisButton = True
      OnElipsisClick = FichierElipsisClick
    end
    object YDO_LIBELLEDOC: THEdit
      Left = 127
      Top = 90
      Width = 349
      Height = 21
      TabOrder = 3
      OnChange = OnDocChange
      OnExit = YDO_LIBELLEDOCExit
      TagDispatch = 0
    end
    object YDO_BLOCNOTE: THRichEditOLE
      Left = 127
      Top = 301
      Width = 349
      Height = 89
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      ScrollBars = ssVertical
      TabOrder = 10
      OnChange = OnDocChange
      Margins.Top = 0
      Margins.Bottom = 0
      Margins.Left = 0
      Margins.Right = 0
      ContainerName = 'Document'
      ObjectMenuPrefix = '&Object'
      LinesRTF.Strings = (
        '{\rtf1\ansi\deff0{\fonttbl{\f0\fnil Arial;}}'
        
          '{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1\pard\lang103' +
          '6\f0\fs16 '
        '\par }')
    end
    object YDO_MOTSCLES: THMemo
      Left = 127
      Top = 240
      Width = 349
      Height = 57
      TabOrder = 9
      OnChange = OnDocChange
      OnExit = YDO_MOTSCLESExit
    end
    object RTD_DATERECEPTION: THCritMaskEdit
      Left = 127
      Top = 115
      Width = 97
      Height = 21
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 4
      Text = '  /  /    '
      OnChange = OnDocChange
      TagDispatch = 0
      OpeType = otDate
      DefaultDate = odDate
      ElipsisButton = True
      ControlerDate = True
    end
    object Dossier: THCritMaskEdit
      Left = 127
      Top = 22
      Width = 125
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 0
      OnChange = OnDocChange
      OnExit = DossierExit
      TagDispatch = 0
      ElipsisButton = True
      OnElipsisClick = DossierElipsisClick
    end
    object YDO_NATDOC: THValComboBox
      Left = 127
      Top = 140
      Width = 328
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 5
      OnChange = OnDocChange
      TagDispatch = 0
      Vide = True
      VideString = '<<Aucun>>'
      DataType = 'YYNATDOC'
      DataTypeParametrable = True
    end
    object NomDossier: THCritMaskEdit
      Left = 256
      Top = 22
      Width = 218
      Height = 21
      ReadOnly = True
      TabOrder = 1
      OnChange = OnDocChange
      TagDispatch = 0
    end
    object RTD_TABLELIBREGED1: THValComboBox
      Left = 127
      Top = 165
      Width = 328
      Height = 21
      ItemHeight = 13
      TabOrder = 6
      OnChange = OnDocChange
      TagDispatch = 0
      Vide = True
      VideString = '<<Aucun>>'
      DataType = 'RTLIBGED1'
    end
    object RTD_TABLELIBREGED2: THValComboBox
      Left = 127
      Top = 190
      Width = 328
      Height = 21
      ItemHeight = 13
      TabOrder = 7
      OnChange = OnDocChange
      TagDispatch = 0
      Vide = True
      VideString = '<<Aucun>>'
      DataType = 'RTLIBGED2'
    end
    object RTD_TABLELIBREGED3: THValComboBox
      Left = 127
      Top = 215
      Width = 328
      Height = 21
      ItemHeight = 13
      TabOrder = 8
      OnChange = OnDocChange
      TagDispatch = 0
      Vide = True
      VideString = '<<Aucun>>'
      DataType = 'RTLIBGED3'
    end
  end
  inherited HMTrad: THSystemMenu
    Left = 444
    Top = 12
  end
end
