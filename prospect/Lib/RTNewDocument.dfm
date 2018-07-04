inherited FRTNewDocument: TFRTNewDocument
  Left = 320
  Top = 135
  Width = 504
  Height = 572
  HelpContext = 111000327
  Caption = 'Nouveau document'
  OldCreateOrder = True
  Position = poOwnerFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 499
    Width = 496
    Height = 39
    inherited PBouton: TToolWindow97
      ClientHeight = 35
      ClientWidth = 496
      ClientAreaHeight = 35
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
      object BMail: TToolbarButton97
        Left = 336
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Envoyer un message'
        ParentShowHint = False
        ShowHint = True
        OnClick = BMailClick
        GlobalIndexImage = 'Z0218_S16G1'
      end
    end
  end
  object PnlProprietes: THPanel [1]
    Left = 0
    Top = 0
    Width = 496
    Height = 499
    Align = alClient
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object LblFichier: THLabel
      Left = 16
      Top = 143
      Width = 31
      Height = 13
      Caption = 'Fichier'
      FocusControl = Fichier
    end
    object LblDossier: THLabel
      Left = 16
      Top = 28
      Width = 23
      Height = 13
      Caption = 'Tiers'
    end
    object LblProprietes: THLabel
      Left = 12
      Top = 119
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
      Top = 167
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
      Top = 194
      Width = 88
      Height = 13
      Caption = 'Date du document'
    end
    object LblNatDocument: THLabel
      Left = 16
      Top = 217
      Width = 32
      Height = 13
      Caption = 'Nature'
    end
    object LblBlocnote: THLabel
      Left = 16
      Top = 373
      Width = 61
      Height = 13
      Caption = 'Commentaire'
      FocusControl = YDO_BLOCNOTE
    end
    object LblMotsCles: THLabel
      Left = 16
      Top = 317
      Width = 45
      Height = 13
      Caption = 'Mots cl'#233's'
    end
    object BDocument: TToolbarButton97
      Left = 456
      Top = 139
      Width = 19
      Height = 21
      Hint = 'Nouvelle s'#233'lection'
      Opaque = False
      ShowBorderWhenInactive = True
      Visible = False
      OnClick = BDocumentClick
      GlobalIndexImage = 'Z0008_S16G1'
    end
    object LBLCHAINAGE: THLabel
      Left = 16
      Top = 49
      Width = 47
      Height = 13
      Caption = 'Cha'#238'nage'
    end
    object TRTD_TABLELIBREGED1: THLabel
      Left = 16
      Top = 243
      Width = 58
      Height = 13
      Caption = 'Table libre 1'
      FocusControl = RTD_TABLELIBREGED1
    end
    object TRTD_TABLELIBREGED2: THLabel
      Left = 16
      Top = 269
      Width = 58
      Height = 13
      Caption = 'Table libre 2'
      FocusControl = RTD_TABLELIBREGED2
    end
    object TRTD_TABLELIBREGED3: THLabel
      Left = 16
      Top = 293
      Width = 58
      Height = 13
      Caption = 'Table libre 3'
      FocusControl = RTD_TABLELIBREGED3
    end
    object LBLACTION: THLabel
      Left = 16
      Top = 73
      Width = 30
      Height = 13
      Caption = 'Action'
    end
    object LBLPERSPECTIVE: THLabel
      Left = 16
      Top = 99
      Width = 52
      Height = 13
      Caption = 'Proposition'
    end
    object LblPrevenir: THLabel
      Left = 16
      Top = 471
      Width = 39
      Height = 13
      Caption = 'Pr'#233'venir'
    end
    object Fichier: THCritMaskEdit
      Left = 107
      Top = 139
      Width = 348
      Height = 21
      TabOrder = 8
      OnChange = OnDocChange
      TagDispatch = 0
      DataType = 'OPENFILE(*.*;*.*)'
      ElipsisButton = True
      OnElipsisClick = FichierElipsisClick
    end
    object YDO_LIBELLEDOC: THEdit
      Left = 107
      Top = 163
      Width = 369
      Height = 21
      TabOrder = 9
      OnChange = OnDocChange
      OnExit = YDO_LIBELLEDOCExit
      TagDispatch = 0
    end
    object YDO_BLOCNOTE: THRichEditOLE
      Left = 107
      Top = 373
      Width = 369
      Height = 89
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      ScrollBars = ssVertical
      TabOrder = 16
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
      Left = 107
      Top = 313
      Width = 369
      Height = 57
      TabOrder = 15
      OnChange = OnDocChange
      OnExit = YDO_MOTSCLESExit
    end
    object RTD_DATERECEPTION: THCritMaskEdit
      Left = 107
      Top = 190
      Width = 97
      Height = 21
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 10
      Text = '  /  /    '
      OnChange = OnDocChange
      TagDispatch = 0
      OpeType = otDate
      DefaultDate = odDate
      ElipsisButton = True
      ControlerDate = True
    end
    object Dossier: THCritMaskEdit
      Left = 107
      Top = 22
      Width = 125
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 0
      OnChange = OnDocChange
      OnExit = DossierExit
      TagDispatch = 0
      DataType = 'OPENFILE(*.*;*.*)'
      ElipsisButton = True
      OnElipsisClick = DossierElipsisClick
    end
    object YDO_NATDOC: THValComboBox
      Left = 107
      Top = 213
      Width = 348
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 11
      OnChange = OnDocChange
      TagDispatch = 0
      Vide = True
      VideString = '<<Aucun>>'
      DataType = 'YYNATDOC'
      DataTypeParametrable = True
    end
    object NomDossier: THCritMaskEdit
      Left = 242
      Top = 22
      Width = 232
      Height = 21
      ReadOnly = True
      TabOrder = 1
      OnChange = OnDocChange
      TagDispatch = 0
    end
    object S_NUMCHAINAGE: THCritMaskEdit
      Left = 107
      Top = 47
      Width = 125
      Height = 21
      TabOrder = 2
      OnChange = OnDocChange
      OnExit = S_NUMCHAINAGEExit
      TagDispatch = 0
      Operateur = Egal
      ElipsisButton = True
    end
    object LIBCHAINAGE: THCritMaskEdit
      Left = 242
      Top = 47
      Width = 232
      Height = 21
      ReadOnly = True
      TabOrder = 3
      TagDispatch = 0
    end
    object RTD_TABLELIBREGED1: THValComboBox
      Left = 107
      Top = 239
      Width = 348
      Height = 21
      ItemHeight = 13
      TabOrder = 12
      OnChange = OnDocChange
      TagDispatch = 0
      Vide = True
      VideString = '<<Aucun>>'
      DataType = 'RTLIBGED1'
    end
    object RTD_TABLELIBREGED2: THValComboBox
      Left = 107
      Top = 264
      Width = 348
      Height = 21
      ItemHeight = 13
      TabOrder = 13
      OnChange = OnDocChange
      TagDispatch = 0
      Vide = True
      VideString = '<<Aucun>>'
      DataType = 'RTLIBGED2'
    end
    object RTD_TABLELIBREGED3: THValComboBox
      Left = 107
      Top = 287
      Width = 348
      Height = 21
      ItemHeight = 13
      TabOrder = 14
      OnChange = OnDocChange
      TagDispatch = 0
      Vide = True
      VideString = '<<Aucun>>'
      DataType = 'RTLIBGED3'
    end
    object S_NUMACTION: THCritMaskEdit
      Left = 107
      Top = 72
      Width = 125
      Height = 21
      TabOrder = 4
      OnChange = OnDocChange
      OnExit = S_NUMACTIONExit
      TagDispatch = 0
      DataType = 'OPENFILE(*.*;*.*)'
      Operateur = Egal
      ElipsisButton = True
      OnElipsisClick = S_NUMACTIONElipsisClick
    end
    object LIBACTION: THCritMaskEdit
      Left = 242
      Top = 72
      Width = 232
      Height = 21
      ReadOnly = True
      TabOrder = 5
      TagDispatch = 0
    end
    object S_PERSPECTIVE: THCritMaskEdit
      Left = 107
      Top = 97
      Width = 125
      Height = 21
      TabOrder = 6
      OnChange = OnDocChange
      OnExit = S_PERSPECTIVEExit
      TagDispatch = 0
      Operateur = Egal
      ElipsisButton = True
    end
    object LIBPERSPECTIVE: THCritMaskEdit
      Left = 242
      Top = 97
      Width = 232
      Height = 21
      ReadOnly = True
      TabOrder = 7
      TagDispatch = 0
    end
    object RTD_RESSOURCE: THCritMaskEdit
      Left = 107
      Top = 469
      Width = 125
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 17
      OnChange = OnDocChange
      OnExit = RTD_RESSOURCEExit
      TagDispatch = 0
      DataType = 'OPENFILE(*.*;*.*)'
      ElipsisButton = True
      OnElipsisClick = RTD_RESSOURCEElipsisClick
    end
    object NomRessource: THCritMaskEdit
      Left = 242
      Top = 469
      Width = 232
      Height = 21
      ReadOnly = True
      TabOrder = 18
      OnChange = OnDocChange
      TagDispatch = 0
    end
  end
  inherited HMTrad: THSystemMenu
    Left = 444
    Top = 12
  end
end
