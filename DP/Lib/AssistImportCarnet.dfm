inherited FAssistImportCarnet: TFAssistImportCarnet
  Caption = 'Assistant importation de carnet d'#39'adresses'
  ClientWidth = 540
  PixelsPerInch = 96
  TextHeight = 13
  inherited lAide: THLabel
    Width = 352
  end
  inherited bPrecedent: TToolbarButton97
    Left = 273
  end
  inherited bSuivant: TToolbarButton97
    Left = 356
  end
  inherited bFin: TToolbarButton97
    Left = 440
  end
  inherited bAnnuler: TToolbarButton97
    Left = 190
  end
  inherited bAide: TToolbarButton97
    Left = 107
  end
  inherited GroupBox1: TGroupBox
    Width = 543
  end
  inherited P: TPageControl
    Width = 353
    ActivePage = PChoixFichier
    object PChoixFichier: TTabSheet
      Caption = 'Choix du fichier'
      object lblSelection: THLabel
        Left = 8
        Top = 12
        Width = 244
        Height = 13
        Caption = 'Importation d'#39'un carnet d'#39'adresses Outlook'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold, fsUnderline]
        ParentFont = False
      end
      object lblAvert1: THLabel
        Left = 14
        Top = 55
        Width = 303
        Height = 34
        AutoSize = False
        Caption = 
          'Vous allez importer un carnet d'#39'adresses issu de Outlook ou Outl' +
          'ook Express.'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        WordWrap = True
      end
      object lblAvert2: THLabel
        Left = 14
        Top = 99
        Width = 303
        Height = 34
        AutoSize = False
        Caption = 
          'Les enregistrements trouvés seront ajoutés dans le répertoire pe' +
          'rsonnel.'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        WordWrap = True
      end
      object lblAvert3: THLabel
        Left = 14
        Top = 139
        Width = 299
        Height = 54
        AutoSize = False
        Caption = 
          'L'#39'export doit avoir été fait dans Outlook (ou Outlook Express) e' +
          'n choisissant format texte (Windows).'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        WordWrap = True
      end
    end
    object PImportation: TTabSheet
      Caption = 'Importation'
      ImageIndex = 1
      object lblFichier: THLabel
        Left = 12
        Top = 48
        Width = 78
        Height = 13
        Caption = 'Choix du fichier :'
      end
      object HLabel1: THLabel
        Left = 8
        Top = 12
        Width = 244
        Height = 13
        Caption = 'Importation d'#39'un carnet d'#39'adresses Outlook'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold, fsUnderline]
        ParentFont = False
      end
      object lblAide: THLabel
        Left = 12
        Top = 112
        Width = 208
        Height = 13
        Caption = 'Cliquer sur <Fin> pour effectuer l'#39'importation.'
      end
      object LeFichier: THCritMaskEdit
        Left = 12
        Top = 68
        Width = 313
        Height = 21
        TabOrder = 0
        TagDispatch = 0
        DataType = 'OPENFILE(*.TXT;*.CSV;*.*)'
        ElipsisButton = True
      end
    end
  end
end
