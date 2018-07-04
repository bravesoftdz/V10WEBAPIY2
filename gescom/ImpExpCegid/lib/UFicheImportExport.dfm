object FexpImpCegid: TFexpImpCegid
  Left = 1094
  Top = 236
  Width = 551
  Height = 597
  Caption = 'Import - Export de donn'#233'es CEGID'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock971: TDock97
    Left = 0
    Top = 519
    Width = 535
    Height = 39
    AllowDrag = False
    Position = dpBottom
    object PBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 535
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 35
      ClientAreaWidth = 535
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        535
        35)
      object BValider: TToolbarButton97
        Left = 439
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Default = True
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ModalResult = 1
        ParentFont = False
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0127_S16G1'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 471
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Cancel = True
        Flat = False
        ModalResult = 2
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object HelpBtn: TToolbarButton97
        Left = 503
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Spacing = -1
        OnClick = HelpBtnClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BImprimer: TToolbarButton97
        Left = 407
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Anchors = [akTop, akRight]
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Visible = False
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
    end
  end
  object PGCTRL: THPageControl2
    Left = 0
    Top = 0
    Width = 535
    Height = 519
    ActivePage = TPSHTCAR
    Align = alClient
    TabOrder = 1
    object TPSHTCAR: TTabSheet
      Caption = 'Traitements'
      object Label1: TLabel
        Left = 7
        Top = 157
        Width = 174
        Height = 13
        Caption = 'R'#233'pertoire de destination des fichiers'
      end
      object Label2: TLabel
        Left = 7
        Top = 180
        Width = 108
        Height = 13
        Caption = 'Nom du fichier d'#39'export'
      end
      object Bevel1: TBevel
        Left = 2
        Top = 132
        Width = 523
        Height = 3
      end
      object Bevel2: TBevel
        Left = 4
        Top = 259
        Width = 523
        Height = 3
      end
      object Label3: TLabel
        Left = 0
        Top = 0
        Width = 527
        Height = 13
        Align = alTop
      end
      object Label5: TLabel
        Left = 0
        Top = 69
        Width = 527
        Height = 20
        Align = alTop
        Alignment = taCenter
        AutoSize = False
        Caption = 'Assurez-vous qu'#39'il reste suffisamment de place sur les disques'
        Font.Charset = ANSI_CHARSET
        Font.Color = clRed
        Font.Height = -16
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label6: TLabel
        Left = 0
        Top = 13
        Width = 527
        Height = 23
        Align = alTop
        Alignment = taCenter
        AutoSize = False
        Caption = 'ATTENTION '
        Font.Charset = ANSI_CHARSET
        Font.Color = clRed
        Font.Height = -16
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label7: TLabel
        Left = 0
        Top = 56
        Width = 527
        Height = 13
        Align = alTop
      end
      object Label4: TLabel
        Left = 0
        Top = 36
        Width = 527
        Height = 20
        Align = alTop
        Alignment = taCenter
        AutoSize = False
        Caption = 'Les temps de traitements peuvent '#234'tre tr'#232's importants'
        Font.Charset = ANSI_CHARSET
        Font.Color = clRed
        Font.Height = -16
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object EXPORTFIC: THCritMaskEdit
        Left = 217
        Top = 153
        Width = 298
        Height = 21
        TabOrder = 0
        TagDispatch = 0
        DataType = 'DIRECTORY'
        ElipsisButton = True
      end
      object NOMFIC: THCritMaskEdit
        Left = 217
        Top = 176
        Width = 288
        Height = 21
        TabOrder = 1
        TagDispatch = 0
      end
      object RDTEXPORT: TRadioButton
        Left = 13
        Top = 124
        Width = 185
        Height = 17
        Caption = 'Exports des donn'#233'es'
        Checked = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 2
        TabStop = True
        OnClick = RDTEXPORTClick
      end
      object RDTIMPORT: TRadioButton
        Left = 13
        Top = 250
        Width = 185
        Height = 17
        Caption = 'Import des donn'#233'es'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 3
        OnClick = RDTIMPORTClick
      end
      object GroupBox1: TGroupBox
        Left = 5
        Top = 284
        Width = 515
        Height = 57
        TabOrder = 4
        object Label8: TLabel
          Left = 7
          Top = 24
          Width = 155
          Height = 13
          Caption = 'Nom du fichier d'#39'export '#224' int'#233'grer'
        end
        object IMPORTFIC: THCritMaskEdit
          Left = 199
          Top = 20
          Width = 298
          Height = 21
          TabOrder = 0
          TagDispatch = 0
          DataType = 'OPENFILE(*.ZIP)'
          ElipsisButton = True
        end
        object CHKZIP: TCheckBox
          Left = 8
          Top = 0
          Width = 105
          Height = 17
          Caption = 'Via un fichier Zip'
          Checked = True
          State = cbChecked
          TabOrder = 1
          OnClick = CHKZIPClick
        end
      end
      object GroupBox2: TGroupBox
        Left = 4
        Top = 344
        Width = 515
        Height = 57
        TabOrder = 5
        object Label9: TLabel
          Left = 7
          Top = 23
          Width = 167
          Height = 13
          Caption = 'R'#233'pertoire de stockage des fichiers'
        end
        object CHKREPERT: TCheckBox
          Left = 10
          Top = 0
          Width = 120
          Height = 17
          Caption = 'Depuis un r'#233'pertoire'
          TabOrder = 0
          OnClick = CHKREPERTClick
        end
        object IMPORTDIR: THCritMaskEdit
          Left = 201
          Top = 19
          Width = 298
          Height = 21
          Enabled = False
          TabOrder = 1
          TagDispatch = 0
          DataType = 'DIRECTORY'
          ElipsisButton = True
        end
      end
      object CHKMSGSTRUCT: TCheckBox
        Left = 8
        Top = 410
        Width = 337
        Height = 17
        Alignment = taLeftJustify
        Caption = 
          'Demande de confirmation de traitement sur structures diff'#233'rentes' +
          ' ?'
        Enabled = False
        TabOrder = 6
      end
      object CHKSTOPONERROR: TCheckBox
        Left = 8
        Top = 431
        Width = 337
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Arr'#234't sur anomalie d'#39'int'#233'gration ?'
        Checked = True
        Enabled = False
        State = cbChecked
        TabOrder = 7
      end
      object CBVIDAGEEXP: TCheckBox
        Left = 8
        Top = 200
        Width = 337
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Vidage du r'#233'pertoire'
        Checked = True
        State = cbChecked
        TabOrder = 8
      end
      object CBVIDAGEIMP: TCheckBox
        Left = 8
        Top = 451
        Width = 337
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Vidage du r'#233'pertoire'
        Checked = True
        Enabled = False
        State = cbChecked
        TabOrder = 9
      end
    end
    object TBSHTCONTROL: TTabSheet
      Caption = 'Cont'#244'les'
      ImageIndex = 1
      object Trace: TListBox
        Left = 0
        Top = 0
        Width = 527
        Height = 491
        Align = alClient
        ItemHeight = 13
        TabOrder = 0
      end
    end
  end
end
