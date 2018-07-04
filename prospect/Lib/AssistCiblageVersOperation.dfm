inherited FAssistCiblageVersOperation: TFAssistCiblageVersOperation
  Left = 223
  Top = 219
  HelpContext = 111000383
  Caption = 'Assistant de passage d'#39'un ciblage vers une op'#233'ration'
  ClientHeight = 389
  ClientWidth = 678
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Top = 366
    Caption = 'Etape 1/5'
  end
  inherited lAide: THLabel
    Top = 336
    Width = 490
  end
  inherited bPrecedent: TToolbarButton97
    Left = 411
    Top = 360
  end
  inherited bSuivant: TToolbarButton97
    Left = 494
    Top = 360
  end
  inherited bFin: TToolbarButton97
    Left = 578
    Top = 360
    Caption = '&D'#233'marrer'
    Enabled = False
  end
  inherited bAnnuler: TToolbarButton97
    Left = 328
    Top = 360
  end
  inherited bAide: TToolbarButton97
    Left = 245
    Top = 360
  end
  object TMODELE_OPERATION: THLabel [7]
    Left = 48
    Top = 300
    Width = 84
    Height = 13
    Caption = 'Mod'#232'le &Op'#233'ration'
    FocusControl = MODELE_OPERATION
  end
  object HLabel1: THLabel [8]
    Left = 263
    Top = 136
    Width = 58
    Height = 13
    Caption = 'Table libre 1'
  end
  inherited GroupBox1: THGroupBox
    Top = 348
    Width = 681
  end
  inherited P: THPageControl2
    Width = 501
    Height = 324
    ActivePage = Operation
    object Operation: TTabSheet
      Caption = 'Op'#233'ration'
      object TROP_OPERATION: THLabel
        Left = 12
        Top = 12
        Width = 74
        Height = 13
        Caption = '&Code Op'#233'ration'
        FocusControl = ROP_OPERATION
      end
      object TROP_LIBELLE: THLabel
        Left = 12
        Top = 36
        Width = 30
        Height = 13
        Caption = '&Libell'#233
        FocusControl = ROP_LIBELLE
      end
      object TROP_DATEDEBUT: THLabel
        Left = 12
        Top = 60
        Width = 55
        Height = 13
        Caption = 'Date &D'#233'but'
        FocusControl = ROP_DATEDEBUT
      end
      object TROP_DATEFIN: THLabel
        Left = 12
        Top = 84
        Width = 40
        Height = 13
        Caption = 'Date &Fin'
        FocusControl = ROP_DATEFIN
      end
      object TROP_BUDGET: THLabel
        Left = 12
        Top = 108
        Width = 34
        Height = 13
        Caption = '&Budget'
        FocusControl = ROP_BUDGET
      end
      object TROP_COUT: THLabel
        Left = 12
        Top = 132
        Width = 22
        Height = 13
        Caption = '&Co'#251't'
        FocusControl = ROP_COUT
      end
      object TROP_BLOCNOTE: THLabel
        Left = 12
        Top = 176
        Width = 29
        Height = 13
        Caption = '&M'#233'mo'
        FocusControl = ROP_BLOCNOTE
      end
      object TROP_OBJETOPE: THLabel
        Left = 12
        Top = 156
        Width = 25
        Height = 13
        Caption = '&Objet'
      end
      object ROP_OPERATION: THCritMaskEdit
        Left = 155
        Top = 8
        Width = 150
        Height = 21
        MaxLength = 17
        TabOrder = 0
        TagDispatch = 0
      end
      object ROP_LIBELLE: THCritMaskEdit
        Left = 155
        Top = 32
        Width = 150
        Height = 21
        MaxLength = 35
        TabOrder = 1
        TagDispatch = 0
      end
      object ROP_DATEDEBUT: THCritMaskEdit
        Left = 155
        Top = 56
        Width = 69
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 2
        Text = '  /  /    '
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = od2099
        ControlerDate = True
      end
      object ROP_DATEFIN: THCritMaskEdit
        Left = 155
        Top = 80
        Width = 69
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 3
        Text = '  /  /    '
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = od2099
        ControlerDate = True
      end
      object ROP_BUDGET: THCritMaskEdit
        Left = 155
        Top = 104
        Width = 69
        Height = 21
        TabOrder = 4
        TagDispatch = 0
        OpeType = otReel
      end
      object ROP_COUT: THCritMaskEdit
        Left = 155
        Top = 128
        Width = 69
        Height = 21
        TabOrder = 5
        TagDispatch = 0
        OpeType = otReel
        ControlerDate = True
      end
      object ROP_BLOCNOTE: THRichEditOLE
        Left = 4
        Top = 192
        Width = 457
        Height = 97
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 7
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
      object ROP_OBJETOPE: THValComboBox
        Left = 155
        Top = 152
        Width = 145
        Height = 21
        ItemHeight = 13
        TabOrder = 6
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'RTOBJETOPE'
      end
    end
    object TsInfoCompl: TTabSheet
      Caption = 'Infos compl'#233'mentaires'
      ImageIndex = 3
      object TROP_ROPTABLELIBRE1: THLabel
        Left = 20
        Top = 26
        Width = 58
        Height = 13
        Caption = 'Table libre 1'
        FocusControl = ROP_ROPTABLELIBRE1
      end
      object TROP_ROPTABLELIBRE2: THLabel
        Left = 20
        Top = 60
        Width = 58
        Height = 13
        Caption = 'Table libre 2'
        FocusControl = ROP_ROPTABLELIBRE2
      end
      object TROP_ROPTABLELIBRE3: THLabel
        Left = 20
        Top = 95
        Width = 58
        Height = 13
        Caption = 'Table libre 3'
        FocusControl = ROP_ROPTABLELIBRE3
      end
      object TROP_ROPTABLELIBRE4: THLabel
        Left = 20
        Top = 129
        Width = 58
        Height = 13
        Caption = 'Table libre 4'
        FocusControl = ROP_ROPTABLELIBRE4
      end
      object TROP_ROPTABLELIBRE5: THLabel
        Left = 20
        Top = 164
        Width = 58
        Height = 13
        Caption = 'Table libre 5'
        FocusControl = ROP_ROPTABLELIBRE5
      end
      object ROP_ROPTABLELIBRE1: THValComboBox
        Left = 173
        Top = 22
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'RTRROPTABLELIBRE1'
      end
      object ROP_ROPTABLELIBRE2: THValComboBox
        Left = 173
        Top = 56
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'RTRROPTABLELIBRE2'
      end
      object ROP_ROPTABLELIBRE3: THValComboBox
        Left = 173
        Top = 91
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'RTRROPTABLELIBRE3'
      end
      object ROP_ROPTABLELIBRE4: THValComboBox
        Left = 173
        Top = 125
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'RTRROPTABLELIBRE4'
      end
      object ROP_ROPTABLELIBRE5: THValComboBox
        Left = 173
        Top = 160
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 4
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'RTRROPTABLELIBRE5'
      end
    end
    object Actiongenerique: TTabSheet
      Caption = 'Action g'#233'n'#233'rique'
      ImageIndex = 1
      object TRAG_TYPEACTION: THLabel
        Left = 12
        Top = 20
        Width = 80
        Height = 13
        Caption = '&Action g'#233'n'#233'rique'
        FocusControl = RAG_TYPEACTION
      end
      object TRAG_LIBELLE: THLabel
        Left = 12
        Top = 48
        Width = 30
        Height = 13
        Caption = '&Libell'#233
        FocusControl = RAG_LIBELLE
      end
      object TRAG_DATEACTION: THLabel
        Left = 12
        Top = 76
        Width = 55
        Height = 13
        Caption = '&Date action'
        FocusControl = RAG_DATEACTION
      end
      object TRAG_DATEECHEANCE: THLabel
        Left = 12
        Top = 128
        Width = 74
        Height = 13
        Caption = 'Date &'#233'ch&'#233'ance'
        FocusControl = RAG_DATEECHEANCE
      end
      object TRAG_COUTACTION: THLabel
        Left = 12
        Top = 156
        Width = 22
        Height = 13
        Caption = '&Co'#251't'
        FocusControl = RAG_COUTACTION
      end
      object BAjoute: TToolbarButton97
        Left = 328
        Top = 264
        Width = 28
        Height = 27
        OnClick = BAjouteClick
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BSupprime: TToolbarButton97
        Left = 360
        Top = 264
        Width = 28
        Height = 27
        OnClick = BSupprimeClick
        GlobalIndexImage = 'M0009_S16G1'
      end
      object BValider: TToolbarButton97
        Left = 392
        Top = 264
        Width = 28
        Height = 27
        OnClick = BValiderClick
        GlobalIndexImage = 'M0012_S16G1'
      end
      object BAnnule: TToolbarButton97
        Left = 424
        Top = 264
        Width = 28
        Height = 27
        OnClick = BAnnuleClick
        GlobalIndexImage = 'M0002_S16G1'
      end
      object bMonter: TToolbarButton97
        Left = 268
        Top = 120
        Width = 28
        Height = 27
        OnClick = bMonterClick
        GlobalIndexImage = 'Z0171_S16G1'
      end
      object bDescendre: TToolbarButton97
        Left = 268
        Top = 148
        Width = 28
        Height = 27
        OnClick = bDescendreClick
        GlobalIndexImage = 'Z0319_S16G1'
      end
      object TRAG_TABLELIBRE1: THLabel
        Left = 12
        Top = 188
        Width = 58
        Height = 13
        Caption = 'Table libre 1'
        FocusControl = RAG_TABLELIBRE1
      end
      object TRAG_TABLELIBRE2: THLabel
        Left = 12
        Top = 216
        Width = 58
        Height = 13
        Caption = 'Table libre 2'
        FocusControl = RAG_TABLELIBRE2
      end
      object TRAG_TABLELIBRE3: THLabel
        Left = 12
        Top = 244
        Width = 58
        Height = 13
        Caption = 'Table libre 3'
        FocusControl = RAG_TABLELIBRE3
      end
      object TRAG_ETATACTION: THLabel
        Left = 12
        Top = 102
        Width = 70
        Height = 13
        Caption = '&Etat de l'#39'action'
      end
      object RAG_TYPEACTION: THValComboBox
        Left = 120
        Top = 16
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 0
        OnExit = RAG_TYPEACTIONExit
        TagDispatch = 0
      end
      object RAG_LIBELLE: THCritMaskEdit
        Left = 120
        Top = 44
        Width = 150
        Height = 21
        MaxLength = 35
        TabOrder = 1
        TagDispatch = 0
      end
      object RAG_DATEACTION: THCritMaskEdit
        Left = 120
        Top = 72
        Width = 69
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 2
        Text = '  /  /    '
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = od2099
        ControlerDate = True
      end
      object RAG_DATEECHEANCE: THCritMaskEdit
        Left = 120
        Top = 124
        Width = 69
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 4
        Text = '  /  /    '
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = od2099
        ControlerDate = True
      end
      object RAG_COUTACTION: THCritMaskEdit
        Left = 120
        Top = 152
        Width = 69
        Height = 21
        TabOrder = 5
        TagDispatch = 0
        OpeType = otReel
        ControlerDate = True
      end
      object GAction: THGrid
        Left = 300
        Top = 0
        Width = 185
        Height = 261
        ColCount = 2
        DefaultRowHeight = 18
        FixedCols = 0
        RowCount = 100
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
        ScrollBars = ssVertical
        TabOrder = 9
        OnClick = GActionClick
        SortedCol = -1
        Titres.Strings = (
          'Type'
          'Libell'#233)
        Couleur = False
        MultiSelect = False
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = clSilver
        ColWidths = (
          58
          122)
      end
      object RAG_TABLELIBRE1: THValComboBox
        Left = 120
        Top = 184
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 6
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'RTRPRLIBACTION1'
      end
      object RAG_TABLELIBRE2: THValComboBox
        Left = 120
        Top = 212
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 7
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'RTRPRLIBACTION2'
      end
      object RAG_TABLELIBRE3: THValComboBox
        Left = 120
        Top = 240
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 8
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'RTRPRLIBACTION3'
      end
      object RAG_ETATACTION: THValComboBox
        Left = 120
        Top = 98
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 3
        TagDispatch = 0
        DrawingMode = cbDMBoth
      end
    end
    object Contacts: TTabSheet
      Caption = 'Contacts'
      ImageIndex = 4
      object TC_FONCTIONCODEE: THLabel
        Left = 12
        Top = 16
        Width = 41
        Height = 13
        Caption = 'Fonction'
        FocusControl = C_FONCTIONCODEE
      end
      object TC_LIBRECONTACT1: THLabel
        Left = 12
        Top = 136
        Width = 58
        Height = 13
        Caption = 'Table libre 1'
        FocusControl = C_LIBRECONTACT1
      end
      object TC_LIBRECONTACT2: THLabel
        Left = 12
        Top = 165
        Width = 58
        Height = 13
        Caption = 'Table libre 2'
        FocusControl = C_LIBRECONTACT2
      end
      object TC_LIBRECONTACT3: THLabel
        Left = 12
        Top = 195
        Width = 58
        Height = 13
        Caption = 'Table libre 3'
        FocusControl = C_LIBRECONTACT3
      end
      object TC_LIBRECONTACT4: THLabel
        Left = 12
        Top = 224
        Width = 58
        Height = 13
        Caption = 'Table libre 4'
        FocusControl = C_LIBRECONTACT4
      end
      object TC_LIBRECONTACT5: THLabel
        Left = 12
        Top = 252
        Width = 58
        Height = 13
        Caption = 'Table libre 5'
        FocusControl = C_LIBRECONTACT5
      end
      object TC_LIBRECONTACT6: THLabel
        Left = 256
        Top = 136
        Width = 58
        Height = 13
        Caption = 'Table libre 6'
        FocusControl = C_LIBRECONTACT6
      end
      object TC_LIBRECONTACT7: THLabel
        Left = 256
        Top = 165
        Width = 58
        Height = 13
        Caption = 'Table libre 7'
        FocusControl = C_LIBRECONTACT7
      end
      object TC_LIBRECONTACT8: THLabel
        Left = 256
        Top = 195
        Width = 58
        Height = 13
        Caption = 'Table libre 8'
        FocusControl = C_LIBRECONTACT8
      end
      object TC_LIBRECONTACT9: THLabel
        Left = 256
        Top = 224
        Width = 58
        Height = 13
        Caption = 'Table libre 9'
        FocusControl = C_LIBRECONTACT9
      end
      object TC_LIBRECONTACTA: THLabel
        Left = 256
        Top = 252
        Width = 64
        Height = 13
        Caption = 'Table libre 10'
        FocusControl = C_LIBRECONTACTA
      end
      object C_FONCTIONCODEE: THMultiValComboBox
        Left = 96
        Top = 12
        Width = 121
        Height = 21
        TabOrder = 0
        Abrege = False
        DataType = 'TTFONCTION'
        Complete = True
        OuInclusif = False
      end
      object C_PRINCIPAL: THCheckbox
        Left = 96
        Top = 44
        Width = 97
        Height = 17
        AllowGrayed = True
        Caption = 'Contact principal'
        State = cbGrayed
        TabOrder = 1
      end
      object C_PUBLIPOSTAGE: THCheckbox
        Left = 96
        Top = 70
        Width = 97
        Height = 17
        AllowGrayed = True
        Caption = 'Publipostage'
        State = cbGrayed
        TabOrder = 2
      end
      object C_FERME: THCheckbox
        Left = 96
        Top = 94
        Width = 97
        Height = 17
        AllowGrayed = True
        Caption = 'Ferm'#233
        State = cbGrayed
        TabOrder = 3
      end
      object C_LIBRECONTACT1: THMultiValComboBox
        Left = 96
        Top = 132
        Width = 121
        Height = 21
        TabOrder = 4
        Abrege = False
        DataType = 'YYLIBRECON1'
        Complete = True
        OuInclusif = False
      end
      object C_LIBRECONTACT2: THMultiValComboBox
        Left = 96
        Top = 161
        Width = 121
        Height = 21
        TabOrder = 5
        Abrege = False
        DataType = 'YYLIBRECON2'
        Complete = True
        OuInclusif = False
      end
      object C_LIBRECONTACT3: THMultiValComboBox
        Left = 96
        Top = 191
        Width = 121
        Height = 21
        TabOrder = 6
        Abrege = False
        DataType = 'YYLIBRECON3'
        Complete = True
        OuInclusif = False
      end
      object C_LIBRECONTACT4: THMultiValComboBox
        Left = 96
        Top = 220
        Width = 121
        Height = 21
        TabOrder = 7
        Abrege = False
        DataType = 'YYLIBRECON4'
        Complete = True
        OuInclusif = False
      end
      object C_LIBRECONTACT5: THMultiValComboBox
        Left = 96
        Top = 248
        Width = 121
        Height = 21
        TabOrder = 8
        Abrege = False
        DataType = 'YYLIBRECON5'
        Complete = True
        OuInclusif = False
      end
      object C_LIBRECONTACT6: THMultiValComboBox
        Left = 340
        Top = 132
        Width = 121
        Height = 21
        TabOrder = 9
        Abrege = False
        DataType = 'YYLIBRECON6'
        Complete = True
        OuInclusif = False
      end
      object C_LIBRECONTACT7: THMultiValComboBox
        Left = 340
        Top = 161
        Width = 121
        Height = 21
        TabOrder = 10
        Abrege = False
        DataType = 'YYLIBRECON7'
        Complete = True
        OuInclusif = False
      end
      object C_LIBRECONTACT8: THMultiValComboBox
        Left = 340
        Top = 191
        Width = 121
        Height = 21
        TabOrder = 11
        Abrege = False
        DataType = 'YYLIBRECON8'
        Complete = True
        OuInclusif = False
      end
      object C_LIBRECONTACT9: THMultiValComboBox
        Left = 340
        Top = 220
        Width = 121
        Height = 21
        TabOrder = 12
        Abrege = False
        DataType = 'YYLIBRECON9'
        Complete = True
        OuInclusif = False
      end
      object C_LIBRECONTACTA: THMultiValComboBox
        Left = 340
        Top = 248
        Width = 121
        Height = 21
        TabOrder = 13
        Abrege = False
        DataType = 'YYLIBRECONA'
        Complete = True
        OuInclusif = False
      end
      object CC_LIBRECONTACT1: THValComboBox
        Left = 96
        Top = 132
        Width = 121
        Height = 21
        ItemHeight = 0
        TabOrder = 14
        TagDispatch = 0
        Vide = True
        DataType = 'YYLIBRECON1'
      end
      object CC_LIBRECONTACT2: THValComboBox
        Left = 96
        Top = 161
        Width = 121
        Height = 21
        ItemHeight = 0
        TabOrder = 15
        TagDispatch = 0
        Vide = True
        DataType = 'YYLIBRECON2'
      end
      object CC_LIBRECONTACT3: THValComboBox
        Left = 96
        Top = 191
        Width = 121
        Height = 21
        ItemHeight = 0
        TabOrder = 16
        TagDispatch = 0
        Vide = True
        DataType = 'YYLIBRECON3'
      end
      object CC_LIBRECONTACT4: THValComboBox
        Left = 96
        Top = 220
        Width = 121
        Height = 21
        ItemHeight = 0
        TabOrder = 17
        TagDispatch = 0
        Vide = True
        DataType = 'YYLIBRECON4'
      end
      object CC_LIBRECONTACT5: THValComboBox
        Left = 96
        Top = 248
        Width = 121
        Height = 21
        ItemHeight = 0
        TabOrder = 18
        TagDispatch = 0
        Vide = True
        DataType = 'YYLIBRECON5'
      end
      object CC_LIBRECONTACT6: THValComboBox
        Left = 340
        Top = 132
        Width = 121
        Height = 21
        ItemHeight = 0
        TabOrder = 19
        TagDispatch = 0
        Vide = True
        DataType = 'YYLIBRECON6'
      end
      object CC_LIBRECONTACT7: THValComboBox
        Left = 340
        Top = 161
        Width = 121
        Height = 21
        ItemHeight = 0
        TabOrder = 20
        TagDispatch = 0
        Vide = True
        DataType = 'YYLIBRECON7'
      end
      object CC_LIBRECONTACT8: THValComboBox
        Left = 340
        Top = 191
        Width = 121
        Height = 21
        ItemHeight = 0
        TabOrder = 21
        TagDispatch = 0
        Vide = True
        DataType = 'YYLIBRECON8'
      end
      object CC_LIBRECONTACT9: THValComboBox
        Left = 340
        Top = 220
        Width = 121
        Height = 21
        ItemHeight = 0
        TabOrder = 22
        TagDispatch = 0
        Vide = True
        DataType = 'YYLIBRECON9'
      end
      object CC_LIBRECONTACTA: THValComboBox
        Left = 340
        Top = 248
        Width = 121
        Height = 21
        ItemHeight = 0
        TabOrder = 23
        TagDispatch = 0
        Vide = True
        DataType = 'YYLIBRECONA'
      end
      object CC_FONCTIONCODEE: THValComboBox
        Left = 96
        Top = 12
        Width = 121
        Height = 21
        ItemHeight = 0
        TabOrder = 24
        TagDispatch = 0
        Vide = True
        DataType = 'TTFONCTION'
      end
    end
    object Recap: TTabSheet
      Caption = 'R'#233'capitulatif'
      ImageIndex = 2
      object HRECAP: THListBox
        Left = 8
        Top = 8
        Width = 477
        Height = 269
        Color = clBtnFace
        ItemHeight = 13
        TabOrder = 0
      end
    end
  end
  inherited cControls: THListBox
    Left = 36
    Top = 228
  end
  object TCodeCiblage: THCritMaskEdit [14]
    Left = 112
    Top = 232
    Width = 65
    Height = 21
    Color = clYellow
    TabOrder = 5
    Visible = False
    TagDispatch = 0
  end
  object MODELE_OPERATION: THCritMaskEdit [15]
    Left = 8
    Top = 316
    Width = 161
    Height = 21
    TabOrder = 6
    TagDispatch = 0
    ElipsisButton = True
    OnElipsisClick = MODELE_OPERATIONElipsisClick
  end
end
