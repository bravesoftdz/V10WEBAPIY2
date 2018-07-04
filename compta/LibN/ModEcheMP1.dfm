inherited FMODIFECHEMP1: TFMODIFECHEMP1
  Height = 390
  Caption = 'Modification des échéances'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: TPageControl
    ActivePage = PModifS
    inherited PCritere: TTabSheet
      object HLabel4: THLabel
        Left = 8
        Top = 8
        Width = 74
        Height = 13
        Caption = 'Compte &général'
        FocusControl = E_GENERAL
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel1: THLabel
        Left = 8
        Top = 32
        Width = 79
        Height = 13
        Caption = 'Compte &auxiliaire'
        FocusControl = E_AUXILIAIRE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TE_JOURNAL: THLabel
        Left = 8
        Top = 56
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = E_JOURNAL
      end
      object TE_EXERCICE: THLabel
        Left = 282
        Top = 8
        Width = 41
        Height = 13
        Caption = '&Exercice'
        FocusControl = E_EXERCICE
      end
      object HLabel10: THLabel
        Left = 281
        Top = 32
        Width = 69
        Height = 13
        Caption = '&Echéances du'
        FocusControl = E_DATEECHEANCE
      end
      object HLabel3: THLabel
        Left = 281
        Top = 56
        Width = 100
        Height = 13
        Caption = '&Dates comptables du'
        FocusControl = E_DATECOMPTABLE
      end
      object HLabel6: THLabel
        Left = 474
        Top = 56
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object HLabel7: THLabel
        Left = 474
        Top = 32
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATEECHEANCE_
      end
      object E_JOURNAL: THValComboBox
        Left = 104
        Top = 52
        Width = 145
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTJOURNAUX'
      end
      object E_AUXILIAIRE: THCpteEdit
        Tag = 1
        Left = 104
        Top = 28
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 1
        OnChange = E_AUXILIAIREChange
        ZoomTable = tzTlettrable
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_NUMECHE: THCritMaskEdit
        Left = 116
        Top = 29
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        Text = '1'
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object XX_WHERE: TEdit
        Left = 140
        Top = 29
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        Text = 
          '(E_ECRANOUVEAU="N" or E_ECRANOUVEAU="H") AND (E_ETATLETTRAGE<>"R' +
          'I")'
        Visible = False
      end
      object E_ECHE: TEdit
        Left = 190
        Top = 30
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        Text = 'X'
        Visible = False
      end
      object E_TRESOLETTRE: THCritMaskEdit
        Left = 212
        Top = 29
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object XX_WHEREPOP: TEdit
        Left = 244
        Top = 29
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 6
        Visible = False
      end
      object E_GENERAL: THCpteEdit
        Tag = 1
        Left = 104
        Top = 4
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 7
        ZoomTable = tzGLettColl
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 388
        Top = 52
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 8
        Text = '01/01/1990'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 496
        Top = 52
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 9
        Text = '31/12/2099'
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATEECHEANCE_: THCritMaskEdit
        Left = 496
        Top = 28
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 10
        Text = '31/12/1999'
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATEECHEANCE: THCritMaskEdit
        Left = 388
        Top = 28
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 11
        Text = '01/01/1990'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_EXERCICE: THValComboBox
        Left = 388
        Top = 4
        Width = 185
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 12
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
    end
    inherited PComplement: TTabSheet
      object Label1: THLabel
        Left = 8
        Top = 8
        Width = 57
        Height = 13
        Caption = '&Numéros de'
        FocusControl = E_NUMEROPIECE
      end
      object Label2: THLabel
        Left = 154
        Top = 8
        Width = 6
        Height = 13
        Caption = 'à'
        FocusControl = E_NUMEROPIECE_
      end
      object TE_NATUREPIECE: THLabel
        Left = 8
        Top = 32
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = E_NATUREPIECE
      end
      object TE_QUALIFPIECE: THLabel
        Left = 8
        Top = 57
        Width = 24
        Height = 13
        Caption = '&Type'
        FocusControl = E_QUALIFPIECE
      end
      object TMP_CODEACCEPT: THLabel
        Left = 290
        Top = 57
        Width = 84
        Height = 13
        Caption = 'Code &acceptation'
        FocusControl = MP_CODEACCEPT
      end
      object Label12: THLabel
        Left = 290
        Top = 32
        Width = 55
        Height = 13
        Caption = 'Ref. &interne'
        FocusControl = E_REFINTERNE
      end
      object HLabel8: THLabel
        Left = 290
        Top = 8
        Width = 33
        Height = 13
        Caption = '&Devise'
        FocusControl = E_DEVISE
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Tag = 2
        Left = 76
        Top = 4
        Width = 65
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 0
        Text = '0'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_NUMEROPIECE_: THCritMaskEdit
        Tag = 2
        Left = 172
        Top = 4
        Width = 65
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 1
        Text = '99999999'
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_NATUREPIECE: THValComboBox
        Left = 76
        Top = 28
        Width = 163
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATUREPIECE'
      end
      object E_QUALIFPIECE: THValComboBox
        Left = 76
        Top = 53
        Width = 163
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        DataType = 'TTQUALPIECE'
      end
      object MP_CODEACCEPT: THMultiValComboBox
        Left = 388
        Top = 53
        Width = 149
        Height = 21
        TabOrder = 4
        Abrege = False
        DataType = 'TTLETTRECHANGE'
        Complete = False
        OuInclusif = False
      end
      object E_REFINTERNE: THCritMaskEdit
        Left = 388
        Top = 28
        Width = 149
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
      end
      object E_DEVISE: THValComboBox
        Tag = 1
        Left = 388
        Top = 4
        Width = 149
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 6
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
    end
    object PModifS: TTabSheet [3]
      Caption = 'Modification en série'
      ImageIndex = 4
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 592
        Height = 83
        Align = alClient
      end
      object Label3: TLabel
        Left = 7
        Top = 12
        Width = 108
        Height = 16
        Caption = 'Zone(s) à modifier'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object ZoomRib: TToolbarButton97
        Left = 566
        Top = 43
        Width = 16
        Height = 21
        Hint = 'Choisir le RIB'
        Caption = '...'
        ParentShowHint = False
        ShowHint = True
        OnClick = ZoomRibClick
      end
      object Label4: TLabel
        Left = 7
        Top = 46
        Width = 108
        Height = 16
        Caption = 'Nouvelles valeurs'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object CMSMP: TCheckBox
        Left = 135
        Top = 13
        Width = 118
        Height = 17
        Caption = '&Mode de paiement'
        TabOrder = 0
        OnClick = CMSMPClick
      end
      object CMSDateEche: TCheckBox
        Left = 290
        Top = 12
        Width = 118
        Height = 17
        Caption = '&Date d'#39'échéance'
        TabOrder = 1
        OnClick = CMSDateEcheClick
      end
      object CMSRIB: TCheckBox
        Left = 466
        Top = 12
        Width = 50
        Height = 17
        Caption = '&RIB'
        TabOrder = 2
        OnClick = CMSRIBClick
      end
      object RIBNEW: TEdit
        Left = 410
        Top = 43
        Width = 155
        Height = 21
        TabOrder = 3
        OnKeyDown = RIBNEWKeyDown
      end
      object DateEcheNew: THCritMaskEdit
        Left = 289
        Top = 43
        Width = 100
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '01/01/1990'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object MPNew: THValComboBox
        Left = 134
        Top = 43
        Width = 118
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        DataType = 'TTMODEPAIE'
      end
    end
  end
  inherited Dock971: TDock97
    inherited PFiltres: TToolWindow97
      inherited BCherche: TToolbarButton97
        Left = 565
      end
      inherited lpresentation: TLabel
        Left = 425
      end
      inherited FFiltres: THValComboBox
        Width = 489
      end
      inherited cbPresentations: THValComboBox
        Left = 496
      end
    end
  end
  inherited FListe: THDBGrid
    Height = 162
  end
  inherited Dock: TDock97
    Top = 327
    inherited PanelBouton: TToolWindow97
      ClientHeight = 36
      ClientAreaHeight = 36
      object BModifSerie: TToolbarButton97 [0]
        Left = 500
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Modification en série'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Modifier en série'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Glyph.Data = {
          E2060000424DE206000000000000360400002800000024000000130000000100
          080000000000AC02000000000000000000000001000000010000000000000000
          80000080000000808000800000008000800080800000C0C0C000C0DCC000F0CA
          A600000000000000000000000000000000000000000000000000000000000000
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
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00030303030304
          0403030303030303030303030303030303F8F8FF030303030303030303030303
          0303040202040303030303030303030303030303F80303F8FF03030303030303
          0303030303040202020204030303030303030303030303F803030303F8FF0303
          0303030303030303040202020202020403030303030303030303F80303030303
          03F8FF030303030303030304020202FA02020202040303030303030303F8FF03
          03FF03030303F8FF03030303030303020202FA02FA0202020403030303030303
          03F8FF03F803FF030303F8FF03030303030303FA02FA020202FA020202040303
          0303030303F8FFF8030303FF030303F8FF03030303030304FA0202020202FA02
          020204030303030303F8F80303030303FF030303F8FF0303030304020202FA02
          020202FA0202020403030303F8FF0303F8FF030303FF030303F8FF0303030202
          02FA03FA02020204FA02020204030303F8FF03F803F8FF0303F8FF030303F8FF
          0303FA02FA030303FA02020204FA020202040303F8FFF8030303F8FF0303F8FF
          030303F8FF0303FA0303030303FA02020204FA020202040303F80303030303F8
          FF0303F8FF030303F8FF0303030303030303FA02020204FA0202040303030303
          03030303F8FF0303F8FF0303F8FF030303030303030303FA02020204FA020203
          030303030303030303F8FF0303F8FFF8030303030303030303030303FA020202
          04FA030303030303030303030303F8FF0303F8FF030303030303030303030303
          03FA0202020403030303030303030303030303F8FF0303F8FF03030303030303
          030303030303FA0202040303030303030303030303030303F8FF03F8FF030303
          0303030303030303030303FA0202030303030303030303030303030303F8FFF8
          FF030303030303030303030303030303FA030303030303030303030303030303
          0303F8030303}
        Margin = 3
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BModifSerieClick
        IsControl = True
      end
      object BZoomPiece: TToolbarButton97
        Tag = 1
        Left = 164
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Voir écriture'
        DisplayMode = dmGlyphOnly
        Caption = 'Zoom'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Glyph.Data = {
          AA040000424DAA04000000000000360000002800000013000000130000000100
          1800000000007404000000000000000000000000000000000000FF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
          FFFFFFFFFFFFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFFFFFFBF0000BF
          0000FFFFFFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFFFFFFBF000000FFFFFFFF00BF00
          00FFFFFFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FFFF00FFFFFFFFBF000000FFFFFFFF80FF8000BF0000FF00FF
          FF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFBF000000FFFFFFFF00BF0000BF0000FF00FFFF00FFFF00FF00
          0000FF00FFFF00FFFF00FFFF00FFFFFFFFFFFFFF000000000000000000000000
          BF000000FFFFFFFF00FF8000BF0000FF00FFFF00FFFF00FFFF00FF000000FF00
          FFFF00FFFFFFFFFFFFFF000000000000FF00FFFF00FFFF00FFFF00FFFF00FFBF
          0000BF0000FF0040FF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FF
          FFFFFF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF0000
          00FF0040FF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFFFFFFFFFFFF00
          0000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFFFFFF000000FF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF
          00FFFF00FFFF00FFFF00FF000000FF00FFFFFFFF000000FF00FFFF00FF00FFFF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00
          FFFF00FFFF00FF000000FF00FFFFFFFF000000FF00FFFF00FF00FFFFFF00FFFF
          00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FF
          FF00FF000000FF00FFFF00FF000000FF00FFFF00FF00FFFF00FFFFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FF00
          0000FF00FFFF00FFFF00FF000000FF00FFFF00FF00FFFF00FFFF00FFFFFF00FF
          FF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00
          FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
          00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FF
          FF00FFFF00FF000000000000FF00FFFF00FFFF00FFFF00FF000000000000FF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FF000000000000000000000000FF00FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FFFF00FF000000}
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BZoomPieceClick
        IsControl = True
      end
    end
  end
  inherited PCumul: TPanel
    Top = 310
  end
  inherited PanVBar: TPanel
    Height = 162
  end
  object HME: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Modification des échéances;Cette échéance d'#39'A-Nouveau n'#39'est pa' +
        's modifiable;W;O;O;O;'
      
        '1;Modification des échéances;Les critères de compte ne sont pas ' +
        'remplis. Confirmez-vous la recherche sur tous les comptes ?;Q;YN' +
        'C;N;N;'
      
        '2;Modification en série sur mode de paiement;La nouvelle valeur ' +
        'du mode de paiement n'#39'est pas renseignée;W;O;O;O;'
      
        '3;Modification en série sur RIB;La nouvelle valeur de RIB n'#39'est ' +
        'pas renseignée;W;O;O;O;'
      
        '4;Modification en série sur date d'#39'échéance;La nouvelle valeur d' +
        'e la date d'#39'échéance n'#39'est pas renseignée;W;O;O;O;'
      '')
    Left = 271
    Top = 255
  end
end
