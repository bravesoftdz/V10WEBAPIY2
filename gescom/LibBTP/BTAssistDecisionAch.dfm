inherited FassistDecisionAch: TFassistDecisionAch
  Left = 325
  Top = 200
  Caption = 'Assistant de g'#233'n'#233'ration de d'#233'cisionnel d'#39'achat'
  ClientWidth = 541
  PixelsPerInch = 96
  TextHeight = 13
  inherited lAide: THLabel
    Width = 353
  end
  inherited bPrecedent: TToolbarButton97
    Left = 274
  end
  inherited bSuivant: TToolbarButton97
    Left = 357
  end
  inherited bFin: TToolbarButton97
    Left = 441
  end
  inherited bAnnuler: TToolbarButton97
    Left = 191
  end
  inherited bAide: TToolbarButton97
    Left = 108
  end
  inherited GroupBox1: THGroupBox
    Width = 544
  end
  inherited P: THPageControl2
    Left = 185
    ActivePage = TassistDec1
    object TassistDec1: TTabSheet
      Caption = 'TassistDec1'
      object TText4: THLabel
        Left = 36
        Top = 19
        Width = 280
        Height = 26
        Caption = 
          'Veuillez indiquer la p'#233'riode de prise en compte des besoins de c' +
          'hantier'
        WordWrap = True
      end
      object TDATEDEB: THLabel
        Left = 36
        Top = 86
        Width = 82
        Height = 13
        Caption = 'D'#233'but de p'#233'riode'
      end
      object TDATEFIN: THLabel
        Left = 36
        Top = 121
        Width = 67
        Height = 13
        Caption = 'Fin de p'#233'riode'
      end
      object DATEDEB: THCritMaskEdit
        Left = 180
        Top = 84
        Width = 69
        Height = 21
        EditMask = '!99 >L<LL 0000;1;_'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 11
        ParentFont = False
        TabOrder = 0
        Text = '           '
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = odDate
        ControlerDate = True
      end
      object DATEDEB_: THCritMaskEdit
        Left = 180
        Top = 119
        Width = 69
        Height = 21
        EditMask = '!99 >L<LL 0000;1;_'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 11
        ParentFont = False
        TabOrder = 1
        Text = '           '
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = odDate
        ControlerDate = True
      end
      object CBSELDOC: TCheckBox
        Left = 179
        Top = 217
        Width = 153
        Height = 17
        Alignment = taLeftJustify
        Caption = 'S'#233'lection des documents'
        TabOrder = 2
      end
    end
    object Tassitdec2: TTabSheet
      Caption = 'Tassitdec2'
      ImageIndex = 3
      object TText2: THLabel
        Left = 35
        Top = 23
        Width = 271
        Height = 52
        Caption = 
          'Vous avez la possibilit'#233' de limiter le nombre de document en fon' +
          'ction des informations ci dessous. Une zone laiss'#233'e vide signifi' +
          'e que tous les '#233'l'#233'ments correspondants seront trait'#233's.'
        WordWrap = True
      end
      object TSR_ARTICLE: THLabel
        Left = 35
        Top = 95
        Width = 72
        Height = 13
        Caption = 'Du code article'
        FocusControl = GZZ_ARTICLE
      end
      object TSR_ARTICLE_: THLabel
        Left = 35
        Top = 119
        Width = 71
        Height = 13
        Caption = 'Au code article'
        FocusControl = GZZ_ARTICLE_
      end
      object TSR_DEPOT: THLabel
        Left = 35
        Top = 172
        Width = 44
        Height = 13
        Caption = 'Du d'#233'p'#244't'
      end
      object TSR_DEPOT_: THLabel
        Left = 35
        Top = 196
        Width = 43
        Height = 13
        Caption = 'Au d'#233'p'#244't'
      end
      object GZZ_ARTICLE: THCritMaskEdit
        Left = 156
        Top = 91
        Width = 145
        Height = 21
        TabOrder = 0
        TagDispatch = 0
        ElipsisButton = True
      end
      object GZZ_ARTICLE_: THCritMaskEdit
        Left = 156
        Top = 115
        Width = 145
        Height = 21
        TabOrder = 1
        TagDispatch = 0
        ElipsisButton = True
      end
      object GZZ_DEPOT: THValComboBox
        Left = 156
        Top = 168
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        VideString = '<<Tous>>'
        DataType = 'GCDEPOT'
      end
      object GZZ_DEPOT_: THValComboBox
        Left = 156
        Top = 192
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        VideString = '<<Tous>>'
        DataType = 'GCDEPOT'
      end
    end
    object TassistDec3: TTabSheet
      Caption = 'TassistDec3'
      ImageIndex = 2
      object HLabel1: THLabel
        Left = 36
        Top = 19
        Width = 120
        Height = 13
        Caption = 'Veuillez d'#233'finir les options'
        WordWrap = True
      end
      object HLabel2: THLabel
        Left = 29
        Top = 86
        Width = 65
        Height = 13
        Caption = 'Etablissement'
      end
      object CBETABLISSEMENT: THValComboBox
        Left = 168
        Top = 86
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 0
        Text = 'CBETABLISSEMENT'
        TagDispatch = 0
        Vide = True
        VideString = '<<Tous>>'
        DataType = 'TTETABLISSEMENT'
      end
    end
    object TabSheetFinal: TTabSheet
      Caption = 'R'#233'capitulatif'
      ImageIndex = 1
      object HLabel7: THLabel
        Left = 14
        Top = 89
        Width = 59
        Height = 13
        Caption = 'R'#233'capitulatif'
      end
      object Recap: THRichEditOLE
        Left = 0
        Top = 111
        Width = 338
        Height = 133
        Align = alBottom
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 0
        Margins.Top = 0
        Margins.Bottom = 0
        Margins.Left = 0
        Margins.Right = 0
        ContainerName = 'Document'
        ObjectMenuPrefix = '&Object'
        LinesRTF.Strings = (
          '{\rtf1\ansi\deff0\nouicompat{\fonttbl{\f0\fnil Arial;}}'
          '{\*\generator Riched20 10.0.15063}\viewkind4\uc1 '
          '\pard\f0\fs16\lang1036 '
          '\par '
          '\par '
          '\par '
          '\par '
          '\par '
          '\par '
          '\par '
          '\par '
          '\par }')
      end
      object PanelFin: TPanel
        Left = 0
        Top = 0
        Width = 338
        Height = 84
        Align = alTop
        TabOrder = 1
        object TTextFin1: THLabel
          Left = 17
          Top = 10
          Width = 342
          Height = 26
          Caption = 
            'Le param'#232'trage est maintenant correctement renseign'#233' pour permet' +
            'tre la g'#233'n'#233'ration du d'#233'cisionnel d'#39'achat.'
          WordWrap = True
        end
        object TTextFin2: THLabel
          Left = 17
          Top = 44
          Width = 324
          Height = 26
          Caption = 
            'Si vous d'#233'sirez revoir le param'#233'trage, il suffit de cliquer sur ' +
            'le bouton Pr'#233'c'#233'dent sinon, le bouton Fin, permet de d'#233'buter le t' +
            'raitement.'
          WordWrap = True
        end
      end
    end
  end
end
