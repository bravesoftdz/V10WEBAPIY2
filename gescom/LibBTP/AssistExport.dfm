inherited FAssistExport: TFAssistExport
  Left = 440
  Top = 217
  BorderIcons = []
  BorderStyle = bsNone
  Caption = 'Exportation Fichier Comptable'
  ClientHeight = 431
  ClientWidth = 767
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Left = 10
    Top = 407
    Caption = 'Etape 1/3'
  end
  inherited lAide: THLabel
    Left = 240
    Top = 373
    Width = 523
  end
  inherited bPrecedent: TToolbarButton97
    Left = 511
    Top = 403
  end
  inherited bSuivant: TToolbarButton97
    Left = 594
    Top = 403
  end
  inherited bFin: TToolbarButton97
    Left = 678
    Top = 403
    Caption = '&Exporter'
  end
  inherited bAnnuler: TToolbarButton97
    Left = 428
    Top = 403
    Caption = '&Annuler'
  end
  inherited bAide: TToolbarButton97
    Left = 345
    Top = 403
    Caption = 'Ai&de'
  end
  inherited Plan: THPanel
    Left = 240
    Top = 53
    Width = 523
    Height = 316
  end
  inherited GroupBox1: THGroupBox
    Top = 390
    Width = 770
  end
  inherited P: THPageControl2
    Left = 248
    Top = 64
    Width = 505
    Height = 297
    ActivePage = TRESUME
    object TGENERALITE: TTabSheet
      Caption = 'G'#233'n'#233'ralit'#233's'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      object LblDu: TLabel
        Left = 11
        Top = 81
        Width = 17
        Height = 13
        Caption = 'Du '
      end
      object LblAu: TLabel
        Left = 126
        Top = 82
        Width = 13
        Height = 13
        Caption = 'Au'
      end
      object HLabel2: THLabel
        Left = 0
        Top = 25
        Width = 497
        Height = 13
        Alignment = taCenter
        AutoSize = False
        Caption = 'Choix de la P'#233'riode'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clHighlight
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object HLabel1: THLabel
        Left = 0
        Top = 0
        Width = 497
        Height = 13
        Alignment = taCenter
        AutoSize = False
        Caption = 'Assistant d'#39'exportation des Fichiers TRA'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clHighlight
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object ChkExport: TCheckBox
        Left = 11
        Top = 127
        Width = 225
        Height = 25
        BiDiMode = bdLeftToRight
        Caption = 'Ecritures d'#233'j'#224' export'#233'es pour cette p'#233'riode'
        ParentBiDiMode = False
        TabOrder = 0
      end
      object DateDeb: TMaskEdit
        Left = 35
        Top = 78
        Width = 73
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 1
        Text = '  /  /    '
      end
      object DateFin: TMaskEdit
        Left = 154
        Top = 78
        Width = 73
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 2
        Text = '  /  /    '
      end
      object ChkEnvoiEcriture: TRadioButton
        Left = 11
        Top = 48
        Width = 120
        Height = 17
        Caption = 'Envoi des Ecritures'
        Checked = True
        TabOrder = 3
        TabStop = True
      end
      object GroupBox2: TGroupBox
        Left = 0
        Top = 15
        Width = 497
        Height = 8
        TabOrder = 4
      end
    end
    object TENVOIFIL: TTabSheet
      Caption = 'Envoi par Fichier'
      ImageIndex = 1
      object PENVOIMAIL: TPanel
        Left = 0
        Top = 73
        Width = 497
        Height = 196
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 2
        object LblAdrMail: TLabel
          Left = 11
          Top = 11
          Width = 110
          Height = 13
          Caption = 'Adresse de Messagerie'
        end
        object LblCorpMail: TLabel
          Left = 11
          Top = 87
          Width = 88
          Height = 13
          Caption = 'Corps du Message'
        end
        object LblAttention: TLabel
          Left = 11
          Top = 183
          Width = 414
          Height = 13
          Caption = 
            'ATTENTION : Le fichier d'#39'echange sera envoy'#233' sous forme de pi'#232'ce' +
            ' jointe au message'
        end
        object AdrMail: TEdit
          Left = 11
          Top = 32
          Width = 374
          Height = 21
          TabOrder = 0
        end
        object ChkMailImportance: TCheckBox
          Left = 11
          Top = 58
          Width = 169
          Height = 17
          Caption = 'Mail avec importance Haute'
          TabOrder = 1
        end
        object CorpMessage: TMemo
          Left = 10
          Top = 102
          Width = 361
          Height = 73
          ScrollBars = ssVertical
          TabOrder = 2
          WantTabs = True
        end
      end
      object PENVOIFIL: TPanel
        Left = 0
        Top = 73
        Width = 497
        Height = 196
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 1
        object LblExchangeFile: TLabel
          Left = 11
          Top = 11
          Width = 180
          Height = 13
          Caption = 'Indiquer le nom du Fichier d'#39'echange :'
        end
        object FileTrF: THCritMaskEdit
          Left = 11
          Top = 32
          Width = 374
          Height = 21
          TabOrder = 0
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = FileTrFElipsisClick
        end
      end
      object Panel1: TPanel
        Left = 0
        Top = 0
        Width = 497
        Height = 73
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 0
        object HLabel4: THLabel
          Left = 0
          Top = 25
          Width = 497
          Height = 13
          Alignment = taCenter
          AutoSize = False
          Caption = 'Envoi par fichier'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clHighlight
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object HLabel3: THLabel
          Left = 0
          Top = 0
          Width = 497
          Height = 13
          Alignment = taCenter
          AutoSize = False
          Caption = 'Assistant d'#39'exportation des Fichiers TRA'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clHighlight
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object HLabel6: THLabel
          Left = 0
          Top = 25
          Width = 497
          Height = 13
          Alignment = taCenter
          AutoSize = False
          Caption = 'Envoi par Messagerie'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clHighlight
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object ChkMail: TRadioButton
          Left = 304
          Top = 48
          Width = 81
          Height = 17
          Caption = 'Messagerie'
          TabOrder = 0
          OnClick = ChkMailClick
        end
        object GroupBox3: TGroupBox
          Left = 0
          Top = 15
          Width = 497
          Height = 8
          TabOrder = 1
        end
        object ChkFile: TRadioButton
          Left = 11
          Top = 48
          Width = 129
          Height = 17
          Caption = 'Fichier'
          Checked = True
          TabOrder = 2
          TabStop = True
          OnClick = ChkFileClick
        end
      end
    end
    object TRESUME: TTabSheet
      Caption = 'R'#233'sum'#233
      ImageIndex = 3
      object HLabel7: THLabel
        Left = 0
        Top = 0
        Width = 497
        Height = 13
        Alignment = taCenter
        AutoSize = False
        Caption = 'Assistant d'#39'exportation des Fichiers TRA'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clHighlight
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object HLabel8: THLabel
        Left = 0
        Top = 25
        Width = 497
        Height = 13
        Alignment = taCenter
        AutoSize = False
        Caption = 'R'#233'sum'#233
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clHighlight
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object GroupBox5: TGroupBox
        Left = 0
        Top = 15
        Width = 497
        Height = 8
        TabOrder = 0
      end
      object PRESUME: TPanel
        Left = 8
        Top = 48
        Width = 481
        Height = 121
        BevelInner = bvLowered
        BevelOuter = bvLowered
        Color = clInactiveCaptionText
        TabOrder = 1
        object LblTraitement: TLabel
          Left = 16
          Top = 8
          Width = 59
          Height = 13
          Caption = 'Traitement : '
        end
        object LblPeriode: TLabel
          Left = 16
          Top = 32
          Width = 42
          Height = 13
          Caption = 'P'#233'riode :'
        end
        object LblFile: TLabel
          Left = 16
          Top = 64
          Width = 40
          Height = 13
          Caption = 'Fichier : '
        end
        object LblMessagerie: TLabel
          Left = 16
          Top = 96
          Width = 63
          Height = 13
          Caption = 'Messagerie : '
        end
      end
      object BlocNote: TMemo
        Left = 8
        Top = 48
        Width = 481
        Height = 209
        Color = clHighlightText
        Lines.Strings = (
          'BlocNote')
        ReadOnly = True
        TabOrder = 2
        Visible = False
      end
    end
  end
  inherited PanelImage: THPanel
    Left = 8
    Top = 81
    Width = 225
    Height = 265
    inherited Image: TToolbarButton97
      Left = 8
      Top = 10
      Width = 209
      Height = 251
      NoBorder = True
    end
  end
  inherited cControls: THListBox
    Left = 112
    Top = 41
  end
  object PTitre: TPanel [12]
    Left = 0
    Top = 0
    Width = 767
    Height = 33
    Align = alTop
    BevelInner = bvLowered
    BevelOuter = bvLowered
    Caption = 'Exportation Fichier Comptable (Format TRA)'
    Color = clHighlight
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clHighlightText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 5
  end
  inherited Msg: THMsgBox
    Left = 71
    Top = 41
  end
  inherited HMTrad: THSystemMenu
    Left = 36
    Top = 41
  end
  object ODGetInfosTRA: TOpenDialog
    Filter = 
      'Fichier R'#233'vision Expert (*.TRA)|*.TRA|Fichier Compress'#233' (*.ZIP)|' +
      '*.ZIP|Tous les fichiers (*.*)|*.*'
    Options = [ofHideReadOnly, ofPathMustExist]
    Left = 187
    Top = 41
  end
end
