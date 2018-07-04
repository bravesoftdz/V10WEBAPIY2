inherited FMessage: TFMessage
  Left = 226
  Top = 169
  Width = 757
  Height = 509
  BorderIcons = [biSystemMenu, biMinimize, biMaximize]
  Caption = 'Message'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  object VSplitter: TSplitter [0]
    Left = 745
    Top = 197
    Width = 4
    Height = 233
    Align = alRight
    Visible = False
  end
  object Splitter1: TSplitter [1]
    Left = 539
    Top = 197
    Width = 4
    Height = 233
    Align = alRight
    Visible = False
  end
  inherited Dock971: TDock97
    Top = 430
    Width = 749
    Height = 45
    inherited PBouton: TToolWindow97
      ClientHeight = 41
      ClientWidth = 749
      ClientAreaHeight = 41
      ClientAreaWidth = 749
      object BClasser: TToolbarButton97 [0]
        Left = 106
        Top = 6
        Width = 78
        Height = 27
        Hint = 'Cl'#244'turer le message'
        Caption = 'Classer'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BClasserClick
        GlobalIndexImage = 'Z0956_S16G1'
      end
      inherited BValider: TToolbarButton97
        Left = 653
        Top = 6
        ParentShowHint = False
        ShowHint = True
      end
      inherited BFerme: TToolbarButton97
        Left = 685
        Top = 6
        ParentShowHint = False
        ShowHint = True
      end
      inherited HelpBtn: TToolbarButton97
        Left = 717
        Top = 6
        ParentShowHint = False
        ShowHint = True
      end
      inherited bDefaire: TToolbarButton97
        Left = 196
        Top = 6
        ParentShowHint = False
        ShowHint = True
      end
      inherited Binsert: TToolbarButton97
        Left = 208
        Top = 6
        ParentShowHint = False
        ShowHint = True
      end
      inherited BDelete: TToolbarButton97
        Left = 67
        Top = 6
        ParentShowHint = False
        ShowHint = True
        Visible = True
      end
      object BWebPrecedent: TToolbarButton97 [7]
        Left = 6
        Top = 6
        Width = 28
        Height = 27
        Hint = 'Page pr'#233'c'#233'dente'
        Enabled = False
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BWebPrecedentClick
        GlobalIndexImage = 'Z0719_S16G1'
      end
      object BExtraireEml: TToolbarButton97 [8]
        Left = 577
        Top = 6
        Width = 82
        Height = 27
        Hint = 'Extraire le mail initial'
        Caption = 'Extraire'
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BExtraireEmlClick
        GlobalIndexImage = 'Z2023_S16G1'
      end
      object BTRANSFERER: TToolbarButton97 [9]
        Left = 480
        Top = 6
        Width = 92
        Height = 27
        Hint = 'Transf'#233'rer'
        Caption = 'Transf'#233'rer '
        ParentShowHint = False
        ShowHint = True
        OnClick = BTRANSFERERClick
        GlobalIndexImage = 'Z1525_S24G1'
      end
      object BREPONDRE: TToolbarButton97 [10]
        Left = 264
        Top = 6
        Width = 89
        Height = 27
        Hint = 'R'#233'pondre'
        Caption = 'R'#233'pondre  '
        ParentShowHint = False
        ShowHint = True
        OnClick = BREPONDREClick
        GlobalIndexImage = 'Z1678_S24G1'
      end
      object BDeclasser: TToolbarButton97 [11]
        Left = 106
        Top = 6
        Width = 78
        Height = 27
        Hint = 'D'#233'cl'#244'turer le message'
        Caption = 'D'#233'classer'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BDeclasserClick
        GlobalIndexImage = 'Z0956_S16G1'
      end
      object BREPONDRETOUS: TToolbarButton97 [12]
        Left = 360
        Top = 6
        Width = 113
        Height = 27
        Hint = 'R'#233'pondre '#224' tous'
        Caption = 'R'#233'pondre '#224' tous'
        ParentShowHint = False
        ShowHint = True
        OnClick = BREPONDRETOUSClick
        GlobalIndexImage = 'Z1690_S16G1'
      end
      inherited BImprimer: TToolbarButton97
        Left = 12
        Top = 6
      end
      object BPrint: TToolbarButton97
        Left = 36
        Top = 6
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BPrintClick
        GlobalIndexImage = 'M0072_S16G1'
      end
    end
  end
  object PnlHaut: TPanel [3]
    Left = 0
    Top = 0
    Width = 749
    Height = 197
    Align = alTop
    TabOrder = 1
    object Image1: TImage
      Left = 12
      Top = 8
      Width = 37
      Height = 37
      Picture.Data = {
        07544269746D6170220E0000424D220E00000000000036000000280000002400
        0000210000000100180000000000EC0D0000C40E0000C40E0000000000000000
        0000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FF0000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FF737573639A9C639A9C639A9C639A9C639A9C639A
        9C639A9C639A9C639A9C639A9C639A9C639A9C639A9C639A9C639A9C639A9C63
        9A9C639A9C639A9C639A9C639A9C639A9C639A9C639A9C639A9C000000FF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF7375739CCFCE9CCFCE9CCF
        CE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9C
        CFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE
        9CCFCE000000639A9C000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FF737573F7F3F79CFFFFF7F3F79CFFFFF7F3F79CFFFFF7F3F79CFFFFF7F3F79C
        FFFFF7F3F79CFFFFF7F3F79CFFFFF7F3F79CFFFFF7F3F79CFFFFF7F3F79CFFFF
        F7F3F79CFFFFF7F3F79CFFFF9CCFCE000000639A9C000000FF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FF7375739CFFFFF7F3F79CFFFFF7F3F79CFFFFF7
        F3F79CFFFFF7F3F79CFFFFF7F3F79CFFFFF7F3F79CFFFFF7F3F79CFFFFF7F3F7
        9CFFFFF7F3F79CFFFFF7F3F79CFFFFF7F3F79CFFFFF7F3F79CCFCE000000639A
        9C000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF737573F7F3F79C
        FFFF000000000000000000000000000000000000000000000000F7F3F79CFFFF
        000000000000000000000000000000000000000000000000000000000000F7F3
        F79CFFFF9CCFCE000000639A9C000000FF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FF7375739CFFFFF7F3F7737573009ACE00659C009ACE00659C009ACE
        00659C0000009CFFFFF7F3F79CFFFFF7F3F79CFFFFF7F3F79CFFFFF7F3F79CFF
        FFF7F3F79CFFFFF7F3F79CFFFFF7F3F79CCFCE000000639A9C000000FF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF737573F7F3F7F7F3F773757300CFFF
        009ACE00CFFF009ACE00CFFF009ACE000000F7F3F79CFFFF0000000000000000
        00000000000000000000000000000000000000000000F7F3F79CFFFF9CCFCE00
        0000639A9C000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF737573
        F7F3F7F7F3F77375739CFFFF00CFFF9CFFFF00CFFF9CFFFF00CFFF000000F7F3
        F7F7F3F79CFFFFF7F3F79CFFFFF7F3F79CFFFFF7F3F79CFFFFF7F3F79CFFFFF7
        F3F79CFFFFF7F3F79CCFCE000000639A9C000000FF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FF737573F7F3F7F7F3F77375739CFFFF9CFFFF9CFFFF9CFF
        FF9CFFFF9CFFFF000000F7F3F79CFFFF73757373757373757373757373757373
        7573737573737573737573737573F7F3F79CFFFF9CCFCE000000639A9C000000
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF7375739CCFCE9CCFCE0065
        9C009ACE009ACE009ACE009ACE009ACE009ACE0000009CCFCE9CCFCE9CCFCE9C
        CFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCE9CCFCEF7F3F7
        9CCFCE000000639A9C000000FF00FFFF00FFFF00FFFF00FF0000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000009CCFCE9CFFFF9CCFCE000000639A9C000000FF00FFFF00FFFF00
        FFFF00FF73757373757373757373757373757373757373757373757373757373
        7573737573737573737573737573737573737573737573737573737573737573
        7375737375737375737375737375730000009CCFCEF7F3F79CCFCE000000639A
        9C000000FF00FFFF00FFFF00FFFF00FF737573F7F3F7F7F3F7F7F3F7F7F3F7F7
        F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7
        F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F77375730000009CCF
        CE9CFFFF9CCFCE000000639A9C000000FF00FFFF00FFFF00FFFF00FF737573F7
        F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7
        F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3
        F7F7F3F77375730000009CCFCEF7F3F79CCFCE000000639A9C000000FF00FFFF
        00FFFF00FFFF00FF737573F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7
        F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3
        F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F77375730000009CCFCE9CFFFF9CCFCE00
        0000639A9C000000FF00FFFF00FFFF00FFFF00FF737573F7F3F7F7F3F7F7F3F7
        F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3
        F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F773757300
        00009CCFCEF7F3F79CCFCE000000639A9C000000FF00FFFF00FFFF00FFFF00FF
        737573F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7CE3000CE30
        00CE3000CE3000CE3000CE3000CE3000CE3000CE3000CE3000F7F3F7F7F3F7F7
        F3F7F7F3F7F7F3F77375730000009CCFCE9CFFFF9CCFCE000000639A9C000000
        FF00FFFF00FFFF00FFFF00FF737573F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3
        F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7
        F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F77375730000009CCFCEF7F3F7
        9CCFCE000000639A9C000000FF00FFFF00FFFF00FFFF00FF737573F7F3F7F7F3
        F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7FF6531CE3000FF6531CE3000FF
        6531CE3000FF6531CE3000FF6531CE3000F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7
        7375730000009CCFCE9CFFFF9CCFCE000000639A9C000000FF00FFFF00FFFF00
        FFFF00FF737573F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7
        F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7
        F7F3F7F7F3F7F7F3F7F7F3F77375730000009CCFCEF7F3F79CCFCE000000639A
        9C000000FF00FFFF00FFFF00FFFF00FF737573F7F3F7F7F3F7F7F3F7F7F3F7F7
        F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7
        F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F77375730000009CCF
        CE9CFFFF9CCFCE000000737573000000FF00FFFF00FFFF00FFFF00FF737573F7
        F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7
        F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3
        F7F7F3F77375730000009CCFCEF7F3F79CCFCE000000FF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FF737573F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7
        F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3
        F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F773757300000073757373757373757300
        0000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF737573F7F3F7F7F3F7F7F3F7
        F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3
        F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F773757300
        0000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        737573F7F3F7CE3000CE3000CE3000CE3000CE3000CE3000F7F3F7F7F3F7F7F3
        F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F700009C00
        009C00009CF7F3F7737573000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FF737573F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3
        F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7
        F3F7F7F3F7F7F3F70065FF3100FF00009CF7F3F7737573000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF737573F7F3F7FF65
        31CE3000FF6531CE3000FF6531CE3000F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7
        F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F70065FF0065FF00009CF7F3F7
        737573000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FF737573F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7
        F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7F7F3F7
        F7F3F7F7F3F7F7F3F7F7F3F7737573000000FF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FF73757373757373757373757373757373
        7573737573737573737573737573737573737573737573737573737573737573
        737573737573737573737573737573737573737573737573737573000000FF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FF}
      Transparent = True
    end
    object TDateMessage: THLabel
      Left = 60
      Top = 31
      Width = 80
      Height = 13
      Caption = 'Modifier/Lu le'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object DateMessage: TLabel
      Left = 154
      Top = 31
      Width = 78
      Height = 13
      Caption = 'DateMessage'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object TYMS_SUJET: THLabel
      Left = 12
      Top = 168
      Width = 25
      Height = 13
      Caption = 'Objet'
      FocusControl = YMS_SUJET
    end
    object TYMS_ALLDESTCC: THLabel
      Left = 12
      Top = 112
      Width = 20
      Height = 13
      Caption = 'CC :'
      WordWrap = True
    end
    object TYMS_TEL: THLabel
      Left = 596
      Top = 167
      Width = 18
      Height = 13
      Caption = 'T'#233'l.'
    end
    object TYMS_LIBFROM: THLabel
      Left = 12
      Top = 60
      Width = 20
      Height = 13
      Caption = 'De :'
    end
    object TYMS_ALLDEST: THLabel
      Left = 12
      Top = 80
      Width = 13
      Height = 13
      Caption = 'A :'
    end
    object YMS_LIBFROM: TLabel
      Left = 52
      Top = 60
      Width = 76
      Height = 13
      Caption = 'YMS_LIBFROM'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object ALLDEST: THLabel
      Left = 52
      Top = 79
      Width = 389
      Height = 30
      Hint = 'Destinataires'
      AutoSize = False
      ParentShowHint = False
      ShowHint = True
      WordWrap = True
      OnDblClick = DetailDestClick
    end
    object EnCopie: TLabel
      Left = 623
      Top = 4
      Width = 139
      Height = 13
      Caption = '(Message adress'#233' en copie)'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsItalic]
      ParentFont = False
      Visible = False
    end
    object ALLDESTCC: THLabel
      Left = 52
      Top = 111
      Width = 389
      Height = 30
      Hint = 'Destinataires'
      AutoSize = False
      ParentShowHint = False
      ShowHint = True
      WordWrap = True
      OnDblClick = DetailDestClick
    end
    object TDateCreationMsg: THLabel
      Left = 60
      Top = 10
      Width = 41
      Height = 13
      Caption = 'Cr'#233#233' le'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object DateCreationMsg: TLabel
      Left = 154
      Top = 10
      Width = 98
      Height = 13
      Caption = 'DateCreationMsg'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object GrpDossier: TGroupBox
      Left = 453
      Top = 17
      Width = 312
      Height = 135
      Caption = ' Dossier '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
      OnMouseMove = GrpDossierMouseMove
      object ANN_NOM2: TLabel
        Left = 8
        Top = 42
        Width = 60
        Height = 13
        Caption = 'ANN_NOM2'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object ANN_ALRUE1: TLabel
        Left = 8
        Top = 58
        Width = 71
        Height = 13
        Caption = 'ANN_ALRUE1'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object ANN_ALCP: TLabel
        Left = 8
        Top = 90
        Width = 56
        Height = 13
        Caption = 'ANN_ALCP'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object ANN_ALVILLE: TLabel
        Left = 72
        Top = 90
        Width = 71
        Height = 13
        Caption = 'ANN_ALVILLE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object ANN_ALRUE2: TLabel
        Left = 8
        Top = 74
        Width = 71
        Height = 13
        Caption = 'ANN_ALRUE2'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object ANN_NOM1: TLabel
        Left = 84
        Top = 19
        Width = 69
        Height = 13
        Caption = 'ANN_NOM1'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold, fsUnderline]
        ParentFont = False
        OnClick = ANN_NOM1Click
        OnMouseMove = ANN_NOM1MouseMove
      end
      object TANN_TEL1: TLabel
        Left = 8
        Top = 112
        Width = 21
        Height = 13
        Caption = 'T'#233'l :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsUnderline]
        ParentFont = False
      end
      object TANN_FAX: TLabel
        Left = 164
        Top = 112
        Width = 23
        Height = 13
        Caption = 'Fax :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsUnderline]
        ParentFont = False
      end
      object ANN_FAX: TLabel
        Left = 192
        Top = 112
        Width = 49
        Height = 13
        Caption = 'ANN_FAX'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsUnderline]
        ParentFont = False
      end
      object ANN_TEL1: TLabel
        Left = 32
        Top = 112
        Width = 55
        Height = 13
        Caption = 'ANN_TEL1'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsUnderline]
        ParentFont = False
      end
      object YMS_NODOSSIER: THCritMaskEdit
        Left = 8
        Top = 16
        Width = 69
        Height = 21
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        OnExit = YMS_NODOSSIERExit
        TagDispatch = 0
        ElipsisButton = True
        OnElipsisClick = YMS_NODOSSIERElipsisClick
      end
    end
    object YMS_URGENT: TCheckBox
      Left = 52
      Top = 144
      Width = 97
      Height = 17
      Caption = 'Urgent'
      TabOrder = 1
    end
    object YMS_PRIVE: TCheckBox
      Left = 196
      Top = 144
      Width = 97
      Height = 17
      Caption = 'Priv'#233
      TabOrder = 2
    end
    object YMS_SUJET: TEdit
      Left = 52
      Top = 164
      Width = 533
      Height = 21
      TabOrder = 3
      OnExit = YMS_SUJETExit
    end
    object YMS_TEL: TEdit
      Left = 620
      Top = 164
      Width = 145
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 4
    end
  end
  object PanelFiles: TPanel [4]
    Left = 543
    Top = 197
    Width = 202
    Height = 233
    Align = alRight
    TabOrder = 2
    Visible = False
    object PanelToolbar: TPanel
      Left = 1
      Top = 1
      Width = 200
      Height = 33
      Align = alTop
      Alignment = taLeftJustify
      BevelOuter = bvNone
      Caption = '  '
      TabOrder = 0
      DesignSize = (
        200
        33)
      object BOuvrirFichier: TToolbarButton97
        Left = 108
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Ouvrir'
        Anchors = [akTop, akRight]
        ImageIndex = 1
        Images = ImageListFile
        ParentShowHint = False
        ShowHint = True
        OnClick = BOuvrirFichierClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BEnregistrerSous: TToolbarButton97
        Left = 139
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Enregistrer sous'
        Anchors = [akTop, akRight]
        Glyph.Data = {
          DE010000424DDE01000000000000760000002800000024000000120000000100
          0400000000006801000000000000000000001000000010000000000084000084
          8400848484000000FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00111111111111
          111111111111111111111111000011112211111111111111111F111111111111
          0000111300211111321111111122F111111F1111000011130002111300211111
          12F12F111F22F11100001113000021300002111112F112F1F2112F1100001111
          300002000002111112F1112F211112F1000011111300000000211111112F1112
          1111F2110000111111300000021111111112F111111F21110000111111100000
          2111111111112F1111121111000011111113000021111111111112F111211111
          00001111113000002111111111111211112F1111000011111300020002111111
          11112111112F111100001111300021300021111111121112F112F11100001111
          3002111300021111112F11212F112F11000011111301111130001111112FF211
          12F112F100001111111111111303111111122111112FFF210000111111111111
          1111111111111111111222110000111111111111111111111111111111111111
          0000}
        ImageIndex = 2
        Images = ImageListFile
        ParentShowHint = False
        ShowHint = True
        OnClick = BEnregistrerSousClick
      end
      object BVisualiserFichier: TToolbarButton97
        Left = 170
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Visualiser'
        Anchors = [akTop, akRight]
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Glyph.Data = {
          BE060000424DBE06000000000000360400002800000024000000120000000100
          0800000000008802000000000000000000000001000000010000000000000000
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
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00030303030303
          03030606030303030303030303030303030303FFFF0303030303030303030303
          0303030303060404060303030303030303030303030303F8F8FF030303030303
          030303030303030303FE06060403030303030303030303030303F8FF03F8FF03
          0303030303030303030303030303FE060603030303030303030303030303F8FF
          FFF8FF0303030303030303030303030303030303030303030303030303030303
          030303F8F8030303030303030303030303030303030304040603030303030303
          0303030303030303FFFF03030303030303030303030303030306060604030303
          0303030303030303030303F8F8F8FF0303030303030303030303030303FE0606
          0403030303030303030303030303F8FF03F8FF03030303030303030303030303
          03FE06060604030303030303030303030303F8FF03F8FF030303030303030303
          030303030303FE060606040303030303030303030303F8FF0303F8FF03030303
          0303030303030303030303FE060606040303030303030303030303F8FF0303F8
          FF030303030303030303030404030303FE060606040303030303030303FF0303
          F8FF0303F8FF030303030303030306060604030303FE06060403030303030303
          F8F8FF0303F8FF0303F8FF03030303030303FE06060604040406060604030303
          030303F8FF03F8FFFFFFF80303F8FF0303030303030303FE0606060606060606
          06030303030303F8FF0303F8F8F8030303F8FF030303030303030303FEFE0606
          060606060303030303030303F8FFFF030303030303F803030303030303030303
          0303FEFEFEFEFE03030303030303030303F8F8FFFFFFFFFFF803030303030303
          0303030303030303030303030303030303030303030303F8F8F8F8F803030303
          0303}
        ImageIndex = 3
        Images = ImageListFile
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BVisualiserFichierClick
        IsControl = True
      end
    end
    object ListViewFiles: TListView
      Left = 1
      Top = 34
      Width = 200
      Height = 198
      Align = alClient
      Columns = <
        item
          Caption = 'Fichiers attach'#233's'
          Width = 180
        end>
      FlatScrollBars = True
      HideSelection = False
      LargeImages = ImageListLargeIcon
      MultiSelect = True
      ReadOnly = True
      ShowColumnHeaders = False
      SmallImages = ImageListLargeIcon
      SortType = stText
      TabOrder = 1
      ViewStyle = vsReport
    end
  end
  object PageControl: TPageControl [5]
    Left = 0
    Top = 197
    Width = 539
    Height = 233
    ActivePage = TabMessage
    Align = alClient
    Images = ImageListFile
    TabOrder = 3
    OnChange = PageControlChange
    object TabMessage: TTabSheet
      Caption = 'Message'
      object Web0: TWebBrowser_V1
        Left = 0
        Top = 0
        Width = 531
        Height = 204
        Align = alClient
        TabOrder = 0
        ControlData = {
          4C000000E1360000161500000000000000000000000000000000000000000000
          000000004C000000000000000000000001000000E0D057007335CF11AE690800
          2B2E12620C000000000000004C0000000114020000000000C000000000000046
          8000000000000000000000000000000000000000000000000000000000000000
          00000000000000000100000000000000000000000000000000000000}
      end
    end
    object TabText: TTabSheet
      Caption = 'Texte'
      ImageIndex = 4
      object HroMessage: THRichEditOLE
        Left = 0
        Top = 0
        Width = 531
        Height = 204
        Align = alClient
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        ScrollBars = ssVertical
        TabOrder = 0
        Margins.Top = 0
        Margins.Bottom = 0
        Margins.Left = 0
        Margins.Right = 0
        ContainerName = 'Document'
        ObjectMenuPrefix = '&Object'
        LinesRTF.Strings = (
          '{\rtf1\ansi\deff0{\fonttbl{\f0\fnil Arial;}}'
          
            '{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1\pard\lang103' +
            '6\f0\fs16 HroMessage'
          '\par }')
      end
    end
  end
  inherited HMTrad: THSystemMenu
    Left = 620
    Top = 344
  end
  object SaveDialog: TSaveDialog
    Left = 584
    Top = 252
  end
  object ImageListFile: THImageList
    GlobalIndexImages.Strings = (
      'Z0680_S16G1'
      'Z1100_S16G1'
      'Z0027_S16G1'
      'Z0056_S16G1'
      'Z0892_S16G1')
    Left = 616
    Top = 252
  end
  object ImageListLargeIcon: THImageList
    GlobalIndexImages.Strings = (
      'Z1019_S32G1')
    Height = 32
    Width = 32
    Left = 616
    Top = 284
  end
  object FilePopupMenu: TPopupMenu
    Images = ImageListFile
    Left = 656
    Top = 252
    object Ouvrir: TMenuItem
      Caption = 'Ouvrir'
      ImageIndex = 0
      OnClick = BOuvrirFichierClick
    end
    object EnregistrerSous: TMenuItem
      Caption = 'Enregistrer Sous'
      ImageIndex = 1
      OnClick = BEnregistrerSousClick
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object Visualiser: TMenuItem
      Caption = 'Visualiser'
      ImageIndex = 2
      OnClick = BVisualiserFichierClick
    end
  end
end
