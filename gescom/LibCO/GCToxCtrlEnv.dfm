object FicheEnveloppe: TFicheEnveloppe
  Left = 218
  Top = 107
  Width = 475
  Height = 326
  BorderIcons = [biSystemMenu]
  Caption = 'Contenu du fichier enveloppe'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCanResize = FormCanResize
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 0
    Top = 0
    Width = 467
    Height = 113
    Align = alTop
    Caption = 'Emetteur'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentColor = False
    ParentFont = False
    TabOrder = 0
    object Label1: TLabel
      Left = 8
      Top = 16
      Width = 18
      Height = 13
      Caption = 'Site'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 197
      Top = 16
      Width = 22
      Height = 13
      Caption = 'Nom'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label3: TLabel
      Left = 8
      Top = 40
      Width = 52
      Height = 13
      Caption = 'Application'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label4: TLabel
      Left = 8
      Top = 64
      Width = 23
      Height = 13
      Caption = 'Date'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label5: TLabel
      Left = 197
      Top = 64
      Width = 43
      Height = 13
      Caption = 'Message'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label6: TLabel
      Left = 8
      Top = 88
      Width = 34
      Height = 13
      Caption = 'Chrono'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label9: TLabel
      Left = 197
      Top = 88
      Width = 30
      Height = 13
      Caption = 'Status'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label10: TLabel
      Left = 197
      Top = 40
      Width = 54
      Height = 13
      Caption = 'Evènement'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object esite: TEdit
      Left = 71
      Top = 12
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 0
    end
    object eLibelle: TEdit
      Left = 257
      Top = 12
      Width = 200
      Height = 21
      Enabled = False
      TabOrder = 1
    end
    object eApplication: TEdit
      Left = 71
      Top = 36
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 2
    end
    object eTypeSend: TEdit
      Left = 257
      Top = 60
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 3
    end
    object DateMsg: TEdit
      Left = 72
      Top = 60
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 4
    end
    object NumChrono: TEdit
      Left = 72
      Top = 84
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 5
    end
    object eTypeMsg: TEdit
      Left = 257
      Top = 84
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 6
    end
    object eEvent: TEdit
      Left = 257
      Top = 36
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 7
    end
  end
  object GroupBox2: TGroupBox
    Left = 0
    Top = 113
    Width = 467
    Height = 40
    Align = alTop
    Caption = 'Destinataire'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    object Label7: TLabel
      Left = 8
      Top = 16
      Width = 18
      Height = 13
      Caption = 'Site'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label8: TLabel
      Left = 197
      Top = 16
      Width = 22
      Height = 13
      Caption = 'Nom'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object dSite: TEdit
      Left = 71
      Top = 12
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 0
    end
    object dLibelle: TEdit
      Left = 257
      Top = 12
      Width = 200
      Height = 21
      Enabled = False
      TabOrder = 1
    end
  end
  object GroupBox3: TGroupBox
    Left = 0
    Top = 153
    Width = 467
    Height = 146
    Align = alClient
    Caption = 'Liste des fichiers'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    object G: THGrid
      Left = 2
      Top = 15
      Width = 463
      Height = 129
      Align = alClient
      ColCount = 2
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 2
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
      ParentFont = False
      TabOrder = 0
      SortedCol = -1
      Titres.Strings = (
        'Fichier'
        'CRC')
      Couleur = False
      MultiSelect = False
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
      ColWidths = (
        326
        132)
    end
  end
end
