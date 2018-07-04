object TFDeseria: TTFDeseria
  Left = 282
  Top = 288
  Width = 576
  Height = 379
  Caption = 
    'Confirmation Electronique d'#39'arr'#234't d'#39'utilisation des produits BUS' +
    'INESS'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnShow = FormShow
  DesignSize = (
    560
    341)
  PixelsPerInch = 96
  TextHeight = 13
  object lEtape: TLabel
    Left = 16
    Top = 313
    Width = 62
    Height = 14
    Anchors = [akLeft, akBottom]
    Caption = 'Etape 1/N'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object P: TPageControl
    Left = 0
    Top = 0
    Width = 560
    Height = 289
    ActivePage = TabSheet4
    Align = alTop
    TabOrder = 0
    OnChange = PChange
    object TabSheet1: TTabSheet
      Caption = 'Informations importantes'
      object MDefinition: TMemo
        Left = 0
        Top = 0
        Width = 552
        Height = 261
        Align = alClient
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlue
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        Lines.Strings = (
          'IMPORTANT.'
          ''
          
            'L'#39'utilisation de cet outil permet de confirmer, l'#39'arr'#234't de l'#39'uti' +
            'lisation des '
          'produits CEGID et L.S.E BUSINESS.'
          ''
          
            'Un message de confirmation d'#39'arr'#234't d'#39'utilisation des produits se' +
            'ra '
          'envoy'#233' '#224' L.S.E.'
          ''
          
            'Cliquer sur "Suivant'#39' pour continuer ou "Annuler" pour ressortir' +
            '.')
        ParentFont = False
        TabOrder = 0
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Votre Identification '
      ImageIndex = 1
      object Label2: TLabel
        Left = 40
        Top = 56
        Width = 54
        Height = 13
        Caption = 'Code Client'
      end
      object Label3: TLabel
        Left = 40
        Top = 78
        Width = 87
        Height = 13
        Caption = 'Nom Responsable'
      end
      object Label4: TLabel
        Left = 40
        Top = 100
        Width = 29
        Height = 13
        Caption = 'E Mail'
      end
      object Label5: TLabel
        Left = 40
        Top = 122
        Width = 36
        Height = 13
        Caption = 'Soci'#233't'#233
      end
      object Adress: TLabel
        Left = 40
        Top = 144
        Width = 38
        Height = 13
        Caption = 'Adresse'
      end
      object Label6: TLabel
        Left = 0
        Top = 245
        Width = 342
        Height = 16
        Align = alBottom
        Alignment = taCenter
        Caption = 'Les champs marqu'#233's par une '#233'toile sont obligatoires'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label7: TLabel
        Left = 288
        Top = 56
        Width = 8
        Height = 16
        Caption = '*'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label8: TLabel
        Left = 440
        Top = 77
        Width = 8
        Height = 16
        Caption = '*'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label9: TLabel
        Left = 440
        Top = 98
        Width = 8
        Height = 16
        Caption = '*'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label10: TLabel
        Left = 440
        Top = 120
        Width = 8
        Height = 16
        Caption = '*'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label11: TLabel
        Left = 440
        Top = 142
        Width = 8
        Height = 16
        Caption = '*'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label12: TLabel
        Left = 440
        Top = 209
        Width = 8
        Height = 16
        Caption = '*'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label19: TLabel
        Left = 0
        Top = 0
        Width = 140
        Height = 19
        Align = alTop
        Alignment = taCenter
        Caption = 'IDENTIFICATION'
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlue
        Font.Height = -16
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object CodeClient: TMaskEdit
        Left = 131
        Top = 52
        Width = 145
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      object Responsable: TMaskEdit
        Left = 131
        Top = 74
        Width = 304
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
      end
      object EmailResp: TMaskEdit
        Left = 131
        Top = 96
        Width = 304
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
      end
      object Societe: TMaskEdit
        Left = 131
        Top = 118
        Width = 304
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
      end
      object Adr1: TMaskEdit
        Left = 131
        Top = 140
        Width = 304
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
      end
      object Adr2: TMaskEdit
        Left = 131
        Top = 162
        Width = 304
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
      end
      object Adr3: TMaskEdit
        Left = 131
        Top = 184
        Width = 304
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 6
      end
      object CodePostal: TMaskEdit
        Left = 131
        Top = 207
        Width = 49
        Height = 22
        EditMask = '00000;1; '
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        MaxLength = 5
        ParentFont = False
        TabOrder = 7
        Text = '     '
      end
      object Ville: TMaskEdit
        Left = 185
        Top = 207
        Width = 250
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 8
      end
    end
    object TabSheet3: TTabSheet
      Caption = 'Messagerie'
      ImageIndex = 2
      object Label13: TLabel
        Left = 40
        Top = 56
        Width = 70
        Height = 13
        Caption = 'Serveur SMTP'
      end
      object Label14: TLabel
        Left = 432
        Top = 54
        Width = 8
        Height = 16
        Caption = '*'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label15: TLabel
        Left = 40
        Top = 79
        Width = 19
        Height = 13
        Caption = 'Port'
      end
      object Luser: TLabel
        Left = 40
        Top = 100
        Width = 36
        Height = 13
        Caption = 'Compte'
        Visible = False
      end
      object LPAssword: TLabel
        Left = 40
        Top = 122
        Width = 64
        Height = 13
        Caption = 'Mot de passe'
        Visible = False
      end
      object Label18: TLabel
        Left = 432
        Top = 77
        Width = 8
        Height = 16
        Caption = '*'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label16: TLabel
        Left = 0
        Top = 245
        Width = 342
        Height = 16
        Align = alBottom
        Alignment = taCenter
        Caption = 'Les champs marqu'#233's par une '#233'toile sont obligatoires'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label20: TLabel
        Left = 0
        Top = 0
        Width = 194
        Height = 19
        Align = alTop
        Alignment = taCenter
        Caption = 'MOYEN DE TRANSPORT'
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlue
        Font.Height = -16
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object SMTPAdress: TMaskEdit
        Left = 131
        Top = 51
        Width = 292
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      object SMTPPort: TMaskEdit
        Left = 131
        Top = 74
        Width = 62
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        MaxLength = 4
        ParentFont = False
        TabOrder = 1
        Text = '25'
      end
      object CBAuthentification: TCheckBox
        Left = 200
        Top = 77
        Width = 145
        Height = 17
        Caption = 'Authentification requise'
        TabOrder = 2
        OnClick = CBAuthentificationClick
      end
      object SMTPUser: TMaskEdit
        Left = 131
        Top = 96
        Width = 292
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        Visible = False
      end
      object SMTPPassword: TMaskEdit
        Left = 131
        Top = 118
        Width = 292
        Height = 24
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
        PasswordChar = '*'
        TabOrder = 4
        Visible = False
      end
    end
    object TabSheet4: TTabSheet
      Caption = 'Informations finale'
      ImageIndex = 3
      object Mresult: TMemo
        Left = 0
        Top = 0
        Width = 552
        Height = 273
        Align = alTop
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        Lines.Strings = (
          'Mresult')
        ParentFont = False
        TabOrder = 0
      end
    end
  end
  object BFIn: TBitBtn
    Left = 471
    Top = 308
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = 'Envoyer'
    TabOrder = 1
    OnClick = BFInClick
  end
  object BSuivant: TBitBtn
    Left = 391
    Top = 308
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = 'Suivant >'
    TabOrder = 2
    OnClick = BSuivantClick
  end
  object BPrecedent: TBitBtn
    Left = 311
    Top = 308
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = '< Pr'#233'c'#233'dent'
    TabOrder = 3
    OnClick = BPrecedentClick
  end
  object BAnnuler: TBitBtn
    Left = 223
    Top = 308
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Cancel = True
    Caption = 'Annuler'
    TabOrder = 4
    OnClick = BAnnulerClick
  end
  object Plan: TPanel
    Left = 8
    Top = 24
    Width = 161
    Height = 25
    TabOrder = 5
  end
  object Ccontrols: TListBox
    Left = 96
    Top = 308
    Width = 121
    Height = 25
    Color = clYellow
    ItemHeight = 13
    TabOrder = 6
    Visible = False
  end
end
