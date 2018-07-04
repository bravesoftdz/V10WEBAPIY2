object FEDI: TFEDI
  Left = 171
  Top = 133
  Width = 390
  Height = 357
  HelpContext = 6370000
  BorderIcons = [biSystemMenu]
  Caption = 'Param'#233'trage des '#233'changes EDIFICAS'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 295
    Width = 382
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    Caption = ' '
    TabOrder = 1
    object BtAide: THBitBtn
      Left = 348
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Aide'
      TabOrder = 2
      OnClick = BtAideClick
      GlobalIndexImage = 'Z1117_S16G1'
    end
    object BtSortir: THBitBtn
      Left = 316
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ModalResult = 2
      TabOrder = 1
      OnClick = BtSortirClick
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BtValider: THBitBtn
      Left = 284
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Valider'
      TabOrder = 0
      OnClick = BtValiderClick
      NumGlyphs = 2
      GlobalIndexImage = 'Z0003_S16G2'
    end
  end
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 382
    Height = 295
    ActivePage = TabSheet3
    Align = alClient
    TabOrder = 0
    object TabSheet1: TTabSheet
      Caption = 'Param'#232'tres'
      object GroupBox1: TGroupBox
        Left = 0
        Top = 0
        Width = 373
        Height = 265
        HelpContext = 6371000
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object LabRefDem: THLabel
          Left = 9
          Top = 20
          Width = 97
          Height = 13
          Caption = '&R'#233'f'#233'rence demande'
          FocusControl = EditRefDem
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LabFctTxt: THLabel
          Left = 9
          Top = 50
          Width = 82
          Height = 13
          Caption = '&Fonction du texte'
          FocusControl = CombFctTxt
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LabRefTxt: THLabel
          Left = 228
          Top = 52
          Width = 91
          Height = 13
          Caption = '&R'#233'f'#233'rence du texte'
          FocusControl = EditRefTxt
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LabTxt: THLabel
          Left = 9
          Top = 80
          Width = 27
          Height = 13
          Caption = '&Texte'
          FocusControl = EditTxt
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LabNumDosEmt: THLabel
          Left = 9
          Top = 112
          Width = 147
          Height = 13
          Caption = '&Num'#233'ro dossier chez l'#39#233'metteur'
          FocusControl = EditNumDosEmt
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LabNumDosRec: THLabel
          Left = 9
          Top = 144
          Width = 158
          Height = 13
          Caption = 'N&um'#233'ro dossier chez le r'#233'cepteur'
          FocusControl = EditNumDosRec
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LabIdentEnv: THLabel
          Left = 9
          Top = 175
          Width = 108
          Height = 13
          Caption = '&Identification de l'#39'envoi'
          FocusControl = EditIdentEnv
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label5: TLabel
          Left = 6
          Top = 0
          Width = 109
          Height = 13
          Caption = ' Param'#232'tres g'#233'n'#233'raux  '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clActiveCaption
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LabRefIntChg: TLabel
          Left = 9
          Top = 208
          Width = 169
          Height = 13
          Caption = 'R'#233'&f'#233'rence contr'#244'le de l'#39'interchange'
          FocusControl = EditRefIntChg
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object EditRefDem: TEdit
          Left = 114
          Top = 16
          Width = 251
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
        end
        object CombFctTxt: TComboBox
          Left = 114
          Top = 47
          Width = 43
          Height = 20
          Style = csOwnerDrawFixed
          Ctl3D = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 14
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 1
          Items.Strings = (
            'AU'
            'EC'
            'CC')
        end
        object EditRefTxt: TEdit
          Left = 338
          Top = 48
          Width = 27
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 3
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 2
        end
        object EditTxt: TEdit
          Left = 114
          Top = 76
          Width = 251
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 3
        end
        object EditNumDosEmt: TEdit
          Left = 183
          Top = 108
          Width = 183
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 17
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 4
        end
        object EditNumDosRec: TEdit
          Left = 183
          Top = 140
          Width = 183
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 17
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 5
        end
        object EditIdentEnv: TEdit
          Left = 183
          Top = 172
          Width = 183
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 17
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 6
        end
        object EditRefIntChg: TEdit
          Left = 183
          Top = 204
          Width = 105
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 14
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 7
        end
        object CBoxEssai: TCheckBox
          Left = 156
          Top = 238
          Width = 133
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Echange '#224' titre d'#39'essai'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 9
        end
        object CBoxAccRec: TCheckBox
          Left = 8
          Top = 239
          Width = 125
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Accus'#233' de r'#233'ception'
          Checked = True
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          State = cbChecked
          TabOrder = 8
        end
      end
    end
    object TabSheet3: TTabSheet
      HelpContext = 6372000
      Caption = 'Dossier'
      object Groupdos: TGroupBox
        Left = 0
        Top = 0
        Width = 373
        Height = 265
        Caption = 'Groupdos'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        object DLabSiret: TLabel
          Left = 10
          Top = 28
          Width = 96
          Height = 13
          Caption = '&Code SIRET dossier'
          FocusControl = DEditSiret
        end
        object DLabFctCont: TLabel
          Left = 275
          Top = 28
          Width = 57
          Height = 13
          Caption = '&Fct. contact'
          FocusControl = DEditFctCont
        end
        object DLabPerChrg: TLabel
          Left = 10
          Top = 56
          Width = 96
          Height = 13
          Caption = '&Personne en charge'
          FocusControl = DEditPerChrg
        end
        object DLabRue1: TLabel
          Left = 10
          Top = 145
          Width = 29
          Height = 13
          Caption = 'R&ue 1'
          FocusControl = DEditRue1
        end
        object DLabDiv: TLabel
          Left = 10
          Top = 234
          Width = 37
          Height = 13
          Caption = '&Division'
          FocusControl = DEditDiv
        end
        object DLabCodPost: TLabel
          Left = 144
          Top = 234
          Width = 56
          Height = 13
          Caption = 'Code &postal'
          FocusControl = DEditCodPost
        end
        object DLabPays: TLabel
          Left = 309
          Top = 234
          Width = 23
          Height = 13
          Caption = 'Pa&ys'
          FocusControl = DEditPays
        end
        object DLabVille: TLabel
          Left = 10
          Top = 205
          Width = 19
          Height = 13
          Caption = '&Ville'
          FocusControl = DEditVille
        end
        object DLabRue2: TLabel
          Left = 10
          Top = 175
          Width = 29
          Height = 13
          Caption = 'Ru&e 2'
          FocusControl = DEditRue2
        end
        object DLabNumCom: TLabel
          Left = 10
          Top = 86
          Width = 38
          Height = 13
          Caption = '&N'#176' com.'
          FocusControl = DEditNumCom
        end
        object DLabModCom: TLabel
          Left = 275
          Top = 86
          Width = 53
          Height = 13
          Caption = '&Mode com.'
          Color = clYellow
          FocusControl = DEditModCom
          ParentColor = False
          Visible = False
        end
        object DLabNom: TLabel
          Left = 10
          Top = 114
          Width = 22
          Height = 13
          Caption = 'No&m'
          FocusControl = DEditNom
        end
        object Label2: TLabel
          Left = 8
          Top = 0
          Width = 231
          Height = 13
          Caption = ' Coordonn'#233'es li'#233'es aux donn'#233'es communiqu'#233'es '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clActiveCaption
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object DEditSiret: TEdit
          Left = 124
          Top = 24
          Width = 125
          Height = 21
          Ctl3D = True
          MaxLength = 17
          ParentCtl3D = False
          TabOrder = 0
        end
        object DEditFctCont: TEdit
          Left = 334
          Top = 24
          Width = 27
          Height = 21
          Ctl3D = True
          MaxLength = 3
          ParentCtl3D = False
          TabOrder = 1
          Text = 'EC'
        end
        object DEditPerChrg: TEdit
          Left = 124
          Top = 53
          Width = 237
          Height = 21
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 2
        end
        object DEditRue1: TEdit
          Left = 52
          Top = 141
          Width = 309
          Height = 21
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 6
        end
        object DEditDiv: TEdit
          Left = 52
          Top = 231
          Width = 69
          Height = 21
          Ctl3D = True
          MaxLength = 9
          ParentCtl3D = False
          TabOrder = 9
        end
        object DEditCodPost: TEdit
          Left = 206
          Top = 230
          Width = 69
          Height = 21
          Ctl3D = True
          MaxLength = 9
          ParentCtl3D = False
          TabOrder = 10
        end
        object DEditPays: TEdit
          Left = 341
          Top = 230
          Width = 20
          Height = 21
          Ctl3D = True
          MaxLength = 2
          ParentCtl3D = False
          TabOrder = 11
        end
        object DEditVille: TEdit
          Left = 52
          Top = 201
          Width = 309
          Height = 21
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 8
        end
        object DEditRue2: TEdit
          Left = 52
          Top = 171
          Width = 309
          Height = 21
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 7
        end
        object DEditNumCom: TEdit
          Left = 52
          Top = 82
          Width = 181
          Height = 21
          Ctl3D = True
          MaxLength = 25
          ParentCtl3D = False
          TabOrder = 3
        end
        object DEditModCom: TEdit
          Left = 334
          Top = 82
          Width = 27
          Height = 21
          Color = clYellow
          Ctl3D = True
          MaxLength = 3
          ParentCtl3D = False
          TabOrder = 4
          Text = 'TE'
          Visible = False
        end
        object DEditNom: TEdit
          Left = 52
          Top = 111
          Width = 309
          Height = 21
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 5
        end
      end
    end
    object TabSheet4: TTabSheet
      HelpContext = 6373000
      Caption = 'Emetteur'
      object GroupEmet: TGroupBox
        Left = 0
        Top = 0
        Width = 373
        Height = 265
        Caption = 'GroupEmet'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        object ELabSiret: TLabel
          Left = 10
          Top = 71
          Width = 104
          Height = 13
          Caption = '&Code SIRET '#233'metteur'
          FocusControl = EEditSiret
        end
        object ELabFctCont: TLabel
          Left = 278
          Top = 71
          Width = 57
          Height = 13
          Caption = '&Fct. contact'
          FocusControl = EEditFctCont
        end
        object ELabPerChrg: TLabel
          Left = 10
          Top = 94
          Width = 96
          Height = 13
          Caption = '&Personne en charge'
          FocusControl = EEditPerChrg
        end
        object ELabNumCom: TLabel
          Left = 10
          Top = 119
          Width = 38
          Height = 13
          Caption = '&N'#176' com.'
          FocusControl = EEditNumCom
        end
        object ELabModCom: TLabel
          Left = 278
          Top = 119
          Width = 53
          Height = 13
          Caption = '&Mode com.'
          Color = clYellow
          FocusControl = EEditModCom
          ParentColor = False
          Visible = False
        end
        object ELabRue1: TLabel
          Left = 10
          Top = 167
          Width = 29
          Height = 13
          Caption = 'R&ue 1'
          FocusControl = EEditRue1
        end
        object ELabRue2: TLabel
          Left = 10
          Top = 192
          Width = 29
          Height = 13
          Caption = 'Ru&e 2'
          FocusControl = EEditRue2
        end
        object ELabVille: TLabel
          Left = 10
          Top = 217
          Width = 19
          Height = 13
          Caption = '&Ville'
          FocusControl = EEditVille
        end
        object ELabDiv: TLabel
          Left = 10
          Top = 242
          Width = 37
          Height = 13
          Caption = '&Division'
          FocusControl = EEditDiv
        end
        object ELabCodPost: TLabel
          Left = 155
          Top = 242
          Width = 56
          Height = 13
          Caption = 'Code po&stal'
          FocusControl = EEditCodPost
        end
        object ELabPays: TLabel
          Left = 312
          Top = 242
          Width = 23
          Height = 13
          Caption = 'Pa&ys'
          FocusControl = EEditPays
        end
        object ELabNom: TLabel
          Left = 10
          Top = 143
          Width = 22
          Height = 13
          Caption = 'No&m'
          FocusControl = EEditNom
        end
        object Label3: TLabel
          Left = 8
          Top = 0
          Width = 264
          Height = 13
          Caption = ' Coordonn'#233'es li'#233'es '#224' la soci'#233't'#233' d'#233'tentrice des donn'#233'es '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clActiveCaption
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LabEmet: TLabel
          Left = 11
          Top = 23
          Width = 42
          Height = 13
          Caption = '&Emetteur'
          FocusControl = EditEmet
        end
        object LabAdrRec: TLabel
          Left = 11
          Top = 47
          Width = 71
          Height = 13
          Caption = '&Adresse retour '
          FocusControl = EditAdrEmet
        end
        object EEditSiret: TEdit
          Left = 124
          Top = 67
          Width = 125
          Height = 21
          Ctl3D = True
          MaxLength = 17
          ParentCtl3D = False
          TabOrder = 0
        end
        object EEditFctCont: TEdit
          Left = 337
          Top = 67
          Width = 27
          Height = 21
          Ctl3D = True
          MaxLength = 3
          ParentCtl3D = False
          TabOrder = 1
          Text = 'EC'
        end
        object EEditPerChrg: TEdit
          Left = 124
          Top = 91
          Width = 240
          Height = 21
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 2
        end
        object EEditNumCom: TEdit
          Left = 52
          Top = 115
          Width = 196
          Height = 21
          Ctl3D = True
          MaxLength = 25
          ParentCtl3D = False
          TabOrder = 3
        end
        object EEditModCom: TEdit
          Left = 337
          Top = 115
          Width = 27
          Height = 21
          Color = clYellow
          Ctl3D = True
          MaxLength = 3
          ParentCtl3D = False
          TabOrder = 4
          Text = 'TE'
          Visible = False
        end
        object EEditRue1: TEdit
          Left = 52
          Top = 164
          Width = 312
          Height = 21
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 6
        end
        object EEditRue2: TEdit
          Left = 52
          Top = 189
          Width = 312
          Height = 21
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 7
        end
        object EEditVille: TEdit
          Left = 52
          Top = 214
          Width = 312
          Height = 21
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 8
        end
        object EEditDiv: TEdit
          Left = 52
          Top = 239
          Width = 69
          Height = 21
          Ctl3D = True
          MaxLength = 9
          ParentCtl3D = False
          TabOrder = 9
        end
        object EEditCodPost: TEdit
          Left = 214
          Top = 238
          Width = 69
          Height = 21
          Ctl3D = True
          MaxLength = 9
          ParentCtl3D = False
          TabOrder = 10
        end
        object EEditPays: TEdit
          Left = 344
          Top = 238
          Width = 20
          Height = 21
          Ctl3D = True
          MaxLength = 2
          ParentCtl3D = False
          TabOrder = 11
        end
        object EEditNom: TEdit
          Left = 52
          Top = 139
          Width = 312
          Height = 21
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 5
        end
        object EditEmet: TEdit
          Left = 88
          Top = 19
          Width = 276
          Height = 21
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 12
        end
        object EditAdrEmet: TEdit
          Left = 88
          Top = 43
          Width = 276
          Height = 21
          Ctl3D = True
          MaxLength = 14
          ParentCtl3D = False
          TabOrder = 13
        end
      end
    end
    object TabSheet5: TTabSheet
      HelpContext = 6374000
      Caption = 'R'#233'cepteur'
      object GroupRec: TGroupBox
        Left = 0
        Top = 0
        Width = 373
        Height = 265
        Caption = 'GroupRec'
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        object RLabSiret: TLabel
          Left = 10
          Top = 71
          Width = 108
          Height = 13
          Caption = '&Code SIRET r'#233'cepteur'
          FocusControl = REditSiret
        end
        object RLabFctCont: TLabel
          Left = 279
          Top = 71
          Width = 57
          Height = 13
          Caption = '&Fct. contact'
          FocusControl = REditFctCont
        end
        object RLabPerChrg: TLabel
          Left = 10
          Top = 94
          Width = 96
          Height = 13
          Caption = '&Personne en charge'
          FocusControl = RComboPerChrg
        end
        object RLabNumCom: TLabel
          Left = 10
          Top = 119
          Width = 38
          Height = 13
          Caption = '&N'#176' com.'
          FocusControl = REditNumCom
        end
        object RLabModCom: TLabel
          Left = 278
          Top = 117
          Width = 53
          Height = 13
          Caption = '&Mode com.'
          Color = clYellow
          FocusControl = REditModCom
          ParentColor = False
          Visible = False
        end
        object RLabRue1: TLabel
          Left = 10
          Top = 167
          Width = 29
          Height = 13
          Caption = 'R&ue 1'
          FocusControl = REditRue1
        end
        object RLabRue2: TLabel
          Left = 10
          Top = 192
          Width = 29
          Height = 13
          Caption = 'Ru&e 2'
          FocusControl = REditRue2
        end
        object RLabVille: TLabel
          Left = 10
          Top = 217
          Width = 19
          Height = 13
          Caption = '&Ville'
          FocusControl = REditVille
        end
        object RLabDiv: TLabel
          Left = 10
          Top = 242
          Width = 37
          Height = 13
          Caption = '&Division'
          FocusControl = REditDiv
        end
        object RLabCodPost: TLabel
          Left = 155
          Top = 242
          Width = 56
          Height = 13
          Hint = 'REditCodPost'
          Caption = 'Code po&stal'
        end
        object RLabPays: TLabel
          Left = 312
          Top = 242
          Width = 23
          Height = 13
          Caption = 'Pa&ys'
          FocusControl = REditPays
        end
        object RLabNom: TLabel
          Left = 10
          Top = 143
          Width = 22
          Height = 13
          Caption = 'No&m'
          FocusControl = REditNom
        end
        object Label4: TLabel
          Left = 8
          Top = 0
          Width = 271
          Height = 13
          Caption = ' Coordonn'#233'es li'#233'es '#224' la soci'#233't'#233' destinatrice des donn'#233'es '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clActiveCaption
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LabAdrEmet: TLabel
          Left = 11
          Top = 47
          Width = 68
          Height = 13
          Caption = '&Adresse retour'
          FocusControl = EditAdrRec
        end
        object LabRec: TLabel
          Left = 11
          Top = 23
          Width = 50
          Height = 13
          Caption = '&R'#233'cepteur'
          FocusControl = EditRec1
        end
        object EEditRec: TLabel
          Left = 194
          Top = 23
          Width = 3
          Height = 13
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object REditSiret: TEdit
          Left = 124
          Top = 67
          Width = 125
          Height = 21
          Ctl3D = True
          MaxLength = 17
          ParentCtl3D = False
          TabOrder = 1
        end
        object REditFctCont: TEdit
          Left = 338
          Top = 67
          Width = 27
          Height = 21
          Ctl3D = True
          MaxLength = 3
          ParentCtl3D = False
          TabOrder = 2
          Text = 'EC'
        end
        object REditNumCom: TEdit
          Left = 52
          Top = 115
          Width = 197
          Height = 21
          Ctl3D = True
          MaxLength = 25
          ParentCtl3D = False
          TabOrder = 4
        end
        object REditModCom: TEdit
          Left = 337
          Top = 115
          Width = 27
          Height = 21
          Color = clYellow
          Ctl3D = True
          MaxLength = 3
          ParentCtl3D = False
          TabOrder = 5
          Text = 'TE'
          Visible = False
        end
        object REditRue1: TEdit
          Left = 52
          Top = 164
          Width = 312
          Height = 21
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 7
        end
        object REditRue2: TEdit
          Left = 52
          Top = 189
          Width = 312
          Height = 21
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 8
        end
        object REditVille: TEdit
          Left = 52
          Top = 214
          Width = 312
          Height = 21
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 9
        end
        object REditDiv: TEdit
          Left = 52
          Top = 239
          Width = 69
          Height = 21
          Ctl3D = True
          MaxLength = 9
          ParentCtl3D = False
          TabOrder = 10
        end
        object REditCodPost: TEdit
          Left = 215
          Top = 238
          Width = 69
          Height = 21
          Ctl3D = True
          MaxLength = 9
          ParentCtl3D = False
          TabOrder = 11
        end
        object REditPays: TEdit
          Left = 344
          Top = 238
          Width = 20
          Height = 21
          Ctl3D = True
          MaxLength = 2
          ParentCtl3D = False
          TabOrder = 12
        end
        object REditNom: TEdit
          Left = 52
          Top = 139
          Width = 312
          Height = 21
          Ctl3D = True
          ParentCtl3D = False
          TabOrder = 6
        end
        object RComboPerChrg: TComboBox
          Left = 124
          Top = 91
          Width = 241
          Height = 21
          ItemHeight = 0
          TabOrder = 3
          OnChange = RComboPerChrgChange
        end
        object RComboTel: TComboBox
          Left = 172
          Top = 92
          Width = 73
          Height = 21
          Color = clYellow
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 0
          Visible = False
        end
        object EditRec1: THCpteEdit
          Left = 88
          Top = 19
          Width = 105
          Height = 21
          CharCase = ecUpperCase
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 14
          OnExit = EditRec1Exit
          ZoomTable = tzTiers
          Vide = False
          Bourre = False
          Libelle = EEditRec
          okLocate = False
          SynJoker = False
        end
        object EditAdrRec: TEdit
          Left = 88
          Top = 43
          Width = 277
          Height = 21
          Ctl3D = True
          MaxLength = 14
          ParentCtl3D = False
          TabOrder = 15
        end
        object EditRec: TEdit
          Left = 108
          Top = 18
          Width = 37
          Height = 21
          Color = clYellow
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 14
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 13
          Text = 'EditRec'
          Visible = False
        end
      end
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Enregistrement inaccessible')
    Left = 88
    Top = 276
  end
  object Query: TQuery
    DatabaseName = 'SOC'
    RequestLive = True
    Left = 40
    Top = 276
  end
  object TFMTCHOIX: THTable
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    IndexName = 'IC_CLE1'
    TableName = 'FMTCHOIX'
    dataBaseName = 'SOC'
    RequestLive = True
    Left = 137
    Top = 274
    object TFMTCHOIXIC_IMPORT: TStringField
      FieldName = 'IC_IMPORT'
      Size = 1
    end
    object TFMTCHOIXIC_NATURE: TStringField
      FieldName = 'IC_NATURE'
      Size = 3
    end
    object TFMTCHOIXIC_FORMAT: TStringField
      FieldName = 'IC_FORMAT'
      Size = 3
    end
    object TFMTCHOIXIC_DETRUIRE: TStringField
      FieldName = 'IC_DETRUIRE'
      Size = 1
    end
    object TFMTCHOIXIC_RESUME: TStringField
      FieldName = 'IC_RESUME'
      Size = 1
    end
    object TFMTCHOIXIC_VALIDE: TStringField
      FieldName = 'IC_VALIDE'
      Size = 1
    end
    object TFMTCHOIXIC_RECYCLE: TStringField
      FieldName = 'IC_RECYCLE'
      Size = 1
    end
    object TFMTCHOIXIC_FICHIER: TStringField
      FieldName = 'IC_FICHIER'
      Size = 100
    end
    object TFMTCHOIXIC_ASCII: TStringField
      FieldName = 'IC_ASCII'
      Size = 1
    end
    object TFMTCHOIXIC_FIXEFORMAT: TStringField
      FieldName = 'IC_FIXEFORMAT'
      Size = 1
    end
    object TFMTCHOIXIC_FIXEFICHIER: TStringField
      FieldName = 'IC_FIXEFICHIER'
      Size = 1
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 32
    Top = 12
  end
end
