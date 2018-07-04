inherited FChangePwdDossier: TFChangePwdDossier
  Left = 343
  Top = 214
  VertScrollBar.Range = 0
  BorderStyle = bsDialog
  Caption = 'Mot de passe du dossier'
  ClientHeight = 228
  ClientWidth = 283
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 193
    Width = 283
    inherited PBouton: TToolWindow97
      ClientWidth = 283
      ClientAreaWidth = 283
      inherited BValider: TToolbarButton97
        Left = 187
        ModalResult = 0
      end
      inherited BFerme: TToolbarButton97
        Left = 219
      end
      inherited HelpBtn: TToolbarButton97
        Left = 251
        Visible = False
      end
      inherited BImprimer: TToolbarButton97
        Left = 155
      end
    end
  end
  object Pages: TPageControl [1]
    Left = 0
    Top = 0
    Width = 283
    Height = 193
    ActivePage = TabPwd
    Align = alClient
    TabOrder = 1
    object TabPwd: TTabSheet
      Caption = 'Mot de passe'
      object TOldPassWord: THLabel
        Left = 16
        Top = 14
        Width = 99
        Height = 13
        Caption = 'Ancien mot de passe'
        FocusControl = FOldPassWord
      end
      object FOldPassWord: TEdit
        Left = 144
        Top = 10
        Width = 109
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 8
        PasswordChar = '*'
        TabOrder = 0
        OnChange = Password_OnChange
      end
      object FPwdGlobal: TCheckBox
        Left = 16
        Top = 136
        Width = 233
        Height = 17
        Caption = 'Mot de passe pour l'#39'acc'#232's au dossier client'
        TabOrder = 2
        OnClick = FPwdGlobalClick
      end
      object GrpPwd: TGroupBox
        Left = 4
        Top = 41
        Width = 265
        Height = 82
        TabOrder = 1
        object TPassWord: THLabel
          Left = 11
          Top = 24
          Width = 110
          Height = 13
          Caption = 'Nouveau mot de passe'
          Enabled = False
          FocusControl = FPassWord
        end
        object TConfirm: THLabel
          Left = 12
          Top = 52
          Width = 58
          Height = 13
          Caption = 'Confirmation'
          Enabled = False
          FocusControl = FConfirm
        end
        object FPassWord: TEdit
          Left = 140
          Top = 48
          Width = 109
          Height = 21
          CharCase = ecUpperCase
          Enabled = False
          MaxLength = 8
          PasswordChar = '*'
          TabOrder = 2
          OnChange = Password_OnChange
        end
        object FConfirm: TEdit
          Left = 140
          Top = 20
          Width = 109
          Height = 21
          CharCase = ecUpperCase
          Enabled = False
          MaxLength = 8
          PasswordChar = '*'
          TabOrder = 1
        end
        object FChgPwd: TCheckBox
          Left = 12
          Top = 0
          Width = 141
          Height = 17
          Caption = 'Changer le mot de passe'
          TabOrder = 0
          OnClick = FChgPwdClick
        end
      end
    end
    object TabApplis: TTabSheet
      Caption = 'Applications'
      ImageIndex = 1
      DesignSize = (
        275
        165)
      object HLabel1: THLabel
        Left = 8
        Top = 10
        Width = 213
        Height = 13
        Caption = 'Le mot de passe s'#39'applique aux applications :'
      end
      object BSelectAll: TToolbarButton97
        Left = 20
        Top = 137
        Width = 108
        Height = 27
        Hint = 'Valider'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Default = True
        Caption = 'Tout s'#233'lectionner'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Spacing = -1
        OnClick = BSelectAllClick
        IsControl = True
      end
      object BDeselectAll: TToolbarButton97
        Left = 147
        Top = 137
        Width = 108
        Height = 27
        Hint = 'Valider'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Default = True
        Caption = 'Tout d'#233's'#233'lectionner'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Spacing = -1
        OnClick = BDeselectAllClick
        IsControl = True
      end
      object GdApplis: THGrid
        Left = 8
        Top = 32
        Width = 258
        Height = 95
        ColCount = 2
        DefaultRowHeight = 18
        FixedCols = 0
        TabOrder = 0
        OnClick = GdApplisClick
        SortedCol = -1
        Titres.Strings = (
          'Application'
          'Prot'#233'g'#233'e')
        Couleur = False
        MultiSelect = False
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = 13224395
        ColWidths = (
          85
          64)
      end
    end
  end
end
