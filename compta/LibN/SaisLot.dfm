object FSaisLot: TFSaisLot
  Left = 301
  Top = 202
  Width = 269
  Height = 156
  BorderIcons = []
  Caption = 'Saisie d'#39'un lot'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 261
    Height = 93
    Align = alClient
    TabOrder = 0
    object HLabel3: THLabel
      Left = 34
      Top = 39
      Width = 36
      Height = 13
      Caption = '&Nom lot'
      FocusControl = E_NOMLOT
    end
    object E_NOMLOT: THCritMaskEdit
      Left = 113
      Top = 35
      Width = 131
      Height = 21
      CharCase = ecUpperCase
      Ctl3D = True
      MaxLength = 14
      ParentCtl3D = False
      TabOrder = 0
      TagDispatch = 0
      DataType = 'CPNOMLOT'
      Operateur = Superieur
      ElipsisButton = True
      OnElipsisClick = E_NOMLOTElipsisClick
    end
    object PrefE_NOMLOT: TEdit
      Left = 80
      Top = 35
      Width = 33
      Height = 21
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
      Text = 'CLG'
    end
  end
  object Dock: TDock97
    Left = 0
    Top = 93
    Width = 261
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object PanelBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 261
      Caption = 'Barre outils multicrit'#232're'
      ClientAreaHeight = 32
      ClientAreaWidth = 261
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BOuvrir: TToolbarButton97
        Left = 192
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Ouvrir'
        DisplayMode = dmGlyphOnly
        Caption = 'Ouvrir'
        Flat = False
        Layout = blGlyphTop
        ModalResult = 1
        NumGlyphs = 2
        OnClick = BOuvrirClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BAnnuler: TToolbarButton97
        Left = 224
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        Layout = blGlyphTop
        ModalResult = 2
        GlobalIndexImage = 'Z0021_S16G1'
      end
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 14
    Top = 102
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?Caption?;Vous n'#39'avez pas renseign'#233' de code lot.;W;O;O;O;'
      '1;?Caption?;Ce lot existe d'#233'j'#224'. Confirmez-vous ?;E;YN;N;N;'
      '2;?Caption?;Confirmez-vous le traitement?;E;YN;N;N;')
    Left = 52
    Top = 100
  end
end
