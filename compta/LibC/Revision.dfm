object Revis: TRevis
  Left = 365
  Top = 218
  HelpContext = 7682000
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Mode r'#233'vision'
  ClientHeight = 136
  ClientWidth = 268
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
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
    Top = 0
    Width = 268
    Height = 96
    Align = alClient
    TabOrder = 0
    object TDateRev: THLabel
      Left = 21
      Top = 72
      Width = 77
      Height = 13
      Caption = 'Date de r'#233'vision'
      FocusControl = FDateRev
    end
    object FDateRev: TMaskEdit
      Left = 141
      Top = 69
      Width = 99
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 1
      Text = '01/01/1900'
      OnEnter = FDateRevEnter
    end
    object EtatRevision: TRadioGroup
      Left = 20
      Top = 12
      Width = 221
      Height = 45
      Caption = ' Etat '
      Columns = 2
      ItemIndex = 1
      Items.Strings = (
        'Activ'#233
        'D'#233'sactiv'#233)
      TabOrder = 0
      OnClick = EtatRevisionClick
    end
  end
  object Dock: TDock97
    Left = 0
    Top = 96
    Width = 268
    Height = 40
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 36
      ClientWidth = 268
      Caption = 'Barre outils multicrit'#232're'
      ClientAreaHeight = 36
      ClientAreaWidth = 268
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        268
        36)
      object BValider: TToolbarButton97
        Left = 172
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Ouvrir'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Ouvrir'
        Enabled = False
        Flat = False
        Layout = blGlyphTop
        NumGlyphs = 2
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BFerme: TToolbarButton97
        Left = 204
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Anchors = [akTop, akRight]
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        Layout = blGlyphTop
        ModalResult = 2
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 236
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        Layout = blGlyphTop
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
    end
  end
  object Confirmation: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Mode r'#233'vision;Confirmez-vous l'#39'activation du mode r'#233'vision?;Q;' +
        'YN;Y;N'
      
        '1;Mode r'#233'vision;Confirmez-vous la d'#233'sactivation du mode r'#233'vision' +
        '?;Q;YN;Y;N'
      '')
    Left = 50
    Top = 53
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 108
    Top = 52
  end
end
