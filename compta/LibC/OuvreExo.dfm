object OuvExo: TOuvExo
  Left = 340
  Top = 193
  HelpContext = 7718000
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Ouverture d'#39'exercice '
  ClientHeight = 154
  ClientWidth = 322
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 322
    Height = 116
    Align = alClient
    TabOrder = 0
    object HLabel1: THLabel
      Left = 16
      Top = 13
      Width = 85
      Height = 13
      Caption = 'Exercice en cours'
    end
    object HLabel2: THLabel
      Left = 16
      Top = 67
      Width = 78
      Height = 13
      Caption = 'Exercice suivant'
    end
    object HLabel3: THLabel
      Left = 141
      Top = 13
      Width = 29
      Height = 13
      Caption = '&D'#233'but'
      FocusControl = FDateEnCDeb
    end
    object HLabel4: THLabel
      Left = 141
      Top = 37
      Width = 14
      Height = 13
      Caption = '&Fin'
      FocusControl = FDateEnCFin
    end
    object HLabel5: THLabel
      Left = 141
      Top = 67
      Width = 29
      Height = 13
      Caption = 'D'#233'&but'
      FocusControl = FDateSuivDeb
    end
    object HLabel6: THLabel
      Left = 141
      Top = 91
      Width = 14
      Height = 13
      Caption = 'Fi&n'
      FocusControl = FDateSuivFin
    end
    object FDateEnCDeb: TMaskEdit
      Left = 201
      Top = 10
      Width = 100
      Height = 21
      Ctl3D = True
      Enabled = False
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 0
      Text = '01/01/1900'
    end
    object FDateEnCFin: TMaskEdit
      Left = 201
      Top = 34
      Width = 100
      Height = 21
      Ctl3D = True
      Enabled = False
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 1
      Text = '01/01/1900'
    end
    object FDateSuivDeb: TMaskEdit
      Left = 201
      Top = 64
      Width = 100
      Height = 21
      Ctl3D = True
      Enabled = False
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 2
      Text = '01/01/1900'
    end
    object FDateSuivFin: TMaskEdit
      Left = 201
      Top = 88
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 3
      Text = '01/01/1900'
      OnChange = FDateSuivFinChange
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 116
    Width = 322
    Height = 38
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 34
      ClientWidth = 322
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 34
      ClientAreaWidth = 322
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BAide: TToolbarButton97
        Left = 287
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BValider: TToolbarButton97
        Left = 225
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Ouvrir l'#39'exercice'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Ouvrir l'#39'exercice'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Glyph.Data = {
          BE060000424DBE06000000000000360400002800000024000000120000000100
          0800000000008802000000000000000000000001000000010000000000000000
          80000080000000808000800000008000800080800000C0C0C000C0DCC000F0C8
          A400000000000000000000000000000000000000000000000000000000000000
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
          0303030303030303030303030303030303030303030303030303030303030303
          03030303030303030303030303030303030303030303FF030303030303030303
          03030303030303040403030303030303030303030303030303F8F8FF03030303
          03030303030303030303040202040303030303030303030303030303F80303F8
          FF030303030303030303030303040202020204030303030303030303030303F8
          03030303F8FF0303030303030303030304020202020202040303030303030303
          0303F8030303030303F8FF030303030303030304020202FA0202020204030303
          0303030303F8FF0303F8FF030303F8FF03030303030303020202FA03FA020202
          040303030303030303F8FF03F803F8FF0303F8FF03030303030303FA02FA0303
          03FA0202020403030303030303F8FFF8030303F8FF0303F8FF03030303030303
          FA0303030303FA0202020403030303030303F80303030303F8FF0303F8FF0303
          0303030303030303030303FA0202020403030303030303030303030303F8FF03
          03F8FF03030303030303030303030303FA020202040303030303030303030303
          0303F8FF0303F8FF03030303030303030303030303FA02020204030303030303
          03030303030303F8FF0303F8FF03030303030303030303030303FA0202020403
          030303030303030303030303F8FF0303F8FF03030303030303030303030303FA
          0202040303030303030303030303030303F8FF03F8FF03030303030303030303
          03030303FA0202030303030303030303030303030303F8FFF803030303030303
          030303030303030303FA0303030303030303030303030303030303F803030303
          0303030303030303030303030303030303030303030303030303030303030303
          0303}
        GlyphMask.Data = {00000000}
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 256
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z1770_S16G1'
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
      
        '0;Ouverture d'#39'exercice;L'#39'exercice suivant est d'#233'j'#224' ouvert.;E;O;O' +
        ';O'
      
        '1;Ouverture d'#39'exercice;Confirmez-vous l'#39'ouverture de l'#39'exercice ' +
        '?;Q;YN;Y;N'
      
        '2;Ouverture d'#39'exercice;La date de fin d'#39'exercice que vous avez r' +
        'enseign'#233'e est incoh'#233'rente.;W;O;O;O'
      
        '3;Ouverture d'#39'exercice;La date de fin d'#39'un exercice doit '#234'tre un' +
        'e date de fin de mois.;W;O;O;O'
      
        '4;Ouverture d'#39'exercice;L'#39'exercice en cours n'#39'est pas ouvert ou n' +
        #39'existe pas. Vous devez le param'#233'trer par l'#39'assistant de param'#233't' +
        'rage de soci'#233't'#233'.;W;O;O;O'
      
        '5;Ouverture d'#39'exercice;Le traitement s'#39'est correctement effectu'#233 +
        ';E;O;O;O')
    Left = 86
    Top = 5
  end
  object Morceaux: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Exo de '
      ' au '
      'Exercice de ')
    Left = 83
    Top = 59
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 28
    Top = 92
  end
end
