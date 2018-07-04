object FAssistExo: TFAssistExo
  Left = 409
  Top = 335
  BorderStyle = bsDialog
  Caption = 'Ouverture d'#39'exercice'
  ClientHeight = 155
  ClientWidth = 292
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Dock: TDock97
    Left = 0
    Top = 119
    Width = 292
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 292
      Caption = 'Actions'
      ClientAreaHeight = 32
      ClientAreaWidth = 292
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BValider: TToolbarButton97
        Left = 194
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Valider'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0119_S16G2'
        IsControl = True
      end
      object BAnnuler: TToolbarButton97
        Left = 226
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ModalResult = 7
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object Baide: TToolbarButton97
        Left = 258
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z1117_S16G1'
      end
    end
  end
  object HPanel1: THPanel
    Left = 0
    Top = 0
    Width = 292
    Height = 119
    Align = alClient
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object HLabel1: THLabel
      Left = 16
      Top = 16
      Width = 259
      Height = 13
      Caption = 'Vous devez sp'#233'cifier des dates pour cr'#233'er un exercice.'
    end
    object HLabel2: THLabel
      Left = 40
      Top = 56
      Width = 68
      Height = 13
      Caption = 'Date de d'#233'but'
    end
    object HLabel3: THLabel
      Left = 40
      Top = 83
      Width = 52
      Height = 13
      Caption = 'Date de fin'
    end
    object DATE_EXO: THCritMaskEdit
      Left = 124
      Top = 52
      Width = 71
      Height = 21
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 0
      Text = '  /  /    '
      TagDispatch = 0
      Operateur = Superieur
      OpeType = otDate
      DefaultDate = odDebAnnee
      ControlerDate = True
    end
    object DATE_EXO_: THCritMaskEdit
      Left = 124
      Top = 79
      Width = 71
      Height = 21
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 1
      Text = '  /  /    '
      TagDispatch = 0
      Operateur = Inferieur
      OpeType = otDate
      DefaultDate = odFinAnnee
      ControlerDate = True
    end
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Remise '#224' z'#233'ro de l'#39'activit'#233';Cette op'#233'ration est impossible : v' +
        'ous devez '#234'tre seul sur le r'#233'seau. ;W;O;O;O;;'
      
        '1;Remise '#224' z'#233'ro de l'#39'activit'#233';Cette op'#233'ration est impossible : l' +
        'e mot de passe est incorrect.;W;O;O;O;;'
      
        '2;Remise '#224' z'#233'ro de l'#39'activit'#233';La remise '#224' z'#233'ro de la soci'#233't'#233' a '#233 +
        't'#233' effectu'#233'e avec succ'#232's.;A;O;O;O;;'
      
        '3;Remise '#224' z'#233'ro de l'#39'activit'#233';Seul un administrateur peut effect' +
        'uer une RAZ d'#39'activit'#233' .;W;O;O;O;;'
      'Remise '#224' z'#233'ro de l'#39'activit'#233
      
        '5;Remise '#224' z'#233'ro de l'#39'activit'#233';Confirmez la remise '#224' z'#233'ro de l'#39'ac' +
        'tivit'#233' ?;Q;YNC;N;N;;'
      
        '6;Remise '#224' z'#233'ro de l'#39'activit'#233';D'#233'sirez-vous abandonner l'#39'op'#233'ratio' +
        'n ?;Q;YNC;Y;Y;;'
      'OUI'
      'ATTENTION'
      'Vous devez sp'#233'cifier des dates '
      'pour ouvrir un nouvel exercice'
      
        '11;Ouverture d'#39'exercice;Vous devez indiquer une dur'#233'e d'#39'exercice' +
        ' comprise entre 1 et 23 mois;W;O;O;O;;'
      'Exercice ouvert par RAZ soci'#233't'#233
      '13;Ouverture d'#39'exercice;L'#39'exercice 001 est ouvert.;A;O;O;O;;'
      'du'
      'au'
      
        '16;Ouverture d'#39'exercice;L'#39'exercice doit commencer le premier jou' +
        'r du mois.;W;O;O;O;;'
      
        '17;Ouverture d'#39'exercice;L'#39'exercice doit se terminer le dernier j' +
        'our du mois.;W;O;O;O;;'
      ''
      '')
    Left = 232
    Top = 47
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 249
    Top = 75
  end
end
