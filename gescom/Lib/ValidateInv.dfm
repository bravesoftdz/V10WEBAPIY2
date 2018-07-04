object ValidateForm: TValidateForm
  Left = 294
  Top = 204
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = 'Validation de listes d'#39'inventaire'
  ClientHeight = 221
  ClientWidth = 473
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
  object Label1: TLabel
    Left = 4
    Top = 12
    Width = 465
    Height = 13
    Alignment = taCenter
    AutoSize = False
    Caption = 'Label1'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object ProgressBar: TProgressBar
    Left = 4
    Top = 200
    Width = 381
    Height = 13
    Step = 1
    TabOrder = 0
  end
  object bCancel: TButton
    Left = 392
    Top = 196
    Width = 75
    Height = 23
    Cancel = True
    Caption = 'bCancel'
    TabOrder = 1
    OnClick = bCancelClick
  end
  object GridSelection: THGrid
    Left = 4
    Top = 32
    Width = 461
    Height = 157
    ColCount = 4
    DefaultRowHeight = 17
    FixedCols = 0
    RowCount = 2
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goColSizing, goRowSelect]
    ParentFont = False
    TabOrder = 2
    SortedCol = -1
    Titres.Strings = (
      ''
      'Code'
      'Libell'#233
      'Etat')
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = False
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = clSilver
    ColWidths = (
      21
      98
      142
      186)
  end
  object HMsg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Validation en cours...'
      'Validation effectu'#233'e'
      '&Annuler'
      '&Fermer'
      'Ordre des prix retenu :'
      'Abandon'
      'Validation des listes interrompue par l'#39'utilisateur'
      '-- Traitement interrompu par l'#39'utilisateur --'
      'En cours...'
      'Non valid'#233'e'
      'Valid'#233'e'
      'Saisie incompl'#232'te'
      '(une ou plusieurs lignes non inventori'#233'es)'
      'Saisie en cours'
      '(en cours de traitement par un autre utilisateur)'
      'Echec'
      'Validation interrompue'
      'Tiers Ecarts inventaire inconnu')
    Left = 4
    Top = 4
  end
  object iml_progress: THImageList
    GlobalIndexImages.Strings = (
      'Z0742_S16G1'
      'Z0050_S16G1'
      'Z0217_S16G1'
      'Z1415_S16G1')
    Left = 32
    Top = 4
  end
end
