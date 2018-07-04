object VariableDlg: TVariableDlg
  Left = 335
  Top = 175
  BorderStyle = bsDialog
  Caption = 'Variable'
  ClientHeight = 281
  ClientWidth = 330
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
  object Label1: TLabel
    Left = 8
    Top = 12
    Width = 22
    Height = 13
    Caption = '&Nom'
    FocusControl = edtValue
  end
  object Label2: TLabel
    Left = 8
    Top = 36
    Width = 46
    Height = 13
    Caption = '&Contenu :'
    FocusControl = edtName
  end
  object Label3: TLabel
    Left = 8
    Top = 124
    Width = 24
    Height = 13
    Caption = 'Type'
    Visible = False
  end
  object Label4: TLabel
    Left = 8
    Top = 60
    Width = 31
    Height = 13
    Caption = '&Intitul'#233
    FocusControl = edtIntitule
  end
  object Label5: TLabel
    Left = 8
    Top = 152
    Width = 87
    Height = 13
    Caption = 'List&e des '#233'l'#233'ments'
    FocusControl = ListBox1
    Visible = False
  end
  object edtValue: TEdit
    Left = 56
    Top = 8
    Width = 249
    Height = 21
    TabOrder = 0
  end
  object edtName: TEdit
    Left = 56
    Top = 32
    Width = 249
    Height = 21
    TabOrder = 1
  end
  object CheckBox1: TCheckBox
    Left = 8
    Top = 88
    Width = 97
    Height = 17
    Alignment = taLeftJustify
    Caption = 'Demandable'
    TabOrder = 3
  end
  object Button1: TButton
    Left = 184
    Top = 88
    Width = 123
    Height = 25
    Caption = 'Plus de d'#233'tails >>>'
    TabOrder = 4
    OnClick = Button1Click
  end
  object ComboBox1: TComboBox
    Left = 56
    Top = 120
    Width = 145
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 5
    Visible = False
    Items.Strings = (
      'Chaine'
      'Num'#233'rique'
      'Liste'
      'Case '#224' cocher'
      'Fichier')
  end
  object edtIntitule: TEdit
    Left = 56
    Top = 56
    Width = 249
    Height = 21
    TabOrder = 2
  end
  object ListBox1: TListBox
    Left = 8
    Top = 168
    Width = 297
    Height = 73
    ItemHeight = 13
    TabOrder = 6
    Visible = False
  end
  object Dock971: TDock97
    Left = 0
    Top = 246
    Width = 330
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object PBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 330
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 31
      ClientAreaWidth = 330
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BValider: TToolbarButton97
        Left = 262
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        AllowAllUp = True
        Default = True
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ModalResult = 1
        NumGlyphs = 2
        ParentFont = False
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 294
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        AllowAllUp = True
        Cancel = True
        Flat = False
        ModalResult = 2
        GlobalIndexImage = 'Z0021_S16G1'
      end
    end
  end
end
