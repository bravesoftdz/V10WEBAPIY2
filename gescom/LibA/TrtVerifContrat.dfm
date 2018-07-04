object FTRTCONTRAT: TFTRTCONTRAT
  Left = 115
  Top = 218
  Width = 639
  Height = 408
  Caption = 'Traitement des contrats'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object HLabel9: THLabel
    Left = 102
    Top = 40
    Width = 35
    Height = 13
    Caption = 'Affaires'
  end
  object HLabel10: THLabel
    Left = 162
    Top = 40
    Width = 35
    Height = 13
    Caption = 'Affaires'
  end
  object HLabel11: THLabel
    Left = 106
    Top = 44
    Width = 35
    Height = 13
    Caption = 'Affaires'
  end
  object HLabel12: THLabel
    Left = 166
    Top = 44
    Width = 35
    Height = 13
    Caption = 'Affaires'
  end
  object PParam: TPanel
    Left = 0
    Top = 0
    Width = 631
    Height = 350
    Align = alClient
    TabOrder = 0
    object infotemps: TLabel
      Left = 6
      Top = 5
      Width = 375
      Height = 13
      Caption = 
        'Attention : Ce programme va traiter l'#39'ensemble des affaires prés' +
        'entes sur la base'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object Label1: TLabel
      Left = 6
      Top = 196
      Width = 71
      Height = 13
      Caption = 'Avancement'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
    end
    object GBRepertoire: TGroupBox
      Left = 12
      Top = 18
      Width = 389
      Height = 178
      Caption = 'Paramétrage '
      TabOrder = 0
      object HLabel1: THLabel
        Left = 5
        Top = 17
        Width = 132
        Height = 13
        Caption = 'Répertoire du compte rendu'
      end
      object HLabel2: THLabel
        Left = 5
        Top = 38
        Width = 76
        Height = 13
        Caption = 'Traitement des :'
      end
      object HLabel7: THLabel
        Left = 5
        Top = 66
        Width = 97
        Height = 13
        Caption = 'Opérations lancées :'
      end
      object HLabel8: THLabel
        Left = 5
        Top = 105
        Width = 139
        Height = 13
        Caption = 'Recalcul des écheances du  '
      end
      object HLabel13: THLabel
        Left = 266
        Top = 105
        Width = 19
        Height = 13
        Caption = 'Au  '
      end
      object LibBorne: THLabel
        Left = 5
        Top = 156
        Width = 91
        Height = 13
        Caption = 'Bornes de Numéro '
      end
      object HLabel14: THLabel
        Left = 245
        Top = 156
        Width = 6
        Height = 13
        Caption = 'à'
      end
      object RepLog: THCritMaskEdit
        Left = 176
        Top = 13
        Width = 154
        Height = 21
        TabOrder = 0
        Text = 'C:\'
        TagDispatch = 0
        DataType = 'DIRECTORY'
        ElipsisButton = True
      end
      object bAff: TCheckBox
        Left = 324
        Top = 36
        Width = 60
        Height = 17
        Caption = 'Affaires '
        Checked = True
        State = cbChecked
        TabOrder = 1
      end
      object bPro: TCheckBox
        Left = 176
        Top = 36
        Width = 95
        Height = 17
        Caption = 'Propositions'
        Checked = True
        State = cbChecked
        TabOrder = 2
      end
      object DebCal: THCritMaskEdit
        Left = 176
        Top = 101
        Width = 81
        Height = 21
        Enabled = False
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 3
        Text = '01/07/2001'
        TagDispatch = 0
        OpeType = otDate
        ControlerDate = True
      end
      object FinCal: THCritMaskEdit
        Left = 298
        Top = 101
        Width = 81
        Height = 21
        Enabled = False
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 4
        Text = '31/12/2001'
        TagDispatch = 0
        OpeType = otDate
        ControlerDate = True
      end
      object LibUneAff: TCheckBox
        Left = 5
        Top = 131
        Width = 165
        Height = 17
        Caption = 'Traitement d'#39'une seule affaire'
        TabOrder = 5
        OnClick = LibUneAffClick
      end
      object UneAff: THCritMaskEdit
        Left = 176
        Top = 127
        Width = 109
        Height = 21
        TabOrder = 6
        Visible = False
        TagDispatch = 0
        DataType = 'DIRECTORY'
        ElipsisButton = True
        OnElipsisClick = UneAffElipsisClick
      end
      object AffDeb: THCritMaskEdit
        Tag = 1
        Left = 115
        Top = 152
        Width = 125
        Height = 21
        TabOrder = 7
        TagDispatch = 0
        DataType = 'DIRECTORY'
        ElipsisButton = True
        OnElipsisClick = UneAffElipsisClick
      end
      object AffFin: THCritMaskEdit
        Tag = 2
        Left = 258
        Top = 152
        Width = 125
        Height = 21
        TabOrder = 8
        TagDispatch = 0
        DataType = 'DIRECTORY'
        ElipsisButton = True
        OnElipsisClick = UneAffElipsisClick
      end
    end
    object MemoTrace: TMemo
      Left = 1
      Top = 211
      Width = 629
      Height = 138
      Align = alBottom
      Lines.Strings = (
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        ''
        '')
      ReadOnly = True
      ScrollBars = ssBoth
      TabOrder = 1
    end
    object GroupBox1: TGroupBox
      Left = 412
      Top = 18
      Width = 212
      Height = 178
      Caption = 'Suivi du traitement '
      TabOrder = 2
      object HLabel4: THLabel
        Left = 18
        Top = 40
        Width = 35
        Height = 13
        Caption = 'Affaires'
      end
      object HLabel3: THLabel
        Left = 18
        Top = 64
        Width = 60
        Height = 13
        Caption = 'Propositions '
      end
      object HLabel5: THLabel
        Left = 97
        Top = 15
        Width = 40
        Height = 13
        Caption = 'Nb total '
      end
      object HLabel6: THLabel
        Left = 157
        Top = 15
        Width = 51
        Height = 13
        Caption = 'Déja traité '
      end
      object NbTotAff: THLabel
        Left = 103
        Top = 40
        Width = 8
        Height = 13
        Caption = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object NbTrtAff: THLabel
        Left = 160
        Top = 40
        Width = 8
        Height = 13
        Caption = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object NbTotPro: THLabel
        Left = 103
        Top = 64
        Width = 8
        Height = 13
        Caption = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object NbTrtPro: THLabel
        Left = 160
        Top = 64
        Width = 8
        Height = 13
        Caption = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
    end
  end
  object PCompteRendu: TPanel
    Left = 0
    Top = 350
    Width = 631
    Height = 31
    Align = alBottom
    TabOrder = 1
    object bFin: TButton
      Left = 547
      Top = 4
      Width = 75
      Height = 23
      Caption = '&Traitement'
      TabOrder = 0
      OnClick = bFinClick
    end
    object bAnnuler: TButton
      Left = 465
      Top = 4
      Width = 75
      Height = 23
      Cancel = True
      Caption = 'Annuler'
      TabOrder = 1
      OnClick = bAnnulerClick
    end
    object BStop: TBitBtn
      Left = 246
      Top = 3
      Width = 81
      Height = 25
      TabOrder = 2
      OnClick = BStopClick
      Glyph.Data = {
        76050000424D7605000000000000360000002800000015000000150000000100
        1800000000004005000000000000000000000000000000000000C6C3C6C6C3C6
        C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3
        C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C600C6C3C6C6C3C6
        C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3
        C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C600C6C3C6C6C3C6
        C6C3C6C6C3C6C6C3C6C6C3C60000840000840000840000840000840000840000
        84000084000084C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C600C6C3C6C6C3C6
        C6C3C6C6C3C6C6C3C60000840000FF0000FF0000FF0000FF0000FF0000FF0000
        FF0000FF0000FF000084C6C3C6C6C3C6C6C3C6C6C3C6C6C3C600C6C3C6C6C3C6
        C6C3C6C6C3C60000840000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF000084C6C3C6C6C3C6C6C3C6C6C3C600C6C3C6C6C3C6
        C6C3C60000840000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF0000FF000084C6C3C6C6C3C6C6C3C600C6C3C6C6C3C6
        0000840000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF0000FF0000FF000084C6C3C6C6C3C600C6C3C6000084
        0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF0000FF0000FF0000FF000084C6C3C600C6C3C6000084
        0000FFFFFFFFFFFFFFFFFFFF0000FF0000FFFFFFFF0000FF0000FFFFFFFFFFFF
        FFFFFFFF0000FFFFFFFF0000FF0000FF0000FF000084C6C3C600C6C3C6000084
        0000FF0000FF0000FFFFFFFF0000FF0000FFFFFFFF0000FF0000FFFFFFFF0000
        FFFFFFFF0000FFFFFFFF0000FF0000FF0000FF000084C6C3C600C6C3C6000084
        0000FFFFFFFFFFFFFFFFFFFF0000FF0000FFFFFFFF0000FF0000FFFFFFFF0000
        FFFFFFFF0000FFFFFFFFFFFFFFFFFFFF0000FF000084C6C3C600C6C3C6000084
        0000FFFFFFFF0000FF0000FF0000FF0000FFFFFFFF0000FF0000FFFFFFFF0000
        FFFFFFFF0000FFFFFFFF0000FFFFFFFF0000FF000084C6C3C600C6C3C6000084
        0000FFFFFFFFFFFFFFFFFFFF0000FFFFFFFFFFFFFFFFFFFF0000FFFFFFFFFFFF
        FFFFFFFF0000FFFFFFFFFFFFFFFFFFFF0000FF000084C6C3C600C6C3C6000084
        0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF0000FF0000FF0000FF000084C6C3C600C6C3C6000084
        0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF0000FF0000FF0000FF000084C6C3C600C6C3C6C6C3C6
        0000840000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF0000FF0000FF000084C6C3C6C6C3C600C6C3C6C6C3C6
        C6C3C60000840000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF0000FF000084C6C3C6C6C3C6C6C3C600C6C3C6C6C3C6
        C6C3C6C6C3C60000840000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
        FF0000FF0000FF0000FF000084C6C3C6C6C3C6C6C3C6C6C3C600C6C3C6C6C3C6
        C6C3C6C6C3C6C6C3C60000840000FF0000FF0000FF0000FF0000FF0000FF0000
        FF0000FF0000FF000084C6C3C6C6C3C6C6C3C6C6C3C6C6C3C600C6C3C6C6C3C6
        C6C3C6C6C3C6C6C3C6C6C3C60000840000840000840000840000840000840000
        84000084000084C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C600C6C3C6C6C3C6
        C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3
        C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C600}
    end
  end
  object bCalcul: TCheckBox
    Left = 188
    Top = 79
    Width = 164
    Height = 17
    Caption = 'Recalcul des pièces / totaux'
    Checked = True
    State = cbChecked
    TabOrder = 2
  end
  object bGener: TCheckBox
    Left = 188
    Top = 98
    Width = 160
    Height = 17
    Caption = 'Génération des écheances '
    Checked = True
    State = cbChecked
    TabOrder = 3
  end
end
