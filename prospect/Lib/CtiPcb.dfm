object Form1: TForm1
  Left = 510
  Top = 157
  Width = 431
  Height = 454
  HelpContext = 111000600
  Caption = 'Couplage T'#233'l'#233'phonie Informatique'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object LOG: TMemo
    Left = 20
    Top = 252
    Width = 382
    Height = 119
    Lines.Strings = (
      'LOG')
    ScrollBars = ssVertical
    TabOrder = 0
  end
  object APPELENTRANT: TGroupBox
    Left = 20
    Top = 19
    Width = 271
    Height = 109
    Caption = 'Appel Entrant '
    TabOrder = 1
    object LNOTELAPPELANT: TLabel
      Left = 5
      Top = 17
      Width = 114
      Height = 13
      Caption = 'T'#233'l'#233'phone de l'#39'appelant'
    end
    object LTIERSAPPELANT: TLabel
      Left = 5
      Top = 41
      Width = 47
      Height = 13
      Caption = 'Code tiers'
    end
    object LNOMAPPELANT: TLabel
      Left = 5
      Top = 85
      Width = 76
      Height = 13
      Caption = 'Nom du contact'
    end
    object LLIBELLEAPPELANT: TLabel
      Left = 5
      Top = 63
      Width = 69
      Height = 13
      Caption = 'Raison sociale'
    end
    object NOTELAPPELANT: TEdit
      Left = 127
      Top = 14
      Width = 136
      Height = 21
      TabOrder = 0
    end
    object TIERSAPPELANT: TEdit
      Left = 127
      Top = 37
      Width = 136
      Height = 21
      ReadOnly = True
      TabOrder = 1
    end
    object LIBELLEAPPELANT: TEdit
      Left = 127
      Top = 60
      Width = 136
      Height = 21
      ReadOnly = True
      TabOrder = 2
    end
    object NATUREAPPELANT: TEdit
      Left = 43
      Top = 37
      Width = 23
      Height = 21
      Color = clYellow
      TabOrder = 3
      Visible = False
    end
    object AUXIAPPELANT: TEdit
      Left = 75
      Top = 37
      Width = 23
      Height = 21
      Color = clYellow
      TabOrder = 4
      Visible = False
    end
    object NUMCONTACTAPPELANT: TEdit
      Left = 95
      Top = 80
      Width = 23
      Height = 21
      Color = clYellow
      TabOrder = 5
      Visible = False
    end
  end
  object ACTIONENTRANT: TGroupBox
    Left = 296
    Top = 19
    Width = 106
    Height = 224
    Caption = 'Action'
    TabOrder = 2
    object DECROCHER: THSpeedButton
      Left = 5
      Top = 14
      Width = 95
      Height = 23
      AllowAllUp = True
      Caption = 'D'#233'crocher  '
      Margin = 3
      NumGlyphs = 3
      OnClick = BoutonDecrocherClick
      GlobalIndexImage = 'Z0577_S16G3'
    end
    object RACCROCHER: THSpeedButton
      Left = 5
      Top = 104
      Width = 95
      Height = 23
      AllowAllUp = True
      Caption = 'Raccrocher'
      Margin = 3
      NumGlyphs = 3
      OnClick = BoutonRaccrocherClick
      GlobalIndexImage = 'Z0577_S16G3'
    end
    object BATTENTE: THSpeedButton
      Left = 5
      Top = 74
      Width = 95
      Height = 23
      AllowAllUp = True
      GroupIndex = 2
      Caption = 'Attente'
      Margin = 3
      NumGlyphs = 2
      OnClick = BoutonAttenteClick
      GlobalIndexImage = 'Z0766_S16G2'
    end
    object ZOOMAPPELANT: THBitBtn
      Left = 5
      Top = 44
      Width = 95
      Height = 23
      Caption = 'D'#233'tail Fiche '
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = ZoomClick
      Margin = 0
      GlobalIndexImage = 'Z0061_S16G1'
    end
    object APPEL: THBitBtn
      Left = 5
      Top = 134
      Width = 95
      Height = 23
      Caption = 'Appel        '
      TabOrder = 1
      OnClick = BoutonAppelClick
      Margin = 3
      NumGlyphs = 3
      GlobalIndexImage = 'Z0577_S16G3'
    end
    object ZOOMAPPELE: THBitBtn
      Left = 5
      Top = 164
      Width = 95
      Height = 23
      Caption = 'D'#233'tail Fiche '
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = ZoomClick
      Margin = 0
      GlobalIndexImage = 'Z0061_S16G1'
    end
    object SERIEAPPELS: THBitBtn
      Left = 5
      Top = 194
      Width = 95
      Height = 23
      Caption = 'S'#233'rie d'#39'appels'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      OnClick = SerieAppelsClick
      Margin = 1
      GlobalIndexImage = 'Z0548_S16G1'
    end
  end
  object APPELSORTANT: TGroupBox
    Left = 20
    Top = 131
    Width = 271
    Height = 112
    Caption = 'Appel Sortant'
    TabOrder = 4
    object LNOTELAPPELE: TLabel
      Left = 5
      Top = 18
      Width = 108
      Height = 13
      Caption = 'T'#233'l'#233'phone de l'#39'appel'#233' '
    end
    object LTIERSAPPELE: TLabel
      Left = 5
      Top = 41
      Width = 47
      Height = 13
      Caption = 'Code tiers'
    end
    object LLIBELLEAPPELE: TLabel
      Left = 5
      Top = 64
      Width = 69
      Height = 13
      Caption = 'Raison sociale'
    end
    object LNOMAPPELE: TLabel
      Left = 5
      Top = 86
      Width = 76
      Height = 13
      Caption = 'Nom du contact'
    end
    object NOTELAPPELE: TEdit
      Left = 127
      Top = 14
      Width = 136
      Height = 21
      TabOrder = 0
    end
    object LIBELLEAPPELE: TEdit
      Left = 127
      Top = 60
      Width = 136
      Height = 21
      ReadOnly = True
      TabOrder = 1
    end
    object TIERSAPPELE: THCritMaskEdit
      Left = 127
      Top = 37
      Width = 136
      Height = 21
      ReadOnly = True
      TabOrder = 2
      TagDispatch = 0
      DataType = 'RTTIERSPRO'
      ElipsisButton = True
      OnElipsisClick = AppelListeTiers
    end
    object NATUREAPPELE: TEdit
      Left = 55
      Top = 38
      Width = 26
      Height = 21
      Color = clYellow
      TabOrder = 3
      Visible = False
    end
    object AUXIAPPELE: TEdit
      Left = 85
      Top = 38
      Width = 26
      Height = 21
      Color = clYellow
      TabOrder = 4
      Visible = False
    end
    object NOMAPPELE: TEdit
      Left = 127
      Top = 83
      Width = 136
      Height = 21
      ReadOnly = True
      TabOrder = 5
    end
    object NUMCONTACTAPPELE: TEdit
      Left = 89
      Top = 83
      Width = 26
      Height = 21
      Color = clYellow
      TabOrder = 6
      Visible = False
    end
  end
  object Dock972: TDock97
    Left = 0
    Top = 385
    Width = 423
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object PBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 423
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 31
      ClientAreaWidth = 423
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BImprimer: TToolbarButton97
        Left = 320
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = ImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BFerme: TToolbarButton97
        Left = 352
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ParentShowHint = False
        ShowHint = True
        OnClick = FermerClick
        GlobalIndexImage = 'Z1770_S16G1'
      end
      object HelpBtn: TToolbarButton97
        Left = 384
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Glyph.Data = {
          BE060000424DBE06000000000000360400002800000024000000120000000100
          0800000000008802000000000000000000000001000000010000000000000000
          80000080000000808000800000008000800080800000C0C0C000C0DCC000F0CA
          A600000000000000000000000000000000000000000000000000000000000000
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
          03030606030303030303030303030303030303FFFF0303030303030303030303
          0303030303060404060303030303030303030303030303F8F8FF030303030303
          030303030303030303FE06060403030303030303030303030303F8FF03F8FF03
          0303030303030303030303030303FE060603030303030303030303030303F8FF
          FFF8FF0303030303030303030303030303030303030303030303030303030303
          030303F8F8030303030303030303030303030303030304040603030303030303
          0303030303030303FFFF03030303030303030303030303030306060604030303
          0303030303030303030303F8F8F8FF0303030303030303030303030303FE0606
          0403030303030303030303030303F8FF03F8FF03030303030303030303030303
          03FE06060604030303030303030303030303F8FF03F8FF030303030303030303
          030303030303FE060606040303030303030303030303F8FF0303F8FF03030303
          0303030303030303030303FE060606040303030303030303030303F8FF0303F8
          FF030303030303030303030404030303FE060606040303030303030303FF0303
          F8FF0303F8FF030303030303030306060604030303FE06060403030303030303
          F8F8FF0303F8FF0303F8FF03030303030303FE06060604040406060604030303
          030303F8FF03F8FFFFFFF80303F8FF0303030303030303FE0606060606060606
          06030303030303F8FF0303F8F8F8030303F8FF030303030303030303FEFE0606
          060606060303030303030303F8FFFF030303030303F803030303030303030303
          0303FEFEFEFEFE03030303030303030303F8F8FFFFFFFFFFF803030303030303
          0303030303030303030303030303030303030303030303F8F8F8F8F803030303
          0303}
        GlyphMask.Data = {00000000}
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        IsControl = True
      end
      object FEUROUGE: TImage
        Left = 3
        Top = 2
        Width = 24
        Height = 24
        Picture.Data = {
          07544269746D617096010000424D960100000000000076000000280000001800
          0000180000000100040000000000200100000000000000000000100000000000
          0000000000000000800000800000008080008000000080008000808000008080
          8000C0C0C0000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
          FF00888888880000000088888888888888800060040008888888888888806600
          0044088888888888880062777274008888888888880462772720408888888888
          00606277727404008888888BB660602222060446688888800004600000004000
          0888888888806604404608888888888888806601374408888888888888806033
          7170088888888888880060313734008888888888006060731314040088888880
          40046007373040040888888BB660600000060446688888888880660440460888
          8888888888806060040008888888888888806609994408888888888888006099
          999400888888888888046099999040888888888800606099999404008888888B
          B660600999060446688888800004600000004000088888888880B60660660888
          8888}
      end
      object FEUVERT: TImage
        Left = 3
        Top = 2
        Width = 24
        Height = 24
        Picture.Data = {
          07544269746D617096010000424D960100000000000076000000280000001800
          0000180000000100040000000000200100000000000000000000100000001000
          0000000000000000800000800000008080008000000080008000808000008080
          8000C0C0C0000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
          FF00888888880000000088888888888888800060040008888888888888806600
          0044088888888888880062AAAB24008888888888880462AABAB0408888888888
          006062AAABA404008888888BB660602222060446688888800004600000004000
          0888888888806604404608888888888888806600004408888888888888806037
          7300088888888888880060033734008888888888006060033734040088888880
          40046000037040040888888BB660600000060446688888888880660440460888
          8888888888806060040008888888888888806600004408888888888888006001
          171400888888888888046010717040888888888800606001171404008888888B
          B660600000060446688888800004600000004000088888888880B60660660888
          8888}
      end
      object INDISPONIBLE: TLabel
        Left = 33
        Top = 7
        Width = 56
        Height = 13
        Caption = 'Indisponible'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object DISPONIBLE: TLabel
        Left = 31
        Top = 7
        Width = 49
        Height = 13
        Caption = 'Disponible'
      end
      object BSTOP: TToolbarButton97
        Left = 256
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Arr'#234't/Reprise du couplage'
        AllowAllUp = True
        GroupIndex = 2
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = StopClick
        GlobalIndexImage = 'Z0107_S16G1'
      end
      object BCONTACT: TToolbarButton97
        Left = 224
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Recherche d'#39'un contact'
        AllowAllUp = True
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = ContactClick
        GlobalIndexImage = 'Z0116_S16G1'
      end
      object BTIERS: TToolbarButton97
        Left = 192
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Recherche d'#39'un tiers'
        AllowAllUp = True
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = AppelListeTiers
        GlobalIndexImage = 'Z0109_S16G1'
      end
      object EFFACER: TToolbarButton97
        Left = 288
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Effacer'
        AllowAllUp = True
        GroupIndex = 2
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = EffacerClick
        GlobalIndexImage = 'Z0204_S16G1'
      end
      object BACTIONS: TToolbarButton97
        Left = 161
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Actions'
        AllowAllUp = True
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = ActionsClick
        GlobalIndexImage = 'Z0200_S16G1'
      end
    end
  end
  object RichEdit1: TRichEdit
    Left = 144
    Top = 256
    Width = 201
    Height = 89
    Lines.Strings = (
      'RichEdit1')
    TabOrder = 6
    Visible = False
  end
  object NOMAPPELANT: TEdit
    Left = 147
    Top = 102
    Width = 136
    Height = 21
    ReadOnly = True
    TabOrder = 5
  end
end
