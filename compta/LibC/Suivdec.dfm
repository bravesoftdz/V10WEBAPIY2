object FSuivDec: TFSuivDec
  Left = 8
  Top = -13
  HelpContext = 1500000
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Circuit de validation des d'#233'caissements'
  ClientHeight = 208
  ClientWidth = 384
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PFen: TPanel
    Left = 0
    Top = 0
    Width = 384
    Height = 172
    Align = alClient
    TabOrder = 1
    object Label1: TLabel
      Left = 44
      Top = 4
      Width = 145
      Height = 26
      Caption = 'Etapes du circuit de validation des d'#233'caissements'
      WordWrap = True
    end
    object Label2: TLabel
      Left = 312
      Top = 4
      Width = 67
      Height = 39
      Caption = 'Affectation pr'#233'visionnelle des banques'
      WordWrap = True
    end
    object e1: TEdit
      Left = 44
      Top = 46
      Width = 263
      Height = 21
      TabOrder = 0
      OnChange = eChange
    end
    object e2: TEdit
      Left = 44
      Top = 71
      Width = 263
      Height = 21
      TabOrder = 1
      OnChange = eChange
    end
    object e3: TEdit
      Left = 44
      Top = 96
      Width = 263
      Height = 21
      TabOrder = 2
      OnChange = eChange
    end
    object e4: TEdit
      Left = 44
      Top = 121
      Width = 263
      Height = 21
      TabOrder = 3
      OnChange = eChange
    end
    object e5: TEdit
      Left = 44
      Top = 146
      Width = 263
      Height = 21
      TabOrder = 4
      OnChange = eChange
    end
    object b1: TCheckBox
      Left = 333
      Top = 46
      Width = 15
      Height = 17
      TabOrder = 5
    end
    object b2: TCheckBox
      Left = 333
      Top = 71
      Width = 15
      Height = 17
      TabOrder = 6
    end
    object b3: TCheckBox
      Left = 333
      Top = 97
      Width = 15
      Height = 17
      TabOrder = 7
    end
    object b4: TCheckBox
      Left = 333
      Top = 122
      Width = 15
      Height = 17
      TabOrder = 8
    end
    object b5: TCheckBox
      Left = 333
      Top = 147
      Width = 15
      Height = 17
      TabOrder = 9
    end
    object c1: TCheckBox
      Left = 7
      Top = 47
      Width = 30
      Height = 17
      Alignment = taLeftJustify
      Caption = '&1'
      TabOrder = 10
      OnClick = cClick
    end
    object c2: TCheckBox
      Left = 7
      Top = 72
      Width = 30
      Height = 17
      Alignment = taLeftJustify
      Caption = '&2'
      TabOrder = 11
      OnClick = cClick
    end
    object c3: TCheckBox
      Left = 7
      Top = 97
      Width = 30
      Height = 17
      Alignment = taLeftJustify
      Caption = '&3'
      TabOrder = 12
      OnClick = cClick
    end
    object c4: TCheckBox
      Left = 7
      Top = 122
      Width = 30
      Height = 17
      Alignment = taLeftJustify
      Caption = '&4'
      TabOrder = 13
      OnClick = cClick
    end
    object c5: TCheckBox
      Left = 7
      Top = 147
      Width = 30
      Height = 17
      Alignment = taLeftJustify
      Caption = '&5'
      TabOrder = 14
      OnClick = cClick
    end
  end
  object Dock: TDock97
    Left = 0
    Top = 172
    Width = 384
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 384
      Caption = 'Actions'
      ClientAreaHeight = 32
      ClientAreaWidth = 384
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BValider: TToolbarButton97
        Left = 286
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
      object BAnnuler: TToolbarButton97
        Left = 318
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
        ModalResult = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object Baide: TToolbarButton97
        Left = 350
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BaideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
    end
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Le circuit de validation ne peut comporter qu'#39'une se' +
        'ule '#233'tape.;W;O;O;O;'
      
        '1;?caption?;Les '#233'tapes du circuit de validation doivent '#234'tre con' +
        's'#233'cutives.;W;O;O;O;'
      
        '2;?caption?;Le libell'#233' d'#39'une '#233'tape du circuit de validation n'#39'es' +
        't pas renseign'#233'.;W;O;O;O;')
    Left = 281
    Top = 293
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 225
    Top = 8
  end
end
