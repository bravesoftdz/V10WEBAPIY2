object PrixMarche: TPrixMarche
  Left = 239
  Top = 156
  Width = 320
  Height = 302
  Caption = 'Gestion des prix march'#233's'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PPanBase: THPanel
    Left = 0
    Top = 57
    Width = 312
    Height = 211
    Align = alClient
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object LMontantHt: THLabel
      Left = 9
      Top = 66
      Width = 125
      Height = 13
      Caption = 'Montant H.T du document'
    end
    object Dock971: TDock97
      Left = 1
      Top = 174
      Width = 310
      Height = 36
      Position = dpBottom
      object ToolWindow971: TToolWindow97
        Left = 0
        Top = 0
        ClientHeight = 32
        ClientWidth = 301
        ActivateParent = False
        Caption = 'ToolWindow971'
        ClientAreaHeight = 32
        ClientAreaWidth = 301
        DockPos = 0
        FullSize = True
        TabOrder = 0
        DesignSize = (
          301
          32)
        object BAide: TToolbarButton97
          Tag = 1
          Left = 262
          Top = 3
          Width = 28
          Height = 27
          Hint = 'Aide'
          Anchors = [akTop, akRight]
          DisplayMode = dmGlyphOnly
          Caption = 'Aide'
          Flat = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Layout = blGlyphTop
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          Spacing = -1
          GlobalIndexImage = 'Z1117_S16G1'
          IsControl = True
        end
        object BAbandon: TToolbarButton97
          Tag = 1
          Left = 232
          Top = 3
          Width = 28
          Height = 27
          Hint = 'Fermer'
          Anchors = [akTop, akRight]
          DisplayMode = dmGlyphOnly
          Caption = 'Annuler'
          Flat = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ModalResult = 2
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          Spacing = -1
          OnClick = BAbandonClick
          GlobalIndexImage = 'Z0021_S16G1'
          IsControl = True
        end
        object BValider: TToolbarButton97
          Left = 202
          Top = 3
          Width = 28
          Height = 27
          Hint = 'Enregistrer la nomenclature'
          Anchors = [akTop, akRight]
          Default = True
          DisplayMode = dmGlyphOnly
          Caption = 'Enregistrer'
          Flat = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
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
          Layout = blGlyphTop
          ModalResult = 1
          NumGlyphs = 2
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          Spacing = -1
          OnClick = BValiderClick
          IsControl = True
        end
      end
    end
    object TMontantHt: THNumEdit
      Left = 175
      Top = 62
      Width = 124
      Height = 21
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      TabOrder = 1
      UseRounding = True
      Validate = False
    end
    object HNumEdit1: THNumEdit
      Left = 183
      Top = 72
      Width = 0
      Height = 21
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      TabOrder = 2
      UseRounding = True
      Validate = False
    end
    object GroupBox1: TGroupBox
      Left = 8
      Top = 104
      Width = 293
      Height = 65
      Caption = 'Pr'#233'f'#233'rences de calcul'
      TabOrder = 3
      object RMontant: TRadioButton
        Left = 12
        Top = 16
        Width = 273
        Height = 17
        Caption = 'Prorata du montant des lignes du document'
        Checked = True
        TabOrder = 0
        TabStop = True
      end
      object RMarge: TRadioButton
        Left = 12
        Top = 36
        Width = 273
        Height = 17
        Caption = 'Prorata de la marge des lignes du document'
        TabOrder = 1
      end
    end
  end
  object HPanel1: THPanel
    Left = 0
    Top = 0
    Width = 312
    Height = 57
    Align = alTop
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object G_APPLICCALC: TGroupBox
      Left = 4
      Top = 4
      Width = 301
      Height = 45
      Caption = 'Application sur'
      TabOrder = 0
      object RDOC: TRadioButton
        Left = 16
        Top = 20
        Width = 113
        Height = 17
        Caption = 'Le document'
        TabOrder = 0
        OnClick = RDOCClick
      end
      object RPARAG: TRadioButton
        Left = 180
        Top = 24
        Width = 113
        Height = 13
        Caption = 'Le paragraphe'
        TabOrder = 1
        OnClick = RPARAGClick
      end
    end
  end
end
