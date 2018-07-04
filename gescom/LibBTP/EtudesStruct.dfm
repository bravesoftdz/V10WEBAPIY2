object FEtudesStruct: TFEtudesStruct
  Left = 102
  Top = 244
  BorderStyle = bsToolWindow
  Caption = 'D'#233'finition de la structure'
  ClientHeight = 311
  ClientWidth = 757
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 53
    Width = 757
    Height = 227
    Align = alClient
    TabOrder = 0
    object GS: THGrid
      Left = 1
      Top = 1
      Width = 755
      Height = 225
      Align = alClient
      ColCount = 4
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 8
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goTabs]
      TabOrder = 0
      OnDblClick = GSDblClick
      OnEnter = GSEnter
      SortedCol = -1
      Couleur = False
      MultiSelect = False
      TitleBold = True
      TitleCenter = True
      OnCellEnter = GSCellEnter
      OnCellExit = GSCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      ColWidths = (
        248
        114
        323
        34)
    end
  end
  object DockBottom: TDock97
    Left = 0
    Top = 280
    Width = 757
    Height = 31
    BackgroundTransparent = True
    Position = dpBottom
    object Valide97: TToolbar97
      Left = 625
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DefaultDock = DockBottom
      DockPos = 625
      DragHandleStyle = dhNone
      TabOrder = 0
      DesignSize = (
        84
        27)
      object BValider: TToolbarButton97
        Left = 0
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Enregistrer la pi'#232'ce'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Enregistrer'
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
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 28
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Anchors = [akTop, akRight]
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Annuler'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ModalResult = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAbandonClick
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 56
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Aide'
        HelpContext = 119000017
        Alignment = taRightJustify
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
    object Outils97: TToolbar97
      Left = 0
      Top = 0
      Caption = 'Actions'
      CloseButton = False
      DefaultDock = DockBottom
      DockPos = 0
      DragHandleStyle = dhNone
      TabOrder = 1
      object BDelete: TToolbarButton97
        Left = 29
        Top = 0
        Width = 36
        Height = 27
        Hint = 'Supprimer'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopupDel
        Caption = 'Supprimer'
        Layout = blGlyphTop
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0005_S16G1'
      end
      object BNewligne: TToolbarButton97
        Left = 0
        Top = 0
        Width = 29
        Height = 27
        Hint = 'Nouvelle feuille'
        DisplayMode = dmGlyphOnly
        Caption = 'Nouvelle feuille'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BNewligneClick
        GlobalIndexImage = 'Z0053_S16G1'
      end
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 757
    Height = 53
    Align = alTop
    TabOrder = 2
    object GroupBox1: TGroupBox
      Left = 8
      Top = 4
      Width = 585
      Height = 41
      TabOrder = 0
      object ApplicAffaire: TRadioButton
        Left = 69
        Top = 15
        Width = 146
        Height = 13
        Caption = 'Applicable '#224' ce document'
        Checked = True
        TabOrder = 0
        TabStop = True
      end
      object ApplicClient: TRadioButton
        Left = 301
        Top = 16
        Width = 200
        Height = 13
        Caption = 'Applicable '#224' cette nature pour le tiers'
        TabOrder = 1
      end
    end
  end
  object PopupDel: TPopupMenu
    Left = 224
    Top = 105
    object LigCurDel: TMenuItem
      Caption = 'Ligne Courante'
      OnClick = LigCurDelClick
    end
    object StructDel: TMenuItem
      Caption = 'Toute la d'#233'finition'
      OnClick = StructDelClick
    end
  end
end
