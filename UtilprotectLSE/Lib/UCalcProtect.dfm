object FCalcProtec: TFCalcProtec
  Left = 393
  Top = 366
  Width = 552
  Height = 184
  ActiveControl = EMPLACESTO
  Caption = 'Calcul des protections'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 11
    Top = 36
    Width = 109
    Height = 16
    Caption = 'Identifiant client'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label2: TLabel
    Left = 11
    Top = 59
    Width = 137
    Height = 16
    Caption = 'Date fin d'#39'utilisation'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label3: TLabel
    Left = 11
    Top = 82
    Width = 162
    Height = 16
    Caption = 'Emplacement du fichier'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Panel1: TPanel
    Left = 0
    Top = 112
    Width = 536
    Height = 33
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      536
      33)
    object BABANDON: TToolbarButton97
      Left = 505
      Top = 3
      Width = 24
      Height = 24
      Anchors = [akTop, akRight]
      Cancel = True
      DisplayMode = dmGlyphOnly
      ModalResult = 2
      GlobalIndexImage = 'Z0021_S24G1'
    end
    object ToolbarButton971: TToolbarButton97
      Left = 476
      Top = 3
      Width = 24
      Height = 24
      Anchors = [akTop, akRight]
      DisplayMode = dmGlyphOnly
      OnClick = ToolbarButton971Click
      GlobalIndexImage = 'Z0127_S24G1'
    end
  end
  object IDCLIENT: TEdit
    Left = 182
    Top = 34
    Width = 295
    Height = 21
    TabOrder = 1
  end
  object EMPLACESTO: THCritMaskEdit
    Left = 182
    Top = 80
    Width = 327
    Height = 21
    TabOrder = 2
    TagDispatch = 0
    ElipsisButton = True
    OnElipsisClick = EMPLACESTOElipsisClick
  end
  object DATEFIN: THCritMaskEdit
    Left = 182
    Top = 57
    Width = 75
    Height = 21
    EditMask = '!99 >L<LL 0000;1;_'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    MaxLength = 11
    ParentFont = False
    TabOrder = 3
    Text = '           '
    TagDispatch = 0
    OpeType = otDate
    DefaultDate = odDate
    ControlerDate = True
  end
end
