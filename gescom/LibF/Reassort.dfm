inherited FReassort: TFReassort
  Left = 299
  Top = 150
  Width = 698
  Caption = 'Commande de réassort'
  PixelsPerInch = 96
  TextHeight = 13
  object HLabel2: THLabel [0]
    Left = 491
    Top = 30
    Width = 53
    Height = 13
    Caption = 'Livraison le'
    Transparent = True
  end
  inherited PEntete: THPanel
    Width = 690
    inherited GP_DATEREFEXTERNE: THCritMaskEdit
      TabOrder = 27
    end
    object PEnteteReassort: THPanel
      Left = 0
      Top = 0
      Width = 690
      Height = 112
      FullRepaint = False
      TabOrder = 25
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      object FTitreReassort: THLabel
        Left = 6
        Top = 6
        Width = 226
        Height = 20
        Alignment = taCenter
        AutoSize = False
        Caption = 'Commande de réassort'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        Ombre = True
        OmbreDecalX = 1
        OmbreDecalY = 1
        OmbreColor = clGray
      end
      object HGP_NUMEROPIECE_: THLabel
        Left = 8
        Top = 46
        Width = 12
        Height = 13
        Caption = 'N°'
        Transparent = True
      end
      object HGP_DATEPIECE_: THLabel
        Left = 120
        Top = 46
        Width = 12
        Height = 13
        Caption = 'du'
        FocusControl = GP_DATEPIECE_
        Transparent = True
      end
      object HGP_REFINTERNE_: THLabel
        Left = 236
        Top = 10
        Width = 50
        Height = 13
        Caption = 'Référence'
        FocusControl = GP_REFINTERNE_
        Transparent = True
      end
      object HLabel1: THLabel
        Left = 236
        Top = 32
        Width = 61
        Height = 13
        Caption = 'Commentaire'
        Transparent = True
      end
      object HGP_DATELIVRAISON_: THLabel
        Left = 520
        Top = 10
        Width = 53
        Height = 13
        Caption = 'Livraison le'
        Transparent = True
      end
      object GP_NUMEROPIECE_: THPanel
        Left = 35
        Top = 42
        Width = 80
        Height = 21
        BevelOuter = bvLowered
        Caption = 'Non affecté'
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        BackGroundEffect = bdFond
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        ColorNb = 2
        TextEffect = tenone
      end
      object GP_DATEPIECE_: THCritMaskEdit
        Left = 139
        Top = 41
        Width = 78
        Height = 21
        EditMask = '!99/99/0000;1;_'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 10
        ParentFont = False
        TabOrder = 1
        Text = '  /  /    '
        OnExit = GP_DATEPIECEExit
        TagDispatch = 0
        Operateur = Egal
        OpeType = otDate
        ElipsisButton = True
        ElipsisAutoHide = True
        ControlerDate = True
      end
      object GP_REFINTERNE_: TEdit
        Left = 310
        Top = 6
        Width = 183
        Height = 21
        TabOrder = 2
      end
      object GP_BLOCNOTE: THRichEditOLE
        Left = 310
        Top = 30
        Width = 183
        Height = 76
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        Margins.Top = 0
        Margins.Bottom = 0
        Margins.Left = 0
        Margins.Right = 0
        ContainerName = 'Document'
        ObjectMenuPrefix = '&Object'
        LinesRTF.Strings = (
          
            '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fnil Aria' +
            'l;}}'
          '\viewkind4\uc1\pard\f0\fs16 GP_BLOCNOTE'
          '\par }')
      end
      object GP_DATELIVRAISON_: THCritMaskEdit
        Left = 585
        Top = 6
        Width = 81
        Height = 21
        EditMask = '!99/99/0000;1;_'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 10
        ParentFont = False
        TabOrder = 4
        Text = '  /  /    '
        OnExit = GP_DATELIVRAISONExit
        TagDispatch = 0
        Operateur = Egal
        OpeType = otDate
        ElipsisButton = True
        ElipsisAutoHide = True
        ControlerDate = True
      end
    end
  end
  inherited PPied: THPanel
    Width = 690
  end
  inherited DockBottom: TDock97
    Width = 690
  end
  inherited SB: TScrollBox
    Width = 690
    inherited Debut: THRichEditOLE
      Width = 670
    end
    inherited Fin: THRichEditOLE
      Width = 670
    end
    inherited GS: THGrid
      Width = 670
    end
  end
  inherited TDescriptif: TToolWindow97
    Left = 22
    Top = 154
  end
end
