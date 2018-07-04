inherited FOutilsLivr: TFOutilsLivr
  Left = 286
  Top = 152
  Width = 465
  Height = 420
  Caption = 'Revalorisation des livraisons'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 347
    Width = 457
    Height = 39
    inherited PBouton: TToolWindow97
      ClientHeight = 35
      ClientWidth = 457
      ClientAreaHeight = 35
      ClientAreaWidth = 457
      inherited BValider: TToolbarButton97
        Left = 361
        Visible = False
      end
      inherited BFerme: TToolbarButton97
        Left = 393
      end
      inherited HelpBtn: TToolbarButton97
        Left = 425
      end
      inherited BImprimer: TToolbarButton97
        Left = 329
      end
      object Blance: TToolbarButton97
        Left = 297
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Valider'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Default = True
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0217_S16G1'
        IsControl = True
      end
    end
  end
  object MRetour: TMemo [1]
    Left = 0
    Top = 0
    Width = 457
    Height = 347
    Align = alClient
    ScrollBars = ssBoth
    TabOrder = 1
  end
end
