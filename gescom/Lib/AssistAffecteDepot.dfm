inherited FAssistAffecteDepot: TFAssistAffecteDepot
  Left = 352
  Top = 172
  Caption = 'Affectation d'#39'un d'#233'p'#244't aux mouvements'
  PixelsPerInch = 96
  TextHeight = 13
  inherited bPrecedent: TToolbarButton97
    Left = 307
    Width = 70
  end
  inherited bSuivant: TToolbarButton97
    Left = 383
    Width = 70
  end
  inherited bFin: TToolbarButton97
    Left = 459
    Width = 70
    Caption = 'Fin'
  end
  inherited bAnnuler: TToolbarButton97
    Left = 232
    Width = 70
  end
  inherited bAide: TToolbarButton97
    Left = 157
    Width = 70
  end
  object PopButton971: TPopButton97 [7]
    Left = 72
    Top = 8
    Width = 95
    Height = 22
    Visible = False
    ItemIndex = 0
  end
  object bImprimer: TToolbarButton97 [8]
    Left = 82
    Top = 308
    Width = 70
    Height = 23
    Anchors = [akRight, akBottom]
    Caption = 'Imprimer'
    Flat = False
    OnClick = bImprimerClick
  end
  inherited P: THPageControl2
    ActivePage = TSExplications
    object TSExplications: TTabSheet
      Caption = 'Explications'
      object GBExplications: TGroupBox
        Left = 0
        Top = 8
        Width = 337
        Height = 233
        TabOrder = 0
        object Lib1: THLabel
          Left = 8
          Top = 20
          Width = 321
          Height = 29
          AutoSize = False
          Caption = 
            'Cet utilitaire vous permet de cr'#233'er un d'#233'p'#244't et d'#39'affecter ce de' +
            'rnier aux mouvements qui n'#39'en ont pas.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object TMCodeDep: THLabel
          Left = 24
          Top = 104
          Width = 25
          Height = 13
          Caption = 'Code'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TMLibelleDep: THLabel
          Left = 24
          Top = 128
          Width = 30
          Height = 13
          Caption = 'Libelle'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Lib2: THLabel
          Left = 8
          Top = 60
          Width = 321
          Height = 21
          AutoSize = False
          Caption = 'Ce d'#233'p'#244't deviendra le d'#233'p'#244't par d'#233'faut'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object DepotCode: THCritMaskEdit
          Left = 60
          Top = 100
          Width = 77
          Height = 21
          TabOrder = 0
          OnChange = DepotCodeChange
          TagDispatch = 0
        end
        object DepotLib: THCritMaskEdit
          Left = 60
          Top = 124
          Width = 253
          Height = 21
          TabOrder = 1
          OnChange = DepotLibChange
          TagDispatch = 0
        end
      end
    end
    object TSConfirmation: TTabSheet
      Caption = 'Confirmation'
      ImageIndex = 1
      object GBCreateDepot: TGroupBox
        Left = 0
        Top = 8
        Width = 337
        Height = 233
        TabOrder = 0
        object Lib5: THLabel
          Left = 8
          Top = 204
          Width = 321
          Height = 13
          AutoSize = False
          Caption = 'Pour valider, cliquez sur le bouton "Fin".'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object Lib4: THLabel
          Left = 8
          Top = 188
          Width = 321
          Height = 13
          AutoSize = False
          Caption = 'Pour modifier le d'#233'p'#244't, cliquez sur le bouton "Pr'#233'c'#233'dent".'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object Lib3: THLabel
          Left = 8
          Top = 172
          Width = 321
          Height = 13
          AutoSize = False
          Caption = 'Pour annuler, cliquez sur le bouton "Annuler".'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object Lib6: THLabel
          Left = 8
          Top = 12
          Width = 65
          Height = 13
          AutoSize = False
          Caption = 'R'#233'capitulatif'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object LBRecap: TListBox
          Left = 6
          Top = 24
          Width = 323
          Height = 145
          BiDiMode = bdLeftToRight
          ItemHeight = 13
          ParentBiDiMode = False
          TabOrder = 0
        end
      end
    end
    object TSRapport: TTabSheet
      Caption = 'TSRapport'
      ImageIndex = 2
      object LBRapport: TListBox
        Left = 6
        Top = 0
        Width = 323
        Height = 241
        BiDiMode = bdLeftToRight
        ItemHeight = 13
        ParentBiDiMode = False
        TabOrder = 0
      end
    end
  end
end
