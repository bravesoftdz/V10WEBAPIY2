inherited FAssistRazActivite: TFAssistRazActivite
  Left = 92
  Top = 131
  Caption = 'Assistant de remise '#224' z'#233'ro de l'#39'activit'#233
  ClientHeight = 315
  ClientWidth = 596
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Left = 10
    Top = 298
  end
  inherited lAide: THLabel
    Left = 181
    Top = 268
    Width = 405
  end
  inherited bPrecedent: TToolbarButton97
    Left = 371
    Top = 290
    Width = 70
  end
  inherited bSuivant: TToolbarButton97
    Left = 446
    Top = 290
    Width = 70
  end
  inherited bFin: TToolbarButton97
    Left = 522
    Top = 290
    Width = 70
    Caption = 'Fin'
  end
  inherited bAnnuler: TToolbarButton97
    Left = 296
    Top = 290
    Width = 70
  end
  inherited bAide: TToolbarButton97
    Left = 221
    Top = 290
    Width = 70
  end
  object bImprimer: TToolbarButton97 [7]
    Left = 145
    Top = 290
    Width = 70
    Height = 23
    Caption = 'Imprimer'
    Anchors = [akRight, akBottom]
    Flat = False
    OnClick = bImprimerClick
  end
  inherited Plan: THPanel
    Left = 175
    Top = 0
  end
  inherited GroupBox1: THGroupBox
    Left = 0
    Top = 280
    Width = 596
  end
  inherited P: THPageControl2
    Left = 184
    Top = 4
    Width = 409
    ActivePage = TSMotDePasse
    object TSMotDePasse: TTabSheet
      Caption = 'TSMotDePasse'
      object GBMotDePasse: TGroupBox
        Left = 0
        Top = 0
        Width = 409
        Height = 241
        Caption = 'Mot de passe'
        TabOrder = 0
        object Lib1: THLabel
          Left = 32
          Top = 28
          Width = 297
          Height = 29
          AutoSize = False
          Caption = 
            'Cet utilitaire supprime d'#233'finitivement toutes les donn'#233'es li'#233'es ' +
            #224' votre activit'#233'.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object TMotDePasse: THLabel
          Left = 32
          Top = 153
          Width = 64
          Height = 13
          Caption = 'Mot de passe'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Lib2: THLabel
          Left = 32
          Top = 68
          Width = 297
          Height = 29
          AutoSize = False
          Caption = 
            'Il faut imp'#233'rativement une sauvegarde r'#233'cente de votre base avan' +
            't de continuer. '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object Lib3: THLabel
          Left = 32
          Top = 108
          Width = 297
          Height = 29
          AutoSize = False
          Caption = 
            'Si vous n'#39'en avez pas, cliquez sur "Annuler" pour en faire une, ' +
            'sinon, veuillez saisir le mot de passe pour continuer.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object MotDePasse: THCritMaskEdit
          Left = 100
          Top = 148
          Width = 157
          Height = 21
          PasswordChar = '*'
          TabOrder = 0
          OnChange = MdpOnChange
          TagDispatch = 0
        end
      end
    end
    object TSOptions: TTabSheet
      Caption = 'TSOptions'
      ImageIndex = 3
      object GBOptions: TGroupBox
        Left = 0
        Top = 0
        Width = 409
        Height = 241
        Caption = 'Option'
        TabOrder = 0
        object CBSuppDepots: TCheckBox
          Left = 35
          Top = 24
          Width = 178
          Height = 17
          Caption = 'Suppression des d'#233'p'#244'ts'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 0
        end
      end
    end
    object TSConfirmation: TTabSheet
      Caption = 'TSConfirmation'
      ImageIndex = 1
      object GBConfirmation: TGroupBox
        Left = 0
        Top = 0
        Width = 409
        Height = 241
        Caption = 'Confirmation'
        TabOrder = 0
        object Lib4: TFlashingLabel
          Left = 32
          Top = 29
          Width = 88
          Height = 13
          Caption = 'ATTENTION !!!'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          Flashing = True
        end
        object Lib5: THLabel
          Left = 32
          Top = 52
          Width = 300
          Height = 29
          AutoSize = False
          Caption = 
            'Vous '#234'tes sur le point du supprimer d'#233'finitivement toutes les do' +
            'nn'#233'es de votre activit'#233' ...'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object Lib7: THLabel
          Left = 32
          Top = 180
          Width = 300
          Height = 37
          AutoSize = False
          Caption = 
            'Cliquez sur "Fin" pour valider l'#39#233'x'#233'cution du traitement, sinon,' +
            ' cliquez sur "Annuler".'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object Lib6: THLabel
          Left = 32
          Top = 84
          Width = 300
          Height = 37
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
      end
    end
    object TSRapport: TTabSheet
      Caption = 'TSRapport'
      ImageIndex = 2
      object RapportExec: TListBox
        Left = 0
        Top = 0
        Width = 401
        Height = 244
        Align = alRight
        BiDiMode = bdLeftToRight
        ItemHeight = 13
        ParentBiDiMode = False
        TabOrder = 0
      end
    end
  end
  inherited PanelImage: THPanel
    Left = 12
    Top = 37
  end
  inherited cControls: THListBox
    Left = 99
    Top = 251
  end
end
