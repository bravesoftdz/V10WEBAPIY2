object FVerImmo: TFVerImmo
  Left = 279
  Top = 174
  Width = 325
  Height = 220
  Caption = 'FVerImmo'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 40
    Top = 88
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Le code n'#39'est pas renseign'#233
      'Le libell'#233' n'#39'est pas renseign'#233
      'La longueur du code est non conforme'
      'Caract'#232're du compte non autoris'#233
      'La m'#233'thode '#233'conomique n'#39'est pas renseign'#233'e'
      'La dur'#233'e '#233'conomique n'#39'est pas renseign'#233'e'
      'Le taux '#233'conomique n'#39'est pas renseign'#233
      'Le qualifiant de l'#39'immobilisation n'#39'est pas renseign'#233
      'Le montant n'#39'est pas renseign'#233
      ' n'#39'existe pas'
      'L'#39'immobilisation n'#39'existe pas'
      'Le compte n'#39'existe pas'
      'L'#39'existence d'#39'une op'#233'ration n'#39'est pas renseign'#233'.'
      'L'#39#233'tat n'#39'est pas renseign'#233'.'
      'L'#39#39'op'#233'ration d'#39#39'acquisition n'#39#39'existe pas'
      'Le plan d'#39'amortissement n'#39'existe pas'
      'Le plan d'#39'amortissement est d'#233'synchronis'#233' de la fiche'
      'Le plan d'#39'amortissement est d'#233'synchronis'#233' de l'#39'historique'
      '')
    Left = 112
    Top = 88
  end
  object MsgRien: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Aucun enregistrement ne correspond aux crit'#232'res que ' +
        'vous avez s'#233'lectionn'#233's.;W;O;O;O;'
      '1;Immobilisations;Aucune erreur n'#39'a '#233't'#233' d'#233'tect'#233'e;E;O;O;O;'
      '2;Comptes associ'#233's;Aucune erreur n'#39'a '#233't'#233' d'#233'tect'#233'e;E;O;O;O;'
      '3;Amortissements;Aucune erreur n'#39'a '#233't'#233' d'#233'tect'#233'e;E;O;O;O;'
      '4;Op'#233'rations;Aucune erreur n'#39'a '#233't'#233' d'#233'tect'#233'e;E;O;O;O;'
      '5;Ech'#233'ances;Aucune erreur n'#39'a '#233't'#233' d'#233'tect'#233'e;E;O;O;O;')
    Left = 164
    Top = 88
  end
  object MsgLib: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Nature d'#39'immobilsation'
      'Compte d'#39'immobilisation'
      'Compte de charge'
      'Compte d'#39'amortissement'
      'Compte de dotation'
      'Compte d'#233'rogatoire'
      'Compte de reprise d'#233'rogatoire'
      'Compte de provision d'#233'rogatoire'
      'Compte de dotation exceptionnelle'
      'Compte de V.O. c'#233'd'#233'e'
      'Compte de V.A. c'#233'd'#233'e'
      'Compte d'#39'amortissement c'#233'd'#233
      'Compte de reprise d'#39'exploitation'
      'Compte de reprise exceptionnelle'
      'M'#233'thode '#233'conomique'
      'Dur'#233'e '#233'conomique'
      'Taux '#233'conomique'
      'Montant d'#39'amortissement'
      'Le plan d'#39'amortissement'
      'Montant d'#39'achat'
      'Organisme CB'
      'Op'#233'ration')
    Left = 232
    Top = 84
  end
end
