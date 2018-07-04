{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  09/02/06   JP   Création de l'unité : Récupération de données et action
 8.01.001.001  24/01/07   JP   FQ 10009 : Ajout des colonnes Prev_Viseur, Next_Viseur,
                               Bap_Datecomptable / Gestion de la GED (AfficheFacture)
 8.01.001.020  20/06/07   JP   Maj du champ BAP_VISEUR lors de la validation
 8.01.001.021  21/06/07   JP   FQ 10019 : Ajout de champs dans les éditions
--------------------------------------------------------------------------------------}
unit CPVIGSUIVIBAP;

interface

uses
  Classes, UTob,
  {$IFNDEF DBXPRESS}
  dbtables,
  {$ELSE}
  uDbxDataSet,
  {$ENDIF}
  uCPVignettePlugIn, HCtrls;

type
  TSuiviBAP = (ab_Mul, ab_RecupMemo);
  TObjetSuiviBap = class(TCPVignettePlugIn)
  private
    TypeAction : TSuiviBAP;
  protected
    procedure RecupDonnees;                 override;
    procedure GetClauseWhere;               override;
    procedure TraiteParam(Param  : string); override;
    procedure DrawGrid   (Grille : string); override;
  public
    procedure ChargerMul;
    procedure RecupMemo;
  end;


implementation

uses
  SysUtils, HEnt1, ParamSoc, UtilPgi, uToolsPlugin, RTFCounter;

  
{---------------------------------------------------------------------------------------}
procedure TObjetSuiviBap.DrawGrid(Grille: string);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  inherited;

  if TobDonnees.Detail.Count = 0 then begin
    T := TOB.Create('****', TobDonnees, -1);
    T.AddChampSupValeur('BAP_JOURNAL'      , '');
    T.AddChampSupValeur('BAP_NUMEROPIECE'  , '');
    T.AddChampSupValeur('BAP_NUMEROORDRE'  , '');
    T.AddChampSupValeur('E_LIBELLE'        , '');
    T.AddChampSupValeur('T_LIBELLE'        , '');
    T.AddChampSupValeur('BAP_DATECOMPTABLE', '');
    T.AddChampSupValeur('E_DATEECHEANCE'   , '');
    T.AddChampSupValeur('E_MONTANT'        , '');
    T.AddChampSupValeur('BAP_VISEUR'      , '');
    T.AddChampSupValeur('BAP_VISEUR1'      , '');
    T.AddChampSupValeur('BAP_ECHEANCEBAP'  , '');
    T.AddChampSupValeur('BAP_STATUTBAP'    , '');
    PutGridDetail('FListe', TobDonnees);
  end;
(*
    BAP_JOURNAL, BAP_EXERCICE, BAP_NUMEROPIECE, BAP_NUMEROORDRE, E_AUXILIAIRE, T_LIBELLE,
    E_LIBELLE, BAP_DATECOMPTABLE, E_DATEECHEANCE, E_MONTANT, E_MODEPAIE,
    U1.US_LIBELLE BAP_VISEUR, U2.US_LIBELLE BAP_VISEUR1,
    U3.US_LIBELLE VISEUR2, BAP_ECHEANCEBAP, BAP_STATUTBAP, BAP_BLOCNOTE
  *)

  ddWriteLN('CPSUIVIBAP : DRAWGRID');
  SetVisibleCol('FListe', 'BAP_EXERCICE', False);
  SetVisibleCol('FListe', 'E_AUXILIAIRE', False);
  SetVisibleCol('FListe', 'E_MODEPAIE', False);
  SetVisibleCol('FListe', 'BAP_VISEUR2', False);
//  SetVisibleCol('GRID', 'E_DATEECHEANCE', False); {FQ 10009 : 24/01/07}
//  SetVisibleCol('GRID', 'BAP_NUMEROORDRE', False);

  SetTitreCol('FListe' , 'BAP_JOURNAL', 'Jal.');
  SetTitreCol('FListe' , 'BAP_NUMEROPIECE', 'Num. Pièce');
  SetTitreCol('FListe' , 'BAP_NUMEROORDRE', 'Étape');
  SetTitreCol('FListe' , 'T_LIBELLE', 'Fournisseur');
  SetTitreCol('FListe' , 'E_LIBELLE', 'Libellé');
  SetTitreCol('FListe' , 'BAP_DATECOMPTABLE', 'Date Cpt.'); {FQ 10009 : 24/01/07}
  SetTitreCol('FListe' , 'E_DATEECHEANCE', 'Date Ech.'); {FQ 10009 : 24/01/07}
  SetTitreCol('FListe' , 'E_MONTANT', 'Mnt. TTC');
  SetTitreCol('FListe' , 'BAP_VISEUR', 'Vis. Eff'); {FQ 10009 : 24/01/07}
  SetTitreCol('FListe' , 'BAP_VISEUR1', 'Princ.');
  SetTitreCol('FListe' , 'BAP_ECHEANCEBAP', 'Eché. du BAP');
  SetTitreCol('FListe' , 'BAP_STATUTBAP', 'Statut');
  SetTitreCol('FListe' , 'BAP_BLOCNOTE', 'Suivi');

  SetWidthCol('FListe' , 'BAP_JOURNAL', 8);
  SetWidthCol('FListe' , 'BAP_NUMEROPIECE', 12);
  SetWidthCol('FListe' , 'BAP_NUMEROORDRE', 12);
  SetWidthCol('FListe' , 'T_LIBELLE', 24);
  SetWidthCol('FListe' , 'E_LIBELLE', 32);
  SetWidthCol('FListe' , 'BAP_DATECOMPTABLE', 13);
  SetWidthCol('FListe' , 'E_DATEECHEANCE', 13);
  SetWidthCol('FListe' , 'E_MONTANT', 13);
  SetWidthCol('FListe' , 'BAP_VISEUR', 13);
  SetWidthCol('FListe' , 'BAP_VISEUR1', 13);
  SetWidthCol('FListe' , 'BAP_ECHEANCEBAP', 13);
  SetWidthCol('FListe' , 'BAP_STATUTBAP', 7);
  SetWidthCol('FListe' , 'BAP_BLOCNOTE', 10);

  SetFormatCol('FListe' , 'BAP_JOURNAL', 'C.0O ---');
  SetFormatCol('FListe' , 'BAP_NUMEROPIECE', 'C.0O ---');
  SetFormatCol('FListe' , 'BAP_NUMEROORDRE', 'C.0O ---');
  SetFormatCol('FListe' , 'T_LIBELLE', 'G.0O ---');
  SetFormatCol('FListe' , 'E_LIBELLE', 'G.0O ---');
  SetFormatCol('FListe' , 'BAP_DATECOMPTABLE', 'C.0O ---');
  SetFormatCol('FListe' , 'E_DATEECHEANCE', 'C.0O ---');
  SetFormatCol('FListe' , 'E_MONTANT', 'D.2O ---');
  SetFormatCol('FListe' , 'BAP_VISEUR', 'C.0O ---');
  SetFormatCol('FListe' , 'BAP_VISEUR1', 'C.0O ---');
  SetFormatCol('FListe' , 'BAP_ECHEANCEBAP', 'C.0O ---');
  SetFormatCol('FListe' , 'BAP_STATUTBAP', 'C.0O ---');
  SetFormatCol('FListe' , 'BAP_BLOCNOTE', 'C.0O ---');

  SetTypeCol('FListe' , 'BAP_BLOCNOTE', 'M');

//  TobResponse.SaveToXMLFile('C:\Documents and Settings\Pasteris\Bureau\Draw.txt', True, True);
end;

{---------------------------------------------------------------------------------------}
procedure TObjetSuiviBap.RecupDonnees;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  try
    ddWriteLN('CPSUIVIBAP : Debut du suivi des BAP');

    case TypeAction of
      ab_Mul        : ChargerMul;
      ab_RecupMemo  : RecupMemo; {Appel de la vignette Memo}
    end;

    ddWriteLN('CPSUIVIBAP : Fin du suivi des BAP');
  except
    on E : Exception do
      MessageErreur := 'Erreur lors du traitement des données avec le message :'#13 + E.Message;
  end;
end;

{13/06/07 : Ajout des left join sur Utilisateur pour afficher les libellés
{---------------------------------------------------------------------------------------}
procedure TObjetSuiviBap.GetClauseWhere;
{---------------------------------------------------------------------------------------}
begin
  ddWriteLN('CPSUIVIBAP : User = ' + CodeUser);
  ClauseWhere :=
    'SELECT BAP_JOURNAL, BAP_EXERCICE, BAP_NUMEROPIECE, BAP_NUMEROORDRE, E_AUXILIAIRE, T_LIBELLE, ' +
    'E_LIBELLE, BAP_DATECOMPTABLE, E_DATEECHEANCE, E_MONTANT, E_MODEPAIE, ' +
    'U1.US_LIBELLE BAP_VISEUR, U2.US_LIBELLE BAP_VISEUR1, ' +
    'U3.US_LIBELLE BAP_VISEUR2, BAP_ECHEANCEBAP, BAP_STATUTBAP, BAP_BLOCNOTE ' + 
    'FROM CPBAPECRITURE B1 ';
  ClauseWhere := ClauseWhere +
    'LEFT JOIN UTILISAT U1 ON U1.US_UTILISATEUR = BAP_VISEUR ' +
    'LEFT JOIN UTILISAT U2 ON U2.US_UTILISATEUR = BAP_VISEUR1 ' +
    'LEFT JOIN UTILISAT U3 ON U3.US_UTILISATEUR = BAP_VISEUR2 ' +
    'WHERE EXISTS (SELECT BAP_JOURNAL FROM CPBONSAPAYER B2 ' +
           'WHERE B2.BAP_JOURNAL = B1.BAP_JOURNAL ' +
           'AND B2.BAP_EXERCICE = B1.BAP_EXERCICE ' +
           'AND B2.BAP_DATECOMPTABLE = B1.BAP_DATECOMPTABLE ' +
           'AND B2.BAP_NUMEROPIECE = B1.BAP_NUMEROPIECE ' + 
           'AND (B2.BAP_VISEUR1 = "' + CodeUser + '" OR B2.BAP_VISEUR2 = "' + CodeUser + '") ' +
           ') AND NOT EXISTS (SELECT BAP_JOURNAL FROM CPBONSAPAYER B2 ' +
           'WHERE B2.BAP_JOURNAL = B1.BAP_JOURNAL ' +
           'AND B2.BAP_EXERCICE = B1.BAP_EXERCICE ' +
           'AND B2.BAP_DATECOMPTABLE = B1.BAP_DATECOMPTABLE ' +
           'AND B2.BAP_NUMEROPIECE = B1.BAP_NUMEROPIECE ' +
           'AND B2.BAP_STATUTBAP = "DEF")';

  ClauseWhere := ClauseWhere + 'ORDER BY BAP_JOURNAL, BAP_NUMEROPIECE, BAP_NUMEROORDRE';
  ddWriteLN('CPSUIVIBAP : Requête = ' + ClauseWhere);
end;

{---------------------------------------------------------------------------------------}
procedure TObjetSuiviBap.TraiteParam(Param : string);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  inherited;
  {Reste des paramètres, en attendant d'autres variables}
  s := ReadTokenSt(FParam);
  if s = 'DETA' then TypeAction := ab_RecupMemo
                else TypeAction := ab_mul;
end;

{---------------------------------------------------------------------------------------}
procedure TObjetSuiviBap.RecupMemo;
{---------------------------------------------------------------------------------------}
var
  F : TOB;
  T : TOB;
begin
  inherited;
  SansInterface := True;
  {On commence par remplir la TobSelection}
  AvecSelection := True;
  GetInterface('FListe');
  AvecSelection := False;
  WarningOk := False;

  if TobSelection.Detail.Count = 0 then
    MessageErreur := TraduireMemoire('Veuillez sélectionner une ligne.')
  else begin
    F := TobSelection.Detail[0];
    T := GetVignetteZoom('CP', 'CPVIGMEMOBAP', TraduireMemoire('Suivi des commentaires'));
    if T.Detail.Count > 0 then begin
      T.Detail[0].AddChampSupValeur('NAME', 'BAP_BLOCNOTE');
      T.Detail[0].AddChampSupValeur('VALUE', F.GetValue('BAP_BLOCNOTE'));
    end
    else
      MessageErreur := TraduireMemoire('Impossible de mémoriser la clef du BAP.')

  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjetSuiviBap.ChargerMul;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  ddwriteln('CPSUIVIBAP : ChargerMul');
  try
    try
      Q := OpenSelect(ClauseWhere);
      TobDonnees.LoadDetailDB('****', '', '', Q, False);
      ddwriteln('CPSUIVIBAP : Nb enreg = ' + IntToStr(TobDonnees.Detail.Count));
      if (TobDonnees.Detail.Count = 0) and (MessageErreur = '') then
        MessageErreur := TraduireMemoire('Vous n''avez pas de bons à payer à viser');
    except
      on E : Exception do
        MessageErreur := E.Message;
    end
  finally
    Ferme(Q);
  end;
end;

end.
 