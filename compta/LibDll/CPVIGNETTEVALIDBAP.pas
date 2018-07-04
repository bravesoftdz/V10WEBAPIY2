{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  09/02/06   JP   Création de l'unité : Récupération de données et action
 8.01.001.001  24/01/07   JP   FQ 10009 : Ajout des colonnes Prev_Viseur, Next_Viseur,
                               Bap_Datecomptable / Gestion de la GED (AfficheFacture)
 8.01.001.020  20/06/07   JP   Maj du champ BAP_VISEUR lors de la validation
 8.01.001.021  21/06/07   JP   FQ 10019 : Ajout de champs dans les éditions
--------------------------------------------------------------------------------------}
unit CPVIGNETTEVALIDBAP;

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
  TActionBAP = (ab_Etat, ab_Validation, ab_Mul, ab_Facture, ab_RecupMemo, ab_ChargeMemo);
  TObjetValidBap = class(TCPVignettePlugIn)
  private
    TypeAction : TActionBAP;
    EnCroisAxe : Boolean;
    sServer    : string;
    sFrom      : string;
  protected
    procedure RecupDonnees;                 override;
    procedure GetClauseWhere;               override;
    procedure TraiteParam(Param  : string); override;
    procedure DrawGrid   (Grille : string); override;
  public
    procedure ChargerMul;
    procedure LanceEtat;
    procedure Valider;
    procedure AfficheFacture; {24/01/07 : FQ 10009}
    procedure RemplitTobDonneesEtat;
    procedure RecupMemo;
    procedure ChargeMemo;
    procedure EnvoieMail(Fact : TOB);
  end;

implementation

uses
  SysUtils, HEnt1, ParamSoc, UtilPgi, uToolsPlugin, RTFCounter, Registry, eHttp, Windows;

{---------------------------------------------------------------------------------------}
procedure TObjetValidBap.DrawGrid(Grille: string);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  inherited;

  if not (TypeAction in [ab_Validation, ab_Mul]) then Exit;

  if TobDonnees.Detail.Count = 0 then begin
    T := TOB.Create('****', TobDonnees, -1);
    T.AddChampSupValeur('BAP_JOURNAL'      , '');
    T.AddChampSupValeur('BAP_NUMEROPIECE'  , '');
    T.AddChampSupValeur('E_LIBELLE'        , '');
    T.AddChampSupValeur('T_LIBELLE'        , '');
    T.AddChampSupValeur('BAP_DATECOMPTABLE', '');
    T.AddChampSupValeur('E_MONTANT'        , '');
    T.AddChampSupValeur('BAP_ECHEANCEBAP'  , '');
    T.AddChampSupValeur('BAP_VISEUR1'      , '');
    T.AddChampSupValeur('BAP_VISEUR2'      , '');
    T.AddChampSupValeur('PREV_VISEUR'      , '');
    T.AddChampSupValeur('NEXT_VISEUR'      , '');
    T.AddChampSupValeur('BAP_STATUTBAP'    , '');
    PutGridDetail('FListe', TobDonnees);
  end;
(*
BAP_JOURNAL, BAP_NUMEROPIECE, E_AUXILIAIRE, T_LIBELLE, E_REFINTERNE, E_LIBELLE, ' +
    'E_MODEPAIE, E_DATEECHEANCE, E_MONTANT, BAP_EXERCICE, BAP_NUMEROORDRE, BAP_DATECOMPTABLE, ' +
    'BAP_TIERSPAYEUR, BAP_CLEFFACTURE, BAP_VISEUR1, BAP_VISEUR2, ' +
    'PREV_VISEUR, NEXT_VISEUR, BAP_ECHEANCEBAP, BAP_STATUTBAP, BAP_BLOCNOTE, BAP_IDGED
  *)

  ddWriteLN('CPVIGNETTEVALIDBAP : DRAWGRID');
  SetVisibleCol('FListe', 'E_AUXILIAIRE', False);
  SetVisibleCol('FListe', 'E_REFINTERNE', False);
  SetVisibleCol('FListe', 'E_MODEPAIE', False);
  SetVisibleCol('FListe', 'E_DATEECHEANCE', False); {FQ 10009 : 24/01/07}
  SetVisibleCol('FListe', 'BAP_EXERCICE', False);
  SetVisibleCol('FListe', 'BAP_NUMEROORDRE', False);
  SetVisibleCol('FListe', 'BAP_TIERSPAYEUR', False);
  SetVisibleCol('FListe', 'BAP_CLEFFACTURE', False);
  SetVisibleCol('FListe', 'US_EMAIL', False);
//  SetVisibleCol('FListe', 'BAP_BLOCNOTE', False);
  SetVisibleCol('FListe', 'BAP_IDGED', False);

  SetTitreCol('FListe' , 'BAP_JOURNAL', 'Jal.');
  SetTitreCol('FListe' , 'BAP_NUMEROPIECE', 'Num. Pièce');
  SetTitreCol('FListe' , 'T_LIBELLE', 'Fournisseur');
  SetTitreCol('FListe' , 'E_LIBELLE', 'Libellé');
  SetTitreCol('FListe' , 'E_MONTANT', 'Mnt. TTC');
  SetTitreCol('FListe' , 'BAP_DATECOMPTABLE', 'Date Cpt.'); {FQ 10009 : 24/01/07}
  SetTitreCol('FListe' , 'BAP_VISEUR1', 'Princ.');
  SetTitreCol('FListe' , 'BAP_VISEUR2', 'Suppl.');
  SetTitreCol('FListe' , 'PREV_VISEUR', 'Préc.'); {FQ 10009 : 24/01/07}
  SetTitreCol('FListe' , 'NEXT_VISEUR', 'Suiv.'); {FQ 10009 : 24/01/07}
  SetTitreCol('FListe' , 'BAP_ECHEANCEBAP', 'Eché. du BAP');
  SetTitreCol('FListe' , 'BAP_STATUTBAP', 'Statut');
  SetTitreCol('FListe' , 'BAP_BLOCNOTE', 'Suivi');

  SetWidthCol('FListe' , 'BAP_JOURNAL', 7);
  SetWidthCol('FListe' , 'BAP_NUMEROPIECE', 12);
  SetWidthCol('FListe' , 'T_LIBELLE', 24);
  SetWidthCol('FListe' , 'E_LIBELLE', 32);
  SetWidthCol('FListe' , 'E_MONTANT', 13);
  SetWidthCol('FListe' , 'BAP_DATECOMPTABLE', 13);
  SetWidthCol('FListe' , 'BAP_VISEUR1', 13);
  SetWidthCol('FListe' , 'BAP_VISEUR2', 13);
  SetWidthCol('FListe' , 'PREV_VISEUR', 13);
  SetWidthCol('FListe' , 'NEXT_VISEUR', 13);
  SetWidthCol('FListe' , 'BAP_ECHEANCEBAP', 13);
  SetWidthCol('FListe' , 'BAP_STATUTBAP', 7);
  SetWidthCol('FListe' , 'BAP_BLOCNOTE', 10);

  SetFormatCol('FListe' , 'BAP_JOURNAL', 'C.0O ---');
  SetFormatCol('FListe' , 'BAP_NUMEROPIECE', 'C.0O ---');
  SetFormatCol('FListe' , 'T_LIBELLE', 'G.0O ---');
  SetFormatCol('FListe' , 'E_LIBELLE', 'G.0O ---');
  SetFormatCol('FListe' , 'E_MONTANT', 'D.2O ---');
  SetFormatCol('FListe' , 'BAP_DATECOMPTABLE', 'C.0O ---');
  SetFormatCol('FListe' , 'BAP_VISEUR1', 'C.0O ---');
  SetFormatCol('FListe' , 'BAP_VISEUR2', 'C.0O ---');
  SetFormatCol('FListe' , 'PREV_VISEUR', 'C.0O ---'); {FQ 10009 : 24/01/07}
  SetFormatCol('FListe' , 'NEXT_VISEUR', 'C.0O ---'); {FQ 10009 : 24/01/07}
  SetFormatCol('FListe' , 'BAP_ECHEANCEBAP', 'C.0O ---');
  SetFormatCol('FListe' , 'BAP_STATUTBAP', 'C.0O ---');
  SetFormatCol('FListe' , 'BAP_BLOCNOTE', 'C.0O ---');

  SetTypeCol('FListe' , 'BAP_BLOCNOTE', 'M');

//  TobResponse.SaveToXMLFile('C:\Documents and Settings\Pasteris\Bureau\Draw.txt', True, True);
end;

{---------------------------------------------------------------------------------------}
procedure TObjetValidBap.RecupDonnees;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  try
    ddWriteLN('CPVIGNETTEVALIDBAP : Debut des traitements BAP');

    case TypeAction of
      ab_Mul        : ChargerMul;
      ab_Etat       : LanceEtat;
      ab_Validation : Valider;
      ab_Facture    : AfficheFacture;
      ab_RecupMemo  : RecupMemo; {Appel de la vignette Memo}
      ab_ChargeMemo : ChargeMemo; {Récupération du Mémo pour la vignette Memo}
    end;

    ddWriteLN('CPVIGNETTEVALIDBAP : Fin des traitements BAP');
  except
    on E : Exception do
      MessageErreur := 'Erreur lors du traitement des données avec le message :'#13 + E.Message;
  end;
end;

{13/06/07 : Ajout des left join sur Utilisateur pour afficher les libellés
{---------------------------------------------------------------------------------------}
procedure TObjetValidBap.GetClauseWhere;
{---------------------------------------------------------------------------------------}
begin
  ddWriteLN('CPVIGNETTEVALIDBAP : User = ' + CodeUser);
  ClauseWhere :=
    'SELECT BAP_JOURNAL, BAP_NUMEROPIECE, E_AUXILIAIRE, T_LIBELLE, E_REFINTERNE, E_LIBELLE, ' +
    'E_MODEPAIE, E_DATEECHEANCE, E_MONTANT, BAP_EXERCICE, BAP_NUMEROORDRE, BAP_DATECOMPTABLE, ' +
    'BAP_TIERSPAYEUR, BAP_CLEFFACTURE, U1.US_LIBELLE BAP_VISEUR1, U2.US_LIBELLE BAP_VISEUR2, ' +
    'U3.US_LIBELLE PREV_VISEUR, U4.US_LIBELLE NEXT_VISEUR, BAP_ECHEANCEBAP, BAP_STATUTBAP, BAP_BLOCNOTE, ' + {24/01/07 : FQ 10009 : }
    'BAP_IDGED, U5.US_EMAIL US_EMAIL FROM CPBAPECRITURE ';
  ClauseWhere := ClauseWhere +
    'LEFT JOIN UTILISAT U1 ON U1.US_UTILISATEUR = BAP_VISEUR1 ' +
    'LEFT JOIN UTILISAT U2 ON U2.US_UTILISATEUR = BAP_VISEUR2 ' +
    'LEFT JOIN UTILISAT U3 ON U3.US_UTILISATEUR = PREV_VISEUR ' +
    'LEFT JOIN UTILISAT U4 ON U4.US_UTILISATEUR = NEXT_VISEUR ' +
    'LEFT JOIN UTILISAT U5 ON U5.US_UTILISATEUR = BAP_CREATEUR ' +
    'WHERE (BAP_VISEUR1 = "' + CodeUser + '" OR BAP_VISEUR2 = "' + CodeUser + '") ' +
    {24/01/07 : FQ 10009 : Modification de la vue
    ' AND NOT (BAP_STATUTBAP IN ("DEF", "VAL")) AND E_TYPEMVT = "TTC"';}
    ' AND NOT (BAP_STATUTBAP IN ("DEF", "VAL"))';
  if GetString('CBFILTREBAP', False) <> '' then
    ClauseWhere := ClauseWhere + 'AND BAP_STATUTBAP = "' + GetString('CBFILTREBAP', False) + '" ';

  ClauseWhere := ClauseWhere + 'ORDER BY E_DATEECHEANCE DESC ,E_AUXILIAIRE';
  ddWriteLN('CPVIGNETTEVALIDBAP : Requête = ' + ClauseWhere);
end;

{---------------------------------------------------------------------------------------}
procedure TObjetValidBap.TraiteParam(Param : string);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  inherited;
  {Reste des paramètres, en attendant d'autres variables}
  s := ReadTokenSt(FParam);
       if s = 'ETAT' then TypeAction := ab_Etat
  else if s = 'VALI' then TypeAction := ab_Validation
  else if s = 'FACT' then TypeAction := ab_Facture
  else if s = 'MEMO' then TypeAction := ab_ChargeMemo
  else if s = 'DETA' then TypeAction := ab_RecupMemo
  else TypeAction := ab_mul;
end;

{---------------------------------------------------------------------------------------}
procedure TObjetValidBap.ChargerMul;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  ddwriteln('PGIVignetteValidBap : ChargerMul');
  Q := OpenSelect(ClauseWhere);
  TobDonnees.LoadDetailDB('****', '', '', Q, False);
  ddwriteln('PGIVignetteValidBap : Nb enreg = ' + IntToStr(TobDonnees.Detail.Count));
  Ferme(Q);
  if (TobDonnees.Detail.Count = 0) and (MessageErreur = '') then
    MessageErreur := TraduireMemoire('Vous n''avez pas de bons à payer à viser');
end;

{---------------------------------------------------------------------------------------}
procedure TObjetValidBap.LanceEtat;
{---------------------------------------------------------------------------------------}
begin
  {Pour ne pas lancer le SetInterface}
  SansInterface := True;
  ddwriteln('PGIVignetteValidBap : LanceEtat');

  RemplitTobDonneesEtat;

  if TobDonnees.Detail.Count = 0 then begin
    MessageErreur := TraduireMemoire('Il n''y a pas de bons à payer à traiter');
    Exit;
  end;

  FTip := 'E';
  FNat := 'VCP';
  if EnCroisAxe then FModele := 'VCA'
                else FModele := 'VBA';
  FLanguePrinc := 'FRA';
  Rapport;
end;

{21/06/07 : FQ 10019 : Ajout de la référence interne
{---------------------------------------------------------------------------------------}
procedure TObjetValidBap.RemplitTobDonneesEtat;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  F : TOB;
  n : Integer;
  Q : TQuery;
  r : string;
begin
  {On commence par remplir la TobSelection}
  AvecSelection := True;
  GetInterface('FListe');
  AvecSelection := False;
  if TobSelection.Detail.Count = 0 then
    MessageErreur := TraduireMemoire('Veuillez sélectionner au moins une ligne.')

  else begin
    EnCroisAxe := GetParamsocDossierSecur('SO_CROISAXE', False, Dossier);
    if EnCroisAxe then
      r := 'SELECT Y_CREDIT, Y_DEBIT, Y_GENERAL, Y_NUMLIGNE, Y_NUMVENTIL, Y_POURCENTAGE, ' +
           'Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5, Y_TOTALECRITURE FROM ANALYTIQ '
    else
      r := 'SELECT Y_AXE, Y_CREDIT, Y_DEBIT, Y_GENERAL, Y_NUMLIGNE, Y_NUMVENTIL, Y_POURCENTAGE, ' +
           'Y_SECTION, Y_TOTALECRITURE, S_LIBELLE FROM ANALYTIQ LEFT JOIN SECTION ON S_SECTION =' +
           'Y_SECTION AND S_AXE = Y_AXE ';

    for n := 0 to TobSelection.Detail.Count - 1 do begin
      F := TobSelection.Detail[n];
      Q := OpenSelect(r + 'WHERE Y_JOURNAL = "' + F.GetString('BAP_JOURNAL') + '" AND Y_EXERCICE = "' + F.GetString('BAP_EXERCICE')
                        + '" AND Y_NUMEROPIECE = ' + F.GetString('BAP_NUMEROPIECE') + ' AND Y_DATECOMPTABLE = "'
                        + UsDateTime(F.GetDateTime('BAP_DATECOMPTABLE')) + '"');
      try
        while not Q.EOF do begin
          T := TOB.Create('****', TobDonnees, -1);
          T.AddChampSupValeur('BAP_DATECOMPTABLE', F.GetDateTime('BAP_DATECOMPTABLE'));
          T.AddChampSupValeur('BAP_ECHEANCEBAP'  , F.GetDateTime('BAP_ECHEANCEBAP'  ));
          T.AddChampSupValeur('BAP_NUMEROPIECE'  , F.GetString('BAP_NUMEROPIECE'));
          T.AddChampSupValeur('E_AUXILIAIRE'     , F.GetString('E_AUXILIAIRE'   ));
          T.AddChampSupValeur('BAP_JOURNAL'      , F.GetString('BAP_JOURNAL'    ));
          T.AddChampSupValeur('E_LIBELLE'        , F.GetString('E_LIBELLE'      ));
          T.AddChampSupValeur('BAP_BLOCNOTE'     , F.GetString('BAP_BLOCNOTE'   ));
          T.AddChampSupValeur('T_LIBELLE'        , F.GetString('T_LIBELLE'      ));
          T.AddChampSupValeur('BAP_TIERSPAYEUR'  , F.GetString('BAP_TIERSPAYEUR'));
          T.AddChampSupValeur('BAP_CLEFFACTURE'  , F.GetString('BAP_CLEFFACTURE'));
          T.AddChampSupValeur('E_DATEECHEANCE'   , F.GetString('E_DATEECHEANCE' ));
          T.AddChampSupValeur('E_MONTANT'        , F.GetString('E_MONTANT'      ));
          T.AddChampSupValeur('E_REFINTERNE'     , F.GetString('E_REFINTERNE'   ));
          T.AddChampSupValeur('Y_TOTALECRITURE', Q.FindField('Y_TOTALECRITURE').AsString);
          T.AddChampSupValeur('Y_AXE'          , Q.FindField('Y_AXE'          ).AsString);
          T.AddChampSupValeur('Y_CREDIT'       , Q.FindField('Y_CREDIT'       ).AsFloat);
          T.AddChampSupValeur('Y_DEBIT'        , Q.FindField('Y_DEBIT'        ).AsFloat);
          T.AddChampSupValeur('Y_GENERAL'      , Q.FindField('Y_GENERAL'      ).AsString);
          T.AddChampSupValeur('Y_NUMLIGNE'     , Q.FindField('Y_NUMLIGNE'     ).AsInteger);
          T.AddChampSupValeur('Y_NUMVENTIL'    , Q.FindField('Y_NUMVENTIL'    ).AsInteger);
          T.AddChampSupValeur('Y_POURCENTAGE'  , Q.FindField('Y_POURCENTAGE'  ).AsFloat);
          if EnCroisAxe then begin
            T.AddChampSupValeur('Y_SOUSPLAN1', Q.FindField('Y_SOUSPLAN1').AsString);
            T.AddChampSupValeur('Y_SOUSPLAN2', Q.FindField('Y_SOUSPLAN2').AsString);
            T.AddChampSupValeur('Y_SOUSPLAN3', Q.FindField('Y_SOUSPLAN3').AsString);
            T.AddChampSupValeur('Y_SOUSPLAN4', Q.FindField('Y_SOUSPLAN4').AsString);
            T.AddChampSupValeur('Y_SOUSPLAN5', Q.FindField('Y_SOUSPLAN5').AsString);
          end
          else begin
            T.AddChampSupValeur('Y_SECTION', Q.FindField('Y_SECTION').AsString);
            T.AddChampSupValeur('Y_LIBELLE', Q.FindField('S_LIBELLE').AsString);
          end;
          Q.Next;
        end;
      finally
        Ferme(Q);
      end;
    end;
  end;
//  TobDonnees.SaveToXmlFile('C:\Draw.xml', True);
end;

{20/06/07 : Mise à jour du champ BAP_VISEUR (qui contient le code du viseur effectif de l'étape)
{---------------------------------------------------------------------------------------}
procedure TObjetValidBap.Valider;
{---------------------------------------------------------------------------------------}
const
  {Création de la constante que l'on utilise dans CCMP et CCS5, mais que je ne peux utiliser ici, 
   faute de pouvoir ajouter ULibBonsAPayer dans le projet. En cas de recherche dans les fichiers ...}
  sbap_Valide = 'VAL';

var
  F : TOB;
  M : TOB;
  n : Integer;
  R : string;
  Statut : string;
  User   : string;
  ValSql : string;
  ValOk  : Boolean;
begin
  // FQ 10036 (portail) Message bloquant si aucun statut sélectionné
  if GetString('CBSTATUTBAP', False) = '' then begin
    MessageErreur := TraduireMemoire('Veuillez sélectionner un statut.') ;
    {On recharge le mul}
    ChargerMul;
    Exit ;
  end;

  {On commence par remplir la TobSelection}
  AvecSelection := True;
  GetInterface('FListe');
  AvecSelection := False;

  {Traitement proprement dit}
  if TobSelection.Detail.Count = 0 then
    MessageErreur := TraduireMemoire('Veuillez sélectionner au moins une ligne.')

  else begin
    Statut := GetString('CBSTATUTBAP', False);
    User   := TobRequest.GetValue('USERNAME');
    ValOk  := Statut = sbap_Valide;
    ValSql := '';
    try
      for n := 0 to TobSelection.Detail.Count - 1 do begin
        F := TobSelection.Detail[n];
        R := GetControlValue('BLOCNOTE') + #13#10#13#10 + F.GetValue('BAP_BLOCNOTE');

        {JP 15/05/08 : FQ 22700 : envoie d'un mail a l'initiateur}
        if (Statut <> 'VAL') and (Statut <> 'ENC') then EnvoieMail(F);
        
        {26/06/07 : On ne peut mettre a jour les blobs par requête qu'avec SQLServer.
                    Le problème pour Oracle et DB2, c'est que la TOB n'est pas Multi sociétés
         02/08/07 : Comme sous SQLServer, les #10#13 sautent, je passe aussi par TOB}
        M := TOB.Create('CPBONSAPAYER', nil, -1);
        try
          (*
          if isMssql then begin
            {20/06/07 : Mise à jour du champ BAP_VISEUR qui contient le code du viseur effectif de l'étape,
                        donc uniquement si le l'on valide l'étape}
            if ValOk then ValSql := 'BAP_VISEUR = "' + User + '", ';

            R := 'UPDATE ' + GetTableDossier(Dossier, 'CPBONSAPAYER') + ' SET BAP_STATUTBAP = "' + Statut + '", ' + ValSql +
                 'BAP_BLOCNOTE = "' + R + '", BAP_DATEMODIF = "' + UsDateTime(Now) + '", BAP_MODIFICATEUR = "' + User + '" ' +
                 'WHERE BAP_JOURNAL = "' + F.GetString('BAP_JOURNAL') + '" AND BAP_EXERCICE = "' + F.GetString('BAP_EXERCICE')
                     + '" AND BAP_NUMEROPIECE = ' + F.GetString('BAP_NUMEROPIECE') + ' AND BAP_DATECOMPTABLE = "'
                     + UsDateTime(F.GetDateTime('BAP_DATECOMPTABLE')) + '" AND BAP_NUMEROORDRE = ' + F.GetString('BAP_NUMEROORDRE');
            ddwriteln('PGIVignetteValidBap : MAJ = ' + R);
            ExecuteSQL(R);
          end
          else begin
            *)
            M.PutValue('BAP_JOURNAL', F.GetString('BAP_JOURNAL'));
            M.PutValue('BAP_EXERCICE', F.GetString('BAP_EXERCICE'));
            M.PutValue('BAP_DATECOMPTABLE', F.GetDateTime('BAP_DATECOMPTABLE'));
            M.PutValue('BAP_NUMEROPIECE', F.GetString('BAP_NUMEROPIECE'));
            M.PutValue('BAP_NUMEROORDRE', F.GetString('BAP_NUMEROORDRE'));

            M.LoadDB(True);
            //M.PutValue('BAP_BLOCNOTE', R);
            M.PutValue('BAP_BLOCNOTE', GetControlValue('BLOCNOTE') + #13#10#13#10 + GetRTFStringText(F.GetValue('BAP_BLOCNOTE')));
            {20/06/07 : Mise à jour du champ BAP_VISEUR qui contient le code du viseur effectif de l'étape,
                        donc uniquement si le l'on valide l'étape}
            if ValOk then M.PutValue('BAP_VISEUR', User);
            M.PutValue('BAP_STATUTBAP', Statut);
            M.PutValue('BAP_MODIFICATEUR', User);
            M.PutValue('BAP_DATEMODIF', Now);
            M.UpdateDB();
//          end;

        finally
          FreeAndNil(M);
        end;
      end;
    except
      on E : Exception do begin
        MessageErreur := TraduireMemoire('Erreur lors de ma validation des bons à payer avec le message') + ' :'#13 + E.Message;
        DDWriteln(E.Message);
      end;
    end;
  end;

  {On vide les zones}
  SetControlValue('BLOCNOTE', '');
  SetControlValue('CBSTATUTBAP', '');
  {On recharge le mul}
  ChargerMul;
end;

{24/01/07 : Affichage de la facture stocké dans la GED
{---------------------------------------------------------------------------------------}
procedure TObjetValidBap.AfficheFacture;
{---------------------------------------------------------------------------------------}
var
  F : TOB;
begin
  inherited;
  {On recharge le mul}
//  ChargerMul;
  SansInterface := True;
  {On commence par remplir la TobSelection}
  AvecSelection := True;
  GetInterface('FListe');
  AvecSelection := False;
  WarningOk := False;

  if TobSelection.Detail.Count = 0 then
    MessageErreur := TraduireMemoire('Veuillez sélectionner au moins une ligne.')
  else begin
    F := TobSelection.Detail[0];
    RecupDocument(F.GetString('BAP_IDGED'));
    WarningOk := True;
    {JP 11/02/08 : Pour gérer le Message d'avertissement s'il n'y pas de facture attachée.
    Sans le SetInterface, je serais obligé d'afficher un message d'erreur ... c'est moyen !}
    SetInterface;
    WarningOk := False;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjetValidBap.RecupMemo;
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
procedure TObjetValidBap.ChargeMemo;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  T := GetLinkedValue('BAP_BLOCNOTE');
  SetControlValue('MMMEMO', GetRTFStringText(T.GetValue('VALUE')));
end;

{JP 15/05/08 : FQ 22700 : envoie d'un mail a l'initiateur
{---------------------------------------------------------------------------------------}
procedure TObjetValidBap.EnvoieMail(Fact : TOB);
{---------------------------------------------------------------------------------------}
var
  sBody : string;
  sTo   : string;
begin
  if sServer = '' then begin
    with TRegistry.Create do begin
      RootKey := HKEY_LOCAL_MACHINE;
      if OpenKey('SOFTWARE\CEGID_RM\PgiService\Journal', False) then begin
        if ValueExists('smtpServer') then sServer := ReadString('smtpServer');
        if ValueExists('smtpFrom')   then sFrom   := ReadString('smtpFrom');
        CloseKey;
      end;
      Free;
    end;
  end;

  if sServer = '' then Exit;
  if sFrom   = '' then Exit;
  sTo := Fact.GetString('US_EMAIL');
  if sTo = '' then Exit;
  sBody := 'Le BAP suivant a été refusé ou bloqué :'#13 +
           'Journal = ' + Fact.GetString('BAP_JOURNAL') + #13 +
           'Exercice = ' + Fact.GetString('BAP_EXERCICE') + #13 +
           'DateComptable = ' + Fact.GetString('BAP_DATECOMPTABLE') + #13 +
           'Numéro de pièce = ' + Fact.GetString('BAP_NUMEROPIECE') + #13 +
           'Montant = ' + Fact.GetString('E_MONTANT');
  SendMailSmtp(sServer, sFrom, sTo, 'BAP bloqué ou refusé', sBody)
end;

end.





