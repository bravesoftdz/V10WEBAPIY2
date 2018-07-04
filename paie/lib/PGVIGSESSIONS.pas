unit PGVIGSESSIONS;
{***********UNITE*************************************************
Auteur  ...... : FLO
Créé le ...... : 18/06/2007
Modifié le ... :
Description .. : Saisie d'une session de formation et des présents
               : Vignette : PG_VIG_SESSIONS
               : Tablette :
               : Table    : SESSIONSTAGE, FORMATIONS
Mots clefs ... : VIGNETTE;SESSION;PRESENCE
******************************************************************
PT1  | 18/01/2008 | FLO | Liste des inscrits à présent dans un panel propre,
     |            |     | ajout de critères pour certaines requêtes,
     |            |     | une session créée par la vignette ne doit pas être clôturée par défaut,
     |            |     | externalisation de l'inscription
PT2  | 01/02/2008 | FLO | Ajout de la colonne Clôture + modif gestion des n° de dossier
PT3  | 14/03/2008 | FLO | Ajout de la coche "Affiche/cacher les salariés déjà inscrits"
PT4  | 14/03/2008 | FLO | Gestion multi-dossier formation
PT5  | 18/03/2008 | FLO | Chargement du nombre d'heures par défaut à la sélection d'un stage
PT6  | 19/03/2008 | FLO | Prise en compte du type utilisateur
PT7  | 03/04/2008 | FLO | Gestion des groupes de travail
PT8  | 29/04/2008 | FLO | Prise en compte des salariés non sortis
PT9  | 15/05/2008 | FLO | Rafraîchissement des vignettes associés
}

interface

uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_SESSIONS = class (TGAVignettePaie)
  private
    Function  DonneMillesime : String;
    Procedure CreeSession;

    Procedure Valide;
    Procedure ValideStagiaires;

    Procedure AfficheListeSessions;
    Procedure AfficheListeStagiaires;
    Procedure AfficheListeSalaries;

    Procedure VisuSessions;
    Procedure VisuStagiaires;
    Procedure VisuCreation;
    Procedure VisuAjoutSal;
    
    Procedure InitialiseSaisie;
    Procedure InitialiseSaisieSal;

    Procedure ChargeParamStage; //PT5

  protected
    procedure RecupDonnees                                       ; override;
    function  GetInterface   (NomGrille: string = ''): Boolean   ; override;
    procedure GetClauseWhere                                     ; override;
    procedure DrawGrid       (Grille: string)                    ; override;
    function  SetInterface : Boolean                             ; override;
  end;

implementation
uses
  SysUtils,
  PGVIGUTIL,
  PGOutilsFormation,
  HEnt1,
  PGOutils,
  PgCalendrier;

{-----Récupération des controls/valeurs depuis la TobRequest----------------------------}

function PG_VIG_SESSIONS.GetInterface (NomGrille: string): Boolean;
begin

     { Cas particulier de l'inscription de stagiaires :
       La grille étant en multi-sélection, il faut pouvoir utiliser la TobSelection. Pour ce faire,
       on positionne le booléen AvecSelection à vrai et on appelle GetInterface avec le nom de la grille.
       En retour, la Tob est alimentée. On repositionne alors le booléen à faux pour éviter des problèmes. }
     If Fparam = 'VALIDERAJOUT' Then
     Begin
          AvecSelection := True;
          Result := inherited GetInterface ('GRSALARIES');
          AvecSelection := False;
     End
     Else
          Result := inherited GetInterface ('');

     // Liste des sessions
     If (Fparam = '') Or (Fparam='PERPLUS') Or (Fparam='PERMOINS') Or (Fparam = 'FERMERSTAGIAIRES') Then
     Begin
          // Pour les changements de périodes, on n'effectue aucune action particulière car c'est le RecupDonnees qui fait tout
          VisuSessions;
     End
     Else
     // Gestion de la création d'une session ou de l'ajout de stagiaires
     Begin
          // Mise à jour du warning à False pour ne pas afficher le bandeau "La sélection ne renvoie aucun enregistrement"
          WarningOk := False;

          If Fparam = 'CREERSESSION'  Then
          Begin
               // Affichage du groupe de création de session
               VisuCreation;
          End;

          If Fparam = 'VALIDERCREATION' Then
          Begin
               Valide;
               // Initialisation des champs de saisie pour éviter de les retourner au serveur à chaque fois
               InitialiseSaisie;
               VisuSessions;
          End;

          If Fparam = 'ANNULERCREATION' Then
          Begin
               // Initialisation des champs de saisie pour éviter de les retourner au serveur à chaque fois
               InitialiseSaisie;
               VisuSessions;
          End;

          // Liste des stagiaires
          If (Fparam = 'VOIRSTAGIAIRES') Then
          Begin
               VisuStagiaires
          End;

          If (Fparam = 'AJOUTERSTAGIAIRES') Or (FParam = 'AJOUTERSTAGIAIRESREFRESH') Then //PT3
          Begin
               VisuAjoutSal;
               If Fparam = 'AJOUTERSTAGIAIRES' Then SetControlValue('CKDEJAINSCRITS', True); //PT3
          End;

          If Fparam = 'VALIDERAJOUT' Then
          Begin
               ValideStagiaires;
               // Initialisation des champs de saisie pour éviter de les retourner au serveur à chaque fois
               InitialiseSaisieSal;
               VisuSessions;
          End;

          If Fparam = 'ANNULERAJOUT' Then
          Begin
               // Initialisation des champs de saisie pour éviter de les retourner au serveur à chaque fois
               InitialiseSaisieSal;
               VisuSessions;
          End;
     End;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_SESSIONS.RecupDonnees;
begin
  inherited;

     { Si aucun paramètre, on est à l'affichage des sessions.
       Ce cas est possible pour l'initialisation, les changements de périodes et de prédéfini }
     If (Fparam = '') Or (Fparam='PERMOINS') Or (Fparam= 'PERPLUS') Or (Fparam='FERMERSTAGIAIRES') Or
        (Fparam='ANNULERCREATION') Or (Fparam='VALIDERCREATION') Or (Fparam = 'ANNULERAJOUT') Or (Fparam='VALIDERAJOUT') Then
          AfficheListeSessions;

     If (Fparam = 'VOIRSTAGIAIRES') Then
          AfficheListeStagiaires;

     If (Fparam = 'AJOUTERSTAGIAIRES') Or ((Fparam = 'AJOUTERSTAGIAIRESREFRESH')) Then
          AfficheListeSalaries;

     If (FParam = 'CHANGESTAGE') Then
          ChargeParamStage;

     // On ne charge la liste des stages qu'à l'initialisation de la vignette afin d'optimiser les chargements
     If (Fparam = '') Then
          ChargeStages(Self, 'STAGE');
end;

{-----Critères de la requète -----------------------------------------------------------}

procedure PG_VIG_SESSIONS.GetClauseWhere;
begin
  inherited;
end;


{-----Affectation des controls/valeurs dans la TobResponse -----------------------------}

function PG_VIG_SESSIONS.SetInterface: Boolean;
begin
 inherited SetInterface;

     result:=true;

     If (Fparam = '') Or (Fparam='PERMOINS') Or (Fparam= 'PERPLUS') Or (Fparam='FERMERSTAGIAIRES') Or
        (Fparam='ANNULERCREATION') Or (Fparam='VALIDERCREATION') Or (Fparam = 'ANNULERAJOUT') Or (Fparam='VALIDERAJOUT') Then
     Begin
          PutGridDetail('GRSESSIONS', TobDonnees);

          // Dissimulation des champs de la grille qui ne sont là que pour des raisons techniques
          SetVisibleCol('GRSESSIONS', 'CODESTAGE', False);
          SetVisibleCol('GRSESSIONS', 'ORDRE',     False);
          SetVisibleCol('GRSESSIONS', 'MILLESIME', False);

          // Format des colonnes ici car ça ne passe pas par le DrawGrid
          SetFormatCol('GRSESSIONS', 'DATEDEBUT', 'C.0 ---');
          SetFormatCol('GRSESSIONS', 'DATEFIN',   'C.0 ---');
          SetFormatCol('GRSESSIONS', 'CLOTURE',   'C.0 ---'); //PT2
          SetFormatCol('GRSESSIONS', 'EFFECTUE',  'C.0 ---');
          SetFormatCol('GRSESSIONS', 'PREDEFINI', 'C.0 ---');
          SetFormatCol('GRSESSIONS', 'NATURE',    'C.0 ---');

          SetWidthCol ('GRSESSIONS', 'FORMATION', 90);
          SetWidthCol ('GRSESSIONS', 'SESSION',   90);
          SetWidthCol ('GRSESSIONS', 'DATEDEBUT', 60);
          SetWidthCol ('GRSESSIONS', 'DATEFIN',   60);
          SetWidthCol ('GRSESSIONS', 'NATURE',    50);
          SetWidthCol ('GRSESSIONS', 'PREDEFINI', 50);
          SetWidthCol ('GRSESSIONS', 'EFFECTUE',  50);
          SetWidthCol ('GRSESSIONS', 'CLOTURE',   50);  //PT2
     End;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_SESSIONS.DrawGrid (Grille: string);
begin
  inherited;
end;

{-----Affichage de la liste des sessions ------------------------------------------------}

procedure PG_VIG_SESSIONS.VisuSessions;
begin
     { Etant donné que la liste des sessions est la liste principale de la vignette,
       on gère tous les champs et tous les boutons car on peut revenir sur la visu
       des sessions à partir du bouton ancètre Rafraîchir, accessible de n'importe où }
        
     // Cache la grille et affiche le groupe
     SetControlVisible ('GRSESSIONS',    True);
     //SetControlVisible ('GRSTAGIAIRES',  False);
     SetControlVisible ('PANINSCRITS',   False);  //PT1
     SetControlVisible ('PANCREATION',   False);
     SetControlVisible ('PANAJOUT',      False);

     // Change la barre d'état
     SetControlVisible ('TPERIODE',      True);
     SetControlVisible ('PERIODE',       True);
     SetControlVisible ('TBPLUS',        True);
     SetControlVisible ('TBMOINS',       True);
     SetControlVisible ('N1',            True);
     SetControlVisible ('TPREDEFINI',    True);
     SetControlVisible ('PREDEFINI',     True);
     SetControlVisible ('BNOUVEAU',      True);
     SetControlVisible ('BVALIDER',      False);
     SetControlVisible ('BANNULER',      False);
     SetControlVisible ('BFERMERSTAG',   False);
     SetControlVisible ('BVALIDERAJOUT', False);
     SetControlVisible ('BANNULERAJOUT', False);
     SetControlVisible ('BAJOUTER',      False);
     SetControlVisible ('TTYPEPLAN',     False);  //PT1
     SetControlVisible ('TYPEPLAN',      False);
     SetControlVisible ('CKDEJAINSCRITS',False); //PT3
end;

{-----Affichage du groupe de champs nécessaires à la création d'une session ------------}

procedure PG_VIG_SESSIONS.VisuCreation;
begin
     // Cache la grille et affiche le groupe
     SetControlVisible ('GRSESSIONS',    False);
     SetControlVisible ('PANCREATION',   True);
     //PT4 - Début
     If PGBundleInscFormation Then
     Begin
     	SetControlVisible ('TNODOSSIER',    PGDroitMultiForm);
     	SetControlVisible ('NODOSSIER',     PGDroitMultiForm);
        //PT7
        If PGDroitMultiForm Then ChargeValeursCombo (Self, 'NODOSSIER', '', 'SELECT DOS_NODOSSIER AS VALUE,DOS_LIBELLE AS ITEM FROM DOSSIER WHERE '+GererCritereGroupeConfTous);
     End;
 	 //PT4 - Fin
     // Change la barre d'état
     SetControlVisible ('TPERIODE',      False);
     SetControlVisible ('PERIODE',       False);
     SetControlVisible ('TBPLUS',        False);
     SetControlVisible ('TBMOINS',       False);
     SetControlVisible ('N1',            False);
     SetControlVisible ('TPREDEFINI',    False);
     SetControlVisible ('PREDEFINI',     False);
     SetControlVisible ('BNOUVEAU',      False);
     SetControlVisible ('BVALIDER',      True);
     SetControlVisible ('BANNULER',      True);

     // Initialisation des champs
     InitialiseSaisie;
end;


{-----Affichage de la liste des stagiaires ----------------------------------------------}

procedure PG_VIG_SESSIONS.VisuStagiaires;
begin
     // Cache la grille sessions et affiche la grille des salariés inscrits
     SetControlVisible ('GRSESSIONS',    False);
     SetControlVisible ('PANAJOUT',      False);

     //SetControlVisible ('GRSTAGIAIRES',  True);
     SetControlVisible ('PANINSCRITS',   True);  //PT1

     // Change la barre d'état
     SetControlVisible ('TPERIODE',      False);
     SetControlVisible ('PERIODE',       False);
     SetControlVisible ('TBPLUS',        False);
     SetControlVisible ('TBMOINS',       False);
     SetControlVisible ('N1',            False);
     SetControlVisible ('TPREDEFINI',    False);
     SetControlVisible ('PREDEFINI',     False);
     SetControlVisible ('BNOUVEAU',      False);

     SetControlVisible ('BVALIDERAJOUT', False);
     SetControlVisible ('BANNULERAJOUT', False);
     SetControlVisible ('TTYPEPLAN',     False);  //PT1
     SetControlVisible ('TYPEPLAN',      False);
     SetControlVisible ('CKDEJAINSCRITS',False); //PT3
     SetControlVisible ('BFERMERSTAG',   True);
     SetControlVisible ('BAJOUTER',      True);
end;

{-----Affichage du groupe d'ajout de stagiaires -----------------------------------------}

procedure PG_VIG_SESSIONS.VisuAjoutSal;
begin
     // Cache la grille et affiche le groupe
     //SetControlVisible ('GRSTAGIAIRES',  False);
     SetControlVisible ('PANINSCRITS',   False);  //PT1
     SetControlVisible ('PANAJOUT',      True);

     // Change la barre d'état
     SetControlVisible ('BFERMERSTAG',   False);
     SetControlVisible ('BAJOUTER',      False);
     SetControlVisible ('BVALIDERAJOUT', True);
     SetControlVisible ('BANNULERAJOUT', True);
     SetControlVisible ('TTYPEPLAN',     True);  //PT1
     SetControlVisible ('TYPEPLAN',      True);
     SetControlVisible ('CKDEJAINSCRITS',True); //PT3

     SetControlValue('TLIBELLESESSION',  GetControlValue('MEMOLIBELLE'));
     SetControlValue('TYPEPLAN',         'PLF');  //PT1
end;

{-----Listing des sessions --------------------------------------------------------------}

procedure PG_VIG_SESSIONS.AfficheListeSessions;
var
  StSQL,SN1              : String;
  DataTob,T              : TOB;
  EnDateDu, EnDateAu     : TDatetime;
  i                      : Integer;
  Q                      : TQuery;
begin
  // Récupération de la période sélectionnée
  DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode,SN1);

  Try
    DataTob := TOB.Create('~TEMP', nil, -1);

    // Création de la requête
    StSQL := 'SELECT PST_LIBELLE,PST_LIBELLE1,PSS_CODESTAGE,PSS_ORDRE,PSS_MILLESIME,PSS_LIBELLE,PSS_DATEDEBUT,PSS_DATEFIN,PSS_NATUREFORM,PSS_EFFECTUE,PSS_CLOTUREINSC,PSS_PREDEFINI '+ //PT2
    'FROM SESSIONSTAGE LEFT JOIN STAGE ON PSS_CODESTAGE=PST_CODESTAGE AND PSS_MILLESIME=PST_MILLESIME '+
    'WHERE PSS_VOIRPORTAIL="X" AND PSS_PGTYPESESSION<>"SOS" AND PSS_DATEDEBUT <= "'+USDATETIME (EnDateAu)+'" AND PSS_DATEDEBUT >= "'+USDATETIME (EnDateDu)+'" ';  //PT1
    //PT4 - Début
    If PGBundleInscFormation Then 
    Begin
    	// Ici : pas de cas spécifique si PGDroitsMultiForm car cela permettrait d'inscrire n'importe qui sur n'importe quel dossier
        If PGDroitMultiForm Then
        	StSQL := StSQL + DossiersAInterroger (GetControlValue('PREDEFINI'), '', 'PSS', True, True)  //PT7
        Else
        	StSQL := StSQL + DossiersAInterroger (GetControlValue('PREDEFINI'), V_PGI.NoDossier, 'PSS', True, True);
    End
    Else
    	If GetControlValue('PREDEFINI') <> '' Then StSQL := StSQL + ' AND PSS_PREDEFINI="'+GetControlValue('PREDEFINI')+'"';
	//PT4 - Fin
    StSQL := StSQL + ' ORDER BY PST_LIBELLE,PSS_DATEDEBUT';

    // Lancement de la requête : pas de requête en cache car on peut en créer
    //DataTob := OpenSelectInCache (StSQL);
    //ConvertFieldValue(DataTob);
    Q := OpenSQL(StSQL, True);
    DataTob.LoadDetailDB('$DATA','','',Q,False);
    Ferme(Q);

    // Parcours des résultats afin de générer la liste des sessions
    For i:=0 To DataTob.Detail.Count-1 Do
    Begin
      T := TOB.Create('£REPONSE',TobDonnees,-1);

      T.AddChampSupValeur('FORMATION', DataTob.Detail[i].GetValue('PST_LIBELLE')+' '+DataTob.Detail[i].GetValue('PST_LIBELLE1'));
      T.AddChampSupValeur('SESSION',   DataTob.Detail[i].GetValue('PSS_LIBELLE'));
      T.AddChampSupValeur('DATEDEBUT', DateToStr(DataTob.Detail[i].GetValue('PSS_DATEDEBUT')));
      T.AddChampSupValeur('DATEFIN',   DateToStr(DataTob.Detail[i].GetValue('PSS_DATEFIN')));
      T.AddChampSupValeur('NATURE',    RechDom('PGNATUREFORM',DataTob.Detail[i].GetValue('PSS_NATUREFORM'),False));
      T.AddChampSupValeur('CLOTURE',   LibBool(DataTob.Detail[i].GetValue('PSS_CLOTUREINSC'))); //PT2
      T.AddChampSupValeur('EFFECTUE',  LibBool(DataTob.Detail[i].GetValue('PSS_EFFECTUE')));
      If DataTob.Detail[i].GetValue('PSS_PREDEFINI') = 'DOS' Then
          T.AddChampSupValeur('PREDEFINI','Dossier')
      Else
          T.AddChampSupValeur('PREDEFINI','Standard');
      T.AddChampSupValeur('CODESTAGE', DataTob.Detail[i].GetValue('PSS_CODESTAGE'));
      T.AddChampSupValeur('ORDRE',     DataTob.Detail[i].GetValue('PSS_ORDRE'));
      T.AddChampSupValeur('MILLESIME', DataTob.Detail[i].GetValue('PSS_MILLESIME'));
    End;

    // Dans le cas où il n'y a pas d'enregistrement, on force le rafraîchissement de la grille
    If (TobDonnees.Detail.Count = 0) Then
      PutGridDetail('GRSESSIONS', TobDonnees);

    // Réinitialisation des grilles pour ne pas avoir de problème de rafraîchissement
    T := TOB.Create('~TOBVIDE',nil,-1);
    PutGridDetail('GRSALARIES', T);
    PutGridDetail('GRSTAGIAIRES', T);

  Finally
    FreeAndNil (DataTob);
    FreeAndNil(t);
  End;
end;

{-----Affichage de la liste des stagiaires inscrits à la session ----------------------}

procedure PG_VIG_SESSIONS.AfficheListeStagiaires;
var
    DataTob,T                             : Tob;
    StSQL,Stage,Session,Millesime,Libelle : String;
    Cloture                               : Boolean;
    Q                                     : TQuery;
begin
   Try
     // Récupération dans le flux de la ligne séléctionnée dans la grille sous forme de TOB
     T := GetLinkedValue('GRSESSIONS');

     // Récupération de la session sélectionnée et de sa clé
     If (T <> Nil) And (T.Detail.Count > 0) Then
     Begin
          Stage     := T.Detail[0].GetValue('CODESTAGE');
          Session   := T.Detail[0].GetValue('ORDRE');
          Millesime := T.Detail[0].GetValue('MILLESIME');
          Libelle   := T.Detail[0].GetValue('SESSION');
          Cloture   := (T.Detail[0].GetValue('CLOTURE') = OUI);
     End
     Else
     Begin
          Stage := ''; Session := ''; Millesime := ''; Cloture := False;
     End;

     // Sauvegarde des valeurs dans les champs cachés de la vignette
     SetControlValue('ORDRE',     Session);
     SetControlValue('MILLESIME', Millesime);
     SetControlValue('CODESTAGE', Stage);
     SetControlValue('MEMOLIBELLE', Libelle);
     SetControlValue('TLIBELLESESSIONINSC',  Libelle);  //PT1

     // Chargement des données
     StSQL := 'SELECT PFO_SALARIE,PFO_NOMSALARIE,PFO_PRENOM FROM FORMATIONS '+
              'WHERE PFO_CODESTAGE="'+Stage+'" AND PFO_ORDRE="'+Session+'" AND PFO_MILLESIME="'+Millesime+'" ORDER BY PFO_NOMSALARIE,PFO_PRENOM' ;
     DataTob := Tob.Create('~TEMP', Nil, -1);
     // Pas de lecture en cache car les informations peuvent être modifiées
     //DataTob := OpenSelectInCache (StSQL);
     //ConvertFieldValue(DataTob);
     Q := OpenSQL(StSQL, True);
     DataTob.LoadDetailDB('$DATA','','',Q,False);
     Ferme(Q);

     PutGridDetail('GRSTAGIAIRES', DataTob);

     // Bouton des inscriptions : Non visible si inscriptions clôturées
     If Cloture Then SetControlVisible ('BAJOUTER', False);

     // Réinitialisation des grilles pour ne pas avoir de problème de rafraîchissement
     T := TOB.Create('~TOBVIDE',Nil,-1);
     PutGridDetail('GRSESSIONS', T);
     PutGridDetail('GRSALARIES', T);

  Finally
     If Assigned (DataTob) Then FreeAndNil(DataTob);
     FreeAndNil(t);
  End;
end;

{-----Affichage de la liste des salariés -----------------------------------------------}

procedure PG_VIG_SESSIONS.AfficheListeSalaries;
var
    T,DataTob : Tob;
    StSQL,Stage,Session,Millesime : String;
    Q         : TQuery;
    Where	  : String;
begin
   Try
     Session   := GetControlValue('ORDRE');
     Millesime := GetControlValue('MILLESIME');
     Stage     := GetControlValue('CODESTAGE');
     
     //PT3 - Début
     If GetControlValue('CKDEJAINSCRITS') = False Then
     	Where  := ' AND PFO_ORDRE="'+Session+'"'
     Else
     	Where := '';
     //PT3 - Fin

     // Chargement des données en filtrant les salariés déjà inscrits à la session
     StSQL := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM SALARIES LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE '+
              'WHERE ' + SalariesNonSortis + ' AND '+ //PT8
              AdaptByTypeResp(RecupTypeUtilisateur('FOR'),V_PGI.UserSalarie,False,'PSE',False) + //PT6 //PT8
              ' AND PSA_SALARIE NOT IN (SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_CODESTAGE="'+Stage+'" AND PFO_MILLESIME="'+Millesime+'"' + Where + ' AND PFO_ETATINSCFOR<>"REP")'+ //PT3
              ' ORDER BY PSA_LIBELLE,PSA_PRENOM' ;
              
	 StSQL := AdaptMultiDosForm (StSQL); //PT4
	 
     DataTob := Tob.Create('~TEMP', Nil, -1);
     Q := OpenSQL(StSQL, True);
     DataTob.LoadDetailDB('$DATA','','',Q,False);
     Ferme(Q);

     PutGridDetail('GRSALARIES', DataTob);

     // Réinitialisation des grilles pour ne pas avoir de problème de rafraîchissement
     T := TOB.Create('~TOBVIDE',Nil,-1);
     PutGridDetail('GRSESSIONS', T);
     PutGridDetail('GRSTAGIAIRES', T);

  Finally
     If Assigned (DataTob) Then FreeAndNil(DataTob);
     FreeAndNil(t);
  End;
end;

{-----Initialisation du groupe de saisie d'une session ---------------------------------}

procedure PG_VIG_SESSIONS.InitialiseSaisie;
begin
     SetControlValue('STAGE',          '');
     SetControlValue('LIBELLESESSION', '');
     SetControlValue('DATEDEBUT',      '');
     SetControlValue('DATEFIN',        '');
     SetControlValue('NBHEURES',       0);
     SetControlValue('NODOSSIER',      '');

     // Réinitialisation des valeurs sauvegardées pour ne pas générer de problème
     SetControlValue('ORDRE',     '');
     SetControlValue('MILLESIME', '');
     SetControlValue('CODESTAGE', '');
     SetControlValue('MEMOLIBELLE', '');

end;

{-----Initialisation du groupe de saisie des stagiaires à une session ------------------}

procedure PG_VIG_SESSIONS.InitialiseSaisieSal;
Begin
     SetControlValue('TLIBELLESESSION', '');
end;

{-----Retourne le millésime correspondant à un stage prévu -----------------------------}

Function PG_VIG_SESSIONS.DonneMillesime : String;
var Q         : TQuery;
    Millesime : String;
    DD,DF     : TDateTime;
begin
     Millesime := '0000';

     Try
          DD := StrToDate(GetControlValue('DATEDEBUT'));
          DF := StrToDate(GetControlValue('DATEFIN'));

          // Recherche d'un stage existant au prévisionnel durant la période. Si trouvé, on utilise le millesime
          Q := OpenSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_ACTIF="X" AND PFE_CLOTURE="-" '+
                       'AND PFE_DATEDEBUT<="'+UsDateTime(DD)+'" AND PFE_DATEFIN>="'+UsDateTime(DF)+'"', True);
          If Not Q.Eof then
          begin
               Millesime := Q.FindField('PFE_MILLESIME').AsString;
               If Not ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_CODESTAGE="'+GetControlValue('STAGE')+'" AND PST_MILLESIME="'+Millesime+'"') Then
                    Millesime := '0000';
          end;
     Finally
          If Q <> Nil Then Ferme(Q);
     End;

     Result := Millesime;
end;

{-----Création ou modification de la session de stage-----------------------------------}

procedure PG_VIG_SESSIONS.CreeSession;
Var
     TSession       : TOB;
     Ordre,LeStage,Millesime  : String;
     Q              : TQuery;
     DD             : TDateTime;
begin
  Try
     // Création de la session
     TSession := TOB.Create('SESSIONSTAGE', Nil, -1);

     // Données de la vignette
     LeStage := GetControlValue('STAGE');
     TSession.PutValue('PSS_CODESTAGE', LeStage);
     DD := StrToDate(GetControlValue('DATEDEBUT'));
     TSession.PutValue('PSS_DATEDEBUT', DD);
     DD := StrToDate(GetControlValue('DATEFIN'));
     TSession.PutValue('PSS_DATEFIN', DD);
     TSession.PutValue('PSS_LIBELLE',   GetControlValue('LIBELLESESSION'));
     TSession.PutValue('PSS_DUREESTAGE',GetControlValue('NBHEURES'));

     Q := OpenSQL('SELECT MAX(PSS_ORDRE) AS ORDREMAX FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+LeStage+'"', True);
     If Not Q.Eof Then
     Begin
          If Q.FindField('ORDREMAX').AsInteger <> 0 Then
               Ordre := IntToStr(Q.FindField('ORDREMAX').AsInteger + 1)
          Else Ordre := '1';
     End
     Else
          Ordre := '1';
     Ferme(Q);
     TSession.PutValue('PSS_ORDRE',     Ordre);
     Millesime := DonneMillesime;
     TSession.PutValue('PSS_MILLESIME', Millesime);

     // Données reprises du stage
     Q := OpenSQL('SELECT PST_CENTREFORMGU,PST_LIEUFORM,PST_INCLUSDECL,PST_NATUREFORM,PST_ACTIONFORM,'+
                  'PST_TYPCONVFORM,PST_ORGCOLLECTPGU,PST_ORGCOLLECTSGU,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,'+
                  'PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8 FROM STAGE WHERE PST_CODESTAGE="'+LeStage+'" AND PST_MILLESIME="'+Millesime+'"',TRUE) ;
     if Not Q.EOF then
     begin
          TSession.PutValue('PSS_CENTREFORMGU',Q.FindField('PST_CENTREFORMGU').AsString);
          TSession.PutValue('PSS_LIEUFORM',Q.FindField('PST_LIEUFORM').AsString);
          TSession.PutValue('PSS_INCLUSDECL',Q.FindField('PST_INCLUSDECL').AsString);
          TSession.PutValue('PSS_NATUREFORM',Q.FindField('PST_NATUREFORM').AsString);
          TSession.PutValue('PSS_ACTIONFORM',Q.FindField('PST_ACTIONFORM').AsString);
          TSession.PutValue('PSS_TYPCONVFORM',Q.FindField('PST_TYPCONVFORM').AsString);
          TSession.PutValue('PSS_ORGCOLLECTPGU',Q.FindField('PST_ORGCOLLECTPGU').AsString);
          TSession.PutValue('PSS_ORGCOLLECTSGU',Q.FindField('PST_ORGCOLLECTSGU').AsString);
          TSession.PutValue('PSS_FORMATION1',Q.FindField('PST_FORMATION1').AsString);
          TSession.PutValue('PSS_FORMATION2',Q.FindField('PST_FORMATION2').AsString);
          TSession.PutValue('PSS_FORMATION3',Q.FindField('PST_FORMATION3').AsString);
          TSession.PutValue('PSS_FORMATION4',Q.FindField('PST_FORMATION4').AsString);
          TSession.PutValue('PSS_FORMATION5',Q.FindField('PST_FORMATION5').AsString);
          TSession.PutValue('PSS_FORMATION6',Q.FindField('PST_FORMATION6').AsString);
          TSession.PutValue('PSS_FORMATION7',Q.FindField('PST_FORMATION7').AsString);
          TSession.PutValue('PSS_FORMATION8',Q.FindField('PST_FORMATION8').AsString);
     end;
     Ferme(Q);

     // Données statiques
     TSession.PutValue('PSS_CLOTUREINSC','-');  // Ne pas clôturer les inscriptions sinon il n'y aura jamais personne d'inscrit //PT1
     TSession.PutValue('PSS_EFFECTUE','X');
     TSession.PutValue('PSS_NUMSESSION',-1);

     If V_PGI.NoDossier <> '000000' Then  //PT2
        TSession.PutValue('PSS_PREDEFINI',     'DOS')
     Else
        TSession.PutValue('PSS_PREDEFINI',     'STD');
     //PT4 - Début
     //TSession.PutValue('PSS_NODOSSIER', V_PGI.NoDossier); 
     If PGBundleInscFormation Then
     Begin
     	If PGDroitMultiForm Then
     	Begin
     		If GetControlValue('NODOSSIER') = '' Then
     		Begin
     			MessageErreur := TraduireMemoire('La saisie du dossier est obligatoire.');
     			If Assigned(TSession) Then FreeAndNil(TSession);
     			Exit;
     		End
     		Else
     			TSession.PutValue('PSS_NODOSSIER', GetControlValue('NODOSSIER'));
     	End
     	Else
     		TSession.PutValue('PSS_NODOSSIER', GetNoDossierSalarie);
     End;
     //PT4 - Fin

     TSession.PutValue('PSS_IDSESSIONFOR',LeStage+Ordre);
     TSession.PutValue('PSS_NOSESSIONMULTI',-1);
     TSession.PutValue('PSS_PGTYPESESSION','AUC');
     TSession.PutValue('PSS_DEBUTDJ','MAT');
     TSession.PutValue('PSS_FINDJ','PAM');
//attention!!!!!!!!          TSession.PutValue('PSS_BLOCNOTE','');
     TSession.PutValue('PSS_VOIRPORTAIL', 'X');

     // En création, on positionne tous les champs comme modifiés
     TSession.SetAllModifie(True);

     // Insertion en base
     TSession.InsertOrUpdateDB;
  Finally
     If Assigned(TSession) Then FreeAndNil(TSession);
  End;
end;

{-----Validation des données de création de session ------------------------------------}

procedure PG_VIG_SESSIONS.Valide;
begin
     // Vérification des éléments saisis
     If (Trim(GetControlValue('STAGE')) = '') Or (Trim(GetControlValue('LIBELLESESSION'))='') Or
        (GetControlValue('DATEDEBUT')='/  /    ') Or (GetControlValue('DATEFIN')='/  /    ') Then
     Begin
        MessageErreur := TraduireMemoire('Toutes les zones sont obligatoires.');
        Exit;
     End;

     // Création ou Modification de la session de formation
     CreeSession;
end;

{-----Validation du rattachement des stagiaires à la session ---------------------------}

procedure PG_VIG_SESSIONS.ValideStagiaires;
Var  i : Integer;
     TRefresh : TOB;
begin
     If TobSelection.Detail.Count > 0 Then
     Begin
          // Vérification de l'état de la session
         //PT1 : code désactivé pour le moment
        (*TobSession := TOB.Create('~TEMP', nil, -1);
          TobSession := OpenSelectInCache ('SELECT * FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+Stage+'" AND PSS_ORDRE="'+Session+'" AND PSS_MILLESIME="'+Millesime+'"');
          ConvertFieldValue(TobSession);
          If TobSession.Detail.Count > 0 Then
          Begin
               // la session doit forcément être clôturée et effectuée
               If TobSession.Detail[0].GetValue('PSS_EFFECTUE') = '-' Then TobSession.Detail[0].PutValue('PSS_EFFECTUE', 'X');
               If TobSession.Detail[0].GetValue('PSS_CLOTUREINSC') = '-' Then TobSession.Detail[0].PutValue('PSS_CLOTUREINSC', 'X');
               TobSession.InsertOrUpdateDB;
          End;
          FreeAndNil(TobSession);*)

          // Inscription des salariés
          For i:=0 To TobSelection.Detail.Count - 1 Do
          Begin
               //PT1
               MessageErreur := InscritSalarieSession (GetControlValue('CODESTAGE'), GetControlValue('ORDRE'), GetControlValue('MILLESIME'), TobSelection.Detail[i].GetValue('PSA_SALARIE'), GetControlValue('TYPEPLAN'), RecupTypeUtilisateur('FOR'));
               If MessageErreur <> '' Then Exit;
          End;
          
		  //PT9
		  If MessageErreur = '' Then
		  Begin
		         // Rafraîchissement de la vignette Liste des stagiaires
		   		TRefresh := TOB.Create('T_REFRESH', TobResponse, -1);
	   			TRefresh.AddChampSupValeur('NAMES','PG_VIG_STAGIAIRES');
		  End;
     End
     Else
          MessageErreur := TraduireMemoire('Vous devez sélectionner au moins un salarié de la liste');
end;     
         
{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/03/2008 / PT5
Modifié le ... :   /  /
Description .. : Récupération du nombre d'heures de la formation
Mots clefs ... :
*****************************************************************}
Procedure PG_VIG_SESSIONS.ChargeParamStage ;
Var Q : TQuery;
Begin
    Try
        Q := OpenSQL('SELECT PST_DUREESTAGE FROM STAGE WHERE PST_CODESTAGE="'+GetControlValue('STAGE')+'" AND PST_MILLESIME="0000"', True);
        If Not Q.EOF Then
        Begin
            SetControlValue('NBHEURES', IntToStr(Q.FindField('PST_DUREESTAGE').AsInteger));
        End;
    Except
    End;
End;

end.

