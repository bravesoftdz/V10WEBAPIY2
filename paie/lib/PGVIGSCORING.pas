{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - FLO
Créé le ...... : 05/03/2007
Modifié le ... :   /  /
Description .. : Saisie du scoring formation
               : Vignette : PG_VIG_SCORING
Mots clefs ... :
*****************************************************************
}
unit PGVIGSCORING;

interface
uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_SCORING= class (TGAVignettePaie)
  protected
    procedure RecupDonnees;                                     override;
    function  GetInterface (NomGrille: string = ''): Boolean;   override;
    procedure GetClauseWhere;                                   override;
    procedure DrawGrid (Grille: string);                        override;
    function  SetInterface : Boolean ;                          override;
  private
    procedure AfficheListeFormations; 
    procedure AfficheQuestion (Init : Boolean);
    procedure Visualisation;
    procedure Initialisation;
    function  Validation : Boolean;
    function  RecupQuestions : TOB;
    procedure SupprimeQuestions;
    procedure SauveQuestions (TobQuestions : TOB);
 end;

implementation
uses
  SysUtils,
  PGVIGUTIL,
  rtfcounter,
  esession,
  ParamSoc,
  HEnt1,
  StrUtils,
  PGOutilsFormation,
  uToolsPlugin;
  
CONST ASK_LIST_NAME = 'LesQuestions';

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_SCORING.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
    Result := true;
    
    If (FParam = '') Or (FParam='ANNULER') Then // Or (FParam='PERMOINS') Or (FParam='PERPLUS') Then
    Begin
        // Initialisation des champs de saisie pour éviter de les retourner au serveur à chaque fois
        Initialisation;
        // Suppression de la liste des questions en mémoire dans le cas du retour à la liste suite aux réponses
		SupprimeQuestions;
    End
    Else If (FParam = 'SELECTION') Then
    Begin
		// Mise à jour du warning à False pour ne pas afficher le bandeau "La sélection ne renvoie aucun enregistrement"
        WarningOk := False;
    End
    Else If (FParam = 'VALIDER') Then
    Begin
		// Mise à jour du warning à False pour ne pas afficher le bandeau "La sélection ne renvoie aucun enregistrement"
        WarningOk := False;
        // Suite à validation du dernier élément (à ce moment Validation retourne True), on initialise Fparam pour forcer
        // le retour à la liste     
        If Validation = True Then FParam := 'FIN';
    End;

	// Change l'affichage de la vignette en fonction des cas    
    Visualisation;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_SCORING.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_SCORING.RecupDonnees;
begin
  inherited;

    If (FParam = '') Or (FParam='ANNULER') Or (FParam='PERMOINS') Or (FParam='PERPLUS') Then
        AfficheListeFormations
    Else If (FParam = 'SELECTION') Then
    	AfficheQuestion (True)
    Else If (FParam = 'VALIDER') Then
    	AfficheQuestion (False);
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_SCORING.DrawGrid (Grille: string);
begin
  inherited;

    SetFormatCol(Grille, 'DEBUTSTAGE',  'C.0 ---');
    SetFormatCol(Grille, 'FINSTAGE',    'C.0 ---');

	SetWidthCol (Grille, 'CODESTAGE',   -1);
    SetWidthCol (Grille, 'ORDRE',       -1);
	SetWidthCol (Grille, 'SESSION',     100);
    SetWidthCol (Grille, 'FORMATION',   100);
    SetWidthCol (Grille, 'DEBUTSTAGE',  50);
    SetWidthCol (Grille, 'FINSTAGE',    50);
end;

{-----Affectation des controls/valeurs dans la TobResponse -----------------------------}

function PG_VIG_SCORING.SetInterface: Boolean;
begin
  inherited SetInterface;
    result:=true;
end;

{-----Initialise les valeurs de la vignette susceptibles de changer --------------------}

Procedure PG_VIG_SCORING.Initialisation ;
Begin
		SetControlValue('TTITRE',		'');
		SetControlValue('TLIBTHEME',	'');
		SetControlValue('TQUESTION',	'');
	    SetControlValue('CBREPONSE',	'');
        SetControlValue('NUMQUESTION',	'0');
        SetControlValue('NBQUESTIONS',	'0');
        SetControlValue('CODESTAGE',	'');
        SetControlValue('ORDRE',		'');      
End;

{-----Affiche les éléments correspondant à l'action lancée -----------------------------}

procedure PG_VIG_SCORING.Visualisation;
Var Visu,Fin : Boolean;
begin
	If (FParam = 'SELECTION') Or (FParam = 'VALIDER') Or (FParam = 'FIN') Then
		Visu := False
	Else
		Visu := True;

	SetControlVisible ('TPERIODE', 		Visu);
	SetControlVisible ('PERIODE',  		Visu);
	SetControlVisible ('TBMOINS',  		Visu);
	SetControlVisible ('TBPLUS',   		Visu);
	SetControlVisible ('N1',       		Visu);
	SetControlVisible ('PANBOUTONS',  	Not Visu);

    If (FParam = 'FIN') Then Fin := True Else Fin := False;
	
	SetControlVisible ('PANFIN',  		(Not Visu) And Fin);
	SetControlVisible ('PANSAISIE', 	(Not Visu) And (Not Fin));
	SetControlVisible ('BVALIDER', 		(Not Visu) And (Not Fin));
	SetControlVisible ('BANNULER', 		(Not Visu) And (Not Fin));
	SetControlVisible ('BFIN', 		    (Not Visu) And Fin);
end;

{-----Affiche les formations déjà réalisées par le salarié -----------------------------}

procedure PG_VIG_SCORING.AfficheListeFormations;
var
  StSQL,sN1              : String;
  DataTob                : TOB;
  i                      : integer;
  T                      : TOB;
  EnDateDu, EnDateAu     : TDatetime;
  Q                      : TQuery;
begin
	DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode, sN1);
	
    Try
        DataTob := TOB.Create('~TEMP', nil, -1);

        StSQL := 'SELECT PFO_CODESTAGE,PFO_ORDRE,PST_LIBELLE,PSS_LIBELLE,PFO_DATEDEBUT,PFO_DATEFIN '+
        'FROM FORMATIONS LEFT JOIN STAGE ON PFO_MILLESIME=PST_MILLESIME AND PFO_CODESTAGE=PST_CODESTAGE '+
        'LEFT JOIN SESSIONSTAGE ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_MILLESIME=PSS_MILLESIME AND PFO_ORDRE=PSS_ORDRE '+
        'LEFT JOIN SCORING ON PFO_CODESTAGE=PSC_CODESTAGE AND PFO_ORDRE=PSC_ORDRE AND PSC_SALARIE="'+V_PGI.UserSalarie+'" '+
        'WHERE PFO_DATEDEBUT <= "'+USDATETIME (EnDateAu)+'" AND '+
        'PFO_DATEDEBUT >= "'+USDATETIME (EnDateDu)+'" AND PFO_SALARIE="'+V_PGI.UserSalarie+'" AND PFO_EFFECTUE="X" AND PSC_CODESTAGE IS NULL '+
        'ORDER BY PFO_DATEDEBUT';

        // Pas de lecture en cache car dès qu'une réponse est donnée, le stage doit disparaître de la liste
        Q := OpenSQL(StSQL, True);
        DataTob.LoadDetailDB('$DATA','','',Q,False);
        Ferme(Q);

        For i:=0 To DataTob.Detail.Count-1 Do
        Begin
            T := TOB.Create('£REPONSE',TobDonnees,-1);
            T.AddChampSupValeur('CODESTAGE',    DataTob.Detail[i].GetValue('PFO_CODESTAGE'));
            T.AddChampSupValeur('ORDRE',        DataTob.Detail[i].GetValue('PFO_ORDRE'));
            T.AddChampSupValeur('FORMATION',    DataTob.Detail[i].GetValue('PST_LIBELLE'));
            T.AddChampSupValeur('SESSION',      DataTob.Detail[i].GetValue('PSS_LIBELLE'));
            T.AddChampSupValeur('DEBUTSTAGE',   DateToStr(DataTob.Detail[i].GetValue('PFO_DATEDEBUT')));
            T.AddChampSupValeur('FINSTAGE',     DateToStr(DataTob.Detail[i].GetValue('PFO_DATEFIN')));
        End;

        If (TobDonnees.Detail.Count = 0) Then
            PutGridDetail('GRSALARIE', TobDonnees);
    Finally
        FreeAndNil (DataTob);
    End;
end;

{-----Charge les informations pour la saisie du scoring --------------------------------}

procedure PG_VIG_SCORING.AfficheQuestion (Init : Boolean);
var
  StSQL,Temp             : String;
  T                      : TOB;
  NumQ, NbQ				 : Integer;
begin
	If Init Then
	Begin
		// Récupération dans le flux de la ligne séléctionnée dans la grille sous forme de TOB
		T := GetLinkedValue('GRSALARIE');
		
		// Récupération de la session sélectionnée et de sa clé
		If (T <> Nil) And (T.Detail.Count > 0) Then
		Begin
            SetControlValue('CODESTAGE', T.Detail[0].GetValue('CODESTAGE'));
            SetControlValue('ORDRE',     T.Detail[0].GetValue('ORDRE'));
        	SetControlValue('TTITRE',    T.Detail[0].GetValue('SESSION') + ' du ' + T.Detail[0].GetValue('DEBUTSTAGE') + ' au ' + T.Detail[0].GetValue('FINSTAGE'));
		End;
		NumQ := 1;

		Try		
			StSQL := 'SELECT POD_NUMSCORING,POD_LIBELLE,POD_THEMESCORE,POD_TABLESCORE,"" AS REPONSE FROM SCORINGDEF ORDER BY POD_THEMESCORE';
        	T := OpenSelectInCache (StSQL);
        	ConvertFieldValue(T);
        	NbQ := T.Detail.Count;

            // Sauvegarde de la liste des questions pour ne pas avoir à les relire à chaque réponse
            SauveQuestions(T);
        Except
        	NbQ := 0;
    	End;

    	SetControlValue('NUMQUESTION', IntToStr(NumQ));
    	SetControlValue('NBQUESTIONS', IntToStr(NbQ));
	End
	Else
	Begin
		NumQ := StrToInt(GetControlValue('NUMQUESTION')) + 1;

        // Récupération de l'élément sauvegardé
        T := RecupQuestions;
		If T = Nil Then Exit;
		NbQ := T.Detail.Count;
    	SetControlValue('NUMQUESTION', IntToStr(NumQ));
	End;
	
	// Cas où le paramétrage n'a pas été effectué
	If NbQ = 0 Then
	Begin
		MessageErreur := 'Aucune question n''est paramétrée.';
		Exit;
	End
	Else
	Begin
		SetControlValue('TLIBTHEME', 	RechDom('PGFTHEMESCORING',T.Detail[NumQ-1].GetValue('POD_THEMESCORE'),False));
        Temp := T.Detail[NumQ-1].GetValue('POD_LIBELLE');
		SetControlValue('TQUESTION', 	Temp);
		ChargeValeursCombo (Self, 'CBREPONSE', 'PGFSCORING'+RightStr(T.Detail[NumQ-1].GetValue('POD_TABLESCORE'),1), '', True, '<<Ne se prononce pas>>');
	End;
	
end;

{-----Enregistrement du choix de l'utilisateur dans la TobDonnees ----------------------}

function PG_VIG_SCORING.Validation : Boolean;
Var TobData,T : TOB;
    i,NbQuest : Integer;
Begin
    Result := False;

	// Cas particulier de la fin des réponses. On ne fait rien pour revenir à la liste des stages
	If GetControlValue('NUMQUESTION') = 'FIN' Then Exit;

    // Récupération de l'élément sauvegardé
	T := RecupQuestions;
	If T = Nil Then Exit;

	T.Detail[StrToInt(GetControlValue('NUMQUESTION'))-1].PutValue('REPONSE', GetControlValue('CBREPONSE'));

    NbQuest := GetControlValue('NBQUESTIONS');

    // Si on a tout renseigné, on insert les données en base
	If GetControlValue('NUMQUESTION') = NbQuest Then
	Begin
	    Try
			TobData := TOB.Create('SCORING', Nil, -1);
			TobData.PutValue('PSC_CODESTAGE', 	GetControlValue('CODESTAGE'));
			TobData.PutValue('PSC_ORDRE', 	  	GetControlValue('ORDRE'));
			TobData.PutValue('PSC_SALARIE', 	V_PGI.UserSalarie);
			For i := 1 To NbQuest Do
			Begin
				TobData.PutValue('PSC_SCFO'+IntToStr(i),	T.Detail[i-1].GetValue('REPONSE'));
			End;
            For i := (NbQuest+1) To 30 Do
            Begin
				TobData.PutValue('PSC_SCFO'+IntToStr(i),	'');
            End;

			TobData.InsertOrUpdateDB;
            FreeAndNil(TobData);

			//SetControlValue('NUMQUESTION', 'FIN');
            Result := True;
    	Finally
			SupprimeQuestions;
    		Initialisation;
	    End;
	End;
End;

{-----Récupère la liste des questions à poser à l'utilisateur dans la session ----------}

Function PG_VIG_SCORING.RecupQuestions : TOB;
Var T         : TOB;
    LaSession : TISession;
Begin
	Try
    	LaSession := LookupCurrentSession;
    	T := TOB(LaSession.UserObjects.Objects[LaSession.UserObjects.IndexOf(ASK_LIST_NAME)]);
    	Result := T;
    Except
    	MessageErreur := TraduireMemoire('Impossible de récupérer la session.');
    	Result := Nil;
	End;
End;

{-----Supprime la liste des questions de la session ------------------------------------}

Procedure PG_VIG_SCORING.SupprimeQuestions;
Var T         : TOB;
    LaSession : TISession;
Begin
	Try
		LaSession := LookupCurrentSession;
		T := TOB(LaSession.UserObjects.Objects[LaSession.UserObjects.IndexOf(ASK_LIST_NAME)]);
		If T <> Nil Then
		Begin
			LaSession.UserObjects.Delete(LaSession.UserObjects.IndexOf(ASK_LIST_NAME));
		    If Assigned(T) Then FreeAndNil(T);
		End;
	Except
	End;
End;

{-----Sauvegarde la liste des questions au sein de la session --------------------------}

Procedure PG_VIG_SCORING.SauveQuestions (TobQuestions : TOB);
Var LaSession : TISession;
Begin
	Try
    	LaSession := LookupCurrentSession;
    	LaSession.UserObjects.AddObject(ASK_LIST_NAME, TobQuestions);
    Except
    	MessageErreur := TraduireMemoire('Impossible de sauvegarder la session.');
	End;
End;

end.


