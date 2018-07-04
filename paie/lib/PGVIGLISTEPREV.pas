{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - FLO
Créé le ...... : 18/01/2008
Modifié le ... :   /  /
Description .. : Liste des inscriptions au prévisionnel
               : Vignette : PG_VIG_LISTEPREV
               : Table    : INSCFORMATION
Mots clefs ... :
*****************************************************************
PT1  | 01/02/2008 | FLO | Pas de lecture en cache
PT2  | 14/03/2008 | FLO | Cocher "Formation dans le plan" par défaut + Affichage libellés emplois associés
PT3  | 19/03/2008 | FLO | Prise en compte du type utilisateur
PT4  | 25/03/2008 | FLO | Ne pas prendre en compte les DIF
PT5  | 29/04/2008 | FLO | Prise en compte des salariés non sortis
PT6  | 30/04/2008 | FLO | Vérouillage de la combo motif de décision si pas de tablette paramétrée
PT7  | 14/05/2008 | FLO | Ajout de la colonne reponsable pour les vignettes accessible par l'assistant ou le secrétaire
PT8  | 30/05/2008 | FLO | Affiche ou cache le panel de boutons et non plus les boutons unitairement
}
unit PGVIGLISTEPREV;

interface
uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_LISTEPREV= class (TGAVignettePaie)
  protected
    procedure RecupDonnees;                                     override;
    function  GetInterface (NomGrille: string = ''): Boolean;   override;
    procedure GetClauseWhere;                                   override;
    procedure DrawGrid (Grille: string);                        override;
    function  SetInterface : Boolean ;                          override;
  private
    procedure AfficheListeInsc;
    procedure Visualisation;
    procedure ChargeDetails;
    procedure ChangeEtatInsc (NvlEtat : String = '');
    procedure Valide;
    procedure InitialiseChamps;
 end;

implementation
uses
  SysUtils,
  PGVIGUTIL,
  uToolsPlugin,
  PGOutilsFormation,
  uToolsOWC,
  HEnt1;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_LISTEPREV.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
    Result := true;

    // Initialisation de la vignette au niveau des périodes pour simuler la sélection d'un millésime
    If FParam = '' Then
    Begin
        Periode := 'A';
        MajLibPeriode();
    End;

    // Rafraîchit les objets visibles et invisibles
    Visualisation;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_LISTEPREV.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_LISTEPREV.RecupDonnees;
begin
  inherited;
    If Fparam = 'VOIRETAT' Then
    Begin
        // Mise à jour du warning à False pour ne pas afficher le bandeau "La sélection ne renvoie aucun enregistrement"
        WarningOk := False;
        ChargeDetails;
    End
    Else If Fparam = 'ETATINSC' Then
    Begin
        // Mise à jour du warning à False pour ne pas afficher le bandeau "La sélection ne renvoie aucun enregistrement"
        WarningOk := False;
        ChangeEtatInsc;
    End;

    If Fparam = 'VALIDETAT' Then
        Valide;

    If (Fparam = '') Or (Fparam = 'VALIDETAT') Or (FParam = 'ANNULEETAT') Then
        InitialiseChamps;

    If (Fparam <> 'VOIRETAT') And (Fparam <> 'ETATINSC') Then
        AfficheListeInsc;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_LISTEPREV.DrawGrid (Grille: string);
Var S : String;
begin
  inherited;
    SetVisibleCol(Grille, 'CODESALARIE',    False);
    SetVisibleCol(Grille, 'CODESTAGE',      False);
    SetVisibleCol(Grille, 'CURSUS',         False);
    SetVisibleCol(Grille, 'RANG',           False);
    SetVisibleCol(Grille, 'ETABLISSEMENT',  False);
    SetVisibleCol(Grille, 'CODETYPEPLAN',   False);

	SetFormatCol(Grille, 'ETAT',     'C 0 ---');
    SetFormatCol(Grille, 'TYPEPLAN', 'C 0 ---');
    SetFormatCol(Grille, 'NATURE',   'C 0 ---');

    SetWidthCol(Grille, 'SALARIE',  100);
    SetWidthCol(Grille, 'STAGE',    150);
    SetWidthCol(Grille, 'ETAT',     60);
    SetWidthCol(Grille, 'TYPEPLAN', 80);
    SetWidthCol(Grille, 'NATURE',   50);
    
    //PT7
    S := RecupTypeUtilisateur('FOR');
    If (Pos('SECRETAIRE', S) > 0) Or (Pos('ADJOINT', S) > 0) Then
    	SetWidthCol (Grille, 'RESPONSABLE', 100)
    Else
    	SetVisibleCol (Grille, 'RESPONSABLE', False);
end;

{-----Initialisation de la vignette ----------------------------------------------------}

function PG_VIG_LISTEPREV.SetInterface: Boolean;
begin
  inherited SetInterface;
    result:=true;
end;

{-----Affiche les objets nécessaires ---------------------------------------------------}

procedure PG_VIG_LISTEPREV.Visualisation;
var Visible : Boolean;
begin
    If (FParam = 'VOIRETAT') Or (FParam = 'ETATINSC') Then Visible := True
    Else Visible := False;

    SetControlVisible('PANVALID',   Visible);
    SetControlVisible('PANBOUTONSDETAIL',   Visible); //PT8
//    SetControlVisible('BVALIDER',   Visible);
//    SetControlVisible('BANNULER',   Visible);

    SetControlVisible('FListe',     Not Visible);
    SetControlVisible('TMILLESIME', Not Visible);
    SetControlVisible('TBMOINS',    Not Visible);
    SetControlVisible('N1',         Not Visible);
    SetControlVisible('TBPLUS',     Not Visible);
    SetControlVisible('LBLETAT',    Not Visible);
    SetControlVisible('ETAT',       Not Visible);
    SetControlVisible('CKCURSUS',   Not Visible);
end;

{-----Affichage des données de la vignette ---------------------------------------------}

procedure PG_VIG_LISTEPREV.AfficheListeInsc;
var
  StSQL,Where,sN1   : String;
  DataTob                : TOB;
  EnDateDu, EnDateAu     : TDatetime;
  i                      : integer;
  T                      : TOB;
  Q                      : TQuery;
begin
    DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode, sN1);

    If GetControlValue('ETAT') <> '' then Where := 'PFI_ETATINSCFOR="'+GetControlValue('ETAT')+'" AND ';
    If GetControlValue('CKCURSUS') = True then
        Where := Where + 'PFI_CODESTAGE<>"--CURSUS--" AND '
    Else
        Where := Where + '(PFI_CODESTAGE="--CURSUS--" OR (PFI_CODESTAGE<>"--CURSUS--" AND PFI_CURSUS="")) AND ';

    Try
        DataTob := TOB.Create('~TEMP', nil, -1);

        StSQL := 'SELECT PFI_SALARIE,PFI_LIBELLE,PFI_RANG,PFI_ETABLISSEMENT,PFI_NBINSC,PFI_CODESTAGE,PFI_CURSUS,';
        //PT7
        If PGBundleHierarchie Then
            StSQL := StSQL + '(SELECT PSI_LIBELLE||" "||PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE=PFI_RESPONSFOR) AS RESPONSABLE,'
        Else
            StSQL := StSQL + '(SELECT PSA_LIBELLE||" "||PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE=PFI_RESPONSFOR) AS RESPONSABLE,';
        StSQL := StSQL + 'PST_LIBELLE,PST_LIBELLE1,PFI_ETATINSCFOR,PFI_TYPEPLANPREV,PFI_NATUREFORM '+
                 'FROM INSCFORMATION '+
                 'LEFT JOIN STAGE ON PFI_CODESTAGE=PST_CODESTAGE AND PST_MILLESIME="0000" '+
                 'WHERE '+Where+
                 AdaptByTypeResp(RecupTypeUtilisateur('FOR'),V_PGI.UserSalarie,False,'PFI') + //PT3
                 ' AND PFI_MILLESIME="'+sN1+'" '+
                 ' AND PFI_TYPEPLANPREV<>"DIF" '+ //PT4
                 'ORDER BY RESPONSABLE,PFI_LIBELLE,PST_LIBELLE';
        StSQL := AdaptMultiDosForm (StSQL); //PT5

        // Pas de lecture en cache car les informations peuvent être modifiées
        //DataTob := OpenSelectInCache (StSQL); //PT1
        //ConvertFieldValue(DataTob);
        Q := OpenSQL(StSQL, True);
        DataTob.LoadDetailDB('$DATA','','',Q,False);
        Ferme(Q);

        For i:=0 To DataTob.Detail.Count-1 Do
        Begin
            T := TOB.Create('£REPONSE',TobDonnees,-1);
            T.AddChampSupValeur('CODESALARIE',  DataTob.Detail[i].GetValue('PFI_SALARIE'));
            T.AddChampSupValeur('SALARIE',      DataTob.Detail[i].GetValue('PFI_LIBELLE'));
            T.AddChampSupValeur('RESPONSABLE',  DataTob.Detail[i].GetValue('RESPONSABLE')); //PT7
            T.AddChampSupValeur('CODESTAGE',  	DataTob.Detail[i].GetValue('PFI_CODESTAGE'));
            T.AddChampSupValeur('CURSUS',  	    DataTob.Detail[i].GetValue('PFI_CURSUS'));
            T.AddChampSupValeur('RANG',  	    DataTob.Detail[i].GetValue('PFI_RANG'));
            T.AddChampSupValeur('ETABLISSEMENT',DataTob.Detail[i].GetValue('PFI_ETABLISSEMENT'));
            If DataTob.Detail[i].GetValue('PFI_CODESTAGE') = '--CURSUS--' Then
      			T.AddChampSupValeur('STAGE',  	'Cursus : '+RechDom('PGCURSUS', DataTob.Detail[i].GetValue('PFI_CURSUS'), False))
            Else
      			T.AddChampSupValeur('STAGE',  	DataTob.Detail[i].GetValue('PST_LIBELLE')+' '+DataTob.Detail[i].GetValue('PST_LIBELLE1'));
			T.AddChampSupValeur('ETAT',  		RechDom('PGETATVALIDATION',DataTob.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
			T.AddChampSupValeur('CODETYPEPLAN', DataTob.Detail[i].GetValue('PFI_TYPEPLANPREV'));
			T.AddChampSupValeur('TYPEPLAN',  	RechDom('PGTYPEPLANPREV',DataTob.Detail[i].GetValue('PFI_TYPEPLANPREV'),False));
			T.AddChampSupValeur('NATURE',  		RechDom('PGNATUREFORM',DataTob.Detail[i].GetValue('PFI_NATUREFORM'),False));
        End;

        If (TobDonnees.Detail.Count = 0) Then
            PutGridDetail('FListe', TobDonnees);
    Finally
        FreeAndNil (DataTob);
    End;
end;

{-----Charge les données en rapport avec le détail de la ligne sélectionnée ------------}

procedure PG_VIG_LISTEPREV.ChargeDetails;
var
    T,TobTemp,DataTob      : Tob;
    StSQL,Stage,Millesime  : String;
    Rang,Etab              : String;
    Cursus                 : String;
    Q                      : TQuery;
    i					   : Integer;
begin
  Try
    // Récupération dans le flux de la ligne séléctionnée dans la grille sous forme de TOB
    T := GetLinkedValue('FListe');

    // Récupération des données
    If (T <> Nil) And (T.Detail.Count > 0) Then
    Begin
         Stage     := T.Detail[0].GetValue('CODESTAGE');
         Rang      := T.Detail[0].GetValue('RANG');
         Etab      := T.Detail[0].GetValue('ETABLISSEMENT');
         Cursus    := T.Detail[0].GetValue('CURSUS');
         Millesime := GetControlValue('N1');
    End
    Else
    Begin
         Stage := ''; Millesime := ''; Rang := ''; Etab := ''; Cursus := '';
    End;

    SetControlValue('TVALSALARIE',     T.Detail[0].GetValue('SALARIE'));
    SetControlVisible('TVALFORMATION', Not (Stage = '--CURSUS--'));
    SetControlVisible('TFORMATION',    Not (Stage = '--CURSUS--'));
    SetControlValue('TVALFORMATION',   T.Detail[0].GetValue('STAGE'));
    SetControlValue('TVALCURSUS',      RechDom('PGCURSUS', Cursus, False));
    SetControlVisible('TVALCURSUS',    (Cursus<>''));
    SetControlVisible('TCURSUS',       (Cursus<>''));
    SetControlValue('SAVESTAGE',       Stage);
    SetControlValue('SAVERANG',        Rang);
    SetControlValue('SAVEETAB',        Etab);
    SetControlValue('SAVECURSUS',      Cursus);
    SetControlValue('SAVESALARIE',     T.Detail[0].GetValue('CODESALARIE'));
    SetControlValue('SAVEPLAN',		T.Detail[0].GetValue('CODETYPEPLAN'));

    // Chargement des données
    StSQL := 'SELECT PFI_DUREESTAGE,PFI_FRAISFORFAIT,PFI_AUTRECOUT,PFI_COUTREELSAL,PFI_ALLOCFORM,PFI_ETATINSCFOR,'+
             'PFI_MOTIFETATINSC,PFI_DATEACCEPT,PFI_REALISE FROM INSCFORMATION '+
             'WHERE PFI_RANG="'+Rang+'" AND PFI_CODESTAGE="'+Stage+'" AND PFI_MILLESIME="'+Millesime+'" AND PFI_ETABLISSEMENT="'+Etab+'"';
    Q := OpenSQL(StSQL, True);
    If Not Q.EOF Then
    Begin
        SetControlValue ('TVALNBHEURES', FloatToStr(Q.FindField('PFI_DUREESTAGE').AsFloat));
        SetControlValue ('COUTFORFAIT',  FloatToStr(Q.FindField('PFI_FRAISFORFAIT').AsFloat));
        SetControlValue ('COUTPEDAG',    FloatToStr(Q.FindField('PFI_AUTRECOUT').AsFloat));
        SetControlValue ('COUTREEL',     FloatToStr(Q.FindField('PFI_COUTREELSAL').AsFloat));
        SetControlValue ('ALLOC',        FloatToStr(Q.FindField('PFI_ALLOCFORM').AsFloat));

        SetControlValue ('CBETATINSC',   Q.FindField('PFI_ETATINSCFOR').AsString);
        // Adaptation de la combo avant de mettre à jour sa valeur
        ChangeEtatInsc (Q.FindField('PFI_ETATINSCFOR').AsString);
        SetControlValue ('CBMOTIFDEC',   Q.FindField('PFI_MOTIFETATINSC').AsString);
        SetControlValue ('DATEACCEPT',   DateToStr(Q.FindField('PFI_DATEACCEPT').AsDateTime));
        SetControlValue ('CKPLAN',       StrToBool(Q.FindField('PFI_REALISE').AsString));  
    End;
    Ferme(Q);
     
    //PT2 - Début
	// Emplois associés
	DataTob := TOB.Create ('~EMPLOIS', Nil, -1);
	StSQL := 'SELECT PMF_LIBELLEEMPLOI FROM LIENEMPFORM WHERE PMF_CODESTAGE="'+Stage+'" AND PMF_ORDRE="-1"';
	TobTemp := OpenSelectInCache (StSQL);
	ConvertFieldValue(TobTemp);
	
	For i := 0 To TobTemp.Detail.Count - 1 Do
	Begin
	    T := TOB.Create('£EMPLOIS', DataTob, -1);
	    T.AddChampSupValeur('LIBELLEEMPLOI', RechDom('PGLIBEMPLOI',TobTemp.Detail[i].GetValue('PMF_LIBELLEEMPLOI'),False));
	End;
	
	PutGridDetail('GREMPLOIS', DataTob);
	//PT2 - Fin
  Finally
  End;
end;

{-----Charge les données de la combo motif de décision ---------------------------------}

procedure PG_VIG_LISTEPREV.ChangeEtatInsc (NvlEtat : String = '');
Var Etat      : String;
    Acces     : Boolean;
    NewData,N : Tob;
    i,Nb      : Integer;
begin
	If NvlEtat <> '' Then
		Etat := NvlEtat
	Else
		Etat := GetControlValue('CBETATINSC');
	Acces := True;

	Try
		If Etat = 'REF' then
		Begin
			ChargeValeursCombo (Self, 'CBMOTIFDEC', 'PGMOTIFETATINSC3');
		End
		Else If Etat = 'REP' then
		Begin
			ChargeValeursCombo (Self, 'CBMOTIFDEC', 'PGMOTIFETATINSC2');
		End
		Else If Etat = 'VAL' then
		Begin
			ChargeValeursCombo (Self, 'CBMOTIFDEC', 'PGMOTIFETATINSC1');
		End
		Else
		Begin
			ChargeValeursCombo (Self, 'CBMOTIFDEC', 'PGMOTIFETATINSC');
			Acces := False;
		End;
	Finally
		//PT6
		// Compte du nombre d'éléments dans la combo. S'il n'y en a pas, on vérouille la combo
		Try
			Nb := 0;
			NewData := GetDataInstance('CBMOTIFDEC');
			If Newdata <> nil then
			Begin
				For i := 0 To Newdata.Detail.Count - 1 Do
				Begin
					n := Newdata.Detail[i];
					if n.NomTable = 'T_ITEMSCOMBOBOX' Then
					Begin
						Nb := n.Detail.count;
						Break;
					End;
				End;
			End;
			If Nb = 0 Then Acces := False;
		Finally
		End;
	End;

	SetControlValue  ('CBMOTIFDEC', '');
	SetControlEnabled('CBMOTIFDEC', Acces);
	
	If GetControlValue('CBETATINSC') <> 'ATT' then
	Begin
		SetControlValue ('DATEACCEPT', DateToStr(Date));
		SetControlValue ('CKPLAN',	   True); //PT2
	End;
end;

{-----Valide les informations sur le motif de décision ---------------------------------}

procedure PG_VIG_LISTEPREV.Valide;
Var StSQL     : String;
    TitreMail : String;
    TexteMail : HTStrings;
    T         : TOB;
    Q         : TQuery;
begin
     // Mise à jour des données
     Try
        If ((GetControlValue('CBMOTIFDEC') = '') And (GetControlProperty('CBMOTIFDEC','Enabled'))) And (GetControlValue('CBETATINSC') <> 'NAN') Then //PT6
        Begin
              MessageErreur := TraduireMemoire('Vous devez renseigner le motif de décision');
              Exit;
        End;

        If GetControlValue('SAVESTAGE') = '--CURSUS--' Then
        Begin
            If GetControlValue('QUESTION') = 'O' Then
            Begin
                StSQL := 'UPDATE INSCFORMATION '+
                         'SET PFI_ETATINSCFOR="'+GetControlValue('CBETATINSC')+'", PFI_MOTIFETATINSC="'+GetControlValue('CBMOTIFDEC')+
                         '", PFI_DATEACCEPT="'+UsDateTime(StrToDate(GetControlValue('DATEACCEPT')))+'", PFI_REALISE="'+BoolToStr(GetControlValue('CKPLAN'))+'" '+
                         'WHERE PFI_RANG="'+GetControlValue('SAVERANG')+'" AND PFI_CURSUS="'+GetControlValue('SAVECURSUS')+
                         '" AND PFI_MILLESIME="'+GetControlValue('N1')+'" AND PFI_ETABLISSEMENT="'+GetControlValue('SAVEETAB')+'"';
                ExecuteSQL(StSQL);
            End;
        End
        Else
        Begin
            StSQL := 'UPDATE INSCFORMATION '+
                     'SET PFI_ETATINSCFOR="'+GetControlValue('CBETATINSC')+'", PFI_MOTIFETATINSC="'+GetControlValue('CBMOTIFDEC')+
                     '", PFI_DATEACCEPT="'+UsDateTime(StrToDate(GetControlValue('DATEACCEPT')))+'", PFI_REALISE="'+BoolToStr(GetControlValue('CKPLAN'))+'" '+
                     'WHERE PFI_RANG="'+GetControlValue('SAVERANG')+'" AND PFI_CODESTAGE="'+GetControlValue('SAVESTAGE')+
                     '" AND PFI_MILLESIME="'+GetControlValue('N1')+'" AND PFI_ETABLISSEMENT="'+GetControlValue('SAVEETAB')+'"';
            ExecuteSQL(StSQL);
        End;

		// Envoi d'un mail au collaborateur        
        If GetControlValue('CBETATINSC') <> 'ATT' Then
        Begin
	        T := TOB.Create('~Infos', Nil, -1);
	        StSQL := 'SELECT * FROM INSCFORMATION WHERE PFI_RANG="'+GetControlValue('SAVERANG')+
                     '" AND PFI_MILLESIME="'+GetControlValue('N1')+'" AND PFI_ETABLISSEMENT="'+GetControlValue('SAVEETAB')+'"';
	        If GetControlValue('SAVESTAGE') <> '--CURSUS--' Then
	        	StSQL := StSQL +' AND PFI_CODESTAGE="'+GetControlValue('SAVESTAGE')+'"';
	        Q := OpenSQL (StSQL, True);
        	T.LoadDetailDB('$DATA','','',Q,False);
        	Ferme(Q);
	        PrepareMailFormation (GetControlValue('SAVESALARIE'),V_PGI.UserSalarie,'DECISION'+GetControlValue('SAVEPLAN'),T,TitreMail,TexteMail);
	        FreeAndNil(T);
	        EnvoieMail (GetControlValue('SAVESALARIE'), TitreMail, TexteMail);
	    End;
     Finally
     End;
end;

{-----Initialise les champs de l'écran de détail ---------------------------------------}

procedure PG_VIG_LISTEPREV.InitialiseChamps;
begin
     SetControlValue('TVALSALARIE',     '');
     SetControlValue('TVALFORMATION',   '');
     SetControlValue('TVALCURSUS',      '');
     SetControlValue('SAVESTAGE',       '');
     SetControlValue('SAVERANG',        '');
     SetControlValue('SAVEETAB',        '');
     SetControlValue('SAVECURSUS',      '');
     SetControlValue ('TVALNBHEURES',   '');
     SetControlValue ('COUTFORFAIT',    '');
     SetControlValue ('COUTPEDAG',      '');
     SetControlValue ('COUTREEL',       '');
     SetControlValue ('ALLOC',          '');
     SetControlValue ('CBETATINSC',     '');
     SetControlValue ('CBMOTIFDEC',     '');
     SetControlValue ('DATEACCEPT',     '');
     SetControlValue ('QUESTION',       '');
     SetControlValue ('CKPLAN',         False);
end;

end.


