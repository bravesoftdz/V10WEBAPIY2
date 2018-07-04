{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - JL
Créé le ...... : 12/02/2007
Modifié le ... :   /  /
Description .. : liste des DIF par responsable
               : Vignette : PG_VIG_LISTEDDIF
               : Tablette : PGPERIODEVIGNETTE
               : Table    : FORMATIONS
Mots clefs ... :
*****************************************************************
PT1  | 14/03/2008 | FLO | Gestion multi-dossier formation
PT2  | 25/03/2008 | FLO | Correctif pour la gestion des types utilisateurs
PT3  | 29/04/2008 | FLO | Prise en compte des salariés non sortis
PT4  | 14/05/2008 | FLO | Ajout de la validation des demandes
PT5  | 14/05/2008 | FLO | Ajout de la colonne reponsable pour les vignettes accessible par l'assistant ou le secrétaire
}
unit PGVIGLISTEDDIF;

interface
uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_LISTEDDIF= class (TGAVignettePaie)
  protected
    procedure RecupDonnees;                                     override;
    function  GetInterface (NomGrille: string = ''): Boolean;   override;
    procedure GetClauseWhere;                                   override;
    procedure DrawGrid (Grille: string);                        override;
    function  SetInterface : Boolean ;                          override;
  private
    procedure AfficheListeDif;
    //PT4
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
  PGOutilsFormation,
  uToolsPlugin,
  uToolsOWC,
  DateUtils,
  HEnt1;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_LISTEDDIF.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
    Result := true;
    
    //PT4
    // Rafraîchit les objets visibles et invisibles
    Visualisation;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_LISTEDDIF.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_LISTEDDIF.RecupDonnees;
begin
  inherited;
    //PT4
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
        AfficheListeDif;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_LISTEDDIF.DrawGrid (Grille: string);
Var
   S :String;
begin
  inherited;
    SetFormatCol(Grille, 'DATEDEMANDE', 'C.0 ---');
    SetFormatCol(Grille, 'ETAT',        'C.0 ---');
    SetFormatCol(Grille, 'NBHEURES',    'D 0 ---');
    SetFormatCol(Grille, 'HEURESTT',    'D 0 ---');
    SetFormatCol(Grille, 'HEURESHT',    'D 0 ---');

    SetWidthCol (Grille, 'NOM',         70);
    SetWidthCol (Grille, 'PRENOM',      70);
    SetWidthCol (Grille, 'FORMATION',   100);
    SetWidthCol (Grille, 'DATEDEMANDE', 60);
    SetWidthCol (Grille, 'ETAT',        60);
    SetWidthCol (Grille, 'NBHEURES',    50);
    SetWidthCol (Grille, 'HEURESTT',    50);
    SetWidthCol (Grille, 'HEURESHT',    50);
    
    //PT4
    SetVisibleCol (Grille, 'CODESALARIE',   False);
    SetVisibleCol (Grille, 'CODESTAGE',     False);
    SetVisibleCol (Grille, 'RANG',     	   	False);
    SetVisibleCol (Grille, 'ETABLISSEMENT', False);
    SetVisibleCol (Grille, 'CURSUS',     	False);
    
    //PT5
    S := RecupTypeUtilisateur('FOR');
    If (Pos('SECRETAIRE', S) > 0) Or (Pos('ADJOINT', S) > 0) Then
    	SetWidthCol (Grille, 'RESPONSABLE', 90)
    Else
    	SetVisibleCol (Grille, 'RESPONSABLE', False);
end;

{-----Initialisation de la vignette ----------------------------------------------------}

function PG_VIG_LISTEDDIF.SetInterface: Boolean;
begin
  inherited SetInterface;
    result:=true;
end;

{-----Affiche les objets nécessaires ---------------------------------------------------}
//PT4
procedure PG_VIG_LISTEDDIF.Visualisation;
var Visible : Boolean;
begin
    If (FParam = 'VOIRETAT') Or (FParam = 'ETATINSC') Then Visible := True
    Else Visible := False;

    SetControlVisible('PANVALID',   Visible);
    SetControlVisible('BVALIDER',   Visible);
    SetControlVisible('BANNULER',   Visible);

    SetControlVisible('FListe',     Not Visible);
    SetControlVisible('TPERIODE',   Not Visible);
    SetControlVisible('PERIODE',    Not Visible);
    SetControlVisible('TBMOINS',    Not Visible);
    SetControlVisible('N1',         Not Visible);
    SetControlVisible('TBPLUS',     Not Visible);
    SetControlVisible('LBLETAT',    Not Visible);
    SetControlVisible('ETATDIF',    Not Visible);
    SetControlVisible('BTOTAL',     Not Visible);
    SetControlVisible('PANTOTAL',   Not Visible);
end;

{-----Affichage des données de la vignette ---------------------------------------------}

procedure PG_VIG_LISTEDDIF.AfficheListeDif;
var
  StSQL,Where,sN1   : String;
  DataTob                : TOB;
  EnDateDu, EnDateAu     : TDatetime;
  i                      : integer;
  T                      : TOB;
  Q                      : TQuery;
begin
    DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode,sN1);

    If GetControlValue('ETATDIF') <> '' then Where := 'PFI_ETATINSCFOR="'+GetControlValue('ETATDIF')+'" AND '
    else Where := '';

    Try
        DataTob := TOB.Create('~TEMP', nil, -1);

        StSQL := 'SELECT PSA_LIBELLE,PSA_PRENOM,PFI_LIBELLE,'+
        		 'PFI_SALARIE,PFI_CODESTAGE,PFI_RANG,PFI_ETABLISSEMENT,PFI_CURSUS,'; //PT4
        //PT5
        If PGBundleHierarchie Then
            StSQL := StSQL + '(SELECT PSI_LIBELLE||" "||PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE=PFI_RESPONSFOR) AS RESPONSABLE,'
        Else
            StSQL := StSQL + '(SELECT PSA_LIBELLE||" "||PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE=PFI_RESPONSFOR) AS RESPONSABLE,';
        StSQL := StSQL + 'PFI_DUREESTAGE,PFI_DATEDIF,PFI_ETATINSCFOR,PST_LIBELLE,PST_LIBELLE1,'+
                 'PFI_HTPSTRAV,PFI_HTPSNONTRAV FROM INSCFORMATION '+
                 'LEFT JOIN STAGE ON PFI_CODESTAGE=PST_CODESTAGE AND PST_MILLESIME="0000" '+
                 'LEFT JOIN SALARIES ON PSA_SALARIE=PFI_SALARIE '+
                 'LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFI_SALARIE '+
                 'WHERE '+Where+'PFI_REALISE="-" AND PFI_TYPEPLANPREV="DIF" '+
                 'AND '+AdaptByTypeResp(RecupTypeUtilisateur('FOR'),V_PGI.UserSalarie,False,'PFI') + //PT2 //PT3
                 ' AND PFI_DATEDIF>="'+UsdateTime(EnDateDu)+'" AND PFI_DATEDIF<="'+UsdateTime(EnDateAu)+'" '+
                 'ORDER BY PFI_DATEDIF';

		StSQL := AdaptMultiDosForm (StSQL); //PT1

        //PT4
        // Pas de lecture en cache car les informations peuvent être modifiées
        //DataTob := OpenSelectInCache (StSQL);
        //ConvertFieldValue(DataTob);
        Q := OpenSQL(StSQL, True);
        DataTob.LoadDetailDB('$DATA','','',Q,False);
        Ferme(Q);

        for i:=0 to DataTob.detail.count-1 do
        begin
            T := TOB.Create('£REPONSE',TobDonnees,-1);
            If DataTob.Detail[i].GetValue('PSA_LIBELLE') <> '' Then
            Begin
                T.AddChampSupValeur('NOM',          DataTob.Detail[i].GetValue('PSA_LIBELLE'));
                T.AddChampSupValeur('PRENOM',       DataTob.Detail[i].GetValue('PSA_PRENOM'));
            End
            Else
            Begin
                T.AddChampSupValeur('NOM',          DataTob.Detail[i].GetValue('PFI_LIBELLE'));
                T.AddChampSupValeur('PRENOM',       '');
            End;
            T.AddChampSupValeur('RESPONSABLE',  DataTob.Detail[i].GetValue('RESPONSABLE')); //PT5
            T.AddChampSupValeur('FORMATION',    DataTob.Detail[i].GetValue('PST_LIBELLE')+' '+DataTob.Detail[i].GetValue('PST_LIBELLE1'));
            T.AddChampSupValeur('DATEDEMANDE',  DateToStr(DataTob.Detail[i].GetValue('PFI_DATEDIF')));
            T.AddChampSupValeur('ETAT',         RechDom('PGETATVALIDATION',DataTob.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
            T.AddChampSupValeur('NBHEURES',     DataTob.Detail[i].GetValue('PFI_DUREESTAGE'));
            T.AddChampSupValeur('HEURESTT',     DataTob.Detail[i].GetValue('PFI_HTPSTRAV'));
            T.AddChampSupValeur('HEURESHT',     DataTob.Detail[i].GetValue('PFI_HTPSNONTRAV'));
            //PT4
            T.AddChampSupValeur('CODESALARIE',  DataTob.Detail[i].GetValue('PFI_SALARIE'));
            T.AddChampSupValeur('CODESTAGE',    DataTob.Detail[i].GetValue('PFI_CODESTAGE'));
            T.AddChampSupValeur('RANG',     	DataTob.Detail[i].GetValue('PFI_RANG'));
            T.AddChampSupValeur('ETABLISSEMENT',DataTob.Detail[i].GetValue('PFI_ETABLISSEMENT'));
            T.AddChampSupValeur('CURSUS',     	DataTob.Detail[i].GetValue('PFI_CURSUS'));
        end;

        if (TobDonnees.Detail.Count = 0) then
            PutGridDetail('FListe', TobDonnees);

        AfficheTotal (Self, TobDonnees);
        DrawGrid ('GRTOTAL');
    finally
        FreeAndNil (DataTob);
    end;
end;

//PT4
{-----Charge les données en rapport avec le détail de la ligne sélectionnée ------------}

procedure PG_VIG_LISTEDDIF.ChargeDetails;
var
    T,TobTemp,DataTob      : Tob;
    StSQL,Stage,Millesime  : String;
    Rang,Etab              : String;
    Cursus,sN1             : String;
    Q                      : TQuery;
    i					   : Integer;
    EnDateDu, EnDateAu     : TDatetime;
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
         DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode, sN1);
         Millesime := IntToStr(YearOf(EnDateDu));
    End
    Else
    Begin
         Stage := ''; Millesime := ''; Rang := ''; Etab := ''; Cursus := '';
    End;

    SetControlValue('TVALSALARIE',     T.Detail[0].GetValue('NOM')+' '+T.Detail[0].GetValue('PRENOM'));
    SetControlVisible('TVALFORMATION', Not (Stage = '--CURSUS--'));
    SetControlVisible('TFORMATION',    Not (Stage = '--CURSUS--'));
    SetControlValue('TVALFORMATION',   T.Detail[0].GetValue('FORMATION'));
    SetControlValue('TVALCURSUS',      RechDom('PGCURSUS', Cursus, False));
    SetControlVisible('TVALCURSUS',    (Cursus<>''));
    SetControlVisible('TCURSUS',       (Cursus<>''));
    SetControlValue('SAVESTAGE',       Stage);
    SetControlValue('SAVERANG',        Rang);
    SetControlValue('SAVEETAB',        Etab);
    SetControlValue('SAVECURSUS',      Cursus);
    SetControlValue('SAVESALARIE',     T.Detail[0].GetValue('CODESALARIE'));

    // Chargement des données
    StSQL := 'SELECT PFI_DUREESTAGE,PFI_ETATINSCFOR,'+
             'PFI_MOTIFETATINSC,PFI_DATEACCEPT,PFI_REALISE FROM INSCFORMATION '+
             'WHERE PFI_RANG="'+Rang+'" AND PFI_CODESTAGE="'+Stage+'" AND PFI_MILLESIME="'+Millesime+'" AND PFI_ETABLISSEMENT="'+Etab+'"';
    Q := OpenSQL(StSQL, True);
    If Not Q.EOF Then
    Begin
        SetControlValue ('TVALNBHEURES', FloatToStr(Q.FindField('PFI_DUREESTAGE').AsFloat));
        SetControlValue ('CBETATINSC',   Q.FindField('PFI_ETATINSCFOR').AsString);
        // Adaptation de la combo avant de mettre à jour sa valeur
        ChangeEtatInsc (Q.FindField('PFI_ETATINSCFOR').AsString);
        SetControlValue ('CBMOTIFDEC',   Q.FindField('PFI_MOTIFETATINSC').AsString);
        SetControlValue ('DATEACCEPT',   DateToStr(Q.FindField('PFI_DATEACCEPT').AsDateTime));
        //SetControlValue ('CKPLAN',       StrToBool(Q.FindField('PFI_REALISE').AsString));  
    End;
    Ferme(Q);
     
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
  Finally
  End;
end;

{-----Charge les données de la combo motif de décision ---------------------------------}

procedure PG_VIG_LISTEDDIF.ChangeEtatInsc (NvlEtat : String = '');
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
		//SetControlValue ('CKPLAN',	   True); 
	End;
end;

{-----Valide les informations sur le motif de décision ---------------------------------}

procedure PG_VIG_LISTEDDIF.Valide;
Var StSQL,TitreMail     : String;
    TexteMail           : HTStrings;
    T                   : TOB;
    Q                   : TQuery;
    Millesime,sN1       : String;
    EnDateDu, EnDateAu  : TDatetime;
begin
     // Mise à jour des données
     Try
        DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode, sN1);
        Millesime := IntToStr(YearOf(EnDateDu));

        If ((GetControlValue('CBMOTIFDEC') = '') And (GetControlProperty('CBMOTIFDEC','Enabled'))) And (GetControlValue('CBETATINSC') <> 'NAN') Then
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
                         '", PFI_DATEACCEPT="'+UsDateTime(StrToDate(GetControlValue('DATEACCEPT')))+'", PFI_REALISE="-" '+
                         'WHERE PFI_RANG="'+GetControlValue('SAVERANG')+'" AND PFI_CURSUS="'+GetControlValue('SAVECURSUS')+
                         '" AND PFI_MILLESIME="'+Millesime+'" AND PFI_ETABLISSEMENT="'+GetControlValue('SAVEETAB')+'"';
                ExecuteSQL(StSQL);
            End;
        End
        Else
        Begin
            StSQL := 'UPDATE INSCFORMATION '+
                     'SET PFI_ETATINSCFOR="'+GetControlValue('CBETATINSC')+'", PFI_MOTIFETATINSC="'+GetControlValue('CBMOTIFDEC')+
                     '", PFI_DATEACCEPT="'+UsDateTime(StrToDate(GetControlValue('DATEACCEPT')))+'", PFI_REALISE="-" '+
                     'WHERE PFI_RANG="'+GetControlValue('SAVERANG')+'" AND PFI_CODESTAGE="'+GetControlValue('SAVESTAGE')+
                     '" AND PFI_MILLESIME="'+Millesime+'" AND PFI_ETABLISSEMENT="'+GetControlValue('SAVEETAB')+'"';
            ExecuteSQL(StSQL);
        End;

		// Envoi d'un mail au collaborateur        
        If GetControlValue('CBETATINSC') <> 'ATT' Then
        Begin
	        T := TOB.Create('~Infos', Nil, -1);
	        StSQL := 'SELECT * FROM INSCFORMATION WHERE PFI_RANG="'+GetControlValue('SAVERANG')+
                     '" AND PFI_MILLESIME="'+Millesime+'" AND PFI_ETABLISSEMENT="'+GetControlValue('SAVEETAB')+'"';
	        If GetControlValue('SAVESTAGE') <> '--CURSUS--' Then
	        	StSQL := StSQL +' AND PFI_CODESTAGE="'+GetControlValue('SAVESTAGE')+'"';
	        Q := OpenSQL (StSQL, True);
        	T.LoadDetailDB('$DATA','','',Q,False);
        	Ferme(Q);
	        PrepareMailFormation (GetControlValue('SAVESALARIE'),V_PGI.UserSalarie,'DECISIONDIF',T,TitreMail,TexteMail);
	        FreeAndNil(T);
	        EnvoieMail (GetControlValue('SAVESALARIE'), TitreMail, TexteMail);
	    End;
     Finally
     End;
end;

{-----Initialise les champs de l'écran de détail ---------------------------------------}

procedure PG_VIG_LISTEDDIF.InitialiseChamps;
begin
     SetControlValue('TVALSALARIE',     '');
     SetControlValue('TVALFORMATION',   '');
     SetControlValue('TVALCURSUS',      '');
     SetControlValue('SAVESTAGE',       '');
     SetControlValue('SAVERANG',        '');
     SetControlValue('SAVEETAB',        '');
     SetControlValue('SAVECURSUS',      '');
     SetControlValue ('TVALNBHEURES',   '');
     SetControlValue ('CBETATINSC',     '');
     SetControlValue ('CBMOTIFDEC',     '');
     SetControlValue ('DATEACCEPT',     '');
     SetControlValue ('QUESTION',       '');
     //SetControlValue ('CKPLAN',         False);
end;

end.


