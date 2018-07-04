{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - JL
Créé le ...... : 13/02/2007
Modifié le ... :   /  /
Description .. : liste des DIF pour un salarié
               : Vignette : PG_VIG_LISTEDDIF
               : Tablette : PGPERIODEVIGNETTE
               : Table    : FORMATIONS
Mots clefs ... :
*****************************************************************
PT1  | 18/01/2008 | FLO | Ajout du récapituliatif des compteurs
PT2  | 29/05/2008 | FLO | Ajout de la possibilité de supprimer une demande de DIF
PT3  | 29/05/2008 | FLO | Pas d'affichage des compteurs si la paie n'est pas gérée
PT4  | 18/06/2008 | FLO | Correction de l'affichage du motif de décision
}
unit PGVIGDDIFSAL;

interface
uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_DDIFSAL= class (TGAVignettePaie)
  protected
    procedure RecupDonnees;                                      override;
    function  GetInterface (NomGrille: string = ''): Boolean;    override;
    procedure GetClauseWhere;                                    override;
    procedure DrawGrid (Grille: string);                         override;
    function  SetInterface : Boolean ;                           override;
  private
    procedure AfficheListeDIF(AParam : string);
    procedure ChargeCompteurs; //PT1
    procedure Visualisation;   //PT2
    procedure ChargeDetails;   //PT2
    procedure SupprimeDemande; //PT2
    procedure InitialiseChamps;//PT2
    procedure ChangeEtatInsc (NvlEtat : String = ''); //PT4
  end;

implementation
uses
  SysUtils,
  PGVIGUTIL,
  DateUtils, //PT2
  ParamSoc,  //PT3
  PGOutilsFormation,
  HEnt1;
  
{-----Lit les critères -----------------------------------------------------------------}

function PG_VIG_DDIFSAL.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
    Result := true;
    
    // Rafraîchit les objets visibles et invisibles
    Visualisation; //PT2
end;

{-----Critères de la requète -----------------------------------------------------------}

procedure PG_VIG_DDIFSAL.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_DDIFSAL.RecupDonnees;
begin
  inherited;
  
  	//PT2
    If Fparam = 'VOIRETAT' Then
    Begin
        // Mise à jour du warning à False pour ne pas afficher le bandeau "La sélection ne renvoie aucun enregistrement"
        WarningOk := False;
        ChargeDetails;
    End;
    
    If (FParam = 'SUPPRIMER') Then
    	SupprimeDemande;

    If (Fparam = '') Or (Fparam = 'SUPPRIMER') Or (FParam = 'ANNULEETAT') Then
    Begin
    	// A l'initialisation de la vignette, on charge le récap des compteurs du salarié
    	// On le charge également après une suppression de demande
    	If (Fparam = '') Or (Fparam = 'SUPPRIMER') Then ChargeCompteurs;
    	
    	InitialiseChamps;
    End;

    If (Fparam <> 'VOIRETAT') Then
        AfficheListeDIF('');
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_DDIFSAL.DrawGrid (Grille: string);
begin
  inherited;
     SetFormatCol(Grille, 'DATEDEMANDE', 'C.0 ---');
     SetFormatCol(Grille, 'ETAT',        'C 0 ---');
     SetFormatCol(Grille, 'NBHEURES',    'D 0 ---');
     SetFormatCol(Grille, 'HEURESTT',    'D 0 ---');
     SetFormatCol(Grille, 'HEURESHT',    'D 0 ---');
     
     SetWidthCol (Grille, 'FORMATION',   150);
     SetWidthCol (Grille, 'DATEDEMANDE', 70);
     SetWidthCol (Grille, 'ETAT',        60);
     SetWidthCol (Grille, 'NBHEURES',    50);
     SetWidthCol (Grille, 'HEURESTT',    50);
     SetWidthCol (Grille, 'HEURESHT',    50);
     
     //PT2
     SetVisibleCol (Grille, 'CODESTAGE',     False);
     SetVisibleCol (Grille, 'RANG',     	 False);
     SetVisibleCol (Grille, 'ETABLISSEMENT', False);
     SetVisibleCol (Grille, 'CURSUS',     	 False);
end;

{-----Initialisation de la vignette ----------------------------------------------------}

function PG_VIG_DDIFSAL.SetInterface: Boolean;
begin
  inherited SetInterface;
     result:=true;
end;

{-----Affiche les objets nécessaires ---------------------------------------------------}

procedure PG_VIG_DDIFSAL.Visualisation; //PT2
var Visible : Boolean;
begin
    If (FParam = 'VOIRETAT') Then Visible := True
    Else Visible := False;

    SetControlVisible('PANVALID',   		Visible);
    SetControlVisible('PANBOUTONSDETAIL',   Visible);
    SetControlVisible('BSUPPRIMER', 		Visible); 

    SetControlVisible('FListe',     Not Visible);
    SetControlVisible('TMILLESIME', Not Visible);
    SetControlVisible('TBMOINS',    Not Visible);
    SetControlVisible('N1',         Not Visible);
    SetControlVisible('TBPLUS',     Not Visible);
    SetControlVisible('LBLETAT',    Not Visible);
    SetControlVisible('ETATDIF',    Not Visible);
    SetControlVisible('PERIODE',    Not Visible);
    SetControlVisible('TPERIODE',   Not Visible);
    SetControlVisible('BCOMPTEURS', Not Visible);
    SetControlVisible('BTOTAL',     Not Visible);
    SetControlVisible('PANTOTAL',   Not Visible);
    
    //PT3
    If GetParamSocSecur('SO_PGCUMULDIFACQUIS', '') = '' Then
        SetControlVisible('BCOMPTEURS', False);
end;

{-----Affichage des données de la vignette ---------------------------------------------}

procedure PG_VIG_DDIFSAL.AfficheListeDIF(AParam : string);
var
  StSQL,Where,sN1        : String;
  DataTob                : TOB;
  EnDateDu, EnDateAu     : TDatetime;
  i                      : integer;
  T                      : TOB;
  Q						 : TQuery;
begin
    DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode, sN1);

    If GetControlValue('ETATDIF') <> '' then Where := 'PFI_ETATINSCFOR="'+GetControlValue('ETATDIF')+'" AND '
    Else Where := '';
    Try
        DataTob := TOB.Create('~TEMP', nil, -1);
        
        StSQL := 'SELECT PFI_CODESTAGE,PFI_RANG,PFI_ETABLISSEMENT,PFI_CURSUS,'+ //PT2
        		 'PFI_DUREESTAGE,PFI_DATEDIF,PFI_ETATINSCFOR,PST_LIBELLE,PST_LIBELLE1,'+
                 'PFI_HTPSTRAV,PFI_HTPSNONTRAV FROM INSCFORMATION '+
                 'LEFT JOIN STAGE ON PFI_CODESTAGE=PST_CODESTAGE AND PST_MILLESIME="0000" '+
                 'WHERE '+Where+'PFI_REALISE="-" AND PFI_TYPEPLANPREV="DIF" '+
                 'AND PFI_SALARIE="'+V_PGI.UserSalarie+'" '+
                 'AND PFI_DATEDIF>="'+UsdateTime(EnDateDu)+'" AND PFI_DATEDIF<="'+UsdateTime(EnDateAu)+'" '+
                 'ORDER BY PFI_DATEDIF';

        //PT2
        // Pas de lecture en cache car les informations peuvent être modifiées
        //DataTob := OpenSelectInCache (StSQL);
        //ConvertFieldValue(DataTob);
        Q := OpenSQL(StSQL, True);
        DataTob.LoadDetailDB('$DATA','','',Q,False);
        Ferme(Q);

        For i:=0 to DataTob.Detail.Count-1 Do
        Begin
            T := TOB.Create('£REPONSE',TobDonnees,-1);
            T.AddChampSupValeur('FORMATION',    DataTob.Detail[i].GetValue('PST_LIBELLE')+' '+DataTob.Detail[i].GetValue('PST_LIBELLE1'));
            T.AddChampSupValeur('DATEDEMANDE',  DateToStr(DataTob.Detail[i].GetValue('PFI_DATEDIF')));
            T.AddChampSupValeur('ETAT',         RechDom('PGETATVALIDATION',DataTob.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
            T.AddChampSupValeur('NBHEURES',     DataTob.Detail[i].GetValue('PFI_DUREESTAGE'));
            T.AddChampSupValeur('HEURESTT',     DataTob.Detail[i].GetValue('PFI_HTPSTRAV'));
            T.AddChampSupValeur('HEURESHT',     DataTob.Detail[i].GetValue('PFI_HTPSNONTRAV'));
            //PT2
            T.AddChampSupValeur('CODESTAGE',    DataTob.Detail[i].GetValue('PFI_CODESTAGE'));
            T.AddChampSupValeur('RANG',     	DataTob.Detail[i].GetValue('PFI_RANG'));
            T.AddChampSupValeur('ETABLISSEMENT',DataTob.Detail[i].GetValue('PFI_ETABLISSEMENT'));
            T.AddChampSupValeur('CURSUS',     	DataTob.Detail[i].GetValue('PFI_CURSUS'));
        End;

        // Vidage de la table si pas de données
        If (TobDonnees.Detail.Count = 0) Then
            PutGridDetail('FListe', TobDonnees);

        AfficheTotal (Self, TobDonnees);
        DrawGrid ('GRTOTAL');
    Finally
        FreeAndNil (DataTob);
    end;
end;

{-----Charge les compteurs actuels du salarié -------------------------------------------------}

procedure PG_VIG_DDIFSAL.ChargeCompteurs();
var
  Acquis,PrisSTT,PrisHTT :Double;
  DdVSTT,DdVHTT,DdASTT,DdAHTT : Double;
  Q                      : TQuery;
begin
    Try
    	If GetParamSocSecur('SO_PGCUMULDIFACQUIS', '') <> '' Then Exit; //PT3

        DdVSTT := 0;
        DdVHTT := 0;
        DdASTT := 0;
        DdAHTT := 0;
        Acquis := 0;
        PrisSTT := 0;
        PrisHTT := 0;

        Q := OpenSQL('SELECT SUM(PHC_MONTANT) ACQUIS '+
        'FROM HISTOCUMSAL WHERE '+
        ' AND PHC_CUMULPAIE=(SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_PGCUMULDIFACQUIS") AND PHC_SALARIE="'+V_PGI.UserSalarie+'"',True);
        If Not Q.eof then Acquis := Q.FindField('ACQUIS').AsFloat;
        Ferme(Q);

        Q := OpenSQL('SELECT SUM (PFO_HTPSTRAV) STT,SUM (PFO_HTPSNONTRAV) HTT FROM FORMATIONS '+
        'WHERE PFO_SALARIE="'+V_PGI.UserSalarie+'" AND PFO_EFFECTUE="X" AND PFO_TYPEPLANPREV="DIF"',True);
        If Not Q.eof then
        begin
            PrisSTT := Q.FindField('STT').AsFloat;
            PrisHTT := Q.FindField('HTT').AsFloat;
        end;
        Ferme(Q);

        Q := OpenSQL('SELECT SUM (PFI_HTPSTRAV) STT,SUM (PFI_HTPSNONTRAV) HTT FROM INSCFORMATION '+
        'WHERE PFI_SALARIE="'+V_PGI.UserSalarie+'" AND PFI_ETATINSCFOR="VAL" AND PFI_REALISE="-" AND PFI_TYPEPLANPREV="DIF"',True);
        If Not Q.eof then
        begin
            DdVSTT := Q.FindField('STT').AsFloat;
            DdVHTT := Q.FindField('HTT').AsFloat;
        end;
        Ferme(Q);

        Q := OpenSQL('SELECT SUM (PFI_HTPSTRAV) STT,SUM (PFI_HTPSNONTRAV) HTT FROM INSCFORMATION '+
        'WHERE PFI_SALARIE="'+V_PGI.UserSalarie+'" AND PFI_ETATINSCFOR="ATT" AND PFI_REALISE="-" AND PFI_TYPEPLANPREV="DIF"',True);
        If Not Q.eof then
        begin
            DdASTT := Q.FindField('STT').AsFloat;
            DdAHTT := Q.FindField('HTT').AsFloat;
        end;
        Ferme(Q);

        SetControlValue('ACQUIS',   FloatToStr(Acquis));
        SetControlValue('PRIS',     FloatToStr(PrisSTT+PrisHTT));
        SetControlValue('SOLDE',    FloatToStr(Acquis - PrisSTT - PrisHTT));
        SetControlValue('VALIDES',  FloatToStr(DdVSTT+DdVHTT));
        SetControlValue('ATTENTE',  FloatToStr(DdASTT+DdAHTT));
        SetControlValue('SOLDETHEO',FloatToStr(Acquis - PrisSTT -PrisHTT -DdVSTT - DdVHTT - DdASTT - DdAHTT));
    finally
    end;
end;

{-----Charge les données en rapport avec le détail de la ligne sélectionnée ------------}

procedure PG_VIG_DDIFSAL.ChargeDetails; //PT2
var
    T                     : Tob;
    StSQL,Stage,Millesime : String;
    LibStage,Rang,Etab    : String;
    Cursus,sN1            : String;
    Q                     : TQuery;
    EnDateDu, EnDateAu    : TDatetime;
begin
   Try
     // Récupération dans le flux de la ligne séléctionnée dans la grille sous forme de TOB
     T := GetLinkedValue('FListe');

     // Récupération des données
     If (T <> Nil) And (T.Detail.Count > 0) Then
     Begin
          Stage     := T.Detail[0].GetValue('CODESTAGE');
          LibStage  := T.Detail[0].GetValue('FORMATION');
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

     SetControlVisible('TVALFORMATION',  Not (Stage = '--CURSUS--'));
     SetControlVisible('TFORMATION',     Not (Stage = '--CURSUS--'));
     SetControlValue  ('TVALFORMATION',  LibStage);
     SetControlValue  ('TVALCURSUS',     RechDom('PGCURSUS', Cursus, False));
     SetControlVisible('TVALCURSUS',     (Cursus<>''));
     SetControlVisible('TCURSUS',        (Cursus<>''));
     SetControlValue  ('SAVESTAGE',      Stage);
     SetControlValue  ('SAVERANG',       Rang);
     SetControlValue  ('SAVEETAB',       Etab);
     SetControlValue  ('SAVECURSUS',     Cursus);

     // Chargement des données
     StSQL := 'SELECT PFI_HTPSTRAV,PFI_HTPSNONTRAV,PFI_ETATINSCFOR,'+
              'PFI_MOTIFETATINSC,PFI_DATEACCEPT,PFI_REALISE FROM INSCFORMATION '+
              'WHERE PFI_RANG="'+Rang+'" AND PFI_CODESTAGE="'+Stage+'" AND PFI_MILLESIME="'+Millesime+'" AND PFI_ETABLISSEMENT="'+Etab+'"';
     Q := OpenSQL(StSQL, True);
     If Not Q.EOF Then
     Begin
         SetControlValue ('TVALNBHEURESSTT', FloatToStr(Q.FindField('PFI_HTPSTRAV').AsFloat));
         SetControlValue ('TVALNBHEURESHTT', FloatToStr(Q.FindField('PFI_HTPSNONTRAV').AsFloat));
         SetControlValue ('CBETATINSC',   Q.FindField('PFI_ETATINSCFOR').AsString);
         // Adaptation de la combo avant de mettre à jour sa valeur
         ChangeEtatInsc (Q.FindField('PFI_ETATINSCFOR').AsString); //PT4
         SetControlValue ('CBMOTIFDEC',   Q.FindField('PFI_MOTIFETATINSC').AsString);
         SetControlValue ('DATEACCEPT',   DateToStr(Q.FindField('PFI_DATEACCEPT').AsDateTime));
         SetControlValue ('CKPLAN',       StrToBool(Q.FindField('PFI_REALISE').AsString));

         If ((Q.FindField('PFI_ETATINSCFOR').AsString <> '') And (Q.FindField('PFI_ETATINSCFOR').AsString = 'VAL')) Or (Stage = '--CURSUS--') Then
         	SetControlVisible('BSUPPRIMER', False);
     End;
     Ferme(Q);
  Finally
  End;
end;

{-----Supprime une demande de formation ------------------------------------------------}

procedure PG_VIG_DDIFSAL.SupprimeDemande; //PT2
Var StSQL,sN1,Millesime : String;
    EnDateDu, EnDateAu  : TDatetime;
    TitreMail           : String;
    TexteMail           : HTStrings;
    DataTob,T           : TOB;
    Q                   : TQuery;
begin
     Try
        DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode, sN1);
        Millesime := IntToStr(YearOf(EnDateDu));

        // Récupération des informations avant suppression pour envoi du mail
	    DataTob := TOB.Create('~Infos', Nil, -1);
	    StSQL := 'SELECT * FROM INSCFORMATION '+
                 'WHERE PFI_CODESTAGE="'+GetControlValue('SAVESTAGE')+'" AND PFI_RANG="'+GetControlValue('SAVERANG')+'" '+
                 'AND PFI_MILLESIME="'+Millesime+'" AND PFI_ETABLISSEMENT="'+GetControlValue('SAVEETAB')+'"';
	    Q := OpenSQL (StSQL, True);
        DataTob.LoadDetailDB('INSCFORMATION','','',Q,False);
        Ferme(Q);

        If DataTob.Detail.Count = 0 Then Begin FreeAndNil(DataTob); Exit; End;

        // Suppression de la demande
        T := DataTob.Detail[0];
        T.DeleteDB();

        // Envoi d'un mail au responsable
	    PrepareMailFormation ('', '','ANNULATIONDIF',T,TitreMail,TexteMail);
	    EnvoieMail (T.GetValue('PFI_RESPONSFOR'), TitreMail, TexteMail);

	    FreeAndNil(DataTob);
     Finally
     End;
end;

{-----Initialise les champs de l'écran de détail ---------------------------------------}

procedure PG_VIG_DDIFSAL.InitialiseChamps; //PT2
begin
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
     SetControlValue ('CKPLAN',         False);
end;

{-----Charge les données de la combo motif de décision ---------------------------------}

procedure PG_VIG_DDIFSAL.ChangeEtatInsc (NvlEtat : String = ''); //PT4
Var Etat      : String;
begin
	If NvlEtat <> '' Then
		Etat := NvlEtat
	Else
		Etat := GetControlValue('CBETATINSC');

	Try
		If Etat = 'REF' then
			ChargeValeursCombo (Self, 'CBMOTIFDEC', 'PGMOTIFETATINSC3')
		Else If Etat = 'REP' then
			ChargeValeursCombo (Self, 'CBMOTIFDEC', 'PGMOTIFETATINSC2')
		Else If Etat = 'VAL' then
			ChargeValeursCombo (Self, 'CBMOTIFDEC', 'PGMOTIFETATINSC1')
		Else
			ChargeValeursCombo (Self, 'CBMOTIFDEC', 'PGMOTIFETATINSC');
	Finally
	End;

	SetControlValue  ('CBMOTIFDEC', '');
end;

end.
