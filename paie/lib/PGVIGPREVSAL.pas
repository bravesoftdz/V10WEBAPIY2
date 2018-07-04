{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - FLO
Créé le ...... : 18/01/2008
Modifié le ... :   /  /
Description .. : Liste des inscriptions au prévisionnel pour le salarié
               : Vignette : PG_VIG_PREVSAL
               : Table    : INSCFORMATION
Mots clefs ... :
*****************************************************************
PT1  | 18/03/2008 | FLO | Suppression de la possibilité d'annuler une demande en attente
PT2  | 15/05/2008 | FLO | Retour arrière sur la précédente demande
PT3  | 30/05/2008 | FLO | Affiche/cache le panel de boutons et non pas les boutons unitairement
PT4  | 18/06/2008 | FLO | Correction de l'affichage du motif de décision
}
unit PGVIGPREVSAL;

interface
uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_PREVSAL= class (TGAVignettePaie)
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
    procedure SupprimeDemande;
    procedure InitialiseChamps;
    procedure ChangeEtatInsc (NvlEtat : String = ''); //PT4
 end;

implementation
uses
  SysUtils,
  PGVIGUTIL,
  uToolsPlugin,
  uToolsOWC,
  HEnt1;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_PREVSAL.GetInterface (NomGrille: string): Boolean;
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

procedure PG_VIG_PREVSAL.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_PREVSAL.RecupDonnees;
begin
  inherited;
    If Fparam = 'VOIRETAT' Then
    Begin
        // Mise à jour du warning à False pour ne pas afficher le bandeau "La sélection ne renvoie aucun enregistrement"
        WarningOk := False;
        ChargeDetails;
    End;
    
    If (FParam = 'SUPPRIMER') Then
    	SupprimeDemande;

    If (Fparam = '') Or (Fparam = 'SUPPRIMER') Or (FParam = 'ANNULEETAT') Then
        InitialiseChamps;

    If (Fparam <> 'VOIRETAT') And (Fparam <> 'ETATINSC') Then
        AfficheListeInsc;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_PREVSAL.DrawGrid (Grille: string);
begin
  inherited;
    SetVisibleCol(Grille, 'CODESTAGE',      False);
    SetVisibleCol(Grille, 'CURSUS',         False);
    SetVisibleCol(Grille, 'RANG',           False);
    SetVisibleCol(Grille, 'ETABLISSEMENT',  False);

	SetFormatCol(Grille, 'ETAT',     'C 0 ---');
    SetFormatCol(Grille, 'NATURE',   'C 0 ---');

    SetWidthCol(Grille, 'STAGE',    150);
    SetWidthCol(Grille, 'ETAT',     60);
    SetWidthCol(Grille, 'NATURE',   50);
end;

{-----Initialisation de la vignette ----------------------------------------------------}

function PG_VIG_PREVSAL.SetInterface: Boolean;
begin
  inherited SetInterface;
    result:=true;
end;

{-----Affiche les objets nécessaires ---------------------------------------------------}

procedure PG_VIG_PREVSAL.Visualisation;
var Visible : Boolean;
begin
    If (FParam = 'VOIRETAT') Or (FParam = 'ETATINSC') Then Visible := True
    Else Visible := False;

    SetControlVisible('PANVALID',   Visible);
    SetControlVisible('PANBOUTONSDETAIL',   Visible); //PT3
    SetControlVisible('BSUPPRIMER', Visible); //PT1 On ne doit jamais pouvoir annuler une demande //PT2
    //SetControlVisible('BANNULER',   Visible);

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

procedure PG_VIG_PREVSAL.AfficheListeInsc;
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

        StSQL := 'SELECT PFI_RANG,PFI_ETABLISSEMENT,PFI_CODESTAGE,PFI_CURSUS,'+
                 'PST_LIBELLE,PST_LIBELLE1,PFI_ETATINSCFOR,PFI_TYPEPLANPREV,PFI_NATUREFORM '+
                 'FROM INSCFORMATION '+
                 'LEFT JOIN STAGE ON PFI_CODESTAGE=PST_CODESTAGE AND PFI_MILLESIME=PST_MILLESIME '+
                 'WHERE '+Where+
                 'PFI_SALARIE="'+V_PGI.UserSalarie+'" '+
                 'AND PFI_TYPEPLANPREV<>"DIF" '+
                 'AND PFI_MILLESIME="'+sN1+'" '+
                 'ORDER BY PST_LIBELLE,PST_LIBELLE1';

        // Pas de lecture en cache car les informations peuvent être modifiées
        //DataTob := OpenSelectInCache (StSQL);
        //ConvertFieldValue(DataTob);
        Q := OpenSQL(StSQL, True);
        DataTob.LoadDetailDB('$DATA','','',Q,False);
        Ferme(Q);

        For i:=0 To DataTob.Detail.Count-1 Do
        Begin
            T := TOB.Create('£REPONSE',TobDonnees,-1);
            T.AddChampSupValeur('CODESTAGE',  	DataTob.Detail[i].GetValue('PFI_CODESTAGE'));
            T.AddChampSupValeur('CURSUS',  	    DataTob.Detail[i].GetValue('PFI_CURSUS'));
            T.AddChampSupValeur('RANG',  	    DataTob.Detail[i].GetValue('PFI_RANG'));
            T.AddChampSupValeur('ETABLISSEMENT',DataTob.Detail[i].GetValue('PFI_ETABLISSEMENT'));
            If DataTob.Detail[i].GetValue('PFI_CODESTAGE') = '--CURSUS--' Then
      			T.AddChampSupValeur('STAGE',  	'Cursus : '+RechDom('PGCURSUS', DataTob.Detail[i].GetValue('PFI_CURSUS'), False))
            Else
      			T.AddChampSupValeur('STAGE',  	DataTob.Detail[i].GetValue('PST_LIBELLE')+' '+DataTob.Detail[i].GetValue('PST_LIBELLE1'));
			T.AddChampSupValeur('ETAT',  		RechDom('PGETATVALIDATION',DataTob.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
			T.AddChampSupValeur('NATURE',  		RechDom('PGNATUREFORM',DataTob.Detail[i].GetValue('PFI_NATUREFORM'),False));
        End;

        If (TobDonnees.Detail.Count = 0) Then
            PutGridDetail('FListe', TobDonnees);
    Finally
        FreeAndNil (DataTob);
    End;
end;

{-----Charge les données en rapport avec le détail de la ligne sélectionnée ------------}

procedure PG_VIG_PREVSAL.ChargeDetails;
var
    T                     : Tob;
    StSQL,Stage,Millesime : String;
    LibStage,Rang,Etab    : String;
    Cursus                : String;
    Q                     : TQuery;
begin
   Try
     // Récupération dans le flux de la ligne séléctionnée dans la grille sous forme de TOB
     T := GetLinkedValue('FListe');

     // Récupération des données
     If (T <> Nil) And (T.Detail.Count > 0) Then
     Begin
          Stage     := T.Detail[0].GetValue('CODESTAGE');
          LibStage  := T.Detail[0].GetValue('STAGE');
          Rang      := T.Detail[0].GetValue('RANG');
          Etab      := T.Detail[0].GetValue('ETABLISSEMENT');
          Cursus    := T.Detail[0].GetValue('CURSUS');
          Millesime := GetControlValue('N1');
     End
     Else
     Begin
          Stage := ''; Millesime := ''; Rang := ''; Etab := '';
          LibStage := ''; Cursus := '';
     End;

     //SetControlValue('TVALSALARIE',     LibSalarie);
     SetControlVisible('TVALFORMATION', Not (Stage = '--CURSUS--'));
     SetControlVisible('TFORMATION',    Not (Stage = '--CURSUS--'));
     SetControlValue('TVALFORMATION',   LibStage);
     SetControlValue('TVALCURSUS',      RechDom('PGCURSUS', Cursus, False));
     SetControlVisible('TVALCURSUS',    (Cursus<>''));
     SetControlVisible('TCURSUS',       (Cursus<>''));
     SetControlValue('SAVESTAGE',       Stage);
     SetControlValue('SAVERANG',        Rang);
     SetControlValue('SAVEETAB',        Etab);
     SetControlValue('SAVECURSUS',      Cursus);

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
         SetControlValue ('CKPLAN',       StrToBool(Q.FindField('PFI_REALISE').AsString));

         //PT1 - Ne pas donner la possibilité d'annuler une demande
         //PT2 - Réactivation
         If ((Q.FindField('PFI_ETATINSCFOR').AsString <> '') And (Q.FindField('PFI_ETATINSCFOR').AsString = 'VAL')) Or (Stage = '--CURSUS--') Then
         	SetControlVisible('BSUPPRIMER', False);
     End;
     Ferme(Q);
  Finally
  End;
end;

{-----Supprime une demande de formation ------------------------------------------------}

procedure PG_VIG_PREVSAL.SupprimeDemande;
Var StSQL : String;
begin
     Try
        StSQL := 'DELETE FROM INSCFORMATION '+
                 'WHERE PFI_RANG="'+GetControlValue('SAVERANG')+'" AND PFI_CODESTAGE="'+GetControlValue('SAVESTAGE')+
                 '" AND PFI_MILLESIME="'+GetControlValue('N1')+'" AND PFI_ETABLISSEMENT="'+GetControlValue('SAVEETAB')+'"';
        ExecuteSQL(StSQL);
     Finally
     End;
end;

{-----Initialise les champs de l'écran de détail ---------------------------------------}

procedure PG_VIG_PREVSAL.InitialiseChamps;
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

procedure PG_VIG_PREVSAL.ChangeEtatInsc (NvlEtat : String = ''); //PT4
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


