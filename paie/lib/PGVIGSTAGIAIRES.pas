{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - JL
Créé le ...... : 13/02/2007
Modifié le ... :   /  /
Description .. : liste des stagiaires apr responsable
               : Vignette : PG_VIG_STAGIAIRES
               : Tablette : PGPERIODEVIGNETTE
               : Table    : FORMATIONS
Mots clefs ... :
*****************************************************************
PT1  | 18/01/2008 | FLO | Ajout de la gestion des modes de visualisation
PT2  | 19/03/2008 | FLO | Prise en compte du type utilisateur
PT3  | 29/04/2008 | FLO | Prise en compte des salariés non sortis
PT4  | 14/05/2008 | FLO | Ajout de la colonne reponsable pour les vignettes accessible par l'assistant ou le secrétaire
PT5  | 15/05/2008 | FLO | Plus de selectInCache car la vignette est rafraîchie automatiquement par l'inscription aux sessions
}
unit PGVIGSTAGIAIRES;

interface
uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_STAGIAIRES= class (TGAVignettePaie)
  protected
    procedure RecupDonnees;                                     override;
    function  GetInterface (NomGrille: string = ''): Boolean;   override;
    procedure GetClauseWhere;                                   override;
    procedure DrawGrid (Grille: string);                        override;
    function  SetInterface : Boolean ;                          override;
  private
    procedure AfficheListeHisto;  //PT1
    procedure AfficheListeAVenir; //PT1
    procedure AfficheListeRealise; //PT1
    procedure ChangeVisu (aMode : String = '');  //PT1
 end;

implementation
uses
  SysUtils,
  PGVIGUTIL,
  HEnt1, uToolsPlugin, PGOutilsFormation;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_STAGIAIRES.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
    Result := true;

    // Gestion des modes de visualisation
    ChangeVisu;  //PT1
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_STAGIAIRES.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_STAGIAIRES.RecupDonnees;
Var ModeVisu : String;  //PT1
begin
  inherited;

    ModeVisu := GetControlValue('VISU');

    // Visualisation des formations futures par défaut
    If (FParam = '') And (GetControlValue('VISU') = '') Then
    Begin
        ModeVisu := 'FUTUR';
        SetControlValue('VISU', ModeVisu);
        ChangeVisu (ModeVisu);
    End;

    If ModeVisu = 'REAL' Then
        AfficheListeRealise
    Else If ModeVisu = 'FUTUR' Then
        AfficheListeAVenir
    Else
        AfficheListeHisto;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_STAGIAIRES.DrawGrid (Grille: string);
Var ModeVisu,S : String;  //PT1
begin
  inherited;
    ModeVisu := GetControlValue('VISU');

    If ModeVisu = 'REAL' Then
    Begin
        SetWidthCol(Grille, 'HORAIRES', -1);
        SetWidthCol(Grille, 'PRESENT',  -1);
    End
    Else If ModeVisu = 'FUTUR' Then
    Begin
        SetWidthCol(Grille, 'HORAIRES', 60);
        SetWidthCol(Grille, 'PRESENT',  -1);
    End
    Else
    Begin
        SetWidthCol(Grille, 'HORAIRES', 60);
        SetWidthCol(Grille, 'PRESENT',  50);
    End;

    SetFormatCol(Grille, 'TYPEPLAN',    'C.0 ---');
    SetFormatCol(Grille, 'DEBUTSTAGE',  'C.0 ---');
    SetFormatCol(Grille, 'FINSTAGE',    'C.0 ---');
    SetFormatCol(Grille, 'HORAIRES',    'C.0 ---');
    SetFormatCol(Grille, 'PRESENT',     'C.0 ---');
    SetFormatCol(Grille, 'HEURESTT',    'D 0 ---');
    SetFormatCol(Grille, 'HEURESHTT',   'D 0 ---');

    SetWidthCol (Grille, 'NOM',         80);
    SetWidthCol (Grille, 'PRENOM',      80);
    SetWidthCol (Grille, 'FORMATION',   120);
    SetWidthCol (Grille, 'TYPEPLAN',    60);
    SetWidthCol (Grille, 'DEBUTSTAGE',  60);
    SetWidthCol (Grille, 'FINSTAGE',    60);
    SetWidthCol (Grille, 'HEURESTT',    50);
    SetWidthCol (Grille, 'HEURESHTT',   50);
    
    //PT4
    S := RecupTypeUtilisateur('FOR');
    If (Pos('SECRETAIRE', S) > 0) Or (Pos('ADJOINT', S) > 0) Then
    	SetWidthCol (Grille, 'RESPONSABLE', 100)
    Else
    	SetVisibleCol (Grille, 'RESPONSABLE', False);
end;

{-----Affectation des controls/valeurs dans la TobResponse -----------------------------}

function PG_VIG_STAGIAIRES.SetInterface: Boolean;
begin
  inherited SetInterface;
    result:=true;
end;

{-----Change le mode de visualisation de la vignette -----------------------------------}
//PT1
procedure PG_VIG_STAGIAIRES.ChangeVisu (aMode : String = '');
Var
  Visible  : Boolean;
  ModeVisu : String;
begin
    If aMode <> '' Then
        ModeVisu := aMode
    Else
        ModeVisu := GetControlValue('VISU');

    Visible := (ModeVisu = 'REAL') Or (ModeVisu = 'FUTUR');

    SetControlVisible('TPERIODE', Not Visible);
    SetControlVisible('PERIODE',  Not Visible);
    SetControlVisible('TBMOINS',  Not Visible);
    SetControlVisible('TBPLUS',   Not Visible);
    SetControlVisible('N1',       Not Visible);
end;

{-----Affiche les données de façon classique -------------------------------------------}

procedure PG_VIG_STAGIAIRES.AfficheListeHisto;
var
  StSQL,sN1              : String;
  DataTob                : TOB;
  EnDateDu, EnDateAu     : TDatetime;
  i                      : integer;
  T                      : TOB;
  Q                      : TQuery;
begin
    DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode,sN1);

    Try
        DataTob := TOB.Create('~TEMP', nil, -1);

        StSQL := 'SELECT PST_LIBELLE,PST_LIBELLE1,PFO_NOMSALARIE,PFO_PRENOM,PFO_DATEDEBUT,PFO_DATEFIN,PFO_EFFECTUE,';
        //PT4
        If PGBundleHierarchie Then
            StSQL := StSQL + '(SELECT PSI_LIBELLE||" "||PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE=PFO_RESPONSFOR) AS RESPONSABLE,'
        Else
            StSQL := StSQL + '(SELECT PSA_LIBELLE||" "||PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE=PFO_RESPONSFOR) AS RESPONSABLE,';
        StSQL := StSQL + 'PFO_HTPSNONTRAV, PFO_HTPSTRAV ,PFO_HEUREDEBUT,PFO_HEUREFIN,PFO_TYPEPLANPREV '+
                 'FROM FORMATIONS LEFT JOIN STAGE ON PFO_MILLESIME=PST_MILLESIME AND PFO_CODESTAGE=PST_CODESTAGE '+
                 'WHERE PFO_DATEDEBUT <= "'+USDATETIME (EnDateAu)+'" AND '+
                 'PFO_DATEDEBUT >= "'+USDATETIME (EnDateDu)+'" AND '+
                 AdaptByTypeResp(RecupTypeUtilisateur('FOR'),V_PGI.UserSalarie,False,'PFO',False) + //PT2
                 //AND PFO_RESPONSFOR="'+V_PGI.UserSalarie+'" '+
                 ' ORDER BY RESPONSABLE,PFO_NOMSALARIE,PFO_DATEDEBUT';
        StSQL := AdaptMultiDosForm (StSQL); //PT3

        // Pas de lecture en cache car les informations peuvent être modifiées
        //DataTob := OpenSelectInCache (StSQL); //PT5
        //ConvertFieldValue(DataTob);
        Q := OpenSQL(StSQL, True);
        DataTob.LoadDetailDB('$DATA','','',Q,False);
        Ferme(Q);

        for i:=0 to DataTob.detail.count-1 do
        begin
            T := TOB.Create('£REPONSE',TobDonnees,-1);
            T.AddChampSupValeur('NOM',        DataTob.Detail[i].GetValue('PFO_NOMSALARIE'));
            T.AddChampSupValeur('PRENOM',     DataTob.Detail[i].GetValue('PFO_PRENOM'));
            T.AddChampSupValeur('RESPONSABLE',DataTob.Detail[i].GetValue('RESPONSABLE')); //PT4
            T.AddChampSupValeur('FORMATION',  DataTob.Detail[i].GetValue('PST_LIBELLE')+' '+DataTob.Detail[i].GetValue('PST_LIBELLE1'));
            T.AddChampSupValeur('TYPEPLAN',   RechDom('PGTYPEPLANPREV',DataTob.Detail[i].GetValue('PFO_TYPEPLANPREV'),False));
            T.AddChampSupValeur('DEBUTSTAGE', DateToStr(DataTob.Detail[i].GetValue('PFO_DATEDEBUT')));
            T.AddChampSupValeur('FINSTAGE',   DateToStr(DataTob.Detail[i].GetValue('PFO_DATEFIN')));
            T.AddChampSupValeur('HORAIRES',   FormatDateTime('hh:mm', DataTob.Detail[i].GetValue('PFO_HEUREDEBUT'))+' - '+FormatDateTime('hh:mm', DataTob.Detail[i].GetValue('PFO_HEUREFIN')));
            T.AddChampSupValeur('HEURESTT',   DataTob.Detail[i].GetValue('PFO_HTPSTRAV'));
            T.AddChampSupValeur('HEURESHTT',  DataTob.Detail[i].GetValue('PFO_HTPSNONTRAV'));
            T.AddChampSupValeur('PRESENT',    LibBool(DataTob.Detail[i].GetValue('PFO_EFFECTUE')));
        end;

        if (TobDonnees.Detail.Count = 0) then
            PutGridDetail('GRSALARIE', TobDonnees);

        AfficheTotal (Self, TobDonnees);
        DrawGrid ('GRTOTAL');
    Finally
        FreeAndNil (DataTob);
    end;
end;

{-----Affiche les formations effectuées ------------------------------------------------}
//PT1
procedure PG_VIG_STAGIAIRES.AfficheListeRealise;
var
  StSQL        : String;
  DataTob      : TOB;
  i            : integer;
  T            : TOB;
begin
    Try
        DataTob := TOB.Create('~TEMP', nil, -1);

        StSQL := 'SELECT PST_LIBELLE,PFO_NOMSALARIE,PFO_PRENOM,PFO_DATEDEBUT,PFO_DATEFIN,PFO_EFFECTUE,';
        //PT4
        If PGBundleHierarchie Then
            StSQL := StSQL + '(SELECT PSI_LIBELLE||" "||PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE=PFO_RESPONSFOR) AS RESPONSABLE,'
        Else
            StSQL := StSQL + '(SELECT PSA_LIBELLE||" "||PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE=PFO_RESPONSFOR) AS RESPONSABLE,';
        StSQL := StSQL + 'PFO_HTPSNONTRAV, PFO_HTPSTRAV ,PFO_HEUREDEBUT,PFO_HEUREFIN,PFO_TYPEPLANPREV '+
                 'FROM FORMATIONS LEFT JOIN STAGE ON PFO_MILLESIME=PST_MILLESIME AND PFO_CODESTAGE=PST_CODESTAGE '+
                 'WHERE PFO_DATEDEBUT <= "'+USDATETIME (Date)+'" AND PFO_EFFECTUE="X" AND '+
                 AdaptByTypeResp(RecupTypeUtilisateur('FOR'),V_PGI.UserSalarie,False,'PFO') + //PT2
                 ' ORDER BY RESPONSABLE,PFO_NOMSALARIE,PFO_DATEDEBUT';
		StSQL := AdaptMultiDosForm (StSQL); //PT3
		
        DataTob := OpenSelectInCache (StSQL);
        ConvertFieldValue(DataTob);

        for i:=0 to DataTob.detail.count-1 do
        begin
            T := TOB.Create('£REPONSE',TobDonnees,-1);
            T.AddChampSupValeur('NOM',        DataTob.Detail[i].GetValue('PFO_NOMSALARIE'));
            T.AddChampSupValeur('PRENOM',     DataTob.Detail[i].GetValue('PFO_PRENOM'));
            T.AddChampSupValeur('RESPONSABLE',DataTob.Detail[i].GetValue('RESPONSABLE')); //PT4
            T.AddChampSupValeur('FORMATION',  DataTob.Detail[i].GetValue('PST_LIBELLE'));
            T.AddChampSupValeur('TYPEPLAN',   RechDom('PGTYPEPLANPREV',DataTob.Detail[i].GetValue('PFO_TYPEPLANPREV'),False));
            T.AddChampSupValeur('DEBUTSTAGE', DateToStr(DataTob.Detail[i].GetValue('PFO_DATEDEBUT')));
            T.AddChampSupValeur('FINSTAGE',   DateToStr(DataTob.Detail[i].GetValue('PFO_DATEFIN')));
            T.AddChampSupValeur('HORAIRES',   0);
            T.AddChampSupValeur('HEURESTT',   DataTob.Detail[i].GetValue('PFO_HTPSTRAV'));
            T.AddChampSupValeur('HEURESHTT',  DataTob.Detail[i].GetValue('PFO_HTPSNONTRAV'));
            T.AddChampSupValeur('PRESENT',    '');
        end;

        if (TobDonnees.Detail.Count = 0) then
            PutGridDetail('GRSALARIE', TobDonnees);

        AfficheTotal (Self, TobDonnees);
        DrawGrid ('GRTOTAL');
    Finally
        FreeAndNil (DataTob);
    end;
end;

{-----Affiche les formations à venir ---------------------------------------------------}
//PT1
procedure PG_VIG_STAGIAIRES.AfficheListeAVenir;
var
  StSQL       : String;
  DataTob     : TOB;
  i           : integer;
  T           : TOB;
begin
    Try
        DataTob := TOB.Create('~TEMP', nil, -1);

        StSQL := 'SELECT PST_LIBELLE,PFO_NOMSALARIE,PFO_PRENOM,PFO_DATEDEBUT,PFO_DATEFIN,PFO_EFFECTUE,';
        //PT4
        If PGBundleHierarchie Then
            StSQL := StSQL + '(SELECT PSI_LIBELLE||" "||PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE=PFO_RESPONSFOR) AS RESPONSABLE,'
        Else
            StSQL := StSQL + '(SELECT PSA_LIBELLE||" "||PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE=PFO_RESPONSFOR) AS RESPONSABLE,';
        StSQL := StSQL + 'PFO_HTPSNONTRAV, PFO_HTPSTRAV ,PFO_HEUREDEBUT,PFO_HEUREFIN,PFO_TYPEPLANPREV '+
                 'FROM FORMATIONS LEFT JOIN STAGE ON PFO_MILLESIME=PST_MILLESIME AND PFO_CODESTAGE=PST_CODESTAGE '+
                 'WHERE PFO_DATEDEBUT >= "'+USDATETIME (Date)+'" AND '+
                 AdaptByTypeResp(RecupTypeUtilisateur('FOR'),V_PGI.UserSalarie,False,'PFO') + //PT2
                 ' ORDER BY RESPONSABLE,PFO_NOMSALARIE,PFO_DATEDEBUT';
                 
        StSQL := AdaptMultiDosForm (StSQL); //PT3

        DataTob := OpenSelectInCache (StSQL);
        ConvertFieldValue(DataTob);

        for i:=0 to DataTob.detail.count-1 do
        begin
            T := TOB.Create('£REPONSE',TobDonnees,-1);
            T.AddChampSupValeur('NOM',        DataTob.Detail[i].GetValue('PFO_NOMSALARIE'));
            T.AddChampSupValeur('PRENOM',     DataTob.Detail[i].GetValue('PFO_PRENOM'));
            T.AddChampSupValeur('RESPONSABLE',DataTob.Detail[i].GetValue('RESPONSABLE')); //PT4
            T.AddChampSupValeur('FORMATION',  DataTob.Detail[i].GetValue('PST_LIBELLE'));
            T.AddChampSupValeur('TYPEPLAN',   RechDom('PGTYPEPLANPREV',DataTob.Detail[i].GetValue('PFO_TYPEPLANPREV'),False));
            T.AddChampSupValeur('DEBUTSTAGE', DateToStr(DataTob.Detail[i].GetValue('PFO_DATEDEBUT')));
            T.AddChampSupValeur('FINSTAGE',   DateToStr(DataTob.Detail[i].GetValue('PFO_DATEFIN')));
            T.AddChampSupValeur('HORAIRES',   FormatDateTime('hh:mm', DataTob.Detail[i].GetValue('PFO_HEUREDEBUT'))+' - '+FormatDateTime('hh:mm', DataTob.Detail[i].GetValue('PFO_HEUREFIN')));
            T.AddChampSupValeur('HEURESTT',   DataTob.Detail[i].GetValue('PFO_HTPSTRAV'));
            T.AddChampSupValeur('HEURESHTT',  DataTob.Detail[i].GetValue('PFO_HTPSNONTRAV'));
            T.AddChampSupValeur('PRESENT',    '');
        end;

        if (TobDonnees.Detail.Count = 0) then
            PutGridDetail('GRSALARIE', TobDonnees);

        AfficheTotal (Self, TobDonnees);
        DrawGrid ('GRTOTAL');
    Finally
        FreeAndNil (DataTob);
    end;
end;

end.


