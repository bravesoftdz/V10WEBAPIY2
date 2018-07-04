{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - MF
Créé le ...... : 03/05/2006
Modifié le ... :   /  /
Description .. : Liste des absence de la période  (Liste)
               : Vignette : PG_VIG_ABSPERIDOE
               : Tablette : PGPERIODEVIGNETTE
               : Table    : ABSENCESALARIE, SALARIES
Mots clefs ... :
*****************************************************************
PT1  | 25/01/2008 | FLO | Modification de la taille des colonnes
PT2  | 29/05/2008 | FLO | Prise en compte du reponsable + 
						  Ajout de la colonne reponsable pour les vignettes accessible par l'assistant ou le secrétaire
}
unit PGVIGABSPERIODE;

interface
uses
  Classes,
  UTob,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  PGVIGUTIL,
  uToolsOWC,
  uToolsPlugin,
  HCtrls;
type
  PG_VIG_ABSPERIODE = class (TGAVignettePaie)
  private
    sN1             : string;
    procedure AfficheChamps();
  protected
    procedure RecupDonnees; override;
    function GetInterface (NomGrille: string = ''): Boolean; override;
    procedure GetClauseWhere; override;
    procedure DrawGrid (Grille: string); override;
    function SetInterface : Boolean ; override;
  public
    function InitPeriode : Char; override;
  end;

implementation
 uses
  SysUtils,
  PgoutilsFormation,
  HEnt1;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_ABSPERIODE.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
  Result := true;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_ABSPERIODE.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_ABSPERIODE.RecupDonnees;
begin
  inherited;
  AfficheChamps();
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_ABSPERIODE.DrawGrid (Grille: string);
var S : String;
begin
  inherited;
    //PT1
    SetWidthCol(Grille, 'NOM',    80);
    SetWidthCol(Grille, 'PRENOM', 80);
    SetWidthCol(Grille, 'LIBABS', 120);
    
    //PT2
    S := RecupTypeUtilisateur('ABS');
    If (Pos('SECRETAIRE', S) > 0) Or (Pos('ADJOINT', S) > 0) Then
    	SetWidthCol (Grille, 'RESPONSABLE', 90)
    Else
    	SetVisibleCol (Grille, 'RESPONSABLE', False);
end;

{-----Initialisation de la vignette ----------------------------------------------------}

function PG_VIG_ABSPERIODE.SetInterface: Boolean;
begin
  inherited SetInterface;
  result:=true;
end;

{-----Affichage des données de la vignette ---------------------------------------------}

procedure PG_VIG_ABSPERIODE.AfficheChamps();
var
  StSQL : String;
  DataTob : TOB;
  EnDateDu, EnDateAu     : TDatetime;
  i: integer;
  T: TOB;
begin
  DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode,sN1);

  Try
    DataTob := TOB.Create('~TEMP', nil, -1);

    StSQL := 'SELECT PSA_LIBELLE,PSA_PRENOM,PCN_LIBELLE,';
	         //PT2 - Début
	         If PGBundleHierarchie Then
	             StSQL := StSQL + '(SELECT PSI_LIBELLE||" "||PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE='
	         Else
	             StSQL := StSQL + '(SELECT PSA_LIBELLE||" "||PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE=';
	         StSQL := StSQL + '(SELECT PSE_RESPONSABS FROM DEPORTSAL WHERE PSE_SALARIE=PCN_SALARIE)) AS RESPONSABLE ';
	         //PT2 - Fin
    		 StSQL := StSQL + 'FROM ABSENCESALARIE '+
             'LEFT JOIN SALARIES ON PCN_SALARIE = PSA_SALARIE '+
             'LEFT JOIN DEPORTSAL ON PCN_SALARIE = PSE_SALARIE '+
             ' WHERE (PCN_TYPEMVT="ABS" OR (PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI" and pcn_mvtduplique<>"X"))'+
             ' AND (((PCN_DATEDEBUTABS >= "'+USDATETIME (EnDateDu)+'") AND (PCN_DATEDEBUTABS <= "'+
             USDATETIME (EnDateAu)+'"))'+
             ' OR ((PCN_DATEFINABS >= "'+USDATETIME (EnDateDu)+'") AND (PCN_DATEFINABS <= "'+
             USDATETIME (EnDateAu)+'"))'+
             ' OR ((PCN_DATEDEBUTABS < "'+USDATETIME (EnDateDu)+'") AND (PCN_DATEFINABS > "'+
             USDATETIME (EnDateAu)+'"))) '+
             'AND '+AdaptByTypeResp(RecupTypeUtilisateur('ABS'),V_PGI.UserSalarie,False,'PSE') + ' ' + //PT2
             'GROUP BY PCN_SALARIE,PCN_LIBELLE,PSA_LIBELLE,PSA_PRENOM '+
             'ORDER BY RESPONSABLE,PSA_LIBELLE,PSA_PRENOM'; //PT2
    StSQL := AdaptMultiDosForm (StSQL); //PT2
    DataTob := OpenSelectInCache (StSQL);
    ConvertFieldValue(DataTob);

    for i:=0 to DataTob.detail.count-1 do
    begin
      T := TOB.Create('£REPONSE',TobDonnees,-1);
      T.AddChampSupValeur('NOM',DataTob.Detail[i].GetValue('PSA_LIBELLE'));
      T.AddChampSupValeur('PRENOM',DataTob.Detail[i].GetValue('PSA_PRENOM'));
      T.AddChampSupValeur('RESPONSABLE',DataTob.Detail[i].GetValue('RESPONSABLE')); //PT2
      T.AddChampSupValeur('LIBABS',DataTob.Detail[i].GetValue('PCN_LIBELLE'));
    end;

    if (TobDonnees.detail.count = 0) then
      PutGridDetail('GRSALARIE', TobDonnees);
  finally
    FreeAndNil (DataTob);
  end;

end;

{-----Retourne la période par défaut de la vignette ------------------------------------}

function PG_VIG_ABSPERIODE.InitPeriode : Char;
begin
  Result := 'J';       // valeur par défaut = J (jour)
end;

end.
