{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - JL
Créé le ...... : 13/02/2007
Modifié le ... :   /  /
Description .. : liste des stages d'un salarié
               : Vignette : PG_VIG_STAGESAL
               : Tablette : PGPERIODEVIGNETTE
Mots clefs ... :
*****************************************************************
PT1  | 18/01/2008 | FLO | Ajout de la prise en compte des modes de visualisation
}
unit PGVIGSTAGESAL;

interface
uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_STAGESAL= class (TGAVignettePaie)
  protected
    procedure RecupDonnees;                                     override;
    function GetInterface (NomGrille: string = ''): Boolean;    override;
    procedure GetClauseWhere;                                   override;
    procedure DrawGrid (Grille: string);                        override;
    function SetInterface : Boolean ;                           override;
  private
    procedure AfficheListeHisto;   //PT1
    procedure AfficheListeAVenir;  //PT1
    procedure AfficheListeRealise; //PT1
    procedure ChangeVisu (aMode : String = '');  //PT1
 end;

implementation
uses
  SysUtils,
  PGVIGUTIL,
  HEnt1;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_STAGESAL.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
    Result := true;

    // Gestion des modes de visualisation
    ChangeVisu;  //PT1
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_STAGESAL.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_STAGESAL.RecupDonnees;
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

procedure PG_VIG_STAGESAL.DrawGrid (Grille: string);
Var ModeVisu : String;  //PT1
begin
  inherited;
    ModeVisu := GetControlValue('VISU');

    If ModeVisu = 'REAL' Then
    Begin
        SetWidthCol(Grille, 'SESSION',  80);
        SetWidthCol(Grille, 'HORAIRES', -1);
        SetWidthCol(Grille, 'NBHEURES', -1);
        SetWidthCol(Grille, 'HEURESTT', 50);
        SetWidthCol(Grille, 'HEURESHTT',50);
        SetWidthCol(Grille, 'PRESENT',  -1);
    End
    Else If ModeVisu = 'FUTUR' Then
    Begin
        SetWidthCol(Grille, 'SESSION',  80);
        SetWidthCol(Grille, 'HORAIRES', 70);
        SetWidthCol(Grille, 'NBHEURES', 50);
        SetWidthCol(Grille, 'HEURESTT', -1);
        SetWidthCol(Grille, 'HEURESHTT',-1);
        SetWidthCol(Grille, 'PRESENT',  -1);
    End
    Else
    Begin
        SetWidthCol(Grille, 'SESSION',  -1);
        SetWidthCol(Grille, 'HORAIRES', -1);
        SetWidthCol(Grille, 'NBHEURES', -1);
        SetWidthCol(Grille, 'HEURESTT', 50);
        SetWidthCol(Grille, 'HEURESHTT',50);
        SetWidthCol(Grille, 'PRESENT',  50);
    End;

    SetFormatCol(Grille, 'TYPEPLAN',    'C.0 ---');
    SetFormatCol(Grille, 'DEBUTSTAGE',  'C.0 ---');
    SetFormatCol(Grille, 'FINSTAGE',    'C.0 ---');
    SetFormatCol(Grille, 'HORAIRES',    'C.0 ---');
    SetFormatCol(Grille, 'PRESENT',     'C.0 ---');
    SetFormatCol(Grille, 'NBHEURES',    'D 0 ---');
    SetFormatCol(Grille, 'HEURESTT',    'D 0 ---');
    SetFormatCol(Grille, 'HEURESHTT',   'D 0 ---');

    SetWidthCol (Grille, 'FORMATION',   100);
    SetWidthCol (Grille, 'TYPEPLAN',    80);
    SetWidthCol (Grille, 'DEBUTSTAGE',  60);
    SetWidthCol (Grille, 'FINSTAGE',    60);
end;

{-----Affectation des controls/valeurs dans la TobResponse -----------------------------}

function PG_VIG_STAGESAL.SetInterface: Boolean;
begin
  inherited SetInterface;
    result:=true;
end;

{-----Change le mode de visualisation de la vignette -----------------------------------}
//PT1
procedure PG_VIG_STAGESAL.ChangeVisu (aMode : String = '');
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

{-----Affiche l'historique de formation réalisées ou non -------------------------------}

procedure PG_VIG_STAGESAL.AfficheListeHisto;
var
  StSQL,sN1              : String;
  DataTob                : TOB;
  EnDateDu, EnDateAu     : TDatetime;
  i                      : integer;
  T                      : TOB;
begin

    DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode, sN1);

    Try
        DataTob := TOB.Create('~TEMP', nil, -1);

        StSQL := 'SELECT PST_LIBELLE,PFO_NOMSALARIE,PFO_PRENOM,PFO_DATEDEBUT,PFO_DATEFIN,PFO_EFFECTUE,'+
        'PFO_HTPSNONTRAV, PFO_HTPSTRAV ,PFO_HEUREDEBUT,PFO_HEUREFIN,PFO_TYPEPLANPREV '+
        'FROM FORMATIONS LEFT JOIN STAGE ON PFO_MILLESIME=PST_MILLESIME AND PFO_CODESTAGE=PST_CODESTAGE '+
        'WHERE PFO_DATEDEBUT <= "'+USDATETIME (EnDateAu)+'" AND '+
        'PFO_DATEDEBUT >= "'+USDATETIME (EnDateDu)+'" AND PFO_SALARIE="'+V_PGI.UserSalarie+'" '+
        'ORDER BY PFO_NOMSALARIE,PFO_DATEDEBUT';

        DataTob := OpenSelectInCache (StSQL);
        ConvertFieldValue(DataTob);

        for i:=0 to DataTob.detail.count-1 do
        begin
            T := TOB.Create('£REPONSE',TobDonnees,-1);
            T.AddChampSupValeur('FORMATION',    DataTob.Detail[i].GetValue('PST_LIBELLE'));
            T.AddChampSupValeur('SESSION',      '');
            T.AddChampSupValeur('TYPEPLAN',     RechDom('PGTYPEPLANPREV',DataTob.Detail[i].GetValue('PFO_TYPEPLANPREV'),False));
            T.AddChampSupValeur('DEBUTSTAGE',   DateToStr(DataTob.Detail[i].GetValue('PFO_DATEDEBUT')));
            T.AddChampSupValeur('FINSTAGE',     DateToStr(DataTob.Detail[i].GetValue('PFO_DATEFIN')));
            T.AddChampSupValeur('HORAIRES',     0);
            T.AddChampSupValeur('NBHEURES',     0);
            T.AddChampSupValeur('HEURESTT',     DataTob.Detail[i].GetValue('PFO_HTPSTRAV'));
            T.AddChampSupValeur('HEURESHTT',    DataTob.Detail[i].GetValue('PFO_HTPSNONTRAV'));
            T.AddChampSupValeur('PRESENT',      LibBool(DataTob.Detail[i].GetValue('PFO_EFFECTUE')));
        end;

        if (TobDonnees.Detail.Count = 0) then
            PutGridDetail('GRSALARIE', TobDonnees);

        AfficheTotal (Self, TobDonnees);
        DrawGrid ('GRTOTAL');
    finally
        FreeAndNil (DataTob);
    end;
end;

{-----Affiche les formations à venir pour le salarié -----------------------------------}
//PT1
procedure PG_VIG_STAGESAL.AfficheListeAVenir;
var
  StSQL      : String;
  DataTob    : TOB;
  i          : integer;
  T          : TOB;
begin
    Try
        DataTob := TOB.Create('~TEMP', nil, -1);

        StSQL := 'SELECT PST_LIBELLE,PSS_LIBELLE,PFO_DATEDEBUT,PFO_DATEFIN,'+
        'PFO_HTPSNONTRAV,PFO_HTPSTRAV,PFO_HEUREDEBUT,PFO_HEUREFIN,PFO_TYPEPLANPREV '+
        'FROM FORMATIONS LEFT JOIN STAGE ON PFO_MILLESIME=PST_MILLESIME AND PFO_CODESTAGE=PST_CODESTAGE '+
        'LEFT JOIN SESSIONSTAGE ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_MILLESIME=PSS_MILLESIME AND PFO_ORDRE=PSS_ORDRE '+
        'WHERE PFO_DATEDEBUT >= "'+USDATETIME (Date)+'" AND PFO_SALARIE="'+V_PGI.UserSalarie+'" '+
        'ORDER BY PFO_NOMSALARIE,PFO_DATEDEBUT';

        DataTob := OpenSelectInCache (StSQL);
        ConvertFieldValue(DataTob);

        For i:=0 To DataTob.Detail.Count-1 Do
        Begin
            T := TOB.Create('£REPONSE',TobDonnees,-1);
            T.AddChampSupValeur('FORMATION',    DataTob.Detail[i].GetValue('PST_LIBELLE'));
            T.AddChampSupValeur('SESSION',      DataTob.Detail[i].GetValue('PSS_LIBELLE'));
            T.AddChampSupValeur('TYPEPLAN',     RechDom('PGTYPEPLANPREV',DataTob.Detail[i].GetValue('PFO_TYPEPLANPREV'),False));
            T.AddChampSupValeur('DEBUTSTAGE',   DateToStr(DataTob.Detail[i].GetValue('PFO_DATEDEBUT')));
            T.AddChampSupValeur('FINSTAGE',     DateToStr(DataTob.Detail[i].GetValue('PFO_DATEFIN')));
            T.AddChampSupValeur('HORAIRES',     FormatDateTime('hh:mm', DataTob.Detail[i].GetValue('PFO_HEUREDEBUT'))+' - '+FormatDateTime('hh:mm', DataTob.Detail[i].GetValue('PFO_HEUREFIN')));
            T.AddChampSupValeur('NBHEURES',     DataTob.Detail[i].GetValue('PFO_HTPSTRAV')+DataTob.Detail[i].GetValue('PFO_HTPSNONTRAV'));
            T.AddChampSupValeur('HEURESTT',     0);
            T.AddChampSupValeur('HEURESHTT',    0);
            T.AddChampSupValeur('PRESENT',      '');
        End;

        if (TobDonnees.Detail.Count = 0) then
            PutGridDetail('GRSALARIE', TobDonnees);

        AfficheTotal (Self, TobDonnees);
        DrawGrid ('GRTOTAL');
    finally
        FreeAndNil (DataTob);
    end;
end;

{-----Affiche les formations déjà réalisées par le salarié -----------------------------}
//PT1
procedure PG_VIG_STAGESAL.AfficheListeRealise;
var
  StSQL      : String;
  DataTob    : TOB;
  i          : integer;
  T          : TOB;
begin
    Try
        DataTob := TOB.Create('~TEMP', nil, -1);

        StSQL := 'SELECT PST_LIBELLE,PSS_LIBELLE,PFO_DATEDEBUT,PFO_DATEFIN,'+
        'PFO_HTPSNONTRAV,PFO_HTPSTRAV,PFO_TYPEPLANPREV '+
        'FROM FORMATIONS LEFT JOIN STAGE ON PFO_MILLESIME=PST_MILLESIME AND PFO_CODESTAGE=PST_CODESTAGE '+
        'LEFT JOIN SESSIONSTAGE ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_MILLESIME=PSS_MILLESIME AND PFO_ORDRE=PSS_ORDRE '+
        'WHERE PFO_DATEDEBUT <= "'+USDATETIME (Date)+'" AND PFO_SALARIE="'+V_PGI.UserSalarie+'" AND PFO_EFFECTUE="X"'+
        'ORDER BY PFO_NOMSALARIE,PFO_DATEDEBUT';

        DataTob := OpenSelectInCache (StSQL);
        ConvertFieldValue(DataTob);

        For i:=0 To DataTob.Detail.Count-1 Do
        Begin
            T := TOB.Create('£REPONSE',TobDonnees,-1);
            T.AddChampSupValeur('FORMATION',    DataTob.Detail[i].GetValue('PST_LIBELLE'));
            T.AddChampSupValeur('SESSION',      DataTob.Detail[i].GetValue('PSS_LIBELLE'));
            T.AddChampSupValeur('TYPEPLAN',     RechDom('PGTYPEPLANPREV',DataTob.Detail[i].GetValue('PFO_TYPEPLANPREV'),False));
            T.AddChampSupValeur('DEBUTSTAGE',   DateToStr(DataTob.Detail[i].GetValue('PFO_DATEDEBUT')));
            T.AddChampSupValeur('FINSTAGE',     DateToStr(DataTob.Detail[i].GetValue('PFO_DATEFIN')));
            T.AddChampSupValeur('HORAIRES',     0);
            T.AddChampSupValeur('NBHEURES',     0);
            T.AddChampSupValeur('HEURESTT',     DataTob.Detail[i].GetValue('PFO_HTPSTRAV'));
            T.AddChampSupValeur('HEURESHTT',    DataTob.Detail[i].GetValue('PFO_HTPSNONTRAV'));
            T.AddChampSupValeur('PRESENT',      '');
        End;

        If (TobDonnees.Detail.Count = 0) Then
            PutGridDetail('GRSALARIE', TobDonnees);

        AfficheTotal (Self, TobDonnees);
        DrawGrid ('GRTOTAL');
    Finally
        FreeAndNil (DataTob);
    End;
end;

end.


