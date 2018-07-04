{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - FLO
Créé le ...... : 24/01/2008
Modifié le ... :   /  /
Description .. : Comparatif prévisionnel / réalisé
               : Vignette : PG_VIG_COMPAPR
               : Table    : INSCFORMATION / FORMATIONS
Mots clefs ... :
*****************************************************************
PT1  | 25/03/2008 | FLO | Correctif pour la gestion des types utilisateurs
PT2  | 29/04/2008 | FLO | Prise en compte des salariés non sortis
}
unit PGVIGCOMPAPR;

interface
uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_COMPAPR= class (TGAVignettePaie)
  protected
    procedure RecupDonnees;                                     override;
    function  GetInterface (NomGrille: string = ''): Boolean;   override;
    procedure GetClauseWhere;                                   override;
    procedure DrawGrid (Grille: string);                        override;
    function  SetInterface : Boolean ;                          override;
  private
    procedure AfficheListe;
 end;

implementation
uses
  SysUtils,
  PGVIGUTIL,
  uToolsPlugin,
  uToolsOWC,
  PGOutilsFormation,
  HEnt1;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_COMPAPR.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
    Result := true;

    // Initialisation de la vignette au niveau des périodes pour simuler la sélection d'un millésime
    If FParam = '' Then
    Begin
        Periode := 'A';
        MajLibPeriode();
    End;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_COMPAPR.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_COMPAPR.RecupDonnees;
begin
  inherited;
    AfficheListe;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_COMPAPR.DrawGrid (Grille: string);
begin
  inherited;

    If Grille <> 'GRTITRE' Then
    Begin
        SetFormatCol(Grille, 'NBHEURESPREV',    'D 0 ---');
        SetFormatCol(Grille, 'COUTSPREV',       'D 0 ---');
        SetFormatCol(Grille, 'NBHEURESREAL',    'D 0 ---');
        SetFormatCol(Grille, 'COUTSREAL',       'D 0 ---');
        SetFormatCol(Grille, 'NBHEURESECAR',    'D 0 ---');
        SetFormatCol(Grille, 'COUTSECAR',       'D 0 ---');

        SetWidthCol (Grille, 'CODESTAGE',       -1);
            SetWidthCol (Grille, 'STAGE',           150);
        SetWidthCol (Grille, 'NBHEURESPREV',    50);
        SetWidthCol (Grille, 'COUTSPREV',       50);
        SetWidthCol (Grille, 'NBHEURESREAL',    50);
        SetWidthCol (Grille, 'COUTSREAL',       50);
        SetWidthCol (Grille, 'NBHEURESECAR',    50);
        SetWidthCol (Grille, 'COUTSECAR',       50);
    End
    Else
    Begin
        // Petite subtilité : Pour avoir 2 lignes de titre pour la grille, il y a 2 grilles
        // dont une qui ne sert que pour la 1ère ligne de titre.
        SetWidthCol ('GRTITRE', 'VIDE',     152);
        SetWidthCol ('GRTITRE', 'PREV',     100);
        SetWidthCol ('GRTITRE', 'REAL',     100);
        SetWidthCol ('GRTITRE', 'ECART',    100);
    End;
end;

{-----Initialisation de la vignette ----------------------------------------------------}

function PG_VIG_COMPAPR.SetInterface: Boolean;
begin
  inherited SetInterface;
    result:=true;
end;

{-----Affichage des données de la vignette ---------------------------------------------}

procedure PG_VIG_COMPAPR.AfficheListe;
var
  StSQL,Resp,sN1         : String;
  DataTobPrev,DataTobReal: TOB;
  EnDateDu, EnDateAu     : TDatetime;
  i                      : integer;
  T,T2                   : TOB;
  Somme                  : Double;
begin
    Resp := V_PGI.UserSalarie;
    DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode,sN1);

    Try
        // Prévisionnel
        DataTobPrev := TOB.Create('~TEMPPREV', nil, -1);

        StSQL := 'SELECT PFI_CODESTAGE,PST_LIBELLE,PST_LIBELLE1,SUM(PFI_NBINSC) AS NBINSC,SUM(PFI_DUREESTAGE) AS DUREESTAGE,SUM(PFI_COUTREELSAL) AS COUTREELSAL,SUM(PFI_AUTRECOUT) AS AUTRECOUT,SUM(PFI_FRAISFORFAIT) AS FRAISFORFAIT '+
                 'FROM INSCFORMATION '+
                 'LEFT JOIN STAGE ON PFI_CODESTAGE=PST_CODESTAGE AND PST_MILLESIME="0000" '+
                 'WHERE PFI_ETATINSCFOR="VAL" AND PFI_CODESTAGE<>"--CURSUS--" '+
                 //'AND PFI_RESPONSFOR="'+Resp+'" '+
                 'AND '+AdaptByTypeResp(RecupTypeUtilisateur('FOR'),V_PGI.UserSalarie,False,'PFI',False) + //PT1
                 ' AND PFI_MILLESIME="'+sN1+'" '+
                 'GROUP BY PFI_CODESTAGE,PST_LIBELLE,PST_LIBELLE1';
		
		StSQL := AdaptMultiDosForm (StSQL); //PT2
		
        DataTobPrev := OpenSelectInCache (StSQL);
        ConvertFieldValue(DataTobPrev);

        // Réalisé
        DataTobReal := TOB.Create('~TEMPREAL', nil, -1);

        StSQL := 'SELECT PFO_CODESTAGE,PST_LIBELLE,PST_LIBELLE1,SUM(PFO_NBREHEURE) AS DUREESTAGE,SUM(PFO_COUTREELSAL) AS COUTREELSAL,SUM(PFO_AUTRECOUT) AS AUTRECOUT,SUM(PFO_FRAISPLAF) AS FRAISPLAF '+
                 'FROM FORMATIONS '+
                 'LEFT JOIN STAGE ON PFO_CODESTAGE=PST_CODESTAGE AND PST_MILLESIME="0000" '+
                 'WHERE PFO_EFFECTUE="X" AND PFO_PGTYPESESSION<>"SOS" '+
                 //'AND PFO_RESPONSFOR="'+Resp+'" '+ 
                 'AND '+AdaptByTypeResp(RecupTypeUtilisateur('FOR'),V_PGI.UserSalarie,False,'PFO',False) + //PT1
                 ' AND PFO_DATEDEBUT>="'+UsDateTime(EnDateDu)+'" AND PFO_DATEDEBUT<="'+UsDateTime(EnDateAu)+'" '+
                 'GROUP BY PFO_CODESTAGE,PST_LIBELLE,PST_LIBELLE1';

		StSQL := AdaptMultiDosForm (StSQL); //PT2
		
        DataTobReal := OpenSelectInCache (StSQL);
        ConvertFieldValue(DataTobReal);

        // Comparatif
        For i:=0 To DataTobPrev.Detail.Count-1 Do
        Begin
            T := TOB.Create('£REPONSE',TobDonnees,-1);
            T.AddChampSupValeur('CODESTAGE',    DataTobPrev.Detail[i].GetValue('PFI_CODESTAGE'));
            T.AddChampSupValeur('STAGE',        DataTobPrev.Detail[i].GetValue('PST_LIBELLE')+' '+DataTobPrev.Detail[i].GetValue('PST_LIBELLE1'));

            T.AddChampSupValeur('NBHEURESPREV', DataTobPrev.Detail[i].GetValue('DUREESTAGE'));
            Somme := DataTobPrev.Detail[i].GetValue('COUTREELSAL') +
                     DataTobPrev.Detail[i].GetValue('AUTRECOUT') +
                     DataTobPrev.Detail[i].GetValue('FRAISFORFAIT');
            T.AddChampSupValeur('COUTSPREV',    Somme);

            T.AddChampSupValeur('NBHEURESREAL', 0);
            T.AddChampSupValeur('COUTSREAL',    0);

            T.AddChampSupValeur('NBHEURESECAR', DataTobPrev.Detail[i].GetValue('DUREESTAGE'));
            T.AddChampSupValeur('COUTSECAR',    Somme);
        End;

        For i:=0 To DataTobReal.Detail.Count-1 Do
        Begin
            T2 := TobDonnees.FindFirst(['CODESTAGE'],[DataTobReal.Detail[i].GetValue('PFO_CODESTAGE')], False);
            If T2 <> Nil Then
            Begin
                T2.PutValue('NBHEURESREAL', DataTobReal.Detail[i].GetValue('DUREESTAGE'));
                T2.PutValue('NBHEURESECAR', T2.GetValue('NBHEURESPREV') - T2.GetValue('NBHEURESREAL'));
                Somme := DataTobReal.Detail[i].GetValue('COUTREELSAL') +
                         DataTobReal.Detail[i].GetValue('AUTRECOUT') +
                         DataTobReal.Detail[i].GetValue('FRAISPLAF');
                T2.PutValue('COUTSREAL',    Somme);
                T2.PutValue('COUTSECAR',    T2.GetValue('COUTSPREV') - T2.GetValue('COUTSREAL'));
            End
            Else
            Begin
                T := TOB.Create('£REPONSE',TobDonnees,-1);
                T.AddChampSupValeur('CODESTAGE',    DataTobReal.Detail[i].GetValue('PFO_CODESTAGE'));
                T.AddChampSupValeur('STAGE',        DataTobReal.Detail[i].GetValue('PST_LIBELLE')+' '+DataTobReal.Detail[i].GetValue('PST_LIBELLE1'));

                T.AddChampSupValeur('NBHEURESPREV', 0);
                T.AddChampSupValeur('COUTSPREV',    0);

                T.AddChampSupValeur('NBHEURESREAL', DataTobReal.Detail[i].GetValue('DUREESTAGE'));
                Somme := DataTobReal.Detail[i].GetValue('COUTREELSAL') +
                         DataTobReal.Detail[i].GetValue('AUTRECOUT') +
                         DataTobReal.Detail[i].GetValue('FRAISPLAF');
                T.AddChampSupValeur('COUTSREAL',    Somme);

                T.AddChampSupValeur('NBHEURESECAR', -DataTobReal.Detail[i].GetValue('DUREESTAGE'));
                T.AddChampSupValeur('COUTSECAR',    -Somme);
            End;
        End;

		TobDonnees.Detail.Sort('STAGE');
		
        If (TobDonnees.Detail.Count = 0) Then
            PutGridDetail('FListe', TobDonnees);

        // Total
        AfficheTotal (Self, TobDonnees, 1);
        DrawGrid ('GRTOTAL');

        // Titres
        T2 := TOB.Create('$DUMMY', Nil, -1);
        T := TOB.Create('$LIGNEDUMMY', T2, -1);
        T.AddChampSupValeur('VIDE',     'x');
        T.AddChampSupValeur('PREV',     'x');
        T.AddChampSupValeur('REAL',     'x');
        T.AddChampSupValeur('ECART',    'x');
        PutGridDetail('GRTITRE', T2);
        DrawGrid ('GRTITRE');
    finally
        FreeAndNil (DataTobPrev);
        FreeAndNil (DataTobReal);
    end;
end;

end.


