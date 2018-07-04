{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - MF
Créé le ...... : 12/06/2006
Modifié le ... :   /  /
Description .. : liste des intérimaires de la période
               : Vignette : PG_VIG_INTERIM
               : Tablette : PGPERIODEVIGNETTE
               : Table    : INTERIMAIRES, EMPLOIINTERIM
Mots clefs ... :
*****************************************************************
PT1  | 18/01/2008 | FLO | Prise en compte des périodes + Ajout total
PT2  | 14/03/2008 | FLO | Gestion multi-dossier formation
PT3  | 25/03/2008 | FLO | Correctif pour la gestion des types utilisateurs
PT4  | 29/04/2008 | FLO | Prise en compte des salariés non sortis
PT5  | 14/05/2008 | FLO | Ajout de la colonne reponsable pour les vignettes accessible par l'assistant ou le secrétaire
}
unit PGVIGCOMPTDIF;

interface
uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_COMPTDIF= class (TGAVignettePaie)
  protected
    procedure RecupDonnees;                                      override;
    function  GetInterface (NomGrille: string = ''): Boolean;    override;
    procedure GetClauseWhere;                                    override;
    procedure DrawGrid (Grille: string);                         override;
    function  SetInterface : Boolean ;                           override;
  private
    procedure AfficheListeDif(AParam : string);
  end;

implementation
uses
  SysUtils,
  PGVIGUTIL,
  PGOutilsFormation,
  uToolsPlugin,
  uToolsOWC,
  HEnt1;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_COMPTDIF.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
    Result := true;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_COMPTDIF.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_COMPTDIF.RecupDonnees;
begin
  inherited;
    AfficheListeDif('');
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_COMPTDIF.DrawGrid (Grille: string);
var S : String;
begin
  inherited;

     SetFormatCol(Grille, 'ACQUIS', 'D 0 ---');
     SetFormatCol(Grille, 'PRISTT', 'D 0 ---');
     SetFormatCol(Grille, 'PRISHT', 'D 0 ---');
     SetFormatCol(Grille, 'SOLDE',  'D 0 ---');
     SetFormatCol(Grille, 'DDVTT',  'D 0 ---');
     SetFormatCol(Grille, 'DDVHT',  'D 0 ---');
     SetFormatCol(Grille, 'DDATT',  'D 0 ---');
     SetFormatCol(Grille, 'DDAHT',  'D 0 ---');
     SetFormatCol(Grille, 'SOLDETHEO', 'D 0 ---');

     SetWidthCol (Grille, 'NOM',    60);
     SetWidthCol (Grille, 'PRENOM', 60);
     SetWidthCol (Grille, 'ACQUIS', 50);
     SetWidthCol (Grille, 'PRISTT', 50);
     SetWidthCol (Grille, 'PRISHT', 50);
     SetWidthCol (Grille, 'SOLDE',  50);
     SetWidthCol (Grille, 'DDVTT',  50);
     SetWidthCol (Grille, 'DDVHT',  50);
     SetWidthCol (Grille, 'DDATT',  50);
     SetWidthCol (Grille, 'DDAHT',  50);
     SetWidthCol (Grille, 'SOLDETHEO', 50);
     
    //PT5
    S := RecupTypeUtilisateur('FOR');
    If (Pos('SECRETAIRE', S) > 0) Or (Pos('ADJOINT', S) > 0) Then
    	SetWidthCol (Grille, 'RESPONSABLE', 80)
    Else
    	SetVisibleCol (Grille, 'RESPONSABLE', False);
end;

{-----Initialisation de la vignette ----------------------------------------------------}

function PG_VIG_COMPTDIF.SetInterface: Boolean;
begin
  inherited SetInterface;
    result:=true;
end;

{-----Affichage des données de la vignette ---------------------------------------------}

procedure PG_VIG_COMPTDIF.AfficheListeDif(AParam : string);
var
  StSQL,Salarie,sN1 : String;
  DataTob                : TOB;
  i                      : integer;
  T                      : TOB;
  Acquis,PrisSTT,PrisHTT :Double;
  DdVSTT,DdVHTT,DdASTT,DdAHTT : Double;
  Q                      : TQuery;
  EnDateDu, EnDateAu     : TDatetime;
begin
    //Resp := V_PGI.UserSalarie;
    DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode,sN1);  //PT1
    
    Try
        DataTob := TOB.Create('~TEMP', nil, -1);
    
        StSQL := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,';
        //PT5
        If PGBundleHierarchie Then
            StSQL := StSQL + '(SELECT PSI_LIBELLE||" "||PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE=PSE_RESPONSFOR) AS RESPONSABLE '
        Else
            StSQL := StSQL + '(SELECT PSA_LIBELLE||" "||PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE=PSE_RESPONSFOR) AS RESPONSABLE ';
        StSQL := StSQL + 'FROM SALARIES LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE '+
                 'WHERE '+ //PSE_RESPONSFOR="'+Resp+'"'+
                 AdaptByTypeResp(RecupTypeUtilisateur('FOR'),V_PGI.UserSalarie,False,'PSE',False) + //PT3
                 ' AND ' + SalariesNonSortis + //PT4
                 ' ORDER BY RESPONSABLE,PSA_LIBELLE';
                 
		StSQL := AdaptMultiDosForm (StSQL); //PT2
		
        DataTob := OpenSelectInCache (StSQL);
        ConvertFieldValue(DataTob);

        for i:=0 to DataTob.detail.count-1 do
        begin
            DdVSTT := 0;
            DdVHTT := 0;
            DdASTT := 0;
            DdAHTT := 0;
            Acquis := 0;
            PrisSTT := 0;
            PrisHTT := 0;

            T := TOB.Create('£REPONSE',TobDonnees,-1);
            Salarie := DataTob.Detail[i].GetValue('PSA_SALARIE');
            T.AddChampSupValeur('NOM',   DataTob.Detail[i].GetValue('PSA_LIBELLE'));
            T.AddChampSupValeur('PRENOM',DataTob.Detail[i].GetValue('PSA_PRENOM'));
            
            T.AddChampSupValeur('RESPONSABLE',  DataTob.Detail[i].GetValue('RESPONSABLE')); //PT5

            Q := OpenSQL('SELECT SUM(PHC_MONTANT) ACQUIS '+
            'FROM HISTOCUMSAL WHERE PHC_DATEDEBUT>="'+UsDateTime(EnDateDu)+'" AND PHC_DATEFIN<="'+UsDateTime(EnDateAu)+'"'+ //PT1
            ' AND PHC_CUMULPAIE=(SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_PGCUMULDIFACQUIS") AND PHC_SALARIE="'+Salarie+'"',True);
            If Not Q.eof then Acquis := Q.FindField('ACQUIS').AsFloat;
            Ferme(Q);

            T.AddChampSupValeur('ACQUIS',Acquis);
            Q := OpenSQL('SELECT SUM (PFO_HTPSTRAV) STT,SUM (PFO_HTPSNONTRAV) HTT FROM FORMATIONS '+
            'WHERE PFO_SALARIE="'+Salarie+'" AND PFO_EFFECTUE="X" AND PFO_TYPEPLANPREV="DIF"'+
            'AND PFO_DATEDEBUT>="'+UsDateTime(EnDateDu)+'" AND PFO_DATEFIN<="'+UsDateTime(EnDateAu)+'"',True);    //PT1
            If Not Q.eof then
            begin
                PrisSTT := Q.FindField('STT').AsFloat;
                PrisHTT := Q.FindField('HTT').AsFloat;
            end;
            Ferme(Q);

            Q := OpenSQL('SELECT SUM (PFI_HTPSTRAV) STT,SUM (PFI_HTPSNONTRAV) HTT FROM INSCFORMATION '+
            'WHERE PFI_SALARIE="'+Salarie+'" AND PFI_ETATINSCFOR="VAL" AND PFI_REALISE="-" AND PFI_TYPEPLANPREV="DIF"'+
            ' AND PFI_DATEDIF>="'+UsDateTime(EnDateDu)+'" AND PFI_DATEDIF<="'+UsDateTime(EnDateAu)+'"',True);     //PT1
            If Not Q.eof then
            begin
                DdVSTT := Q.FindField('STT').AsFloat;
                DdVHTT := Q.FindField('HTT').AsFloat;
            end;
            Ferme(Q);

            Q := OpenSQL('SELECT SUM (PFI_HTPSTRAV) STT,SUM (PFI_HTPSNONTRAV) HTT FROM INSCFORMATION '+
            'WHERE PFI_SALARIE="'+Salarie+'" AND PFI_ETATINSCFOR="ATT" AND PFI_REALISE="-" AND PFI_TYPEPLANPREV="DIF"'+
            ' AND PFI_DATEDIF>="'+UsDateTime(EnDateDu)+'" AND PFI_DATEDIF<="'+UsDateTime(EnDateAu)+'"',True);    //PT1
            If Not Q.eof then
            begin
                DdASTT := Q.FindField('STT').AsFloat;
                DdAHTT := Q.FindField('HTT').AsFloat;
            end;
            Ferme(Q);

            T.AddChampSupValeur('PRISTT',PrisSTT);
            T.AddChampSupValeur('PRISHT',PrisHTT);
            T.AddChampSupValeur('SOLDE',Acquis - PrisSTT -PrisHTT);
            T.AddChampSupValeur('DDVTT',DdVSTT);
            T.AddChampSupValeur('DDVHT',DdVHTT);
            T.AddChampSupValeur('DDATT',DdASTT);
            T.AddChampSupValeur('DDAHT',DdAHTT);
            T.AddChampSupValeur('SOLDETHEO',Acquis - PrisSTT -PrisHTT -DdVSTT - DdVHTT - DdASTT - DdAHTT);
        end;

        if (TobDonnees.detail.count = 0) then
            PutGridDetail('GRSALARIE', TobDonnees);

        AfficheTotal (Self, TobDonnees); //PT1
        DrawGrid ('GRTOTAL');
    
    finally
        FreeAndNil (DataTob);
    end;
end;

end.


