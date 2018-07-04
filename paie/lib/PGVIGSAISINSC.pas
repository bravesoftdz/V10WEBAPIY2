unit PGVIGSAISINSC;
{***********UNITE*************************************************
Auteur  ...... : Paie -JL
Créé le ...... : 13/02/2007
Modifié le ... :
Description .. : Saisie d'une formation
               : Vignette : PG_VIG_SAISINSC
               : Tablette : PGPERIODEVIGNETTE
               : Table    : FORMATIONS
Mots clefs ... :
*****************************************************************
PT1  | 18/01/2008 | FLO | Externalisation de l'inscription et ajout de critères dans certaines requêtes
}

interface

uses
  Classes,
  UTob,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  PG_VIG_SAISINSC = class (TGAVignettePaie)
  protected
    procedure RecupDonnees;                                     override;
    function  GetInterface (NomGrille: string = ''): Boolean;   override;
    procedure GetClauseWhere;                                   override;
    procedure DrawGrid (Grille: string);                        override;
    function  SetInterface : Boolean ;                          override;
  private
    procedure PGValiderInsc;
    procedure AfficheSession ;
    procedure InitialiseSaisie;
    Procedure ChargeDonnees;
  end;

implementation
uses
  SysUtils,
  PGVIGUTIL,
  HEnt1,
  PGCalendrier,
  PGOutilsFormation;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_SAISINSC.GetInterface (NomGrille: string): Boolean;
begin
    WarningOk := False;
    Result := inherited GetInterface ('');
    MessageErreur := '';
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_SAISINSC.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_SAISINSC.RecupDonnees;
begin
  inherited;
    If ParamFich = '' then
    Begin
        ChargeDonnees;
        InitialiseSaisie;
    End
    Else If ParamFich = 'VALIDER' then
        PGValiderInsc ()
    Else if (ParamFich = 'STAGE') then
        AfficheSession ;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_SAISINSC.DrawGrid (Grille: string);
begin
  inherited;
end;

{-----Initialisation de la vignette ----------------------------------------------------}

function PG_VIG_SAISINSC.SetInterface: Boolean;
begin
 inherited SetInterface;
    result:=true;
end;

{-----Initialisation de la vignette ----------------------------------------------------}

procedure PG_VIG_SAISINSC.InitialiseSaisie;
begin
    SetControlValue('TYPEPLAN', 'PLF');
    SetControlValue('SALARIE',  '');
    SetControlValue('STAGE',    '');
    SetControlValue('SESSION',  '');
    SetControlEnabled('SESSION',False);
end;

{-----Chargement des salariés sous la responsabilité de l'utilisateur ------------------}

procedure PG_VIG_SAISINSC.ChargeDonnees;
begin
    ChargeSalaries(Self, 'SALARIE', 'RESPONSFOR');
    ChargeStages  (Self, 'STAGE', ' AND PST_CODESTAGE IN (SELECT PSS_CODESTAGE FROM SESSIONSTAGE WHERE PSS_CLOTUREINSC="-")');
end;

{-----Validation d'une inscription -----------------------------------------------------}

procedure PG_VIG_SAISINSC.PGValiderInsc;
var
   MillesimeEC : String;
   Session     : String;
   Temp        : String;
Begin
    Temp        := GetControlValue('SESSION');
    Session     := ReadTokenSt(Temp);
    MillesimeEC := ReadTokenSt(Temp);

    // Contrôles de saisie
    If (GetControlValue('SALARIE') = '') Or ((Session = '-1') Or (Session = '')) Or (GetControlValue('STAGE') = '') Then
    Begin
        MessageErreur := TraduireMemoire('Toutes les zones sont obligatoires.');
        Exit;
    End;

    //PT1
    MessageErreur := InscritSalarieSession (GetControlValue('STAGE'), Session, MillesimeEC, GetControlValue('SALARIE'), GetControlValue('TYPEPLAN'));
    If MessageErreur <> '' Then Exit;

    InitialiseSaisie;
end;

{-----Renseignement des sessions en cours pour la formation choisie---------------------}

procedure  PG_VIG_SAISINSC.AfficheSession ();
var Q : TQuery;
    TobSession,TS : Tob;
    Lib,SQL : String;
begin
    TobSession := Tob.Create('LesSessions',Nil,-1);
    SetControlEnabled('SESSION',True);

    {$IFDEF MULTIDOSSIER}
    SQL := 'SELECT PSS_LIBELLE,PSS_DATEDEBUT,PSS_ORDRE,PSS_MILLESIME,PSS_DATEFIN FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetControlValue('STAGE')+'"'+
           ' AND PSS_CLOTUREINSC="-" AND PSS_VOIRPORTAIL="X" AND PSS_PGTYPESESSION<>"SOS" AND (PSS_NBRESTAGPREV=0 OR PSS_NBRESTAGPREV>(SELECT COUNT (PFO_CODESTAGE) NB FROM FORMATIONS '+ //PT1
           'WHERE PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE))'+
           ' AND NOT (PSS_PREDEFINI="DOS" AND PSS_NODOSSIER<>"'+V_PGI.NoDossier+'")';
    {$ELSE}
    SQL := 'SELECT PSS_LIBELLE,PSS_DATEDEBUT,PSS_ORDRE,PSS_MILLESIME,PSS_DATEFIN FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+GetControlValue('STAGE')+'"'+
           ' AND PSS_CLOTUREINSC="-" AND PSS_VOIRPORTAIL="X" AND PSS_PGTYPESESSION<>"SOS" AND (PSS_NBRESTAGPREV=0 OR PSS_NBRESTAGPREV>(SELECT COUNT (PFO_CODESTAGE) NB FROM FORMATIONS '+ //PT1
           'WHERE PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE))';
    {$ENDIF}
    Q := OpenSQL(SQL,True);
    Q.First;
    If not Q.eof then
    begin
        While Not Q.eof do
        begin
            TS := Tob.Create('~Session',TobSession,-1);
            TS.AddChampSupValeur('VALUE',IntToStr(Q.FindField('PSS_ORDRE').AsInteger)+';'+Q.FindField('PSS_MILLESIME').AsString);
            Lib := Q.FindField('PSS_LIBELLE').AsString+' du '+DateToStr(Q.FindField('PSS_DATEDEBUT').AsDateTime)+' au '+DateToStr(Q.FindField('PSS_DATEFIN').AsDateTime);
            TS.AddChampSupValeur('ITEM',lib);
            Q.Next;
        end;
    end
    else
    begin
        TS := Tob.Create('~Session',TobSession,-1);
        TS.AddChampSupValeur('VALUE','-1');
        TS.AddChampSupValeur('ITEM','Aucune session');
        SetControlEnabled('SESSION',False);
    end;
    Ferme(Q);
    SetComboDetail('SESSION',TobSession);
    FreeAndNil(TobSession);
end;

end.


