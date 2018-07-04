{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 07/08/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULCDEXIST ()
Mots clefs ... : TOF;PGMULCDEXIST
*****************************************************************}
Unit PGMULCDEXIST_TOF ;

Interface

Uses
//unused     StdCtrls, 
//unused     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
//unused     db,
//unused     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     HDB,
     FE_Main,
     HQry, //unused
{$ELSE}
     MaineAGL,UTOB,
{$ENDIF}
     AGLInit,
//unused     forms,
     sysutils,
//unused     ComCtrls,
     HTB97,
     HCtrls,
//unused     HEnt1,
//unused     HMsgBox,
//unused     HQry,
     UTOF;
Type
  TOF_PGMULCDEXIST = Class (TOF)
    private
    WW                                  : THEdit;
    TypeTraitement                      : String;
    DateDebut,DateFin                   : TDateTime;
    BtnCherche                          : TToolbarButton97;
    procedure ActiveWhere (Sender: TObject);
    procedure GrilleDblClick (Sender: TObject);

    public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_PGMULCDEXIST.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGMULCDEXIST.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGMULCDEXIST.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGMULCDEXIST.OnLoad ;
begin
  Inherited ;
  ActiveWhere (NIL);
end ;

procedure TOF_PGMULCDEXIST.OnArgument (S : String ) ;
var
  St                                    : string;
{$IFDEF EAGLCLIENT}
  Grille                 : THGrid;
{$ELSE}
  Grille                 : THDBGrid;
{$ENDIF}

begin
  Inherited ;
  St := Trim (S);
  DateDebut := StrToDate(ReadTokenSt(St));  // Recup de la date de debut
  DateFin := StrToDate (ReadTokenSt(St));   // Recup Date fin
  TypeTraitement := (ReadTokenSt(St));      // Récup type de traitement
                                           // (S=Saisie Cde, I=Intégr., C=visu Cdes)

{$IFDEF EAGLCLIENT}
  Grille:=THGrid (GetControl ('Fliste'));
{$ELSE}
  Grille:=THDBGrid (GetControl ('Fliste'));
{$ENDIF}
  if Grille <> NIL then Grille.OnDblClick := GrilleDblClick;
  BtnCherche:=TToolbarButton97 (GetControl ('BCherche'));

  WW := THEdit (GetControl ('XX_WHERE'));
end ;

procedure TOF_PGMULCDEXIST.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGMULCDEXIST.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGMULCDEXIST.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PGMULCDEXIST.ActiveWhere(Sender: TObject);
begin
  WW.Text:= '(NOT '+
        '((PRT_DATEDEBUT < "'+UsDateTime(DateDebut)+'" and '+
        ' PRT_DATEFIN < "'+UsDateTime(DateDebut)+'") or '+
        '(PRT_DATEDEBUT > "'+UsDateTime(DateFin)+'" and '+
        'PRT_DATEFIN > "'+UsDateTime(DateFin)+'"))) AND'+
        '(PRT_DATEDEBUT <> "'+UsDateTime(DateDebut)+'" OR '+
        'PRT_DATEFIN <>"'+UsDateTime(DateFin)+'")';
end;
procedure TOF_PGMULCDEXIST.GrilleDblClick(Sender: TObject);
var
  St                                            : String;
begin

{$IFDEF EAGLCLIENT}
  TheMulQ := TOB(Ecran.FindComponent('Q'));
{$ELSE}
  TheMulQ := THQuery(Ecran.FindComponent('Q'));
{$ENDIF}

  St := DateToStr(DateDebut)+';'+DateToStr(DateFin)+';C';
  AGLLanceFiche ('PAY','SAITICKRESTAU',  '', '', St);
  if BtnCherche <> NIL then BtnCherche.click;
          TheMulQ := NIL;
end;
Initialization
  registerclasses ( [ TOF_PGMULCDEXIST ] ) ;
end.
