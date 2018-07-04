{***********UNITE*************************************************
Auteur  ...... : PAIE-MF
Créé le ...... : 18/06/2003
Modifié le ... : 01/07/2003
Description .. : Source TOF de la FICHE : ANALYSE_CDETCK ()
Suite ........ : TobViewer d'analyse de la table CDETICKETS
Mots clefs ... : TOF;PGANALYSECDETCK
*****************************************************************}
Unit PGANALYSECDETCK_TOF ;

Interface

Uses
//unused     StdCtrls, 
//unused     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
//unused     db,
//unused     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
//unused     forms,
     sysutils,
//unused     ComCtrls,
     HCtrls,
//unused     HEnt1,
//unused     HMsgBox,
     ParamSoc,
     PgOutils2,
     P5Def,
     UTOF ;

Type
  TOF_PGANALYSECDETCK = Class (TOF)
    private
    DateDeb,DateFin                     : THEdit;
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

procedure TOF_PGANALYSECDETCK.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSECDETCK.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSECDETCK.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSECDETCK.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 01/07/2003
Modifié le ... : 01/07/2003
Description .. : procédure OnArgument   
Suite ........ : Les date de période sont proposée par défaut
Suite ........ : Sur l'onglet complément présentation des champs libres 
Suite ........ : éventuels
Mots clefs ... : PAIE; PGANALYSECDETCK
*****************************************************************}
procedure TOF_PGANALYSECDETCK.OnArgument (S : String ) ;
var
  DebPer,FinPer,ExerPerEncours              : String;
  OkOk                                      : Boolean;
  Num                                       : integer;

begin
  Inherited ;
  SetControltext('CODECLIENT',GetParamSocSecur('SO_PGCODECLIENT',''));
  DateDeb := ThEdit(getcontrol('XX_VARIABLEDEB'));
  DateFin := ThEdit(getcontrol('XX_VARIABLEFIN'));


  OkOk := RendPeriodeEnCours (ExerPerEncours,DebPer,FinPer);
  if OkOk then
  begin
   if  DateDeb <> NIL   then DateDeb.text:= DebPer;
   if  DateFin <> NIL   then DateFin.text:= FinPer;
  end;

  for Num := 1 to 4 do
  begin
    VisibiliteChampSalarie (IntToStr(Num),
                            GetControl ('PRT_TRAVAILN'+IntToStr(Num)),
                            GetControl ('TPRT_TRAVAILN'+IntToStr(Num)));
    VisibiliteChampLibreX(IntToStr(Num),GetControl ('PRT_LIBRE'+IntToStr(Num)),
                            GetControl ('TPRT_LIBRE'+IntToStr(Num)));
    VisibiliteBoolLibreSal (IntToStr(Num),
                            GetControl ('PRT_BOOLIBRE'+IntToStr(Num)));
  end;

end ;

procedure TOF_PGANALYSECDETCK.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSECDETCK.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSECDETCK.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_PGANALYSECDETCK ] ) ; 
end.
