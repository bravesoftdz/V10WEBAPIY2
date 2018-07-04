{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/03/2006
Modifié le ... :   /  /    
Description .. : Source TOF de la FICHE : PGDADSDATE ()
Mots clefs ... : TOF;PGDADSDATE
*****************************************************************}
unit UTOFPG_DADSDATE;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
{$else}
     eMul,
     uTob,
{$ENDIF}
     sysutils,
     UTOF,
     PgDADSCommun,
     hmsgbox,
     HEnt1;

Type
  TOF_PGDADSDATE = Class (TOF)
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Arg : string;

  end ;

Implementation

procedure TOF_PGDADSDATE.OnNew ;
begin
Inherited ;
end ;

procedure TOF_PGDADSDATE.OnDelete ;
begin
Inherited ;
end ;

procedure TOF_PGDADSDATE.OnUpdate;
var
DateValide : boolean;
DateDeb, DateFin : TDateTime;
begin
Inherited;
DateValide:= True;

DateDeb:= StrToDate (GetControlText ('DD1'));
DateFin:= StrToDate (GetControlText ('DD2'));

if (DateValide=False) then
   PGIBox ('La saisie est invalide', Ecran.Caption)
else
   begin
   if (DateDeb<PlusMois (DebExer, -1)) then
      begin
      PGIBox ('La date de début saisie est invalide', Ecran.Caption);
      end
   else
   if (DateFin>FinExer) then
      begin
      PGIBox ('La date de fin saisie est invalide', Ecran.Caption);
      end
   else
      begin
      DebExer:= StrToDate (GetControlText ('DD1'));
      FinExer:= StrToDate (GetControlText ('DD2'));
      end;
   end;
end ;

procedure TOF_PGDADSDATE.OnLoad ;
begin
Inherited ;
end ;

procedure TOF_PGDADSDATE.OnArgument (S : String ) ;
begin
Inherited;
Arg:= S;
SetControlText ('DD1', DateToStr(DebExer));
SetControlText ('DD2', DateToStr(FinExer));
end ;

procedure TOF_PGDADSDATE.OnClose ;
begin
Inherited;
end ;

procedure TOF_PGDADSDATE.OnDisplay () ;
begin
Inherited ;
end ;

procedure TOF_PGDADSDATE.OnCancel () ;
begin
Inherited ;
end ;

Initialization
registerclasses ( [ TOF_PGDADSDATE ] ) ;
end.
