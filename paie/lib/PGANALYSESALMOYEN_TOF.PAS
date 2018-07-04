{***********UNITE*************************************************
Auteur  ...... :  NA
Créé le ...... : 30/05/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGANALYSESALMOYEN ()
                 Analyse des salaires moyens
Mots clefs ... : TOF;PGANALYSESALMOYEN
*****************************************************************}
Unit PGANALYSESALMOYEN_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     P5def,
     sysutils,
     ComCtrls,
     UTOF ;

Type
  TOF_PGANALYSESALMOYEN = Class (TOF)
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

procedure TOF_PGANALYSESALMOYEN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSESALMOYEN.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSESALMOYEN.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSESALMOYEN.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 30/05/2006
Modifié le ... :   /  /    
Description .. : on Argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGANALYSESALMOYEN.OnArgument (S : String ) ;
var
num : integer;
begin
  Inherited ;
 // Recherche des libellés des zones TRAVAIL, CODESTAT, LIBREPCMB
  for Num := 1 to 4 do
  begin
  VisibiliteChampSalarie(IntToStr(Num), GetControl('PSY_TRAVAILN' + IntToStr(Num)), GetControl('TPSY_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSY_CODESTAT'), GetControl('TPSY_CODESTAT'));
   
  for Num := 1 to 4 do
  begin
  VisibiliteChamplibresal(Inttostr(Num), GetControl('PSY_LIBREPCMB'+ Inttostr(Num)),GetControl('TPSY_LIBREPCMB'+ IntToStr(Num)));
  end;
end ;

procedure TOF_PGANALYSESALMOYEN.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSESALMOYEN.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSESALMOYEN.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_PGANALYSESALMOYEN ] ) ; 
end.
