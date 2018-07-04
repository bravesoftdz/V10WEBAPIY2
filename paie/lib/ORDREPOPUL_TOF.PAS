

{***********UNITE*************************************************
Auteur  ...... : NA
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ORDREPOPUL ()
Mots clefs ... : TOF;ORDREPOPUL
*****************************************************************}
Unit ORDREPOPUL_TOF ;

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
     sysutils,
     ComCtrls,
     HCtrls, 
     UTOF ;

Type
  TOF_ORDREPOPUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Arg : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    selectpop : string;
   end ;

Implementation

procedure TOF_ORDREPOPUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_ORDREPOPUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_ORDREPOPUL.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /    
Description .. : Chargement
Mots clefs ... :
*****************************************************************}
procedure TOF_ORDREPOPUL.OnLoad ;
var
listewhere, popul : string;
begin
  Inherited ;
  listewhere := '';
  popul := READTOKENST(selectpop);
  while popul <> '' do                                // lecture de chaque population
  begin
    listewhere :=  listewhere + '"'+popul+'" ,' ;
    popul := READTOKENST(selectpop);
  end;

     if listewhere <> '' then                  // si au moins une population
     listewhere := '(' + copy(listewhere, 1, length(listewhere) - 1) + ')';

  SetControlText('XX_WHERE','PPC_POPULATION in  '+listewhere+'');

end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 09/03/2006
Modifié le ... :   /  /    
Description .. : Alimentation des zones PREDEFINI, NODOSSIER, 
Suite ........ : CODEPOP passés en paramètre
Mots clefs ... : 
*****************************************************************}
procedure TOF_ORDREPOPUL.OnArgument (Arg : String ) ;
var
QQ: Tquery;
Argsimul,Libsimul  : string;
begin
  Inherited ;
Argsimul := ReadTokenSt(Arg);

QQ:= opensql('Select PSL_LIBELLE, PSL_SELECTPOP from SIMULATION where PSL_CODESIMUL = "'+Argsimul+'"', true);
if not (QQ.EOF) then begin
Libsimul := QQ.findfield('PSL_LIBELLE').asstring;
selectpop := QQ.findfield('PSL_SELECTPOP').asstring;
Ecran.caption := 'Populations sélectionnées pour la simulation ' + ' ' +  Argsimul + ' ' + libsimul;
end;
ferme(QQ);
end ;

procedure TOF_ORDREPOPUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_ORDREPOPUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_ORDREPOPUL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_ORDREPOPUL ] ) ; 
end.
