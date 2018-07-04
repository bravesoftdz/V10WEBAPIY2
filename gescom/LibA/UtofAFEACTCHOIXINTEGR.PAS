{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/06/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFEACTCHOIXINTEGR ()
Mots clefs ... : TOF;AFEACTCHOIXINTEGR
*****************************************************************}
Unit UtofAFEACTCHOIXINTEGR ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFDEF EAGLCLIENT}
       Maineagl,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db,Fe_Main ,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ; 

Type
  TOF_AFEACTCHOIXINTEGR = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Function AFLanceFiche_AFEACTCHOIXINTEGR(Argument:string):Variant;

Implementation

procedure TOF_AFEACTCHOIXINTEGR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFEACTCHOIXINTEGR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFEACTCHOIXINTEGR.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFEACTCHOIXINTEGR.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFEACTCHOIXINTEGR.OnArgument (S : String ) ;
var
    Critere, Champ, valeur  : String;
    x : integer;
begin
Inherited ;


Critere:=(Trim(ReadTokenSt(S)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        X:=pos(':',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy(Critere,X+1,length(Critere)-X);
           end;
        if Champ = 'DATEDEB' then
            begin
            SetControlText('DATEDEBINTEGRATION', Valeur);
            SetControlText('DATEDEBSUPP', Valeur);
            end;
        if Champ = 'DATEFIN' then
            begin
            SetControlText('DATEFININTEGRATION', Valeur);
            SetControlText('DATEFINSUPP', Valeur);
            end;
        END;
    Critere:=(Trim(ReadTokenSt(S)));
    END;

// Réduction des listes de choix
THMultiValComboBox(GetControl('ACTIVITEREPRISE')).Plus := 'AND CO_CODE<>"FAC"';
THValComboBox(GetControl('ORIGINE')).Plus := 'AND CC_CODE<>"PLA" AND CC_CODE<>"ACH" AND CC_CODE<>"REP" AND CC_CODE<>"SDE"';

// Initialisation des champs critères de suppression
THMultiValComboBox(GetControl('ACTIVITEREPRISE')).Text := 'A';
THValComboBox(GetControl('ORIGINE')).value := 'EAC';
end;

procedure TOF_AFEACTCHOIXINTEGR.OnClose ;
begin
  Inherited ;
end ;

Function AFLanceFiche_AFEACTCHOIXINTEGR(Argument:string):Variant;
begin
result:= AglLanceFiche( 'AFF', 'AFEACTCHOIXINTEGR', '','',Argument) ;
end;

Initialization
  registerclasses ( [ TOF_AFEACTCHOIXINTEGR ] ) ;
end.
