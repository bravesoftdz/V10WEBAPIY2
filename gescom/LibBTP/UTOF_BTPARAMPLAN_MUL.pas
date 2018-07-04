{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/10/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARAMPLAN_MUL ()
Mots clefs ... : TOF;BTPARAMPLAN_MUL
*****************************************************************}
Unit UTOF_BTPARAMPLAN_MUL ;

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
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ; 

Type
  TOF_BTPARAMPLAN_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    
private
    ChxPlanning       : THValComboBox;
    TypePlanning      : String;
    ChxTypeRessource  : THValComboBox;
    StTypeRessource   : THLabel;
    ChxCadencement    : THValComboBox;
    procedure Controlechamp(Champ, Valeur: String);

  end ;

Implementation

procedure TOF_BTPARAMPLAN_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPLAN_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPLAN_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPLAN_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPLAN_MUL.OnArgument (S : String ) ;
var x       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin
  Inherited ;

   ChxPlanning      := ThValcombobox(GetControl('HPP_MODEPLANNING'));
   ChxTypeRessource := ThValcombobox(GetControl('HPP_FAMRES'));
   StTypeRessource  := THLabel(GetControl('THPP_FAMRES'));
   ChxCadencement   := ThValcombobox(GetControl('HPP_CADENCEMENT'));

  Critere := uppercase(Trim(ReadTokenSt(S)));
  while Critere <> '' do
  begin
     x := pos('=', Critere);
     if x <> 0 then
        begin
        Champ  := copy(Critere, 1, x - 1);
        Valeur := copy(Critere, x + 1, length(Critere));
        end
     else
        Champ := Critere;
     ControleChamp(Champ, Valeur);
     Critere:= uppercase(Trim(ReadTokenSt(S)));
  end;

  ChxPlanning.Plus := 'AND CO_ABREGE="' + TypePlanning + '"';

  if TypePlanning = 'PMA' then
  begin
    ChxTypeRessource.DataType := 'BTFAMILLEMAT';
    ChxCadencement.Plus       := ' AND CO_LIBRE="PMA"';
    StTypeRessource.Caption   := 'Famille Matériel :';
  end
  else
  begin
    ChxTypeRessource.DataType := 'BTFAMRES';
    StTypeRessource.Caption   := 'Famille Ressource :';
  end

end ;

Procedure TOF_BTPARAMPLAN_MUL.Controlechamp(Champ, Valeur : String);
begin

  if Champ='TYPEPLANNING' then
  begin
    TypePlanning := Valeur;
  end;

end;

procedure TOF_BTPARAMPLAN_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPLAN_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPLAN_MUL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTPARAMPLAN_MUL ] ) ; 
end.

