{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 25/09/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GCDISPOCONTR_MUL ()
Mots clefs ... : TOF;GCDISPOCONTR_MUL
*****************************************************************}
Unit UTOFGCDISPOCONTR_MUL ;

Interface

Uses
     StdCtrls, Controls, Classes,  forms, sysutils, ComCtrls,
{$IFDEF EAGLCLIENT}
      eMul,
{$ELSE}
      mul, db,  dbTables,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, Hqry, UTOF, UTOB ;

Type
  TOF_GCDISPOCONTR_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Implementation

procedure TOF_GCDISPOCONTR_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCDISPOCONTR_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCDISPOCONTR_MUL.OnUpdate ;
var F : TFMul ;
    st1, st2, st3, Where : string ;
begin
F:=TFMul(Ecran);
Where := RecupWhereCritere(F.Pages) ;
st1 := Copy(Where, 0, Pos('GQC_CLIENT', Where) - 1);
st2 := Copy(Where, Pos('GQC_CLIENT', Where), Length(Where));
st3 := Copy(st2, 0, Pos('AND', st2) - 1);
st2 := Copy(st2, Pos('AND', st2), Length(st2));
Where := st1 + '(' + st3 + ' OR GQC_CLIENT="") ' + st2;
  Inherited ;
end ;

procedure TOF_GCDISPOCONTR_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_GCDISPOCONTR_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_GCDISPOCONTR_MUL.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_GCDISPOCONTR_MUL ] ) ; 
end.
