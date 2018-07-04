{***********UNITE*************************************************
Auteur  ...... : YMO
Créé le ...... : 20/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPECRCREPARAM ()
Mots clefs ... : TOF;CPECRCREPARAM
*****************************************************************}
Unit CPECRCREPARAM_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uTOB,
     HTB97,
     AGLInit,
     Ent1,
     SaisUtil,
     ParamSoc;

Type
  TOF_CPECRCREPARAM = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  end ;

Implementation

procedure TOF_CPECRCREPARAM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPECRCREPARAM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPECRCREPARAM.OnUpdate ;
var T : TOB;
begin
    T := TOB.Create ('', nil, -1);
    T.AddChampSupValeur('EMP_DATEDEB',StrToDate(GetControlText('EMP_DATEDEB')));
    T.AddChampSupValeur('EMP_DATEFIN',GetControlText('EMP_DATEFIN'));
    T.AddChampSupValeur('EMP_TYPE', GetControlText('EMP_TYPE'));
    T.AddChampSupValeur('EMP_INTER', GetControlText('EMP_INTER'));
    T.AddChampSupValeur('EMP_ASSUR', GetControlText('EMP_ASSUR'));

  TheTOB := T;

  Ecran.Close ;
end ;

procedure TOF_CPECRCREPARAM.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPECRCREPARAM.OnArgument (S : String ) ;
begin
  Inherited ;
  SetControlText('EMP_DATEDEB', DateToStr(V_PGI.DateEntree));
  SetControlText('EMP_DATEFIN', DateToStr(V_PGI.DateEntree));
  SetControlText('EMP_TYPE', 'A');
  SetControlText('EMP_INTER', 'X');
  SetControlText('EMP_ASSUR', 'X');

  TheTOB := nil;
end ;

procedure TOF_CPECRCREPARAM.OnClose ;
begin
  Inherited ;
end ;


Initialization
  registerclasses ( [ TOF_CPECRCREPARAM ] ) ;
end.
