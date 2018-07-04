unit UFonctionsCBP;

interface

uses  variants,
  Windows, Messages,SysUtils,Classes,Graphics,Controls,Dialogs,hmsgbox,
{$IFNDEF EAGLCLIENT}
  MajTable,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  OLEAuto,OleDicoPGI,UTob,HEnt1,HCtrls,Forms;

function ExJaiLeDroitConcept (cc : TConcept; Msg : boolean) : boolean;

implementation
                                                                                  
function ExJaiLeDroitConcept (cc : TConcept; Msg : boolean) : boolean;        
{$IFDEF V10}                    
var MnTag : string;                                                           
{$ENDIF}
begin
{$IFNDEF V10}                       
  result := JaiLeDroitConcept (cc, Msg );
{$ELSE}
  MnTag := InttoStr(26000+integer(cc));
  result := true;
  if not ExisteSql ('SELECT 1 FROM MENU WHERE MN_1=26 AND MN_TAG='+MnTag) then exit;
  //
  result :=  ExisteSql ('SELECT 1 FROM MENU WHERE SUBSTRING(MN_ACCESGRP,'+IntToStr(V_PGI.UserGrp)+',1)="0" AND MN_1=26 AND MN_TAG='+MnTag) ;
  if (not Result) and (Msg) then PgiError ('Vous n''avez pas le droit d''effectuer cette opération',Application.name);
{$ENDIF}
end;

end.
