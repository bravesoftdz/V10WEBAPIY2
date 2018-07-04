unit UtilPerspective;

interface

uses UTOB,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     Fe_Main,
{$ENDIF}
     SysUtils, HCtrls, Hent1,Paramsoc;

Function RTAffecterPerspective ( TobPiece : Tob ) : boolean;
Procedure PERSPECTIVESTOBCopieChamp(FromTOB, ToTOB : TOB);

implementation
uses  CbpMCD
	  ,CbpEnumerator
    ;


Function RTAffecterPerspective ( TobPiece : Tob ) : boolean;
var StRetour,NoPersp,Maj : string ;
begin
result:=false;
if GetParamsocSecur('SO_RTPROPCHAINDEVIS',False) = true then
    begin
    StRetour:=AGLLanceFiche ('RT','RTPERSPECTIVE_REC','RPE_TIERS='+TobPiece.getValue('GP_TIERS'),'','MONTANTPROPAL='+StrFPoint(TobPiece.GetValue('GP_TOTALHT'))) ;
    NoPersp:=ReadTokenSt(StRetour);
    Maj:= ReadTokenSt(StRetour);
    If NoPersp<>'' then
      begin
      if Maj = '1' then
         ExecuteSQL('UPDATE PERSPECTIVES SET RPE_MONTANTPER=RPE_MONTANTPER+'+StrFPoint(TobPiece.GetValue('GP_TOTALHT'))+' WHERE RPE_PERSPECTIVE='+NoPersp);
      TOBPiece.PutValue('GP_PERSPECTIVE',strToInt(NoPersp)) ;
      result:=True;
      end
    else
      if GetParamsocSecur('SO_RTPROPDEVISSSPROP',True) = true then
         result:=True;
    end
else
    result:=True;
end;

{****************************************************************
Auteur  ...... : AB
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : Copie les valeurs de champs d'une TOB vers une autre
Mots clefs ... : TOB;COPIE;CHAMPS
*****************************************************************}
procedure PERSPECTIVESTOBCopieChamp(FromTOB, ToTOB : TOB);
var iChamp , iTableLigne, i_pos: integer;
    FieldNameTo,FieldNameFrom,St:string;
		Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();
Table := Mcd.GetTable(Mcd.PrefixeToTable('RPE'));
FieldList := Table.Fields;
FieldList.Reset();

While FieldList.MoveNext do
  begin
  FieldNameFrom := (FieldList.Current as IFieldCOM).name;
  St := FieldNameFrom ;
  i_pos := Pos ('_', St) ;
  Delete (St, 1, i_pos-1) ;
  FieldNameTo := 'RPH'+ St ;
  ToTOB.PutValue(FieldNameTo, FromTOB.GetValue(FieldNameFrom));
  end;
end;

end.
