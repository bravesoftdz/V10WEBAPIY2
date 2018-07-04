unit UTofParfou;

interface
uses  M3FP, StdCtrls,Controls,Classes,db,forms,sysutils,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOB,
{$IFDEF EAGLCLIENT} // DBR Fiche 10098
      MaineAgl,
{$ELSE} // EAGLCLIENT
      MajTable, Fe_Main,
{$ENDIF} // EAGLCLIENT
      UTOF, AglInit, Agent,EntGC,
      HTB97,RecupTarifFour, TarifAutoFour,TofParfouDuplic;

Type

     TOF_Parfou = Class (TOF)
       private
       public
       procedure OnArgument(stArgument : String ) ; override ;

     END ;


implementation
/////////////////////////////////////////////////////////////////////////////
procedure TOF_Parfou.OnArgument (stArgument : String ) ;
BEGIN
END;

////////// Paramètrage Recup Auto Tarif Four ///////////////////////////////////
procedure TOF_Parfou_ParamTarifFou (parms: array of variant; nb: integer );
begin
if parms [1] = '' then EntreeTarifAutoFour (taCreat, parms [1])
else EntreeTarifAutoFour (taModif, parms [1]);
end;

procedure TOF_Parfou_RecupTarifFour (parms: array of variant; nb: integer );
begin
EntreeRecupTarifFour (parms [1]);
end;

procedure TOF_Parfou_DuplicTarifFour (parms: array of variant; nb: integer );
begin
DuplicRecupTarifFour (parms [1]);
end;

/////////////////////////////////////////////////////////////////////////////

procedure InitTOFParfou ();
begin
RegisterAglProc ( 'TOF_Parfou_ParamTarifFou', TRUE , 1, TOF_Parfou_ParamTarifFou);
RegisterAglProc ( 'TOF_Parfou_RecupTarifFour', TRUE , 1, TOF_Parfou_RecupTarifFour);
RegisterAglProc ( 'TOF_Parfou_DuplicTarifFour', TRUE , 1, TOF_Parfou_DuplicTarifFour);
end;

Initialization
registerclasses([TOF_Parfou]) ;
InitTOFParfou();
end.
