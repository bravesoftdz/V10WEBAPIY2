unit UTofTVNouvClients;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF,EntGC,stat,
{$IFNDEF EAGLCLIENT}
      db,mul, dbTables,DBGrids,
{$ENDIF}
      HDimension,M3FP, UTobView, TVProp,HQry
      ;
Type
     TOF_TVNvxClients = Class (TOF)
     private
         TobViewer1: TTobViewer;
         procedure TVOnDblClickCell(Sender: TObject ) ;
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure Onload ; override ;
     END ;



implementation

procedure TOF_TVNvxClients.OnArgument(Arguments : String ) ;
begin
inherited ;
TobViewer1:=TTobViewer(getcontrol('TV'));
TobViewer1.OnDblClick:= TVOnDblClickCell ;
end;

procedure TOF_TVNvxClients.OnLoad ;
var STDATEDPI,STDATEFPI,Requete,Requete1,Requete2,Requete3,Requete4,Requete5,stWhere :string;
    listeIn,listeCode : string;
begin
inherited ;

  ExecuteSQL( 'IF (object_id("tempdb..#NVX") IS not Null) Drop Table #NVX ');
  ExecuteSQL( 'IF (object_id("tempdb..#NVX1") IS not Null) Drop Table #NVX1 ');
  ExecuteSQL( 'IF (object_id("tempdb..#NVX2") IS not Null) Drop Table #NVX2 ');
  ExecuteSQL( 'IF (object_id("tempdb..#NVX3") IS not Null) Drop Table #NVX3 ');
  ExecuteSQL( 'IF (object_id("tempdb..#NVX4") IS not Null) Drop Table #NVX4 ');

  Requete1 := getcontroltext('REQUETE1');
  stWhere:=RecupWhereCritere(TFStat(Ecran).Pages);
  stWhere :=  StringReplace( stWhere,'WHERE',' AND ',[rfIgnoreCase]);
  Requete1 := Requete1 + stWhere;
  ExecuteSQL(Requete1);
  
  STDATEDPI := GetControlText('DATEPIECE');
  STDATEFPI := GetControlText('DATEPIECE_');

  ListeCode := THMultiValcomboBox(GetControl('FAMILLENIV1')).value  ;
  if (trim(ListeCode)<>'') then
  begin
  listeIn := ' GL_FAMILLENIV1 IN ("'+ReadtokenSt(ListeCode)+'"';
  while (trim(ListeCode)<>'') do listeIn := listeIn+',"'+ReadtokenSt(ListeCode)+'"';
  listeIn := listeIn + ')';
  end
  else listeIn := '1=1';

  Requete2 := getcontroltext('REQUETE2');
  Requete2 := StringReplace(Requete2, '#DATEDPI#', USDateTime(StrToDate(STDATEDPI)), [rfReplaceAll]);
  Requete2 := StringReplace(Requete2, '#DATEFPI#', USDateTime(StrToDate(STDATEFPI)), [rfReplaceAll]);
  Requete2 :=  StringReplace( Requete2,'#FAMILLENIV1#',listeIn,[rfIgnoreCase]);
  ExecuteSQL(Requete2);

  Requete3 := getcontroltext('REQUETE3');
  Requete3 := StringReplace(Requete3, '#DATEDPI#', USDateTime(StrToDate(STDATEDPI)), [rfReplaceAll]);
  Requete3 := StringReplace(Requete3, '#DATEFPI#', USDateTime(StrToDate(STDATEFPI)), [rfReplaceAll]);
  ExecuteSQL(Requete3);

  Requete4 := getcontroltext('REQUETE4');
  Requete4 := StringReplace(Requete4, '#DATEDPI#', USDateTime(StrToDate(STDATEDPI)), [rfReplaceAll]);
  Requete4 := StringReplace(Requete4, '#DATEFPI#', USDateTime(StrToDate(STDATEFPI)), [rfReplaceAll]);
  Requete4 :=  StringReplace( Requete4,'#FAMILLENIV1#',listeIn,[rfIgnoreCase]);
  ExecuteSQL(Requete4);

  Requete5 := getcontroltext('REQUETE5');
  ExecuteSQL(Requete5);

  Requete := getcontroltext('REQUETE');
  TFStat(Ecran).FSQL.lines.clear;
  TFStat(Ecran).FSQL.lines[0] := Requete +#$D#$A;
end;

procedure TOF_TVNvxClients.TVOnDblClickCell(Sender: TObject );
begin
with TTobViewer(sender) do
    begin
         V_PGI.DispatchTT (8,taConsult ,AsString[ColIndex( 'T_TIERS' ), CurrentRow], '','');
    end;
end;

Initialization
registerclasses([TOF_TVNvxClients]);
end.
