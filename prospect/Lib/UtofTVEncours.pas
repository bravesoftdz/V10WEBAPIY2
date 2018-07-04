unit UtofTVEncours;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF,stat,HStatus,AglInit,
{$IFDEF EAGLCLIENT}
      Maineagl,emul,
{$ELSE}
      db,mul, {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}DBGrids,Fe_Main,
{$ENDIF}
      HDimension,M3FP, UTobView, TVProp, UTob,TiersUtil
      ;
      
Type
     TOF_TVEncours = Class (TOF)
     private
         FTV : TFStat ;
         TobViewer1 : TTobViewer;
         procedure TVOnDblClickCell(Sender: TObject ) ;
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnUpdate ; override ;
        procedure OnClose ; override ;
     END ;



implementation

procedure TOF_TVEncours.OnArgument(Arguments : String ) ;
begin
inherited ;
FTV := TFStat(Ecran);
TobViewer1:=TTobViewer(getcontrol('TV'));
TobViewer1.OnDblClick:= TVOnDblClickCell ;
end;

procedure TOF_TVEncours.OnClose;
begin
inherited  ;
FTV.LaTob.free;
end;

procedure TOF_TVEncours.OnUpdate;
var Q : TQuery ;
    i : integer ;
    TB :TOB ;
    RisqueTiers, P1, P2 : double;
    EtatRisque : string;
begin
inherited ;
Q:=OpenSQL(FTV.stSQL,true);
FTV.LaTob:=TOB.create( 'Encours Clients',nil,-1);
FTV.LaTob.LoaddetailDB('TIERS','','', Q, False,true,-1,0) ;
FTV.LaTob.detail[0].AddChampSup('TVENCOURS',true);
FTV.LaTob.detail[0].AddChampSup('TVFORCERISQUE',true);
initmove (FTV.Latob.detail.count,'');
For i:=FTV.Latob.detail.count-1 downto 0 do
    begin
    MoveCur(False);
    TB := FTV.Latob.detail[i];
    RisqueTiers:= RisqueTiersGC(TB)+RisqueTiersCPTA(TB,V_PGI.DateEntree);
    TB.putvalue('TVENCOURS',RisqueTiers );
    P1:=TB.getValue('T_CREDITACCORDE') ;
    P2:=TB.getValue('T_CREDITPLAFOND') ;
    EtatRisque:='V' ;
    if ((P1>0 ) and (RisqueTiers > P1)) then EtatRisque:='O' ;
    if ((P2>0 ) and (RisqueTiers > P2)) then EtatRisque:='R' ;
    if TB.getvalue('T_ETATRISQUE')='' then
       begin
       TB.putvalue('TVFORCERISQUE',False);
       TB.putvalue('T_ETATRISQUE',EtatRisque )
       end else TB.putvalue('TVFORCERISQUE',True );
    if (EtatRisque='V') and  (GetControlText('CBVERT')='-') then TB.free
    else if (EtatRisque='O') and  (GetControlText('CBORANGE')='-') then TB.free
    else if (EtatRisque='R') and  (GetControlText('CBROUGE')='-') then TB.free;
    end;
ferme(Q);
Finimove;
end;

procedure TOF_TVEncours.TVOnDblClickCell(Sender: TObject );
var EtatRisque : string;

begin
with TTobViewer(sender) do
    begin
    if (ColName[CurrentCol]='TVENCOURS') or (ColName[CurrentCol]='TVFORCERISQUE') or (ColName[CurrentCol]='T_ETATRISQUE') then
       begin
       EtatRisque:=AsString[ColIndex( 'T_ETATRISQUE' ), CurrentRow];
       if not AsBoolean[ColIndex('TVFORCERISQUE'),CurrentRow] then CurrentTob.putvalue('T_ETATRISQUE','');      // pour retrouver la valeur tel que dans la base .... bof
       AfficheRisqueClient(CurrentTob);
       CurrentTob.putvalue('T_ETATRISQUE',EtatRisque);
       end
    else if (ColName[CurrentCol]='T_TIERS') or (ColName[CurrentCol]='RAISONSOCIALE')  or (copy(ColName[CurrentCol],1,2)='T_') then
         V_PGI.DispatchTT (8,taConsult ,AsString[ColIndex( 'T_TIERS' ), CurrentRow], '','');
    end;
end;


Initialization
registerclasses([TOF_TVEncours]);
end.
