unit UTofTVVentes;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF,EntGC,
{$IFNDEF EAGLCLIENT}
      db,mul, dbTables,DBGrids,
{$ENDIF}
      HDimension,M3FP, UTobView, TVProp, UTob,FactUtil,Facture
      ;
Type
     TOF_TVVentes = Class (TOF)
     private
         TobViewer1: TTobViewer;
         procedure TVOnDblClickCell(Sender: TObject ) ;
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override ;
     END ;



implementation

procedure TOF_TVVentes.OnArgument(Arguments : String ) ;
var Iind : integer;
begin
inherited ;
TobViewer1:=TTobViewer(getcontrol('TV'));
TobViewer1.OnDblClick:= TVOnDblClickCell ;
Iind:=THValComboBox(GetControl ('NATUREPIECEG')).Values.IndexOf('FAC');
if Iind<0 then Iind:=0;
THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(Iind,'Facture + Avoir clients');
THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(Iind,'ZZ1');
SetControlText('NATUREPIECEG','ZZ1');

end;

procedure TOF_TVVentes.OnLoad ;
var stWhere : string ;
begin
if GetControlText('NATUREPIECEG')='ZZ1' then
     stWhere := ' AND (GL_NATUREPIECEG="FAC" OR GL_NATUREPIECEG="AVC" OR GL_NATUREPIECEG="AVS") '
else stWhere := ' AND GL_NATUREPIECEG="' + GetControlText ('NATUREPIECEG') + '"';
SetControlText('XX_WHERE_NAT',stWhere);
end;

procedure TOF_TVVentes.TVOnDblClickCell(Sender: TObject );
var CleDoc : R_Cledoc ;
begin
with TTobViewer(sender) do
    begin
    if (ColName[CurrentCol]='GL_ARTICLE') or (ColName[CurrentCol]='GL_CODEARTICLE') or (ColName[CurrentCol]='GL_LIBELLE')or (copy(ColName[CurrentCol],1,3)='GA_') then
         V_PGI.DispatchTT (7,taConsult ,AsString[ColIndex( 'GL_ARTICLE' ), CurrentRow], '','')
    else if (ColName[CurrentCol]='GL_TIERS') or (ColName[CurrentCol]='RAISONSOCIALE')  or (copy(ColName[CurrentCol],1,2)='T_') then
         V_PGI.DispatchTT (8,taConsult ,AsString[ColIndex( 'GL_TIERS' ), CurrentRow], '','')
    else if (ColName[CurrentCol]='GL_APPORTEUR') or (ColName[CurrentCol]='GL_TIERSPAYEUR') or (ColName[CurrentCol]='GL_TIERSFACTURE') or (ColName[CurrentCol]='GL_TIERSLIVRE')then
         V_PGI.DispatchTT (8,taConsult ,AsString[CurrentCol, CurrentRow], '','')
    else if (ColName[CurrentCol] = 'GL_REPRESENTANT') then
         V_PGI.DispatchTT (9,taConsult ,AsString[ColIndex('GL_REPRESENTANT'), CurrentRow], '','')
    else begin
         FillChar(CleDoc,Sizeof(CleDoc),#0) ;
         CleDoc.NaturePiece:=AsString[ColIndex( 'GL_NATUREPIECEG' ), CurrentRow] ;
         CleDoc.DatePiece:=AsDateTime[ColIndex( 'GL_DATEPIECE' ), CurrentRow]  ;
         CleDoc.Souche:=AsString[ColIndex( 'GL_SOUCHE' ), CurrentRow] ;
         CleDoc.NumeroPiece:=AsInteger[ColIndex( 'GL_NUMERO' ), CurrentRow] ;
         CleDoc.Indice:=AsInteger[ColIndex( 'GL_INDICEG' ), CurrentRow]  ;
         SaisiePiece(CleDoc,taConsult) ;
         end ;
    end;
end;

Initialization
registerclasses([TOF_TVVentes]);
end.
