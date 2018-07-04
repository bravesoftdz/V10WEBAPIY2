unit TofDpoin;

interface
uses  StdCtrls,Controls,Classes,sysutils,
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     db, Hent1,HCtrls,HDB,FE_MAIN,Mul,HMsgBox,UTOF,UTOB;

Type
     TOF_DPointe = Class (TOF)
     private
       procedure MonDoubleClick(Sender: TObject) ;
       procedure DepointerUnReleve(Sender: TObject) ;
       procedure VisualiserUnReleve(Sender: TObject) ;
       procedure DepointeAvecReference(RefPointage : string) ;
       function RecupererChamp(Champ : string) : string;       
     public
       procedure OnLoad ; override ;
     end ;
const
     SR_GENERAL  = 1 ;
     SR_REF      = SR_GENERAL + 1 ;
     {Grid Visualisation}
     SV_DATE        = 0 ;
     SV_DATEVAL     = SV_DATE + 1 ;
     SV_LIBELLE     = SV_DATEVAL + 1 ;
     SV_DEBIT       = SV_LIBELLE + 1 ;
     SV_CREDIT      = SV_DEBIT + 1 ;
     SV_DEVISE      = SV_CREDIT + 1 ;
     SV_TYPE        = SV_DEVISE + 1 ;

implementation
procedure Tof_DPointe.OnLoad ;
var BOuvrir,BDepointer,BVisualiser: TButton ;G : THDBGrid ;
begin
BOuvrir:=TButton(GetControl('BOUVRIR')) ;
BDepointer:=TButton(GetControl('BDEPOINTER')) ;
BVisualiser:=TButton(GetControl('BVISUALISER')) ;
G:=THDBGrid(GetControl('FLISTE')) ;
if BOuvrir<>nil then BOuvrir.OnClick:=MonDoubleClick ;
if BDepointer<>nil then BDepointer.OnClick:=DepointerUnReleve ;
if BVisualiser<>nil then BVisualiser.OnClick:=VisualiserUnReleve ;
if G<>nil then G.OnDblClick:=VisualiserUnReleve ;
end;

procedure Tof_DPointe.MonDoubleClick(Sender : TObject) ;
var TheTob : TOB ;FMul : TFMul ; RefPointage,Msg : String ; i : integer ;
begin
inherited ;
FMul:=TFMul(Ecran) ;
TheTob:=TOB.create('_EEXBQ',nil,-1) ;
TheTob.LoadDetailDB('EEXBQ','','',FMul.Q,FALSE,TRUE) ;
if TheTob.Detail.count > 1 then Msg:='les '+IntToStr(TheTob.Detail.count)+' Relevés' else Msg:='ce relevé';
if HShowMessage('1;Dépointage des Relevés bancaires;Etes-vous sur de vouloir dépointer '+Msg+';Q;YN;N;N','','')=mrNo then Exit ;
for i:=0 to TheTob.Detail.count-1 do
  begin
  RefPointage:=TheTob.Detail[i].GetValue('EE_REFPOINTAGE') ;
  DepointeAvecReference(RefPointage) ;
  end ;
TheTob.free ;
end ;

procedure Tof_DPointe.DepointeAvecReference(RefPointage : string) ;
begin
ExecuteSql('UPDATE ECRITURE '+
           'SET E_DATEPOINTAGE="'+UsDateTime(StrToDate('01/01/1900'))+'", '+
           'E_REFPOINTAGE=""'+
           'WHERE E_REFPOINTAGE="'+RefPointage+'"') ;

ExecuteSql('UPDATE EEXBQLIG '+
           'SET CEL_DATEPOINTAGE="'+UsDateTime(StrToDate('01/01/1900'))+'" '+
           'WHERE CEL_REFPOINTAGE="'+RefPointage+'"') ;
end;

function Tof_DPointe.RecupererChamp(Champ : string) : string;
var
F : TFMul ;
begin
F:=TFMul(Ecran) ;
if F<>nil then
  begin
  if not F.Q.Eof then
    begin
    Result:=F.Q.FindField(Champ).AsString ;
    end ;
  end;
end ;

procedure Tof_DPointe.DepointerUnReleve(Sender : TObject) ;
var RefPointage : string  ;
begin
RefPointage:=RecupererChamp('EE_REFPOINTAGE') ;
if RefPointage<>'' then DepointeAvecReference(RefPointage) ;
end;

procedure Tof_DPointe.VisualiserUnReleve(Sender : TObject) ;
var RefPointage : string  ;
begin
RefPointage:=RecupererChamp('EE_REFPOINTAGE') ;
if RefPointage<>'' then AglLanceFiche('CP','RLVVISU','','','NOMREF='+RefPointage) ;
end;


Initialization
registerclasses([Tof_DPointe]);
end.
