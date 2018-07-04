unit UTofTransfert_MUL;

interface

uses  UTOF,HMsgBox,Classes,HCtrls,Hqry,SysUtils,FactUtil,
      HEnt1,Transfert,M3FP,Utob,
{$IFDEF EAGLCLIENT}
      emul,
{$ELSE}
      MUL, DBTables,                   
{$ENDIF}
      EntGC, Forms;

Procedure AppelTransfert ( Parms : array of variant ; nb : integer ) ;
Procedure AllerRetourSurControl ( Parms : array of variant ; nb : integer ) ;

Type
     TOF_Transfert_MUL = Class (TOF)
        private
        ValidationTRV : boolean;
        public
        procedure OnArgument (S : String ) ; override ;
     end ;
var
  Transfert_MUL : TOF_Transfert_MUL;

implementation

procedure TOF_Transfert_MUL.OnArgument (S : String ) ;
var Critere : String;
    F : TFMul ;
begin
Inherited ;
F:=TFMul(Ecran);
ValidationTRV:=False;
Repeat
  Critere:=uppercase(Trim(ReadTokenSt(S))) ;
  if Critere='CONSULTATION' then F.Caption:='Consultation d''un transfert'
  else if Critere='MODIFICATION' then F.Caption:='Modification d''un transfert'
    else if Critere='VALIDATION' then
      begin
      F.Caption:='Validation d''un transfert';
      ValidationTRV:=True;
      SetControlProperty('GP_NATUREPIECEG','Tag',-9971);
      end;
until  Critere='';
// paramétrage du libellé etablissement quand on est en multi-dépôts
if VH_GC.GCMultiDepots then
   begin
   SetControlText('TGP_DEPOT',TraduireMemoire('Dépôt émetteur'));
   SetControlText('TGP_DEPOTDEST',TraduireMemoire('Dépôt récepteur'));
   end
else
   begin
   SetControlText('TGP_DEPOT',TraduireMemoire('Etabliss. émetteur'));
   SetControlText('TGP_DEPOTDEST',TraduireMemoire('Etabliss. récepteur'));
   end;
UpdateCaption(Ecran);
end ;

Procedure AppelTransfert( Parms : array of variant ; nb : integer ) ;
var Critere,Arguments,ChampMul,ValMul : string ;
    F : TFMul;
    Latof : TOF ;
    CleDoc : R_CleDoc ;
    x : integer;
    Action : TActionFiche;
    QSurSite : TQuery;
begin
Action:=taConsult;
F:=TFMul(Integer(Parms[0])) ;
if (F is TFMul) then Latof:=TFMul(F).Latof else exit;
if (Latof is TOF_Transfert_MUL) then
  begin
  Arguments:=string(Parms[1]);
  Repeat
   Critere:=UpperCase(Trim(ReadTokenSt(Arguments))) ;
   if Critere<>'' then
      begin
      x:=pos('=',Critere);
      if x<>0 then
         begin
         ChampMul:=copy(Critere,1,x-1);
         ValMul:=copy(Critere,x+1,length(Critere));
         end;
         if ChampMul='NATUREPIECE' then CleDoc.NaturePiece := ValMul;
         if ChampMul='DATEPIECE' then CleDoc.DatePiece := StrToDate(ValMul);
         if ChampMul='SOUCHE' then CleDoc.Souche := ValMul;
         if ChampMul='NUMEROPIECE' then CleDoc.NumeroPiece := StrToInt(ValMul);
         if ChampMul='INDICE' then CleDoc.Indice := StrToInt(ValMul);
         if ChampMul='ACTION' then
           begin
           if ValMul='CONSULTATION' then Action:=taConsult;
           if ValMul='MODIFICATION' then Action:=taModif;
           end;
      end;
  until  Critere='';
  if (Action=taModif) and (CleDoc.NaturePiece='TEM') then
  begin
    QSurSite := OpenSQL('select GDE_SURSITE from Piece Left '+
    'Join DEPOTS ON GDE_DEPOT=GP_DEPOT where GP_NATUREPIECEG="TEM" and '+
    'GP_SOUCHE="'+ CleDoc.Souche +'" and '+
    'GP_NUMERO="'+ IntToStr(CleDoc.NumeroPiece) +'"', True);
    if Not QSurSite.EOF then
      if QSurSite.FindField('GDE_SURSITE').AsString = '-' then Action:=taConsult;
    Ferme(QSurSite);
  end;
  if TOF_Transfert_MUL(LaTof).ValidationTRV then TransformeTRVenTRE(CleDoc)
  //else SaisieTransfert(CleDoc,Action);
  else SaisieTransfert(CleDoc,Action,'TEM_TRV_NUMIDENTIQUE');  
  end;
end;

Procedure AllerRetourSurControl(Parms:array of variant;nb:integer) ;
var F : TFMul;
    Latof : TOF ;
begin
F:=TFMul(Integer(Parms[0])) ;
if (F is TFMul) then Latof:=TFMul(F).Latof else exit;
if (Latof is TOF_Transfert_MUL) then
  begin
  NextPrevControl(F);
  end;
end;


Initialization
RegisterClasses([TOF_Transfert_MUL]);
RegisterAglProc('AppelTransfert', True , 5, AppelTransfert);
RegisterAglProc('ARControl', True , 0, AllerRetourSurControl);
end.
