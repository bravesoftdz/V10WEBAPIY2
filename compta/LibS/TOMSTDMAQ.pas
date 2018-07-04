unit TOMSTDMAQ;

interface
uses  StdCtrls,Controls,Classes,db,forms,sysutils,dbTables,ComCtrls,
      HCtrls,HEnt1,HMsgBox,HDB,UTOM,DBCtrls,UTOB,HTB97,Fiche,
      Spin,Dialogs,FE_Main, MUL,FichList,ParamSoc,galOutil,
      Menus,Windows, Messages,  Graphics,PGIEnv;
type
      TOM_STDMAQ = Class (TOM)
       procedure OnArgument (stArgument : String ) ; override ;
       procedure DbClickDelete(Sender: TObject);
       procedure OnLoadRecord; override;
       procedure FormKeyDownOrga(Sender: TObject; var Key: Word; Shift: TShiftState);
       private
      Suppression  : string;
     END ;

implementation

procedure TOM_STDMAQ.OnArgument(stArgument: String);
var
typmaq    : string;
begin
     Suppression := ReadTokenSt (stArgument);
     TToolbarButton97(GetControl('BDelete')).Onclick := DbClickDelete;
     Ecran.OnKeyDown :=  FormKeyDownOrga;
end;

procedure TOM_STDMAQ.OnLoadRecord;
begin
if Suppression = 'OK' then
setControlEnabled ('NOUVEAU', FALSE);
end;

procedure TOM_STDMAQ.DbClickDelete(Sender: TObject);
var
OM       : TOM;
NoStd    : integer;
Fichier  : string;
Typemaq  : string;
act      : TCloseAction;
begin
if not IsSuperviseur(TRUE) then exit;
NoStd :=  GetField('STM_NUMPLAN');

if not EstSpecif('51502') then
if noStd <= 20 then
begin PGIInfo('Maquette CEGID,impossible de supprimer','') ;  exit;end;

Typemaq :=  GetField('STM_TYPEMAQ');
if not (TFFicheListe(Ecran).Bouge(nbDelete)) then exit;
// CA - 27/11/2001
if noStd <= 20 then
  Fichier := V_PGI_Env.PathStd + '\'+Typemaq + Format('%.02d',[NoStd])+ '.txt'
else Fichier := V_PGI_Env.PathDat + '\'+Typemaq + Format('%.02d',[NoStd])+ '.txt';
DeleteFile (Pchar(Fichier));
TFMul(Ecran).ModalResult := 1;
end;

procedure TOM_STDMAQ.FormKeyDownOrga(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
Case (ORD(Key)) of
     VK_DELETE : // suppression
               DbClickDelete(Sender);
end;
end;

Initialization
registerclasses([TOM_STDMAQ]) ;

end.
