{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 21/11/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GCFONCTIONPARAM ()
Mots clefs ... : TOF;GCFONCTIONPARAM
*****************************************************************}
Unit GcFonctionParam_Tof ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL,eTablette,
{$ELSE}
     db,dbtables,FE_Main,Tablette,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,HTB97,ExtCtrls,Graphics,
     UTOF,UTOB,AGLInit,Vierge,UTilFonctionCalcul,VoirTob ;



Type
  TOF_GCFONCTIONPARAM = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (StArg : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    FormuleInitiale,FormuleRetour : String;
    MaxNbre: integer ;
    NombreCour : integer ;
    BFONCTIONNEW : TToolbarButton97;
    BFONCTIONPREC : TToolbarButton97;
    BFONCTIONSUIV : TToolbarButton97;
    BFONCTIONANNUL : TToolbarButton97;
    BVOIRTOB : TToolbarButton97;
    BCALCUL : TToolbarButton97;
    TobFonctTab,TobFonct,TobFCour : TOB; 
    // -- Evenements liées aux champs
    procedure NombreEnter(Sender: TObject) ;
    procedure NombreExit(Sender: TObject) ;
    procedure BParamClick(Sender: TObject) ;
    // -- Fonctions liées aux Choix Fonctions
    procedure AfficheSaisie;
    procedure InitSaisie;
    function  recapFormule(TOBF : TOB) : string;
    procedure EtablirFormule ;
    // -- Fonctions liées aux Boutons
    procedure BFONCTIONNEWClick(Sender: TObject);
    procedure BFONCTIONPRECClick(Sender: TObject);
    procedure BFONCTIONSUIVClick(Sender: TObject);
    procedure BFONCTIONANNULClick(Sender: TObject);
    procedure BVOIRTOBClick(Sender: TObject);
    procedure BCALCULClick(Sender: TObject);
  end ;


Type
  TOF_GCFONCTIONCHOIX = Class (TOF)
    LBCHOIXFONCTION : TListBox;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    // -- Evenements liées aux champs
    procedure LBCHOIXFONCTIONClick(Sender: TObject);
    procedure LBCHOIXFONCTIONDblClick(Sender: TObject);
  end ;


Implementation

{==============================================================================================}
{========================== Procédure de la TOF GCFONCTIONPARAM ===============================}
{==============================================================================================}
procedure TOF_GCFONCTIONPARAM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCFONCTIONPARAM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCFONCTIONPARAM.OnUpdate ;
begin
  Inherited ;
NextPrevControl(Ecran) ;
FormuleRetour:=GetControlText('FORMULE') ;
end ;

procedure TOF_GCFONCTIONPARAM.OnLoad ;
begin
  Inherited ;
if FormuleInitiale<>'' then
  begin
  if DechiffrageFormule(FormuleInitiale,TobFonct,TobFonctTab) then
    begin
    SetControlVisible('PCHAMPS',True);
    SetControlVisible('PTEST',True);
    NombreCour:=1;
    AfficheSaisie;
    end;
  end;
end ;

procedure TOF_GCFONCTIONPARAM.OnArgument (StArg : String ) ;
Var QQ : TQuery;
    SCB : TScrollBox;
    Ph : TPanel;
    i_ind : integer ;
    Nombre : THEdit;
    BParam : TToolbarButton97;
begin
  Inherited ;
FormuleInitiale:=StArg; FormuleRetour:=StArg;
// FormuleInitiale:='PRODUIT(SOMME(Large;DIVISION(9;5);45;PRODUIT(2.5;3;Epaisseur));SOMME(3;6);Poids)';
MaxNbre:=10; NombreCour:=-1;
BFONCTIONNEW:=TToolbarButton97(GetControl('BFONCTIONNEW'));
BFONCTIONNEW.OnClick:=BFONCTIONNEWClick;
BFONCTIONPREC:=TToolbarButton97(GetControl('BFONCTIONPREC'));
BFONCTIONPREC.OnClick:=BFONCTIONPRECClick;
BFONCTIONSUIV:=TToolbarButton97(GetControl('BFONCTIONSUIV'));
BFONCTIONSUIV.OnClick:=BFONCTIONSUIVClick;
BFONCTIONANNUL:=TToolbarButton97(GetControl('BFONCTIONANNUL'));
BFONCTIONANNUL.OnClick:=BFONCTIONANNULClick;
BVOIRTOB:=TToolbarButton97(GetControl('BVOIRTOB'));
BVOIRTOB.OnClick:=BVOIRTOBClick;
BCALCUL:=TToolbarButton97(GetControl('BCALCUL'));
BCALCUL.OnClick:=BCALCULClick;
for i_ind:=1 to MaxNbre do
   begin
   Nombre:=THEdit(GetControl('NOMBRE'+IntToStr(i_ind)));
   Nombre.OnEnter:=NombreEnter;
   Nombre.OnExit:=NombreExit;
   end;
for i_ind:=1 to MaxNbre do
   begin
   BParam:=TToolbarButton97(GetControl('BPARAM'+IntToStr(i_ind)));
   BParam.OnClick:=BParamClick;
   end;
TobFonct:=Tob.Create('COMMUN',Nil,-1); TobFCour:=TobFonct;
TobFonctTab:=Tob.Create('',Nil,-1);
QQ:=OpenSql('Select * from Commun Where CO_TYPE="GFX"',True);
if not QQ.Eof then TobFonctTab.LoadDetailDB('COMMUN','','',QQ,False);
Ferme(QQ);
// Création de la ScrollBox associée
SCB:=TScrollBox.create(Ecran) ;
SCB.Parent:=TTabSheet(getcontrol('ONGLETFONCTION')) ;
SCB.Name:='SCB';
SCB.ParentFont:=False; SCB.Font.Color:=clWindowText;
SCB.Align:=alClient;
Ph:=TPanel(GetControl('PCHAMPS'));
Ph.Parent:=SCB;
SetControlVisible('PCHAMPS',False);
SetControlVisible('PTEST',False);
SetControlProperty('RESULTAT','Masks.PositiveMask','# ##0.0000');
end ;

procedure TOF_GCFONCTIONPARAM.OnClose ;
begin
  Inherited ;
TFVierge(Ecran).Retour:=FormuleRetour ;
TobFonctTab.Free; TobFonct.Free; TobFonct:=Nil;
end ;

{==============================================================================================}
{============================== Evenements liées aux champs ===================================}
{==============================================================================================}
procedure TOF_GCFONCTIONPARAM.NombreEnter(Sender: TObject) ;
var ChpActif : THEdit;
    MaxNbFx,i_pos : integer;
    St,StMax : string ;
begin
SetControlProperty('PCHAMPS','AutoSize',False);
ChpActif:=THEdit(Sender) ;
if ChpActif=nil then exit;
i_pos:=Pos('NOMBRE',ChpActif.Name)+6;
NombreCour:=StrToInt(Copy(ChpActif.Name,i_pos,Length(ChpActif.Name)+1-i_pos));
BFONCTIONPREC.Enabled:=(TobFCour.Parent<>Nil);
if TobFCour.Detail.Count>=NombreCour then
   BFONCTIONSUIV.Enabled:=(TobFCour.Detail[NombreCour-1].GetValue('TYPE')='F')
   else BFONCTIONSUIV.Enabled:=False;
if BFONCTIONSUIV.Enabled then
  begin
  BFONCTIONNEW.Enabled:=False;
  SetControlProperty('NOMBRE'+IntToStr(NombreCour),'ReadOnly',True);
  SetControlProperty('NOMBRE'+IntToStr(NombreCour),'DataType','');
  end else
  begin
  BFONCTIONNEW.Enabled:=True;
  SetControlProperty('NOMBRE'+IntToStr(NombreCour),'ReadOnly',False);
  SetControlProperty('NOMBRE'+IntToStr(NombreCour),'DataType','GCFONCTIONVARIABLE');
  end;
BFONCTIONANNUL.Enabled:=BFONCTIONSUIV.Enabled ;
St:=TobFCour.GetValue('CO_ABREGE');
ReadTokenSt(St);StMax:=ReadTokenSt(St);
MaxNbFx:= StrToInt(StMax);
if (NombreCour<MaxNbre) and (NombreCour<MaxNbFx) then
   begin
   SetControlVisible('NOMBRE'+IntToStr(NombreCour+1),True);
   SetControlVisible('TNOMBRE'+IntToStr(NombreCour+1),True);
   SetControlVisible('BPARAM'+IntToStr(NombreCour+1),True);
   end;
SetControlProperty('PCHAMPS','AutoSize',True);
end;

procedure TOF_GCFONCTIONPARAM.NombreExit(Sender: TObject) ;
var ChpActif : THEdit;
    ind : integer;
    St : string ;
    TOBN : TOB ;
    Fonct : boolean;
begin
if TobFonct=Nil then Exit;
ChpActif:=THEdit(Sender) ;
if ChpActif=nil then exit;
if TobFCour.Detail.Count>=NombreCour then
  Fonct:=(TobFCour.Detail[NombreCour-1].GetValue('TYPE')='F')
  else Fonct:=False;

St:=trim(GetControltext(ChpActif.Name));
St:=StringReplace(St,' ','_',[rfReplaceAll]);
if not Fonct then
  begin
  St:=StringReplace(St,';','',[rfReplaceAll]);
  St:=VireParenthese(St);
  end;
St:=StringReplace(St,',','.',[rfReplaceAll]);
St:=ReplaceCaractereAccentue(St);
St:=SupprimeCaracteresSpeciaux(St,False,True,True);
SetControltext(ChpActif.Name,St);
if Not VerifieParentheses(St) then
  begin
  ChpActif.SetFocus;
  exit;
  end;
for ind:=TobFCour.Detail.Count+1 to NombreCour do
  begin
  TOBN:=Tob.Create('_Nombre'+IntToStr(ind),TobFCour,-1);
  TOBN.AddChampSup('TYPE',False);
  TOBN.AddChampSup('VALEUR',False);
  TOBN.PutValue('TYPE','N');
  TOBN.PutValue('VALEUR','');
  end;
if not Fonct then
  if (St='') or (isnumeric(St,True)) then TobFCour.Detail[NombreCour-1].PutValue('TYPE','N')
                                     else TobFCour.Detail[NombreCour-1].PutValue('TYPE','V');
TobFCour.Detail[NombreCour-1].PutValue('VALEUR',St);
EtablirFormule;
end;

procedure TOF_GCFONCTIONPARAM.BParamClick(Sender: TObject) ;
begin
ParamTable('GCFONCTIONVARIABLE',taCreat,0,Nil) ;
end;

{==============================================================================================}
{========================== Fonctions liées aux Choix Fonctions ===============================}
{==============================================================================================}
procedure TOF_GCFONCTIONPARAM.AfficheSaisie;
Var i_ind : integer;
begin
SetControlText('TFONCTION','Fonction : '+TobFCour.GetValue('CO_LIBELLE'));
For i_ind:=1 to TobFCour.Detail.Count do
  begin
  SetControlText('NOMBRE'+IntToStr(i_ind),TobFCour.Detail[i_ind-1].GetValue('VALEUR')) ;
  SetControlVisible('NOMBRE'+IntToStr(i_ind),True);
  SetControlVisible('TNOMBRE'+IntToStr(i_ind),True);
  SetControlVisible('BPARAM'+IntToStr(i_ind),True);
  end;
For i_ind:=TobFCour.Detail.Count+1 to MaxNbre do
  begin
  SetControlText('NOMBRE'+IntToStr(i_ind),'') ;
  SetControlVisible('NOMBRE'+IntToStr(i_ind),False);
  SetControlVisible('TNOMBRE'+IntToStr(i_ind),False);
  SetControlVisible('BPARAM'+IntToStr(i_ind),False);
  end;
EtablirFormule;
THEdit(GetControl('NOMBRE'+IntToStr(NombreCour))).SetFocus;
THEdit(GetControl('NOMBRE1')).OnEnter(GetControl('NOMBRE'+IntToStr(NombreCour)));
end;

procedure TOF_GCFONCTIONPARAM.InitSaisie;
var i_ind,MinNbre : integer;
    St, StMin: string;
begin
SetControlVisible('PCHAMPS',True);
SetControlVisible('PTEST',True);
SetControlText('TFONCTION','Fonction : '+TobFCour.GetValue('CO_LIBELLE'));
St:=TobFCour.GetValue('CO_ABREGE');
StMin:=ReadTokenSt(St);
MinNbre:=StrToInt(StMin);
For i_ind:=1 to MinNbre do
  begin
  SetControlText('NOMBRE'+IntToStr(i_ind),'') ;
  SetControlVisible('NOMBRE'+IntToStr(i_ind),True);
  SetControlVisible('TNOMBRE'+IntToStr(i_ind),True);
  SetControlVisible('BPARAM'+IntToStr(i_ind),True);
  end;
For i_ind:=MinNbre+1 to MaxNbre do
  begin
  SetControlText('NOMBRE'+IntToStr(i_ind),'') ;
  SetControlVisible('NOMBRE'+IntToStr(i_ind),False);
  SetControlVisible('TNOMBRE'+IntToStr(i_ind),False);
  SetControlVisible('BPARAM'+IntToStr(i_ind),False);
  end;
EtablirFormule;
THEdit(GetControl('NOMBRE1')).SetFocus;
end ;

function TOF_GCFONCTIONPARAM.recapFormule(TOBF : TOB) :  string ;
Var St,Sepa, Val : string;
    ind : integer;
begin
Result:='';
if TOBF=Nil then exit;
St:=TOBF.GetValue('CO_LIBELLE')+'(';
For ind:=0 to TOBF.Detail.Count-1 do
  begin
  Val:=TOBF.Detail[ind].GetValue('VALEUR');
  if Val='' then continue;
  if ind=0 then Sepa:='' else Sepa:=';';
  St:=St+Sepa+Val;
  end;
Result:=St+')';
end ;

procedure TOF_GCFONCTIONPARAM.EtablirFormule ;
Var St: string;
    TOBPere,TOBFPere : TOB;
begin
if TobFCour=Nil then exit;
St:=recapFormule(TobFCour);
TOBPere:=TobFCour.Parent;
while TOBPere<>nil do
  begin
  TOBPere.PutValue('TYPE','F');
  TOBPere.PutValue('VALEUR',St);
  TOBFPere:=TOBPere.Parent; if TOBFPere=Nil then break;
  St:=recapFormule(TOBFPere);
  TOBPere:=TOBFPere.Parent;
  end ;
St:=recapFormule(TobFonct);
SetControlText('FORMULE',St);
end;

{==============================================================================================}
{============================== Fonctions liées aux boutons ===================================}
{==============================================================================================}
procedure TOF_GCFONCTIONPARAM.BFONCTIONNEWClick(Sender: TObject);
var FctChoix : integer ;
begin
if NombreCour>0 then THEdit(GetControl('NOMBRE1')).OnExit(GetControl('NOMBRE'+IntToStr(NombreCour)));
TheTob:=TobFonctTab;
FctChoix:=StrToInt(AGLLanceFiche('GC','GCFONCTIONCHOIX','','',''));
if FctChoix>=0 then
   begin
   if NombreCour>0 then TobFCour:=TOB.Create('COMMUN',TobFCour.Detail[NombreCour-1],-1);
   TobFCour.Dupliquer(TobFonctTab.Detail[FctChoix],False,True);
   InitSaisie;
   end;
end ;

procedure TOF_GCFONCTIONPARAM.BFONCTIONPRECClick(Sender: TObject);
Var TOBN: TOB;
begin
if NombreCour>0 then THEdit(GetControl('NOMBRE1')).OnExit(GetControl('NOMBRE'+IntToStr(NombreCour)));
TOBN:=TobFCour.Parent;
NombreCour:=StrToInt(Copy(TOBN.NomTable,Length(TOBN.NomTable),1));
TobFCour:=TOBN.Parent;
AfficheSaisie;
end ;

procedure TOF_GCFONCTIONPARAM.BFONCTIONSUIVClick(Sender: TObject);
Var TOBN: TOB;
begin
if NombreCour>0 then THEdit(GetControl('NOMBRE1')).OnExit(GetControl('NOMBRE'+IntToStr(NombreCour)));
TOBN:=TobFCour.Detail[NombreCour-1];
if TOBN.Detail.Count=0 then exit;
TobFCour:=TOBN.Detail[0];
NombreCour:=1;
AfficheSaisie;
end ;

procedure TOF_GCFONCTIONPARAM.BFONCTIONANNULClick(Sender: TObject);
Var TOBN: TOB;
begin
if NombreCour>0 then
  begin
  TOBN:=TobFCour.Detail[NombreCour-1];
  if TOBN.Detail.Count>0 then TOBN.Detail[0].Free;
  TOBN.PutValue('TYPE','N');
  TOBN.PutValue('VALEUR','');
  end;
NombreCour:=1;
AfficheSaisie;
end ;

procedure TOF_GCFONCTIONPARAM.BVOIRTOBClick(Sender: TObject);
begin
GCVoirTob(TobFonct);
end ;

procedure TOF_GCFONCTIONPARAM.BCALCULClick(Sender: TObject);
Var FoncQte : R_FonctCal ;
begin
NextPrevControl(Ecran);
FoncQte:=EvaluationFormule('',GetControlText('FORMULE'));
SetControlText('RESULTAT',FloatToStr(FoncQte.Resultat));
SetControlText('EXPRESSION',FoncQte.Expression);
end ;

{==============================================================================================}
{========================== Procédure de la TOF GCFONCTIONCHOIX ===============================}
{==============================================================================================}
procedure TOF_GCFONCTIONCHOIX.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCFONCTIONCHOIX.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCFONCTIONCHOIX.OnUpdate ;
begin
  Inherited ;
TFVierge(Ecran).Retour:=IntToStr(LBCHOIXFONCTION.ItemIndex) ;
end ;

procedure TOF_GCFONCTIONCHOIX.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_GCFONCTIONCHOIX.OnArgument (S : String ) ;
Var i_ind : integer ;
begin
  Inherited ;
TFVierge(Ecran).Retour:='-1';
LBCHOIXFONCTION:=TListBox(GetControl('LBCHOIXFONCTION'));
LBCHOIXFONCTION.OnClick:=LBCHOIXFONCTIONClick;
LBCHOIXFONCTION.OnDblClick:=LBCHOIXFONCTIONDblClick;
for i_ind:=0 to LaTob.Detail.Count-1 do
  LBCHOIXFONCTION.Items.Add (LaTob.Detail[i_ind].GetValue('CO_LIBELLE'));
LBCHOIXFONCTION.ItemIndex:=0 ;
SetControlText('TLIBELLE',LaTob.Detail[0].GetValue('CO_LIBRE'));
end ;

procedure TOF_GCFONCTIONCHOIX.OnClose ;
begin
  Inherited ;
end ;

{==============================================================================================}
{============================== Evenements liées aux champs ===================================}
{==============================================================================================}
procedure TOF_GCFONCTIONCHOIX.LBCHOIXFONCTIONClick(Sender: TObject);
Var St : string;
    ind : integer ;
begin
ind:=LBCHOIXFONCTION.ItemIndex ;
St:=LaTob.Detail[ind].GetValue('CO_LIBRE');
SetControlText('TLIBELLE',St);
end;

procedure TOF_GCFONCTIONCHOIX.LBCHOIXFONCTIONDblClick(Sender: TObject);
begin
TFVierge(Ecran).BValiderClick (Sender);
TFVierge(Ecran).Close;
end;

Initialization
  registerclasses([TOF_GCFONCTIONPARAM]) ;
  registerclasses([TOF_GCFONCTIONCHOIX]) ;
end.

