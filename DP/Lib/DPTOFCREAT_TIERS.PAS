{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 30/05/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DPCREAT_TIERS ()
Mots clefs ... : TOF;DPCREAT_TIERS
*****************************************************************}
Unit DPTOFCREAT_TIERS ;

Interface

Uses StdCtrls,
     Classes, Hstatus, M3FP,
     Controls,
{$IFNDEF EAGLCLIENT}
    FE_Main,
     mul,
{$else}
     MaineAGL,
     eMul,
{$ENDIF}
     forms,
     HMsgBox, UtilMulTrt,AnnOutils,
     UTOF, Utob, Hent1, HTB97 ;

Type
  TOF_DPCREAT_TIERS = Class (TOF)
    procedure CreatTiers;
    procedure Bzoom_OnClick(Sender: TObject);
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
  end ;

Procedure DPLanceFiche_Creat_Tiers(Argument:string);
Implementation

Procedure DPLanceFiche_Creat_Tiers(Argument:string);
begin
AGLLanceFiche('DP','DPCREAT_TIERS','','',argument);
end;

procedure TOF_DPCREAT_TIERS.OnArgument (S : String ) ;
begin
  Inherited ;
  setcontrolVisible ('PCOMPLEMENT',false);
  TToolBarButton97(GetControl('BZOOM', true)).OnClick:= Bzoom_OnClick;
end ;

procedure TOF_DPCREAT_TIERS.CreatTiers ;
Var TobAnn, TobDet : tob;
    Wi : integer;
    F : TFMul ;
begin
  F:=TFMul(Ecran);
  if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
    begin
    MessageAlerte('Aucun élément sélectionné');
    V_PGI.ExtendedFieldSelection:='' ;
    exit;
    end;
  if PGIAsk(TraduireMemoire('Confirmez-vous le traitement ?'),Ecran.Caption) <> mrYes  then  Exit;
  TobAnn := Tob.Create('les annuaire',NIL,-1);
  TraiteEnregMulListe (F, 'ANN_GUIDPER','ANNUAIRE', TobAnn, True);
  InitMove(TobAnn.Detail.count,' ');
  wi := 0;
  while wi < TobAnn.Detail.count  do
    Begin
    MoveCur(False);
    TobDET := TobAnn.Detail[wi];
    if TobDet.GetString ('ANN_TIERS') = ''
     then DpCreerTiers(TobDet.getString ('ANN_GUIDPER'),TobDet.getString('ANN_NATUREAUXI')
          ,TobDet.getString('ANN_NOM1'),TobDet.getString('ANN_NOMPER'));
    inc(wi);
    End;  // fin while, boucle sur la sélection
  TobAnn.free; TobAnn:=Nil;
  FiniMove;
end ;

procedure TOF_DPCREAT_TIERS.Bzoom_OnClick(Sender: TObject);
var Fonction, FamPer: String;
begin
  AGLLanceFiche('YY','ANNUAIRE','', GetField('ANN_GUIDPER'),
   'ACTION=MODIFICATION;;;'+'DOS')  ;
end;

procedure TOF_DPCREAT_TIERS.OnLoad ;
begin
  Inherited ;
  if GetCheckBoxState('PASTIERS') = cbChecked  then
   setControlText ('XX_WHERE','(and ANN_NATUREAUXI<>"" and ANN_TIERS="")')
   else SetControlText ('XX_WHERE','');
end ;


procedure AGLCreatTiers(parms: array of variant; nb: integer );
var F : TForm;
    LaTof : TOF;
begin
  F:=TForm(Longint(Parms[0]));
  if (F is TFMul) then Latof:=TFMul(F).Latof;
  if (Latof is TOF_DPCreat_tiers) then TOF_DPCreat_tiers(LaTof).CreatTiers else exit;
end;

Initialization
  registerclasses ( [ TOF_DPCREAT_TIERS ] ) ; 
  RegisterAglProc( 'CreatTiers',True,0,AGLCreatTiers);
end.
