{***********UNITE*************************************************
Auteur  ...... :  DESSEIGNET
Créé le ...... : 22/08/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFMODIFBONI_MUL ()
Mots clefs ... : TOF;AFMODIFBONI_MUL
*****************************************************************}
Unit UtofAFMODIFBONI_MUL ;

Interface


Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     Maineagl,emul,
{$ELSE}
     db,dbTables,FE_Main,mul,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1, HSysMenu,
     HMsgBox,
     UTOF,M3FP,Hqry,
     UtilMulTrt, utob,
     dicoaf, htb97,UtofAfModifBoni,
     UTofAfBaseCodeAffaire, UtofAppreciation  ;

Type
  TOF_AFMODIFBONI_MUL = Class (TOF_AFBASECODEAFFAIRE)
  private
  public
    Action   :  TActionFiche ;
    procedure OnArgument (S : String ) ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
    procedure AfModifBoni;
    function ModifBoni(TobBoniSelec:TOB):integer;
    procedure AppelAppreciation (Affaire, DateApprec : string);
  end ;

Procedure AFLanceFiche_ModifBoni_Mul(Argument:String);

Implementation


procedure TOF_AFMODIFBONI_MUL.OnArgument (S : String ) ;
Var
    Critere, Champ, valeur : String;
    x : integer;
    bProtegeAffaire : boolean;
begin
bProtegeAffaire := false;
// Recup des critères
Critere:=(Trim(ReadTokenSt(S)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        x:=pos('=',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if Champ='ACTION' then
             begin
             if Valeur='CREATION' then BEGIN Action:=taCreat ; END ;
             if Valeur='MODIFICATION' then BEGIN Action:=taModif ; END ;
             if Valeur='CONSULTATION' then BEGIN Action:=taConsult ; END ;
             end;
        if Champ = 'TIERS' then
         begin
         SetControlText('ACT_TIERS',Valeur);
         SetControlEnabled ('ACT_TIERS',False);
         end
        else if Champ = 'AFFAIRE' then
          begin
          SetControlText('ACT_AFFAIRE',Valeur);
          If Valeur <>'' then
            begin
            bProtegeAffaire := true;
(*            setControlEnabled ('BSELECTAFF1',False);
            setControlEnabled ('BEFFACEAFF1',False);
            SetControlEnabled ('ACT_TIERS',False);
            SetControlEnabled ('ACT_AFFAIRE1',False);
            SetControlEnabled ('ACT_AFFAIRE2',False);
            SetControlEnabled ('ACT_AFFAIRE3',False);
            SetControlEnabled ('ACT_AVENANT',False);*)
            end;
          BchangeTiers :=False;
          end
        else if Champ = 'DATE' then SetControlText('ACT_DATEACTIVITE',Valeur)
        else if Champ = 'NUM' then SetControlText('ACT_NUMAPPREC',Valeur);
        END;
    Critere:=(Trim(ReadTokenSt(S)));
    END;

// Init du code affaire dans la tof ancêtre
Inherited;

// Force la protection des champs Affaire
if (bProtegeAffaire = true) then
  begin
    EditAff.ReadOnly := true;
    EditAff1.ReadOnly := true;
    EditAff2.ReadOnly := true;
    EditAff3.ReadOnly := true;
    EditAff4.ReadOnly := true;
    SetControlProperty ('ACT_DATEACTIVITE', 'ReadOnly', true);
  end;    
end ;

procedure TOF_AFMODIFBONI_MUL.AfModifBoni;
var
  TobMAct : TOB;
  NbTrait : integer;
begin

  TobMAct := TOB.Create('', Nil, -1);

  try
  SourisSablier;
  TraiteEnregMulListe (TFMul(Ecran), 'ACT_AFFAIRE', 'AFMODIFBONI', TobMAct, True);

  NbTrait := TobMAct.Detail.count;

  if (NbTrait <> 0) then
     ModifBoni (TobMAct)
  else
      PGIInfoAf('Vous devez sélectionner au-moins un Bonus/Malus', Ecran.Caption);
  // Rafraichit la liste
  TFMul(Ecran).ChercheClick;

finally
  TobMAct.Free;
  SourisNormale;
end;

TToolBarButton97(GetControl('bSelectAll')).Down := false;
end;


function TOF_AFMODIFBONI_MUL.ModifBoni(TobBoniSelec:TOB):integer;
var
i,iNbAct:integer;
RepEF, Arg:string;
begin
Result:=0;

try
//
iNbAct := TobBoniSelec.Detail.Count;
try
i:=0;
while (i < iNbAct) and (RepEF<>'STOP') do
    begin
    Arg := 'AFFAIRE=' + TobBoniSelec.Detail[i].GetValue('ACT_AFFAIRE') + ';';
    Arg := arg+ 'NUM=' + IntToStr(TobBoniSelec.Detail[i].GetValue('ACT_NUMAPPREC')) + ';';
    Arg := Arg+ 'DATE=' + DateToStr(TobBoniSelec.Detail[i].GetValue('ACT_DATEACTIVITE')) + ';';
    Arg := Arg+'TIERS=' + TobBoniSelec.Detail[i].GetValue('ACT_TIERS') + ';';
    RepEF:=AFLanceFiche_ModifBoni(Arg) ;
    Inc(i);
    end;
finally
end;

except
Result:=1;
end;
end;



procedure TOF_AFMODIFBONI_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('ACT_AFFAIRE'));
Aff0:=nil;
Aff1:=THEdit(GetControl('ACT_AFFAIRE1'));
Aff2:=THEdit(GetControl('ACT_AFFAIRE2'));
Aff3:=THEdit(GetControl('ACT_AFFAIRE3'));
Aff4:=THEdit(GetControl('ACT_AVENANT'));
Tiers:=THEdit(GetControl('ACT_TIERS'));
end;

procedure TOF_AFMODIFBONI_MUL.AppelAppreciation (Affaire, DateApprec : string);
begin
  AFLanceFiche_AppreciationNiv0 ('ZORI:SIM1;ZAFF:' + Affaire + ';ZDATAPP:' + DateApprec + ';');

end;


/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLAfModifBoni(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_AFMODIFBONI_MUL) then TOF_AFMODIFBONI_MUL(MaTOF).AfModifBoni else exit;
end;

procedure AGLAppelAppreciation(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_AFMODIFBONI_MUL) then TOF_AFMODIFBONI_MUL(MaTOF).AppelAppreciation (Parms[1], Parms[2]) else exit;
end;


Procedure AFLanceFiche_ModifBoni_Mul(Argument:String);
begin
AGLLanceFiche ('AFF','AFMODIFBONI_MUL','','',Argument);
end;

Initialization
registerclasses ( [ TOF_AFMODIFBONI_MUL ] ) ;
RegisterAglProc('AfModifBoni',TRUE,1,AGLAfModifBoni);
RegisterAglProc('AppelAppreciation',TRUE,2,AGLAppelAppreciation);

end.
