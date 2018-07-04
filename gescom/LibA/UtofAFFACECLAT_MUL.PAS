{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 12/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFFACECLAT_MUL ()
Mots clefs ... : TOF;AFFACECLAT_MUL
*****************************************************************}
Unit UtofAFFACECLAT_MUL ;

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
     HEnt1,
     HMsgBox,
     UTOF, M3FP,Hqry,ed_tools,
     UtilMulTrt, utob, dicoaf, htb97,dialogs,utilcutoff,UtofAFMODIFECLFAC,
     UTofAfBaseCodeAffaire,UtofAFMODIFCUTOFF,Affaireutil,SaisUtil,ParamSoc,AppreciationUtil ;

Type
  TOF_AFFACECLAT_MUL = Class (TOF_AFBASECODEAFFAIRE)
  private
    //GS : THGRID;
    //TobAfcumul :TOB;
    TypeCumul : String; //,TypeSaisie,LesColonnes,Titre :String;
    //DEV : RDEVISE;
    //FindLigne: TFindDialog;
    //FindDebut,ModifGlobal :boolean;
  public  
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (stArgument : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
    procedure AfModifEclatFact;
    function ModifEclatFact(TobAffairesSelec:TOB):integer;
  end ;

Procedure AFLanceFiche_ModifEclatFactMul(Argument:String);

Implementation

procedure TOF_AFFACECLAT_MUL.OnArgument (stArgument : String ) ;
var Critere, Champ, valeur: String;
    X : integer;
begin
Inherited ;
  Critere:=(Trim(ReadTokenSt(stArgument)));
  While (Critere <>'') do
  begin
    X:=pos(':',Critere);
    if X<>0 then
       begin
       Champ:=copy(Critere,1,X-1);
       Valeur:=Copy (Critere,X+1,length(Critere)-X);
       end;
    if (Champ = 'TYPE')   then TypeCumul:=valeur;
    Critere:=(Trim(ReadTokenSt(stArgument)));
  end;
end ;


procedure TOF_AFFACECLAT_MUL.AfModifEclatFact;
var
  TobMAff : TOB; // , TOBLigne
  iNbAffaire:integer;
begin

  TobMAff:=TOB.Create('',Nil,-1);
  try
  SourisSablier;
  TraiteEnregMulListe (TFMul(Ecran), 'ACU_AFFAIRE','AFFACECLAT', TobMAff, True);

  iNbAffaire:=TobMAff.Detail.count;

  //
  if (iNbAffaire<>0) then
      ModifEclatFact(TobMAff)
  else
      PGIInfoAf('Vous devez sélectionner au-moins une affaire', Ecran.Caption);

  (*if (RepCutOff=0) then
     PgiInfoAf('Les modifications  été effectuée avec succès', Ecran.Caption)
  else
     PgiInfoAf('Un problème est survenu lors de la revalorisation', Ecran.Caption);*)

  // Rafraichit la liste
  TFMul(Ecran).ChercheClick;

finally
  TobMAff.Free;
  SourisNormale;
end;

TToolBarButton97(GetControl('bSelectAll')).Down:=false;
end;


function TOF_AFFACECLAT_MUL.ModifEclatFact(TobAffairesSelec:TOB):integer;
var
i,iNbAffaires:integer;
RepEF, Arg:string;
begin
Result:=0;

if DetermineModeEclatFact=tmeSans then exit;

try
//
iNbAffaires := TobAffairesSelec.Detail.Count;
try
i:=0;
while (i < iNbAffaires) and (RepEF<>'STOP') do
    begin
    Arg := 'AFFAIRE=' + TobAffairesSelec.Detail[i].GetValue('ACU_AFFAIRE') + ';';
    Arg := Arg + 'TIERS=' + TobAffairesSelec.Detail[i].GetValue('ACU_TIERS')+ ';';
    Arg := Arg + 'NATURE=' + TobAffairesSelec.Detail[i].GetValue('ACU_NATUREPIECEG')+ ';';
    Arg := Arg + 'NUMECLAT=' + inttostr(TobAffairesSelec.Detail[i].GetValue('ACU_NUMECLAT'))+ ';';
    Arg := Arg + 'NUMERO=' + inttostr(TobAffairesSelec.Detail[i].GetValue('ACU_NUMERO'))+ ';';
    Arg := Arg + 'DATE=' + datetostr(TobAffairesSelec.Detail[i].GetValue('ACU_DATE')) + ';';
    RepEF:=AFLanceFiche_ModifEclatFact(Arg) ;
    Inc(i);
    end;

finally
end;

except
Result:=1;
end;
end;

procedure TOF_AFFACECLAT_MUL.OnLoad ;
begin
  Inherited ;
  SetControlText('ACU_TYPEAC',TypeCumul);
end ;

procedure TOF_AFFACECLAT_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFFACECLAT_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFFACECLAT_MUL.OnUpdate ;
begin
  Inherited ;
end ;


procedure TOF_AFFACECLAT_MUL.OnClose ;
begin
  Inherited ;
end ;


procedure TOF_AFFACECLAT_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('ACU_AFFAIRE'));
Aff0:=nil;
Aff1:=THEdit(GetControl('ACU_AFFAIRE1'));
Aff2:=THEdit(GetControl('ACU_AFFAIRE2'));
Aff3:=THEdit(GetControl('ACU_AFFAIRE3'));
Aff4:=THEdit(GetControl('ACU_AVENANT'));
Tiers:=THEdit(GetControl('ACU_TIERS'));
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLAfModifEclatFact(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_AFFACECLAT_MUL) then TOF_AFFACECLAT_MUL(MaTOF).AfModifEclatFact else exit;
end;


Procedure AFLanceFiche_ModifEclatFactMul(Argument:String);
begin
AGLLanceFiche ('AFF','AFFACECLAT_MUL','','',Argument);
end;

Initialization
registerclasses ( [ TOF_AFFACECLAT_MUL ] ) ;
RegisterAglProc('AfModifEclatFact',TRUE,1,AGLAfModifEclatFact);

end.
