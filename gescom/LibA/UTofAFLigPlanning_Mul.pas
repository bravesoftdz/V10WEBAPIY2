unit UTofAFLIGPLANNING_Mul;

interface

uses  StdCtrls,Controls,Classes,M3FP,HTB97,
{$IFDEF EAGLCLIENT}
      maineagl,eMul,utob,
{$ELSE}
      FE_Main,db,dbTables,DBGrids,mul,
{$ENDIF}
      forms,sysutils,HDB,
      ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,ConfidentAffaire,
      AffaireUtil,AffaireRegroupeUtil,DicoAF,SaisUtil,EntGC,
      utofAfBaseCodeAffaire,utilpgi,AglInit,UtilGc,TraducAffaire,
      UtilTaches, afplanning, paramsoc,GCMZSUtil;
      
Type
    TOF_AFLIGPLANNING_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnArgument(stArgument : String ); override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    private
      procedure TraiteRetour;
      function ExisteTache(pStAffaire : String) : Boolean;
      procedure ModifParLotLignesPlanning(Sender: TObject);
    End;

procedure AFLanceFiche_DetailPlanning;
procedure AFLanceFicheAFPlanningModiflot;
const
   TexteMsgTache: array[1..1] of string 	= (
      {1}        'Cette affaire ne contient aucune tâche');

implementation

procedure AFLanceFiche_DetailPlanning;
begin
  AGLLanceFiche('AFF','AFLIGPLANNING_MUL','','',''); // mul des lignes de planning
end;

procedure AFLanceFicheAFPlanningModiflot;
begin
  AGLLanceFiche('AFF','AFLIGPLANNING_MUL','','','LOTS'); // Modification des lignes de planning
end;

procedure TOF_AFLIGPLANNING_MUL.OnArgument(stArgument : String);
Begin
  Inherited;
  if not(GereSousAffaire) then
  begin
    SetcontrolVisible('TAFF_AFFAIREREF',False);
    SetcontrolVisible('AFFAIREREF1',False); SetcontrolVisible('AFFAIREREF2',False);
    SetcontrolVisible('AFFAIREREF3',False); SetcontrolVisible('AFFAIREREF4',False);
    SetcontrolVisible('AFF_ISAFFAIREREF',False);
    SetcontrolVisible('BSELECTAFF2',False);
  end;

{$IFDEF EAGLCLIENT}
  TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
  TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}
  if stArgument='LOTS' then
  begin
  {$IFDEF EAGLCLIENT}
    SetControlProperty('Fliste','MultiSelect',true);
  {$ELSE}
    SetControlProperty('Fliste','Multiselection',true);
  {$ENDIF}
    setcontrolVisible('BSELECTALL', true);
    setcontrolVisible('BINSERT', false);
    setcontrolVisible('BOuvrir1', true);
    setcontrolVisible('BOuvrir', false);
    setcontrolText('SELECTION','LOTS');
    Ecran.Caption := TraduitGA('Modification des lignes de planning');
    UpdateCaption(Ecran);
    TToolBarButton97 (GetControl('BOuvrir1')).OnClick  := ModifParLotLignesPlanning;
  end;
  {$IFDEF CCS3}
  if (getcontrol('PZONE') <> Nil) then SetControlVisible ('PZONE', False);
  if (getcontrol('PSTAT') <> Nil) then SetControlVisible ('PSTAT', False);
  {$ENDIF}
End;

procedure TOF_AFLIGPLANNING_MUL.OnUpdate;
Begin
  inherited;
  if Not (VH_GC.CleAffaire.Co2Visible) then SetControlText('APL_AFFAIRE2','');
  if Not (VH_GC.CleAffaire.Co3Visible) then SetControlText('APL_AFFAIRE3','');
End;

procedure TOF_AFLIGPLANNING_MUL.TraiteRetour;
Var stRetour  : string;
    Champ     : string;
begin
  stRetour := GetControlText ('RETOUR');
  if stRetour <> '' then
  begin
    Champ:=(Trim(ReadTokenSt(stRetour))); if Champ <> '' then SetControltext('APL_TIERS',Champ);
    Champ:=(Trim(ReadTokenSt(stRetour))); if Champ <> '' then SetControltext('APL_AFFAIRE1',Champ);
    Champ:=(Trim(ReadTokenSt(stRetour))); if (Champ <> '') And (VH_GC.CleAffaire.Co2Visible) then SetControltext('APL_AFFAIRE2',Champ);
    Champ:=(Trim(ReadTokenSt(stRetour))); if (Champ <> '') And (VH_GC.CleAffaire.Co3Visible) then SetControltext('APL_AFFAIRE3',Champ);
  end;
  SetControlText('RETOUR', '');
end;

procedure TOF_AFLIGPLANNING_MUL.OnLoad;
Var Affaire,Affaire0, Affaire1  ,Affaire2 ,Affaire3,Avenant : string;
begin
  inherited;
  Ecran.Caption := TraduitGA(Ecran.Caption);
  if GereSousAffaire then
  begin
    if (GetControlText('AFFAIREREF1')='') then SetControlText('AFF_AFFAIREREF','')
    else
    begin
      Affaire0 := 'A';
      Affaire1 := GetControlText('AFFAIREREF1'); Affaire2 := GetControlText('AFFAIREREF2');
      Affaire3 := GetControlText('AFFAIREREF3'); Avenant := GetControlText('AFFAIREREF4');
      Affaire:=CodeAffaireRegroupe(Affaire0, Affaire1  ,Affaire2 ,Affaire3,Avenant, taModif, false,false,false);
      if not ExisteAffaire(Affaire,'') then SetControlText('AFF_AFFAIREREF',Trim(Copy(Affaire,1,15)));
    end;
  end;
end;

procedure TOF_AFLIGPLANNING_MUL.ModifParLotLignesPlanning(Sender: TObject);
Var F : TFMul ;
    Parametrages : String;
    TheModifLot : TO_ModifParLot;
begin
  F:=TFMul(Ecran);
  if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
    begin MessageAlerte('Aucun élément sélectionné'); exit; end;

  TheModifLot := TO_ModifParLot.Create;
  TheModifLot.F := F.FListe;
  TheModifLot.Q := F.Q;
  TheModifLot.NatureTiers := '';
  TheModifLot.Nature := 'AFF';
  TheModifLot.Titre := Ecran.Caption;
  TheModifLot.TableName:='AFPLANNING';
  TheModifLot.FCode := 'APL_AFFAIRE;APL_NUMEROLIGNE';
  TheModifLot.FicheAOuvrir := 'AFPLANNING';
  V_PGI.ExtendedFieldSelection:='1';
  ModifieEnSerie(TheModifLot, Parametrages) ;
  if F.FListe.AllSelected then F.FListe.AllSelected:=False
                          else F.FListe.ClearSelected;
  F.bSelectAll.Down := False ;
end;

procedure TOF_AFLIGPLANNING_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff:=THEdit(GetControl('APL_AFFAIRE'));
  Aff1:=THEdit(GetControl('APL_AFFAIRE1')); Aff2:=THEdit(GetControl('APL_AFFAIRE2'));
  Aff3:=THEdit(GetControl('APL_AFFAIRE3')); Aff4:=THEdit(GetControl('APL_AVENANT'));
  Tiers:=THEdit(GetControl('APL_TIERS'));

  // affaire de référence pour recherche
  Aff_:=THEdit(GetControl('AFF_AFFAIREREF'));
  Aff1_:=THEdit(GetControl('AFFAIREREF1'));
  Aff2_:=THEdit(GetControl('AFFAIREREF2'));Aff3_:=THEdit(GetControl('AFFAIREREF3'));
  Aff4_:=THEdit(GetControl('AFFAIREREF4'));
end;

function TOF_AFLIGPLANNING_MUL.ExisteTache(pStAffaire : String) : Boolean;
var vQr : TQuery;
    S   : String;
begin
  Result := True;
  S := 'select ATA_LIBELLETACHE1, ATA_QTEINITIALE, ATA_FONCTION, ';
  S := S + 'ATA_QTEAPLANIFIER, ATA_NUMEROTACHE, ATA_ARTICLE ';
  S := S + 'FROM AFFAIRE, TACHE ';
  S := S + 'where AFF_AFFAIRE = "' + pStAffaire + '"';
  S := S + 'and AFF_AFFAIRE = ATA_AFFAIRE ';
  S := S + 'ORDER BY ATA_NUMEROTACHE';
  vQr := nil;
  Try
    vQr := OpenSql(S, True);
    if vQr.Eof then
      begin
        PGIBoxAF (TexteMsgTache[1],'');
        Result := false;
      end;
  Finally
    if vQr <> Nil then Ferme(vQr);
  End;
end;


//******************************************************************************
//************* Fonctions appellées depuis le Scripts ***************************
//******************************************************************************
procedure MajMulAFLIGPLANNING( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFMul) then Latof:=TFMul(F).Latof else exit;
  if (LaTof is TOF_AFLIGPLANNING_MUL) then TOF_AFLIGPLANNING_MUL(LaTof).TraiteRetour else exit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 21/05/2002
Modifié le ... :   /  /
Description .. : tester si une tache existe pour cette affaire
               : pour pouvoir ouvrir le plan de charge
Mots clefs ... :
*****************************************************************}
function AGLExisteTache( parms: array of variant; nb: integer ) : Variant;
var  F : TForm;
     LaTof : TOF;
begin
  result := false;
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFMul) then Latof:=TFMul(F).Latof else exit;
  if (LaTof is TOF_AFLIGPLANNING_MUL) then
    result := TOF_AFLIGPLANNING_MUL(LaTof).ExisteTache(String(Parms[1])) ;
end;


Initialization
  registerclasses([TOF_AFLIGPLANNING_MUL]);
  RegisterAglProc('MajMulAFLIGPLANNING', TRUE , 0, MajMulAFLIGPLANNING);
  RegisterAglFunc('ExisteTache', TRUE, 1, AGLExisteTache);
end.
