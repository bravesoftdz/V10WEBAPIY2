unit UTofAFTACHELIG_Mul;

interface

uses  StdCtrls,Controls,Classes,M3FP,HTB97,
{$IFDEF EAGLCLIENT}
      maineagl,eMul,
{$ELSE}
      FE_main,db,dbTables,mul,
{$ENDIF}
      forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,utob,paramsoc,
      utofAfBaseCodeAffaire,AffaireUtil,TraducAffaire,AffaireRegroupeUtil,DicoAF,UtilTaches;

Type
  TOF_AFTACHELIG_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnArgument(stArgument : String ); override;
    procedure OnLoad; override;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
  End;

procedure AFLanceFiche_DetailTaches;

implementation

procedure AFLanceFiche_DetailTaches;
begin
  AGLLanceFiche('AFF','AFTACHELIG_MUL','','',''); // mul des lignes de taches;
end;

procedure TOF_AFTACHELIG_MUL.OnArgument(stArgument : String);
Begin
  Inherited;
  // et les sous affaires
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
  {$IFDEF CCS3}
  if (getcontrol('PZONE') <> Nil) then SetControlVisible ('PZONE', False);
  if (getcontrol('PSTAT') <> Nil) then SetControlVisible ('PSTAT', False);
  {$ENDIF}
End;

procedure TOF_AFTACHELIG_MUL.OnLoad;
var vStExiste,Affaire,Affaire0,Affaire1,Affaire2 ,Affaire3,Avenant : string;
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

  if (GetCheckBoxState('CBPLANIFIE') = cbChecked) then
        vStExiste := ' AND EXISTS (SELECT APL_NUMEROLIGNE FROM AFPLANNING ' +
                          'WHERE ATA_NUMEROTACHE = APL_NUMEROTACHE AND APL_AFFAIRE = ATA_AFFAIRE)'
  else if (GetCheckBoxState('CBPLANIFIE') = cbUnchecked) then
        vStExiste := ' AND NOT EXISTS (SELECT APL_NUMEROLIGNE FROM AFPLANNING ' +
                          'WHERE ATA_NUMEROTACHE = APL_NUMEROTACHE AND APL_AFFAIRE = ATA_AFFAIRE)'
  else  vStExiste := '';
  if (THEdit(GetControl('XX_WHERE'))<> nil) then  SetControlText('XX_WHERE',vStExiste);
end;

procedure TOF_AFTACHELIG_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff:=THEdit(GetControl('ATA_AFFAIRE'));
  Aff1:=THEdit(GetControl('ATA_AFFAIRE1')); Aff2:=THEdit(GetControl('ATA_AFFAIRE2'));
  Aff3:=THEdit(GetControl('ATA_AFFAIRE3')); Aff4:=THEdit(GetControl('ATA_AVENANT'));
  Tiers:=THEdit(GetControl('ATA_TIERS'));

  // affaire de référence pour recherche
  Aff_:=THEdit(GetControl('AFF_AFFAIREREF'));
  Aff1_:=THEdit(GetControl('AFFAIREREF1'));
  Aff2_:=THEdit(GetControl('AFFAIREREF2'));Aff3_:=THEdit(GetControl('AFFAIREREF3'));
  Aff4_:=THEdit(GetControl('AFFAIREREF4'));
end;

Initialization
  registerclasses([TOF_AFTACHELIG_MUL]);
end.
