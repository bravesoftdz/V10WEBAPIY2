unit UTofEtatsTache;

interface
uses Classes,
{$IFDEF EAGLCLIENT}
  Maineagl,eMul,
{$ELSE}
   FE_Main,dbTables, db,  Mul,
{$ENDIF}
      //HEnt1,HMsgBox,UTOF,
       HCtrls,UtofEtatsAffaire
       //,      utob
       ;


Type
     TOF_ETATS_TACHE = Class (TOF_ETATS_AFF)
      procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
     public
     END ;

Procedure AFLanceFiche_FicheTache;


implementation


procedure TOF_ETATS_TACHE.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff0:=THEdit(GetControl('ATA_AFFAIRE0'));
Aff1:=THEdit(GetControl('ATA_AFFAIRE1'));
Aff2:=THEdit(GetControl('ATA_AFFAIRE2'));
Aff3:=THEdit(GetControl('ATA_AFFAIRE3'));
Aff4:=THEdit(GetControl('ATA_AVENANT'));
Aff0_:=THEdit(GetControl('ATA_AFFAIRE0_'));
Aff1_:=THEdit(GetControl('ATA_AFFAIRE1_'));
Aff2_:=THEdit(GetControl('ATA_AFFAIRE2_'));
Aff3_:=THEdit(GetControl('ATA_AFFAIRE3_'));
Aff4_:=THEdit(GetControl('ATA_AVENANT_'));
Tiers:=THEdit(GetControl('ATA_TIERS'));
Tiers_:=THEdit(GetControl('ATA_TIERS_'));
end;

Procedure AFLanceFiche_FicheTache;
begin
AGLLanceFiche ('AFF','FICHETACHE','','','');
end;

Initialization
registerclasses([TOF_ETATS_TACHE]);
end.
