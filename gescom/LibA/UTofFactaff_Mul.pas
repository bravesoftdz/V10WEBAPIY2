unit UTofFactaff_Mul;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
   eMul,
{$ELSE}
   dbTables, db,mul,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOF,
      DicoAF,SaisUtil,EntGC,utofAfBaseCodeAffaire;
Type
     TOF_FACTAFF_MUL = Class (TOF_AFBASECODEAFFAIRE)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; override ;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
     END ;
implementation


procedure TOF_FACTAFF_MUL.OnArgument(stArgument : String );
Begin
Inherited;
UpdateCaption(Ecran);
End;

procedure TOF_FACTAFF_MUL.OnUpdate;
Begin
inherited;
if Not (VH_GC.CleAffaire.Co2Visible) then SetControlText('AFF_AFFAIRE2','');
if Not (VH_GC.CleAffaire.Co3Visible) then SetControlText('AFF_AFFAIRE3','');
End;

procedure TOF_FACTAFF_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFA_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Tiers:=THEdit(GetControl('AFA_TIERS'));
end;

Initialization
registerclasses([TOF_FACTAFF_MUL]);
end.
