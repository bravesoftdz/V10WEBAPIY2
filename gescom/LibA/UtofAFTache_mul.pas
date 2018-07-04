{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/12/2000
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFTACHE_MUL ()
Mots clefs ... : TOF;AFTACHE_MUL
*****************************************************************}
Unit UTOFAFTACHE_MUL ;

Interface

Uses StdCtrls, Controls, Classes,  forms,Graphics, sysutils,  ComCtrls,
{$IFDEF EAGLCLIENT}
   Maineagl,
{$ELSE}
   dbTables, db, FE_Main,
{$EndIF}
     HCtrls,CalcOleGenericAff, HEnt1, HMsgBox, UTOF, AffaireUtil, utofAfBaseCodeAffaire ;

Type
  TOF_AFTACHE_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (StArgument : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
  End ;

  Function AFLanceFicheAFTache_Mul(Lequel, Argument : String) : String;

Implementation

Function AFLanceFicheAFTache_Mul(Lequel, Argument : String) : String;
begin
  result := AGLLanceFiche('AFF','AFTACHE_MUL','', '',Argument);
end;

procedure TOF_AFTACHE_MUL.OnNew ;
begin
  Inherited ;
End ;

procedure TOF_AFTACHE_MUL.OnDelete ;
begin
  Inherited ;
End ;

procedure TOF_AFTACHE_MUL.OnUpdate ;
begin
  Inherited ;
End ;

procedure TOF_AFTACHE_MUL.OnLoad ;
begin
  inherited;
End ;

procedure TOF_AFTACHE_MUL.OnArgument (StArgument : String ) ;
Var
  X:integer;
  Tmp,champ, valeur,Affaire,Tiers,Aff0,Aff1,Aff2,Aff3,Avenant : String;

begin

  Inherited ;

  Affaire := '';
  Tiers := '';

  // traitement des arguments
  Tmp:=(Trim(ReadTokenSt(stArgument)));
  While (Tmp <>'') do
    Begin
      if Tmp<>'' then
        Begin
          X:=pos(':',Tmp);
          if x = 0 then X:=pos('=',Tmp);
          if x<>0 then
            begin
              Champ:=copy(Tmp,1,X-1);
              Valeur:=Copy (Tmp,X+1,length(Tmp)-X);
            End
          else Champ := Tmp;

          if Champ= 'ATA_AFFAIRE'    then Affaire:=valeur
          else if Champ= 'ATA_TIERS' then Tiers := valeur;
        End;

      Tmp:=(Trim(ReadTokenSt(stArgument)));
      End;

  // Paramétrage Affaire + Tiers
  if (Affaire <> '') then
    Begin
      SetControlText('ATA_AFFAIRE',Affaire);

      SetControlProperty ('ATA_AFFAIRE1','Enabled',False);
      SetControlProperty ('ATA_AFFAIRE1','Color',clBtnFace);
      SetControlProperty ('ATA_AFFAIRE2','Enabled',False);
      SetControlProperty ('ATA_AFFAIRE2','Color',clBtnFace);
      SetControlProperty ('ATA_AFFAIRE3','Enabled',False);
      SetControlProperty ('ATA_AFFAIRE3','Color',clBtnFace);
      SetControlProperty ('ATA_AVENANT' ,'Enabled',False);
      SetControlProperty ('ATA_AVENANT' ,'Color',clBtnFace);
      SetControlProperty ('ATA_TIERS','Enabled',False);
      SetControlProperty ('ATA_TIERS','Color',clBtnFace);
      SetControlvisible('BSELECTAFF1',False);
      SetControlvisible('BEFFACEAFF1',False);

      CodeAffaireDecoupe(Affaire,Aff0,Aff1,Aff2,Aff3,Avenant,taModif,false);

      SetControlText('ATA_AFFAIRE1',Aff1);
      SetControlText('ATA_AFFAIRE2',Aff2);
      SetControlText('ATA_AFFAIRE3',Aff3);
      setControlText('ATA_TIERS',Tiers);
    End;

  {$IFDEF CCS3}
  if (getcontrol('PTEXTLIBRE') <> Nil)  then SetControlVisible ('PTEXTLIBRE', False);
  if (getcontrol('PZONELIBRE') <> Nil)  then SetControlVisible ('PZONELIBRE', False);
  {$ENDIF}



End ;

procedure TOF_AFTACHE_MUL.OnClose ;
Begin
  Inherited ;
End ;

procedure TOF_AFTACHE_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
Begin
  Aff   := THEdit(GetControl('ATA_AFFAIRE'));
  Aff1  := THEdit(GetControl('ATA_AFFAIRE1'));
  Aff2  := THEdit(GetControl('ATA_AFFAIRE2'));
  Aff3  := THEdit(GetControl('ATA_AFFAIRE3'));
  Aff4  := THEdit(GetControl('ATA_AVENANT'));
  Tiers := THEdit(GetControl('ATA_TIERS'));
End;

Initialization
  registerclasses ( [ TOF_AFTACHE_MUL ] ) ;
End.

