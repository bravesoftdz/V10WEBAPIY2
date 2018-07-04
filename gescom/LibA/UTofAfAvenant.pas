unit UTofAfAvenant;
{***********UNITE*************************************************
Auteur  ...... : P ARANEGA
Créé le ...... : 06/06/2001
Modifié le ... :   /  /
Description .. : Source TOF de la creation d'avenant
Mots clefs ... : TOF;AFFAIRE;AVENANT
*****************************************************************}
Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls
     ,HCtrls, HEnt1, HMsgBox, UTOF,UtofAfBaseCodeAffaire,M3FP,AffaireDuplic
{$IFDEF EAGLCLIENT}
     ,MaineAGL
{$ELSE}
     ,Fe_Main, db {$IFNDEF DBXPRESS} ,dbTables {$ELSE} ,uDbxDataSet {$ENDIF}
{$ENDIF}
  ,UtomAffaire
   ;
Type
  TOF_AFAVENANT = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (stArgument : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    private
    AffaireRef,Tiers : string;
  end ;


Implementation
//************************* Evènements de la TOF ******************************/
procedure TOF_AFAVENANT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFAVENANT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFAVENANT.OnUpdate ;
Var AvenantRemplace, RepriseSurAvenant : Boolean;
    CB : TCheckBox;
    AffAv, Arg, Statut : String;

begin
  Inherited ;
  CB := TCheckBox(Getcontrol('AVENANTREMPLACE'));
  if CB <> Nil then
  begin
    AvenantRemplace := CB.checked;
    if AvenantRemplace then arg := arg + 'AVENANTREMPLACE:X;' ;
  end;
  CB := TCheckBox(Getcontrol('REPRISESURAVENANT'));
  if CB <> Nil then
  begin
    RepriseSurAvenant := CB.checked;
    if RepriseSurAvenant then arg := arg + 'REPRISESURAVENANT:X;' ;
  end;

AffAv := DuplicationAffaire (tdaAvenant, AffaireRef, Arg, Nil,True, False, False);

if (AffAv = '') then Exit;
// Lancement de la fiche Avenant
if Copy(AffAv,1,1) = 'P' then Statut := 'PRO' else Statut := 'AFF';
// mcd 12/06/02 AGLLanceFiche('AFF','AFFAIRE','',AffAv,'STATUT:' + Statut + ';ETAT:ENC');
AFLanceFiche_Affaire  (AffAv,'STATUT:' + Statut + ';ETAT:ENC');
end ;

procedure TOF_AFAVENANT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFAVENANT.OnArgument (stArgument : String ) ;
Var Tmp,Champ,Valeur : string;
    X : integer;
begin
Tmp:=(Trim(ReadTokenSt(stArgument)));
While (Tmp <> '') do
    BEGIN
    X:=pos(':',Tmp);
    if x<>0 then begin Champ:=copy(Tmp,1,X-1); Valeur:=Copy (Tmp,X+1,length(Tmp)-X); end;
    if Champ='AFF_AFFAIRE' then
      begin AffaireRef:=valeur; SetControlText('AFF_AFFAIRE',AffaireRef); end;
    if Champ='AFF_TIERS' then
      begin Tiers:=valeur; SetControlText('AFF_TIERS',Tiers); end;
    Tmp:=(Trim(ReadTokenSt(stArgument)));
    END;
// Attention alimenter le champs AFF_AFFAIRE avant l'inherited
Inherited ;
end ;

procedure TOF_AFAVENANT.OnClose;
begin
  Inherited ;
end ;


//************************** Gestion du code Affaire ancètre *******************
procedure TOF_AFAVENANT.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));   Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2')); Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));  Tiers:=THEdit(GetControl('AFF_TIERS'));
end;


// *********************** Création d'avenant sur script **********************************
Procedure AGLAvenantAffaire( parms: array of variant; nb: integer );
begin
     // inutile de faire une fct, cette fiche à cette tof associée ...
AGLLanceFiche('AFF','AFNEWAVENANT','','',Parms[0]);
end;

Initialization
registerclasses ( [ TOF_AFAVENANT] ) ;
RegisterAglProc( 'AvenantAffaire', False ,1, AGLAvenantAffaire);
end.

