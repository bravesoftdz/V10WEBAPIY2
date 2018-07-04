{***********UNITE*************************************************
Auteur  ...... :  MCD
Créé le ...... : 01/02/2001                                   
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFDUPLIC ()
Mots clefs ... : TOF;AFDUPLIC
*****************************************************************}
Unit UTOFAFDUPLIC ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils,  ComCtrls,
{$IFDEF EAGLCLIENT}
   eMul, Maineagl,
{$ELSE}
   dbTables, db,Mul, HDB,  FE_Main,
{$ENDIF}
     HCtrls,Hstatus, HEnt1, M3FP,HMsgBox, UTOF,Utob,DicoAf,
     affaireUtil,Entgc,uTofAfBaseCodeAffaire,
     AffaireDuplic,UtilMulTrt;

Type
  TOF_AFDUPLIC = Class (TOF_AFBASECODEAFFAIRE)
     procedure Duplic_Mission;
     function GenerationMis : boolean;
     procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override  ;
  end ;
  
var titre: string;
Procedure AFLanceFiche_Mul_DuplicMiss(Range:string);

Implementation


procedure TOF_AFDUPLIC.Duplic_Mission;
var  St,Sttexte : String;
     ret : boolean;

begin
  St := 'Confirmez vous la génération de ces Affaires ';
  If (PGIAskAf(st,titre)<> mrYes) then exit;

  Ret := GenerationMis;

  if (not Ret) then StTexte := 'Le traitement ne s''est pas fait en totalité'
   else    StTexte := 'Le traitement s''est terminé correctement';
  PGIInfoAf(StTexte,Titre);

END;

function  TOF_AFDUPLIC.GenerationMis: boolean;
var  wi{, NbPiece} : Integer;
     TobMAff, TobDet : TOB; // Toutes les affaires sélectionnées
     LastAffaire : string;
     dDateDebItv, dDateFinItv : TDateTime;
begin
  result := true;
  // Gestion d'une Tob pour stocker toutes les affaires à traiter
  TobMAff := Tob.Create('les affaires',NIL,-1);
  // PA le 29/08/2001 - Fonction de traitement des enreg du mul externalisée
  // mcd 10/10 mise vue plante si select all TraiteEnregMulListe (TFMul(Ecran), 'AFF_AFFAIRE','AFFAIRE', TobMAff, True);
  TraiteEnregMulListe (TFMul(Ecran), 'AFF_AFFAIRE','AFAFFTIERS', TobMAff, True);

  InitMove(TobMaff.Detail.count,' ');
  wi := 0;
  while wi < TobMAff.Detail.count  do
      Begin
      MoveCur(False);
      TobDet := TobMAff.Detail[wi];
      if (BlocageAffaire( 'MAF', TobDet.GetValue('AFF_AFFAIRE'), V_PGI.groupe, V_PGI.DateEntree, False, False, false, dDateDebItv, dDateFinItv, nil)<>tbaAucun) then
         begin
         PGIInfoAf(Format('Affaire %s %s non modifiable',[TobMAff.Detail[wi].getValue('AFF_TIERS'),TobMAff.Detail[wi].getValue('AFF_AFFAIRE')]),Titre);
         end
      else
         begin
         LastAffaire := DuplicationAffaire (tdaDuplic, TobDet.GetValue('AFF_AFFAIRE'), '',NIL,True, False, False);
         If LastAffaire ='' then result :=False;
         end; // fin si pas blocage affaire
      inc(wi);
      MoveCur (False);
      End;  // fin while  ,boucle sur la sélection

  TobMaff.free;
  FiniMove;
END;


 procedure TOF_AFDUPLIC.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Aff_:=THEdit(GetControl('AFF_AFFAIRE_'));
Aff0_:=THEdit(GetControl('AFF_AFFAIRE0_'));
Aff1_:=THEdit(GetControl('AFF_AFFAIRE1_'));
Aff2_:=THEdit(GetControl('AFF_AFFAIRE2_'));
Aff3_:=THEdit(GetControl('AFF_AFFAIRE3_'));
Aff4_:=THEdit(GetControl('AFF_AVENANT_'));
Tiers:=THEdit(GetControl('AFF_TIERS'));
Tiers_:=THEdit(GetControl('AFF_TIERS_'));
end;

procedure AGLDuplic_Mission( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
F:=TForm(Longint(Parms[0]));
if (F is TFMul) then Latof:=TFMul(F).Latof else laTof := nil;
if (Latof is TOF_AFDUPLIC) then TOF_AFDUPLIC(LaTof).Duplic_Mission else exit;
end;

Procedure AFLanceFiche_Mul_DuplicMiss(Range:string);
begin
AGLLanceFiche ('AFF','AFDUPLIC_MUL',Range,'','');
end;

Initialization
registerclasses ( [ TOF_AFDUPLIC ] ) ;
RegisterAglProc( 'Duplic_Mission',True,0,AGLDuplic_Mission);
end.
