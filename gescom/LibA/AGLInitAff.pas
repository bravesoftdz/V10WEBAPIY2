unit AGLInitAff;

interface
uses M3VM, Hent1,UTOM, Forms, vierge,Stat, AglInit,HCtrls,M3FP,EntGC,UtilPGI,AffaireUtil;


implementation

uses
     UTomRessourcePR, {UTomActivitePaie,} UTofAfProfilGener,UTofAfTableauBord,Dicobtp,
     TiersUtil,
{$IFDEF EAGLCLIENT}
     Maineagl, eMul, eFiche, eFichList,
{$ELSE}
     Fe_Main, Mul, Fiche, FichList,
{$ENDIF}
     UTOF,UtofAfBaseLigne_Mul;

procedure AGLListePieceAffaire( parms: array of variant; nb: integer );
var
Tiers, Affaire,NatAUxi:string;
begin
Tiers := Parms[0]; Affaire := Parms[1];   NatAUxi := Parms[2];
If (NatAUxi='' ) then NatAuxi := 'CLI';
if (NatAuxi='CLI') and (ctxscot in V_PGI.PGIContexte) then NatAuxi:=NatAuxi+';NATURE=FR'; //mcd 13/02/02
// mcd 31/05/02 passe par ligne au lieu piece AGLLanceFiche('AFF','AFPIECE_MUL','GP_TIERS='+Tiers+'; GP_AFFAIRE='+Affaire,'','NOVISIBLE_AFFAIRE;NATUREAUXI:'+NatAuxi);
//mcd 12/06/02 AGLLanceFiche('AFF','AFREGRLIGNE_MUL','GL_TIERS='+Tiers+'; GL_AFFAIRE='+Affaire,'','NOAFFAIRE;NATUREAUXI:'+NatAuxi+';TABLE:GL');
 AFLanceFiche_Mul_RegroupLigne('GL_TIERS='+Tiers+'; GL_AFFAIRE='+Affaire,'NOAFFAIRE;NATUREAUXI:'+NatAuxi+';TABLE:GL;AFFAIRE:'+affaire);
End;


// Fiche PR Ressource recupération des zones du PopUp Menu + intégration dans la formule
procedure AGLRecupItemPopup( parms: array of variant; nb: integer );
var  F : TForm;
     OM : TOM;
     LaTof : TOF;

begin
OM := Nil; LaTof := Nil;
F:=TForm(Longint(Parms[0]));
if (F is TFFicheListe) then OM:=TFFicheListe(F).OM;
if (F is TFMul) then Latof:=TFMul(F).Latof;
if (F is TFVierge) then Latof:=TFVierge(F).Latof;
if (F is TFStat) then Latof:=TFStat(F).Latof;

If (LaTof <> nil) then
    Begin
    if (Latof is TOF_AFPROFILGENER) then   TOF_AFPROFILGENER(LaTof).RecupLibAcpte(Parms[1],Parms[2]);
    if (Latof is TOF_TBAffViewer)   then   TOF_TBAffViewer(LaTof).RecupFormuleEcart(Parms[1],Parms[2]);
    End
Else
    If (OM <> Nil) Then
    Begin
    if (OM is TOM_RessourcePR) then TOM_RessourcePR(OM).FormPRRessource(Parms[1]);
    End;
end;

procedure AGLGrisageFamilles( parms: array of variant; nb: integer );
var  F : TForm;
     OM : TOM;
begin
  F:=TForm(Longint(Parms[0]));
  if (F is TFFicheListe) then OM:=TFFicheListe(F).OM
  else if (F is TFFiche) then OM:=TFFiche(F).OM
  else OM := nil;

  if (OM is TOM_RessourcePR) then TOM_RessourcePR(OM).GereGrisageFamilles ;
//  if (OM is TOM_ActivitePaie) then TOM_ActivitePaie(OM).GereGrisageFamilles;
end;


procedure AGLGrisageArticle( parms: array of variant; nb: integer );
var  F : TForm;
     OM : TOM;
begin
  F:=TForm(Longint(Parms[0]));
  if (F is TFFicheListe) then OM:=TFFicheListe(F).OM
  else if (F is TFFiche) then OM:=TFFiche(F).OM
  else OM := nil;

  if (OM is TOM_RessourcePR) then TOM_RessourcePR(OM).GereGrisageArticle
//  else if (OM is TOM_ActivitePaie) then TOM_ActivitePaie(OM).GereGrisageArticle;
end;

procedure AGLGrisageProfil( parms: array of variant; nb: integer );
var  F : TForm;
     OM : TOM;
begin
  F:=TForm(Longint(Parms[0]));
  if (F is TFFicheListe) then OM:=TFFicheListe(F).OM
  else if (F is TFFiche) then OM:=TFFiche(F).OM
  else OM := nil;

//  if (OM is TOM_ActivitePaie) then TOM_ActivitePaie(OM).GereGrisageProfil;
end;


procedure AGLGrisageRessource( parms: array of variant; nb: integer );
var  F : TForm;
     OM : TOM;
begin
  F:=TForm(Longint(Parms[0]));
  if (F is TFFicheListe) then OM:=TFFicheListe(F).OM
  else if (F is TFFiche) then OM:=TFFiche(F).OM
  else OM := nil;

//  if (OM is TOM_ActivitePaie) then TOM_ActivitePaie(OM).GereGrisageRessource;
end;




Function AGLActionToString( parms: array of variant; nb: integer ): variant;
begin
result := ActionToString(Parms[0]);
end;

Function AGLStringToAction( parms: array of variant; nb: integer ): variant;
begin
result := StringToAction(Parms[0]);
end;

Function AGLReadToken( parms: array of variant; nb: integer ) : variant ;
var  tmp, tmp1 : string;
    indice,ii : integer;
begin
    // fct qui recoit l'indice + le champ. Recherche dans la chaine
    // la xième valeur de celle-ci et la retourne.
Indice := Integer(Parms[0]) ;
tmp := String(Parms[1]) ;
result:='';
ii:=0;
Repeat
    tmp1 := ReadTokenSt(tmp);
    if (ii= indice) then begin result:=tmp1; break; end; ;
    Inc (ii);
    until (tmp1 = '') ;

end;

 Function AGLTiersAUxi( parms: array of variant; nb: integer ) : variant ;
var  tmp  : string;

begin
tmp := String(Parms[0]) ;
result:= TiersAuxiliaire (tmp, False);  // rend auxiliaire sur code tiers fournit
end;

procedure AGLPGIInfoAF( parms: array of variant; nb: integer );
var
Msg, Titre : string;
begin
Msg := Parms[0];
Titre := Parms[1];
PGIInfoAF (Msg, Titre);
End;

Function GetNatureAuxiAff( parms: array of variant; nb: integer ) : variant ;
var Statut : string;
begin
Statut := Parms[0];
Result := FabriqueWhereNatureAuxiAff (Statut);
end;

//////////////////////////////////////////////////////////////////////////////
procedure initM3Affaire();
begin

 RegisterAglProc( 'ListePieceAffaire', False ,3, AGLListePieceAffaire);
 RegisterAglProc( 'AFRecupItemPopup',True,2,AGLRecupItemPopup);
 RegisterAglProc( 'GrisageFamilles',True,0,AGLGrisageFamilles);
 RegisterAglProc( 'GrisageArticle',True,0,AGLGrisageArticle);
 RegisterAglProc( 'GrisageProfil',True,0,AGLGrisageProfil);
 RegisterAglProc( 'GrisageRessource',True,0,AGLGrisageRessource);
 RegisterAglFunc( 'ActionToString',False,1,AGLActionToString);
 RegisterAglFunc( 'StringToAction',False,1,AGLStringToAction);
 RegisterAglFunc( 'ReadToken', FALSE , 2, AGLReadToken);
 RegisterAglFunc( 'TiersAuxi', FALSE , 2, AGLTiersAUxi);
 RegisterAglProc( 'PGIInfoAF', FALSE , 2, AGLPGIInfoAF);
 RegisterAglFunc( 'GetNatureAuxiAff', FALSE , 2, GetNatureAuxiAff);

end;


Initialization
initM3Affaire();
// report modif PCS 13/05/03 finalization
end.
