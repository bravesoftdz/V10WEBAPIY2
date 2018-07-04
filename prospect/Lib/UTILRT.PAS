unit UtilRT;

interface

uses  ComCtrls,Sysutils, HCtrls, Forms, HEnt1, Controls,UtilTOM,STDCtrls,
{$IFDEF EAGLCLIENT}
     eFiche,eMul,
{$ELSE}
{$IFNDEF EAGLSERVER}
     {$IFNDEF ERADIO}
       Fiche,Mul,
     {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     Menus, M3FP,Paramsoc,HMsgBox,utob,Classes, EntRT, UtilGC,EntGC,TiersUtil
{$IFNDEF EAGLSERVER}
     {$IFNDEF ERADIO}
       ,cube,utof,KPMGUTIL
     {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
     ;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
  Procedure RTMajChampsLibres( FF : Tform ) ;
  Procedure RTMajChampsLibresFou( FF : Tform ) ;
  Procedure RTMajChampsLibresPersp( FF : Tform ) ;
  Procedure RTMajChampsLibresAction( FF : Tform ) ;
  Procedure RTMajChampsLibresActionF( FF : Tform ) ;
  Procedure RTMajChampsLibresChainage( FF : Tform ) ;
  Procedure RTMajChampsLibresChainageF( FF : Tform ) ;
  Procedure RTMajChampsLibresSuspect( FF : Tform ) ;
  Procedure RTMajChampsLibresContact( FF : Tform ) ;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

function  RTXXWhereConfident (typeconf:string;Istiers : boolean =false) :string;
function  RTDroitModifTiers (CodeTiers : string ): Boolean;
function  RTDroitModifFou (CodeTiers : string ): Boolean;
Procedure RTMajPopMenu( FF : Tform; MenuPop : String; NoLigne : Integer; AConserver : boolean) ;
function RTControleModifTiers (FF : Tform ;TobProspect: TOB; ActionModif :Boolean): Boolean;
function RTControleModifFou (FF : Tform ;TobProspect: TOB; ActionModif :Boolean): Boolean;
function RTValeurChampLibre(NomChamp,ValeurChamp: string):string;
Procedure CommercialToRessource ; // pour test CEGID charge la table des commerciaux dans la table Ressource
function RTDroitModifActions (CodeTiers,typeAction,Responsable :string): Boolean;
function RTDroitModifActionsF (CodeTiers,typeAction,Responsable :string): Boolean;
Procedure RTCorrespondSuspectProspect(CleChoixCode : string; FromTOB, ToTOB : TOB);
Function RTRechResponsable (Tiers: string) : string;
Function RTFormateDateHeure (Date: TDateTime; Heure : String) : String;
Function RTChangeLibre2 ( NomChamp : string ; FF : Tform ) : boolean;
Function RTExisteConfidentF : string ;

function RTAttribNewCodeSuspect (DernierChrono : string):string;
Procedure RTCreerLiensConcurrents ;  // déplace les concurrents des propositions dans fichier lien
function RTDroitModifTypeAction( ProduitPgi : String) : Boolean;

Function RTAccesMenu(NoMenu,NoTag,NoGrp:integer) : Boolean;
Function RTCalculEch (DateAct : TDateTime;Nbj : Integer; WE : String) : TDateTime;

{$IFNDEF GIGI} // $$$ JP 09/08/07: pas de CTI en GI (PCL), car fait dans le Bureau exclusivement
// $$$ JP 22/08/07: strRef: CHAMPCLE=VALEURCLE pour retrouver/sélectionner le tiers dans TFCtiAlerte.OnLoadContacts
procedure FonctionCTI (fonction, notel : string; strRef:string='');
{$ENDIF}

Function AGLRTExisteConfident ( parms: array of variant; nb: integer ): variant ;

implementation


{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Procedure RTMajChampsLibres( FF : Tform ) ;
var i: integer;
begin
{$IFDEF GCGC}
if FF=Nil then  exit;
For i:=1 to 10 do if i<10 then ChangeLibre2('TYTC_TABLELIBRETIERS'+intToStr(i),FF)else ChangeLibre2('TYTC_TABLELIBRETIERSA',FF);
For i:=1 to 3 do ChangeBoolLibre('YTC_BOOLLIBRE'+intToStr(i),FF);
For i:=1 to 3 do ChangeLibre2('TYTC_VALLIBRE'+intToStr(i),FF);
For i:=1 to 3 do ChangeLibre2('TYTC_DATELIBRE'+intToStr(i),FF);
{$ENDIF GCGC}
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Procedure RTMajChampsLibresFou( FF : Tform ) ;
var i: integer;
begin
{$IFDEF GCGC}
if FF=Nil then  exit;
For i:=1 to 3 do ChangeLibre2('TYTC_TABLELIBREFOU'+intToStr(i),FF);
For i:=1 to 3 do ChangeBoolLibre('YTC_BOOLLIBREFOU'+intToStr(i),FF);
For i:=1 to 3 do ChangeLibre2('TYTC_VALLIBREFOU'+intToStr(i),FF);
For i:=1 to 3 do ChangeLibre2('TYTC_DATELIBREFOU'+intToStr(i),FF);
{$ENDIF GCGC}
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Procedure RTMajChampsLibresContact( FF : Tform ) ;
begin
  if FF=Nil then  exit;
{$IFDEF GCGC}
  GCMAJChampLibre(FF, False, 'COMBO', 'C_LIBRECONTACT', 10, '');
  GCMAJChampLibre(FF, False, 'EDIT', 'C_VALLIBRE', 3, '_') ;
  GCMAJChampLibre(FF, False, 'EDIT', 'C_DATELIBRE', 3, '_');
  GCMAJChampLibre(FF, False, 'BOOL', 'C_BOOLLIBRE', 3, '_');
{$ENDIF GCGC}  
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Procedure RTMajChampsLibresContactSSelect( FF : Tform ) ;
var i: integer;
begin
{$IFDEF GCGC}
if FF=Nil then  exit;
//For i:=1 to 3 do RTChangeLibre2('TC_LIBRECONTACT'+intToStr(i),FF);
For i:=1 to 10 do if i<10 then RTChangeLibre2('TC_LIBRECONTACT'+intToStr(i),FF)else RTChangeLibre2('TC_LIBRECONTACTA',FF);
//For i:=1 to 3 do  RTChangeBoolLibre('C_BOOLLIBRE'+intToStr(i),FF);
//For i:=1 to 3 do RTChangeLibre2('TC_VALLIBRE'+intToStr(i),FF);
//For i:=1 to 3 do RTChangeLibre2('TC_DATELIBRE'+intToStr(i),FF);
{$ENDIF GCGC}
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Procedure RTMajChampsLibresPersp( FF : Tform ) ;
var i: integer;
begin
{$IFDEF GCGC}
if FF=Nil then  exit;
For i:=1 to 3 do ChangeLibre2('TRPE_TABLELIBREPER'+intToStr(i),FF);
{$ENDIF GCGC}
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Procedure RTMajChampsLibresAction( FF : Tform ) ;
var i: integer;
begin
{$IFDEF GCGC}
if FF=Nil then  exit;
For i:=1 to 3 do ChangeLibre2('TRAC_TABLELIBRE'+intToStr(i),FF);
{$ENDIF GCGC}
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Procedure RTMajChampsLibresActionF( FF : Tform ) ;
var i: integer;
begin
{$IFDEF GCGC}
if FF=Nil then  exit;
For i:=1 to 3 do ChangeLibre2('TRAC_TABLELIBREF'+intToStr(i),FF);
{$ENDIF GCGC}
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Procedure RTMajChampsLibresChainage( FF : Tform ) ;
var i: integer;
begin
{$IFDEF GCGC}
if FF=Nil then  exit;
For i:=1 to 3 do ChangeLibre2('TRCH_TABLELIBRECH'+intToStr(i),FF);
{$ENDIF GCGC}
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Procedure RTMajChampsLibresChainageF( FF : Tform ) ;
var i: integer;
begin
{$IFDEF GCGC}
if FF=Nil then  exit;
For i:=1 to 3 do ChangeLibre2('TRCH_TABLELIBRECHF'+intToStr(i),FF);
{$ENDIF GCGC}
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

Procedure RTMajChampsTypeAction( FF : Tform ) ;
var i: integer;
    CC : THLabel;
begin
if FF=Nil then  exit;
For i:=1 to 3 do
  begin
  CC:=THLabel(TomTofGetControl(FF,'TRPA_TABLELIBRE'+intToStr(i)));
  if CC=Nil then Exit;
  if CC.FocusControl = nil then Exit;
  CC.Caption:=ChampToLibelle('RAC_TABLELIBRE'+intToStr(i));
  if (Length(CC.Caption)>0) and (copy(CC.Caption,1,2)='.-') then
     begin
     CC.visible:=false;
     TomTofSetControlVisible(FF,CC.FocusControl.Name, False);
     end;
  end;
end;

Procedure RTMajChampsTypeActionF( FF : Tform ) ;
var i: integer;
    CC : THLabel;
begin
if FF=Nil then  exit;
For i:=1 to 3 do
  begin
  CC:=THLabel(TomTofGetControl(FF,'TRPA_TABLELIBREF'+intToStr(i)));
  if CC=Nil then Exit;
  if CC.FocusControl = nil then Exit;
  CC.Caption:=ChampToLibelle('RAC_TABLELIBREF'+intToStr(i));
  if (Length(CC.Caption)>0) and (copy(CC.Caption,1,2)='.-') then
     begin
     CC.visible:=false;
     TomTofSetControlVisible(FF,CC.FocusControl.Name, False);
     end;
  end;
end;

Procedure RTMajChampsTypeChainage( FF : Tform ) ;
var i: integer;
    CC : THLabel;
begin
if FF=Nil then  exit;
For i:=1 to 3 do
  begin
  CC:=THLabel(TomTofGetControl(FF,'TRPG_TABLELIBRECH'+intToStr(i)));
  if CC=Nil then Exit;
  if CC.FocusControl = nil then Exit;
  CC.Caption:=ChampToLibelle('RCH_TABLELIBRECH'+intToStr(i));
  if (Length(CC.Caption)>0) and (copy(CC.Caption,1,2)='.-') then
     begin
     CC.visible:=false;
     TomTofSetControlVisible(FF,CC.FocusControl.Name, False);
     end;
  end;

end;
Procedure RTMajChampsTypeChainageF( FF : Tform ) ;
var i: integer;
    CC : THLabel;
begin
if FF=Nil then  exit;
For i:=1 to 3 do
  begin
  CC:=THLabel(TomTofGetControl(FF,'TRPG_TABLELIBRECHF'+intToStr(i)));
  if CC=Nil then Exit;
  if CC.FocusControl = nil then Exit;
  CC.Caption:=ChampToLibelle('RCH_TABLELIBRECHF'+intToStr(i));
  if (Length(CC.Caption)>0) and (copy(CC.Caption,1,2)='.-') then
     begin
     CC.visible:=false;
     TomTofSetControlVisible(FF,CC.FocusControl.Name, False);
     end;
  end;

end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
//modif pour ciblage For i:=0 to 2  et for i:=0 to 9 pour les tables
Procedure RTMajChampsLibresSuspect( FF : Tform ) ;
var i: integer;
    AuMoinsUn,AuMoinsUn2,AuMoinsUn3,AuMoinsUnZL,AuMoinsUn6 : boolean;
    MaTablette : string;
    TobParSuspectCor, TobFillePSC : TOB;
begin
{$IFDEF GCGC}
if FF=Nil then  exit;
// Pl le 19/05/07 : gestion des champs libres seulement pour KPMG pour l'instant
// Pl le 05/07/07 : FQ10622 gestion des champs libres pour tous mais avec paramsoc
if (GetParamSocSecur ('SO_RTPARAMAVANCES', false) = false ) then
  begin
  // on cache les nouveaux onglets elements libres
  TomTofSetControlVisible(FF,'RSU_BLOCNOTE_', False);

  if (tTabSheet(TomTofGetControl(FF,'PMULTITABLESLIBRES')) <> nil) then
    tTabSheet(TomTofGetControl(FF,'PMULTITABLESLIBRES')).TabVisible := False;
  if (tTabSheet(TomTofGetControl(FF,'PTABLESLIBRESNEW')) <> nil) then
    tTabSheet(TomTofGetControl(FF,'PTABLESLIBRESNEW')).TabVisible := False;
  if (tTabSheet(TomTofGetControl(FF,'PTABLESLIBRESNEW2')) <> nil) then
    tTabSheet(TomTofGetControl(FF,'PTABLESLIBRESNEW2')).TabVisible := False;
  if (tTabSheet(TomTofGetControl(FF,'PTABLESLIBRESNEW3')) <> nil) then
    tTabSheet(TomTofGetControl(FF,'PTABLESLIBRESNEW3')).TabVisible := False;
  if (tTabSheet(TomTofGetControl(FF,'PTIERSLIBRESNEW')) <> nil) then
    tTabSheet(TomTofGetControl(FF,'PTIERSLIBRESNEW')).TabVisible := False;
  if (tTabSheet(TomTofGetControl(FF,'PTEXTESLIBRES')) <> nil) then
    tTabSheet(TomTofGetControl(FF,'PTEXTESLIBRES')).TabVisible := False;
  if (tTabSheet(TomTofGetControl(FF,'PBOOLEENSLIBRES')) <> nil) then
    tTabSheet(TomTofGetControl(FF,'PBOOLEENSLIBRES')).TabVisible := False;
  if (tTabSheet(TomTofGetControl(FF,'PDATESLIBRES')) <> nil) then
    tTabSheet(TomTofGetControl(FF,'PDATESLIBRES')).TabVisible := False;
  if (tTabSheet(TomTofGetControl(FF,'PVALEURSLIBRES')) <> nil) then
    tTabSheet(TomTofGetControl(FF,'PVALEURSLIBRES')).TabVisible := False;

  AuMoinsUn:=False;
  For i:=0 to 9 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBTABLE'+intToStr(i),FF) or AuMoinsUn);
  if FF is TFMul then TTabSheet(FF.FindComponent('PTABLELIBRE')).TabVisible := AuMoinsUn;
  AuMoinsUnZL:=False;
  For i:=0 to 9 do AuMoinsUnZL:=(ChangeBoolLibre('RSC_RSCLIBBOOL'+intToStr(i),FF) or AuMoinsUnZL);
  For i:=0 to 9 do AuMoinsUnZL:=(ChangeLibre2('TRSC_RSCLIBVAL'+intToStr(i),FF) or AuMoinsUnZL);
  For i:=0 to 9 do AuMoinsUnZL:=(ChangeLibre2('TRSC_RSCLIBDATE'+intToStr(i),FF) or AuMoinsUnZL);
  if FF is TFMul then TTabSheet(FF.FindComponent('PZONELIBRE')).TabVisible := AuMoinsUnZL;
  if FF is TFCube then TTabSheet(FF.FindComponent('PCOMPLEMENT')).TabVisible := (AuMoinsUn or AuMoinsUnZL);
  if FF is TFFiche then
    begin
    AuMoinsUn:=False;
    For i:=0 to 9 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBTEXTE'+intToStr(i),FF) or AuMoinsUn);
    end
  end
else  // seulement pour KPMG pour l'instant
  begin
  // on cache les anciens onglets elements libres
  if (tTabSheet(TomTofGetControl(FF,'CL1')) <> nil) then
    tTabSheet(TomTofGetControl(FF,'CL1')).TabVisible := False;
  if (tTabSheet(TomTofGetControl(FF,'PLIBRE2')) <> nil) then
    tTabSheet(TomTofGetControl(FF,'PLIBRE2')).TabVisible := False;
  if (tTabSheet(TomTofGetControl(FF,'PTABLELIBRE')) <> nil) then
    tTabSheet(TomTofGetControl(FF,'PTABLELIBRE')).TabVisible := False;
  if (tTabSheet(TomTofGetControl(FF,'PZONELIBRE')) <> nil) then
    tTabSheet(TomTofGetControl(FF,'PZONELIBRE')).TabVisible := False;
  if FF is TFCube then
    if (tTabSheet(TomTofGetControl(FF,'PCOMPLEMENT')) <> nil) then
      TTabSheet(FF.FindComponent('PCOMPLEMENT')).TabVisible := False;

  // On renomme tous les champs qui étaient sur les anciens onglets
  IntervertirNomChampsXX ( FF, 'RSC_RSCLIBTABLE', 0, 9, true);
  IntervertirNomChampsXX ( FF, 'TRSC_RSCLIBTABLE', 0, 9, true);
  IntervertirNomChampsXX ( FF, 'RSC_RSCLIBBOOL', 0, 9, true);
  IntervertirNomChampsXX ( FF, 'RSC_RSCLIBVAL', 0, 9, true);
  IntervertirNomChampsXX ( FF, 'TRSC_RSCLIBVAL', 0, 9, true);
  IntervertirNomChampsXX ( FF, 'RSC_RSCLIBTEXTE', 0, 9, true);
  IntervertirNomChampsXX ( FF, 'TRSC_RSCLIBTEXTE', 0, 9, true);
  IntervertirNomChampsXX ( FF, 'RSC_RSCLIBDATE', 0, 9, true);
  IntervertirNomChampsXX ( FF, 'TRSC_RSCLIBDATE', 0, 9, true);

  // Gestion des libellés libres des champs libres
  AuMoinsUn:=False;
  AuMoinsUn6:=false;
  if (TControl(FF.FindComponent('TRSC_RSCLIBTABLE0')) <> nil) then
    For i:=0 to 2 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBTABLE'+intToStr(i),FF) or AuMoinsUn);
  if (TControl(FF.FindComponent('TRSC_RSCLIBTABLE3')) <> nil) then
    For i:=3 to 11 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBTABLE'+intToStr(i),FF) or AuMoinsUn);
  if (FF is TFMul) or (FF is TFFiche) then
    if (TTabSheet(FF.FindComponent('PTABLESLIBRESNEW')) <> nil) then
      TTabSheet(FF.FindComponent('PTABLESLIBRESNEW')).TabVisible := AuMoinsUn;

  AuMoinsUn:=False;
  if (TControl(FF.FindComponent('TRSC_RSCLIBTABLE12')) <> nil) then
    For i:=12 to 23 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBTABLE'+intToStr(i),FF) or AuMoinsUn);
  if (FF is TFMul) or (FF is TFFiche) then
    if (TTabSheet(FF.FindComponent('PTABLESLIBRESNEW2')) <> nil) then
      TTabSheet(FF.FindComponent('PTABLESLIBRESNEW2')).TabVisible := AuMoinsUn;

  AuMoinsUn:=False;
  if (TControl(FF.FindComponent('TRSC_RSCLIBTABLE24')) <> nil) then
    For i:=24 to 25 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBTABLE'+intToStr(i),FF) or AuMoinsUn);
  if (FF is TFMul) or (FF is TFFiche) then
    if (TTabSheet(FF.FindComponent('PTABLESLIBRESNEW3')) <> nil) then
      TTabSheet(FF.FindComponent('PTABLESLIBRESNEW3')).TabVisible := AuMoinsUn;

  if (TControl(FF.FindComponent('TRSC_6RSCLIBTABLE26')) <> nil) then
    For i:=26 to 35 do AuMoinsUn6:=(ChangeLibre2('TRSC_6RSCLIBTABLE'+intToStr(i),FF) or AuMoinsUn6);
  if (FF is TFFiche) or (FF is TFCube) then
    if (TGroupBox(FF.FindComponent('GBTABLELIBRESUR6C')) <> nil) then
      TGroupBox(FF.FindComponent('GBTABLELIBRESUR6C')).Visible := AuMoinsUn6;


  AuMoinsUn:=False;
  if (TControl(FF.FindComponent('RSC_RSCLIBBOOL0')) <> nil) then
    For i:=0 to 2 do AuMoinsUn:=(ChangeBoolLibre('RSC_RSCLIBBOOL'+intToStr(i),FF,true) or AuMoinsUn);
  if (TControl(FF.FindComponent('RSC_RSCLIBBOOL3')) <> nil) then
    For i:=3 to 9 do AuMoinsUn:=(ChangeBoolLibre('RSC_RSCLIBBOOL'+intToStr(i),FF,true) or AuMoinsUn);
  if (TControl(FF.FindComponent('RSC_RSCLIBBOOL10')) <> nil) then
    For i:=10 to 19 do AuMoinsUn:=(ChangeBoolLibre('RSC_RSCLIBBOOL'+intToStr(i),FF,true) or AuMoinsUn);
  if (FF is TFMul) or (FF is TFFiche) or (FF is TFCube) then
    if (TTabSheet(FF.FindComponent('PBOOLEENSLIBRES')) <> nil) then
      TTabSheet(FF.FindComponent('PBOOLEENSLIBRES')).TabVisible := AuMoinsUn;

  AuMoinsUn:=False;
  if (TControl(FF.FindComponent('TRSC_RSCLIBDATE0')) <> nil) then
    For i:=0 to 2 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBDATE'+intToStr(i),FF) or AuMoinsUn);
  if (TControl(FF.FindComponent('TRSC_RSCLIBDATE3')) <> nil) then
    For i:=3 to 9 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBDATE'+intToStr(i),FF) or AuMoinsUn);
  if (TControl(FF.FindComponent('TRSC_RSCLIBDATE10')) <> nil) then
    For i:=10 to 19 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBDATE'+intToStr(i),FF) or AuMoinsUn);
  if (FF is TFMul) or (FF is TFFiche) or (FF is TFCube) then
    if (TTabSheet(FF.FindComponent('PDATESLIBRES')) <> nil) then
      TTabSheet(FF.FindComponent('PDATESLIBRES')).TabVisible := AuMoinsUn;

  AuMoinsUn:=False;
  if (TControl(FF.FindComponent('TRSC_RSCLIBVAL0')) <> nil) then
    For i:=0 to 2 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBVAL'+intToStr(i),FF) or AuMoinsUn);
  if (TControl(FF.FindComponent('TRSC_RSCLIBVAL3')) <> nil) then
    For i:=3 to 9 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBVAL'+intToStr(i),FF) or AuMoinsUn);
  if (TControl(FF.FindComponent('TRSC_RSCLIBVAL10')) <> nil) then
    For i:=10 to 19 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBVAL'+intToStr(i),FF) or AuMoinsUn);
  if (FF is TFMul) or (FF is TFFiche) or (FF is TFCube) then
    if (TTabSheet(FF.FindComponent('PVALEURSLIBRES')) <> nil) then
      TTabSheet(FF.FindComponent('PVALEURSLIBRES')).TabVisible := AuMoinsUn;

  AuMoinsUn:=False;
  if (TControl(FF.FindComponent('TRSC_RSCLIBTEXTE0')) <> nil) then
    For i:=0 to 2 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBTEXTE'+intToStr(i),FF) or AuMoinsUn);
  if (TControl(FF.FindComponent('TRSC_RSCLIBTEXTE3')) <> nil) then
    For i:=3 to 9 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBTEXTE'+intToStr(i),FF) or AuMoinsUn);
  if (TControl(FF.FindComponent('TRSC_RSCLIBTEXTE10')) <> nil) then
    For i:=10 to 29 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBTEXTE'+intToStr(i),FF) or AuMoinsUn);
  if (FF is TFMul) or (FF is TFFiche) or (FF is TFCube) then
    if (TTabSheet(FF.FindComponent('PTEXTESLIBRES')) <> nil) then
      TTabSheet(FF.FindComponent('PTEXTESLIBRES')).TabVisible := AuMoinsUn;
  AuMoinsUn:=False;

  if (TControl(FF.FindComponent('TRSC_RSCLIBMUL0')) <> nil) then
    begin
    For i:=0 to 19 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBMUL'+intToStr(i),FF) or AuMoinsUn);
    if (FF is TFMul) or (FF is TFFiche) then
      if (TTabSheet(FF.FindComponent('PMULTITABLESLIBRES')) <> nil) then
        TTabSheet(FF.FindComponent('PMULTITABLESLIBRES')).TabVisible := AuMoinsUn;
    end;


  // On gère l'affectation des tables libres prospect dans le cas où il y a une correspondance de
  // mise en place entre les champs suspectscompl et prospect ou tierscompl :
  // on cache les champs THedit (table libre : RSC_RSCLIBTABLE10_,  multi libres RSC_RSCLIBMUL0_) et on rend visible
  // les champs THValComboBox qui sont dessous (table libre : RSC_RSCLIBTABLE10, multi libres RSC_RSCLIBMUL0_) en leur affectant
  // la bonne Datatype (celle du prospect ou tierscompl)
  TobParSuspectCor := nil;
  // pour les 10 premières tables libres de RSC_RSCLIBTABLE0 à 9, ce sont deja des THValComboBox, il faut juste leur
  // affecter la datatype du champ prospect ou tierscompl, et passer la propriété DataTypeParametrable à false
  AuMoinsUn:=False;
  For i:=0 to 9 do
    begin
    MaTablette := RenvoiTabletteCor (TobParSuspectCor, 'RSC_RSCLIBTABLE'+intToStr(i));
    if (MaTablette <> '') then
      begin
      THValComboBox(TomTofGetControl(FF, 'RSC_RSCLIBTABLE'+intToStr(i))).DataType := MaTablette;
      THValComboBox(TomTofGetControl(FF, 'RSC_RSCLIBTABLE'+intToStr(i))).DataTypeParametrable := false;
      TobFillePSC := TobParSuspectCor.FindFirst(['RSP_CHSUSPECT'], ['RSC_RSCLIBTABLE'+intToStr(i)], TRUE);
      THLabel(TomTofGetControl(FF, 'TRSC_RSCLIBTABLE'+intToStr(i))).caption := TobFillePSC.GetString('RSP_LIBELLE');
      AuMoinsUn:=true;
      end;
    end;
    // la suite des tables libres
  AuMoinsUn2:=False;
  AuMoinsUn3:=False;
  if (TControl(FF.FindComponent('RSC_RSCLIBTABLE10')) <> nil) then
    begin
    For i:=10 to 25 do
      begin
      MaTablette := RenvoiTabletteCor (TobParSuspectCor, 'RSC_RSCLIBTABLE'+intToStr(i));
      if (MaTablette <> '') then
      if (THValComboBox(FF.FindComponent('RSC_RSCLIBTABLE'+intToStr(i))) <> nil) then
        begin
        IntervertirNomChampsXX ( FF, 'RSC_RSCLIBTABLE', i, i, true);

        THValComboBox(TomTofGetControl(FF,'RSC_RSCLIBTABLE'+intToStr(i))).DataType := MaTablette;
        THValComboBox(TomTofGetControl(FF,'RSC_RSCLIBTABLE'+intToStr(i))).DataTypeParametrable := false;
        TobFillePSC := TobParSuspectCor.FindFirst(['RSP_CHSUSPECT'], ['RSC_RSCLIBTABLE'+intToStr(i)], TRUE);
        THLabel(TomTofGetControl(FF, 'TRSC_RSCLIBTABLE'+intToStr(i))).caption := TobFillePSC.GetString('RSP_LIBELLE');
        AuMoinsUn:=true;
        if (i > 11) and (i < 24) then
          AuMoinsUn2:=true
        else
        if (i > 23) then
          AuMoinsUn3:=true;
        end;
      end;
    end;
  if (TTabSheet(FF.FindComponent('PTABLESLIBRESNEW')) <> nil) then
    if AuMoinsUn and (TTabSheet(FF.FindComponent('PTABLESLIBRESNEW')).TabVisible=false)
                  and ((FF is TFMul) or (FF is TFFiche) or (FF is TFCube)) then
      TTabSheet(FF.FindComponent('PTABLESLIBRESNEW')).TabVisible := true;

  if (TTabSheet(FF.FindComponent('PTABLESLIBRESNEW2')) <> nil) then
    if AuMoinsUn2 and (TTabSheet(FF.FindComponent('PTABLESLIBRESNEW2')).TabVisible=false)
                  and ((FF is TFMul)) then
      TTabSheet(FF.FindComponent('PTABLESLIBRESNEW2')).TabVisible := true;

  if (TTabSheet(FF.FindComponent('PTABLESLIBRESNEW3')) <> nil) then
    if AuMoinsUn3 and (TTabSheet(FF.FindComponent('PTABLESLIBRESNEW3')).TabVisible=false)
                  and ((FF is TFMul)) then
      TTabSheet(FF.FindComponent('PTABLESLIBRESNEW3')).TabVisible := true;

  AuMoinsUn:=False;
  if (TControl(FF.FindComponent('RSC_6RSCLIBTABLE26')) <> nil) then
    begin
    For i:=26 to 35 do
      begin
      MaTablette := RenvoiTabletteCor (TobParSuspectCor, 'RSC_6RSCLIBTABLE'+intToStr(i));
      if (MaTablette <> '') then
        begin
        IntervertirNomChampsXX ( FF, 'RSC_6RSCLIBTABLE', i, i, true);

        THValComboBox(TomTofGetControl(FF,'RSC_6RSCLIBTABLE'+intToStr(i))).DataType := MaTablette;
        THValComboBox(TomTofGetControl(FF,'RSC_6RSCLIBTABLE'+intToStr(i))).DataTypeParametrable := false;
        TobFillePSC := TobParSuspectCor.FindFirst(['RSP_CHSUSPECT'], ['RSC_6RSCLIBTABLE'+intToStr(i)], TRUE);
        THLabel(TomTofGetControl(FF, 'TRSC_6RSCLIBTABLE'+intToStr(i))).caption := TobFillePSC.GetString('RSP_LIBELLE');
        AuMoinsUn:=true;
        end;
      end;
    end;
  if (FF is TFFiche) or (FF is TFCube) then
    if (TTabSheet(FF.FindComponent('PTABLESLIBRESNEW')) <> nil) then
      if AuMoinsUn and (TTabSheet(FF.FindComponent('PTABLESLIBRESNEW')).TabVisible=false)  then
        TTabSheet(FF.FindComponent('PTABLESLIBRESNEW')).TabVisible := true;
  if (FF is TFMul) then
    if (TTabSheet(FF.FindComponent('PTIERSLIBRESNEW')) <> nil) then
      if AuMoinsUn and (TTabSheet(FF.FindComponent('PTIERSLIBRESNEW')).TabVisible=false)  then
        TTabSheet(FF.FindComponent('PTIERSLIBRESNEW')).TabVisible := true;

  if (TGroupBox(FF.FindComponent('GBTABLELIBRESUR6C')) <> nil) then
    if AuMoinsUn and (TGroupBox(FF.FindComponent('GBTABLELIBRESUR6C')).Visible=false)
                  and ((FF is TFFiche) or (FF is TFCube)) then
      TGroupBox(FF.FindComponent('GBTABLELIBRESUR6C')).Visible := true;

  AuMoinsUn:=False;
  if (TControl(FF.FindComponent('RSC_RSCLIBMUL0')) <> nil) then
    begin
    For i:=0 to 19 do
      begin
      MaTablette := RenvoiTabletteCor (TobParSuspectCor, 'RSC_RSCLIBMUL'+intToStr(i));
      if (MaTablette <> '') then
        begin
        IntervertirNomChampsXX ( FF, 'RSC_RSCLIBMUL', i, i, true);

        THMultiValComboBox(TomTofGetControl(FF,'RSC_RSCLIBMUL'+intToStr(i))).DataType := MaTablette;
//        THValComboBox(TomTofGetControl(FF,'RSC_RSCLIBMUL'+intToStr(i))).DataTypeParametrable := true;
        TobFillePSC := TobParSuspectCor.FindFirst(['RSP_CHSUSPECT'], ['RSC_RSCLIBMUL'+intToStr(i)], TRUE);
        THLabel(TomTofGetControl(FF, 'TRSC_RSCLIBMUL'+intToStr(i))).caption := TobFillePSC.GetString('RSP_LIBELLE');
        end;
      end;
    end;
  if (TTabSheet(FF.FindComponent('PMULTITABLESLIBRES')) <> nil) then
    if AuMoinsUn and (TTabSheet(FF.FindComponent('PMULTITABLESLIBRES')).TabVisible=false)
                  and (FF is TFMul) or (FF is TFFiche) then
      TTabSheet(FF.FindComponent('PMULTITABLESLIBRES')).TabVisible := true;
  end;

  FreeAndNil(TobParSuspectCor);
{$ENDIF GCGC}
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

Procedure RTMasqueChampsLibresTiers( FF : Tform ) ;
var i: integer;
begin
if FF=Nil then  exit;
For i:=1 to 3 do
    begin
    TOMTOFSetControlVisible(FF,'YTC_DATELIBRE'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'TYTC_DATELIBRE'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'YTC_DATELIBRE'+intToStr(i)+'_',false);
    TOMTOFSetControlVisible(FF,'TYTC_DATELIBRE'+intToStr(i)+'_',false);
    TOMTOFSetControlVisible(FF,'YTC_BOOLLIBRE'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'YTC_VALLIBRE'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'TYTC_VALLIBRE'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'YTC_VALLIBRE'+intToStr(i)+'_',false);
    TOMTOFSetControlVisible(FF,'TYTC_VALLIBRE'+intToStr(i)+'_',false);
    end;
end;

Procedure RTMasqueChampsLibresFou( FF : Tform ) ;
var i: integer;
begin
if FF=Nil then  exit;
For i:=1 to 3 do
    begin
    TOMTOFSetControlVisible(FF,'YTC_DATELIBREFOU'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'TYTC_DATELIBREFOU'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'YTC_DATELIBREFOU'+intToStr(i)+'_',false);
    TOMTOFSetControlVisible(FF,'TYTC_DATELIBREFOU'+intToStr(i)+'_',false);
    TOMTOFSetControlVisible(FF,'YTC_VALLIBREFOU'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'TYTC_VALLIBREFOU'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'YTC_VALLIBREFOU'+intToStr(i)+'_',false);
    TOMTOFSetControlVisible(FF,'TYTC_VALLIBREFOU'+intToStr(i)+'_',false);
    end;
end;

function RTChangeLibre2 ( NomChamp : string ; FF : Tform ) : boolean;
var CC : THLabel;
begin
result:=True; {Visible}
CC:=THLabel(TomTofGetControl(FF,NomChamp));
if CC=Nil then Exit;
if CC.FocusControl = nil then Exit;
CC.Caption:=ChampToLibelle(copy(NomChamp,2,length(NomChamp)));
if (Length(CC.Caption)>0) and (copy(CC.Caption,1,2)='.-') then
   begin
   CC.visible:=false; TomTofSetControlVisible(FF,CC.Name+'_', False);
   TomTofSetControlVisible(FF,CC.FocusControl.Name, False);  TomTofSetControlVisible(FF,CC.FocusControl.Name+'_', False);
   Result:=False;
   end;
end;

/////// Pour Script AGL ///////////////////////
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure AGLRTMajChampsLibres( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibres(TForm(Longint(Parms[0]))) ;
end;
procedure AGLRTMajChampsLibresFou( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresFou(TForm(Longint(Parms[0]))) ;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure AGLRTMajChampsLibresContact( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresContact(TForm(Longint(Parms[0]))) ;
end;
procedure AGLRTMajChampsLibresContactSSelect( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresContactSSelect(TForm(Longint(Parms[0]))) ;
end;
procedure AGLRTMajChampsLibresPersp( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresPersp(TForm(Longint(Parms[0]))) ;
end;
procedure AGLRTMajChampsLibresAction( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresAction(TForm(Longint(Parms[0]))) ;
end;
procedure AGLRTMajChampsLibresActionF( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresActionF(TForm(Longint(Parms[0]))) ;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure AGLRTMajChampsLibresChainage( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresChainage(TForm(Longint(Parms[0]))) ;
end;
procedure AGLRTMajChampsLibresChainageF( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresChainageF(TForm(Longint(Parms[0]))) ;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure AGLRTMajChampsTypeAction( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsTypeAction(TForm(Longint(Parms[0]))) ;
end;
procedure AGLRTMajChampsTypeActionF( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsTypeActionF(TForm(Longint(Parms[0]))) ;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure AGLRTMajChampsTypeChainage( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsTypeChainage(TForm(Longint(Parms[0]))) ;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure AGLRTMajChampsTypeChainageF( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsTypeChainageF(TForm(Longint(Parms[0]))) ;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Function AGLRTXXWhereConfident( parms: array of variant; nb: integer ): variant ;
begin
  if nb >1 then result:=RTXXWhereConfident(Parms[0],Parms[1])    //mcd 01/06/07 ajout d'un paramètre
  else result:=RTXXWhereConfident(Parms[0],Parms[1]) ;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Function AGLRTDroitModifTiers( parms: array of variant; nb: integer ): variant ;
begin
  result:=RTDroitModifTiers(string(Parms[0])) ;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Function AGLRTDroitModifFou( parms: array of variant; nb: integer ): variant ;
begin
  result:=RTDroitModifFou(string(Parms[0])) ;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Function AGLRTDroitModifActions( parms: array of variant; nb: integer ): variant ;
begin
  result:=RTDroitModifActions(string(Parms[0]),string(parms[1]),string(parms[2])) ;
end;
Function AGLRTDroitModifActionsF( parms: array of variant; nb: integer ): variant ;
begin
  result:=RTDroitModifActionsF(string(Parms[0]),string(parms[1]),string(parms[2])) ;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure AGLRTMajChampsLibresSuspect( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresSuspect(TForm(Longint(Parms[0]))) ;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Function AGLRTRechResponsable( parms: array of variant; nb: integer ): variant ;
begin
  result:=RTRechResponsable(Parms[0]) ;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

Function  AGLRTRecupResponsable( parms: array of variant; nb: integer ): variant ;
begin
if Parms[0] = 'C' then
   result:=VH_RT.RTResponsable
else
   result:=VH_RT.RTNomResponsable;
end;


Function AGLRTFormateDateHeure( parms: array of variant; nb: integer ): variant ;
begin
  result:=RTFormateDateHeure(Parms[0],Parms[1]) ;
end;

Function AGLRTRecupTime( parms: array of variant; nb: integer ): variant ;
begin
//{$IFNDEF EAGLCLIENT}
 result:= FormatDateTime ('hh:nn',Time);
//{$ENDIF}
end;

Function AGLRTRecupDateJour( parms: array of variant; nb: integer ): variant ;
begin
 result:= DateToStr(Date);
end;

procedure AGLRTMasqueChampsLibresTiers( parms: array of variant; nb: integer ) ;
begin
  RTMasqueChampsLibresTiers(TForm(Longint(Parms[0]))) ;
end;

procedure AGLRTMasqueChampsLibresFou( parms: array of variant; nb: integer ) ;
begin
  RTMasqueChampsLibresFou(TForm(Longint(Parms[0]))) ;
end;

function RTXXWhereConfident (typeconf:string;Istiers:boolean =false) :string;
(*{$ifdef GIGI}
var vargroupeconf,str, groupeconf : string;
{$endif} *)
begin
  result:='';
  if typeconf ='CON' then result := VH_RT.RTConfWhereConsult
  else if typeconf ='MOD' then result := VH_RT.RTConfWhereModif
       else if typeconf ='CONF' then result := VH_RT.RFConfWhereConsult
            else if typeconf ='MODF' then result := VH_RT.RFConfWhereModif
                else if typeconf ='CREATP' then
                  begin
                  if VH_RT.RTCreatPropositions then result :='X' else result :='-';
                  end
                  else if typeconf ='CREATACT' then
                    begin
                    if VH_RT.RTCreatActions then result :='X' else result :='-';
                    end
                  else if typeconf ='CREATIC' then
                    begin
                    if VH_RT.RTCreatInfos then result :='X' else result :='-';
                    end
                    else if typeconf ='CREATICF' then
                      begin
                      if VH_RT.RFCreatInfos then result :='X' else result :='-';
                      end
                      else if typeconf ='CREATACTF' then
                        begin
                        if VH_RT.RFCreatActions then result :='X' else result :='-';
                        end
  ;
{$ifdef GIGI}
if Istiers then
begin //mcd 25/05/07 : ajout test appel depusi fiche tiers, car si fait depusi GRC (perspective, action
      //on est sur une vue qui n'a pas le champ tiers et tout plante
(* mcd 31/10/07 tout revu pour essayer d'optimiser
  vargroupeConf := VH_GC.AfGereCritGroupeConf;
  vargroupeConf :=ReadTokenPipe(vargroupeConf, ';ComboPlus;');
  if vargroupeConf = '' then
    exit;
  str := Trim(ReadTokenst(VarGroupeConf));
  while str <> '' do
  begin
    if GroupeConf <>'' then GroupeConf := GroupeConf + ',';
    GroupeConf :=GroupeConf + '"'  + str + '"';
    str := Trim(ReadTokenst(VarGroupeConf));
  end;
  if ((typeconf ='MOD') or (typeconf ='CON')) and (GetParamsocSecur('SO_AFGROUPECLIENT',false)) then
  begin   //mcd 28/03/2007 si GI et gestion groupe travail dossier, je prend en compte
          //mcd 01/06/07 on prend en plus, tous les tiers qui ne sont pas affecté à des dossiers pour OK PRO
          //PL 13/07/07 : correction bug : remplacement des '' par des ""  ici : AND ANN_TIERS <>"") or Exists
   result:= result + ' AND ( not exists (SELECT DOS_NODOSSIER FROM DOSSIER,ANNUAIRE WHERE ANN_GUIDPER=DOS_GUIDPER '
    +  'AND ANN_TIERS= T_TIERS AND ANN_TIERS <>"") or Exists (select 1 from dossiergrp where dog_nodossier='
    + '(select dos_nodossier from dossier,annuaire where ann_guidper=dos_guidper '
    + ' and ann_tiers= t_tiers and ann_tiers <>"") and dog_groupeconf in ('+ GroupeConf + ')))';
  end;  *)
  if ((typeconf ='MOD') or (typeconf ='CON')) and (GetParamsocSecur('SO_AFGROUPECLIENT',false)) then
  begin   //si GI et gestion groupe travail dossier, je prend en compte  le groupe de travail affecté au dossier
          //on prend en plus, tous les tiers qui ne sont pas affecté à des dossiers pour OK PRO
          //par contre, on n'a  pas les tiers qui n'ont pas d'annuaire associé
   result:= result + 'AND EXISTS (SELECT 1 FROM ANNUAIRE LefT JOIN DOSSIER ON ANN_GUIDPER=DOS_GUIDPER'
   +' WHERE ANN_TIERS= T_TIERS'
   +' AND (NOT EXISTS (SELECT 1 FROM DOSSIER WHERE ann_guidper= DOS_guidper )';
    // mcd 14/11/07 en fct paramsoc Bureau pour voir (ou pas)  dossier sans groupe
   If GetParamsocDPSecur ('SO_MDDOSSANSGRP',true) then
   begin
     Result := result+ ' or (not EXISTS (SELECT 1 FROM grpdonnees'
      +' left join liendosgrp on grp_nom=ldo_nom and grp_id=ldo_grpid'
      +' left join liendonnees on grp_nom=lnd_nom and grp_id=lnd_grpid'
      +'  WHERE grp_nom = "GROUPECONF" AND LDO_NODOSSIER = DOS_NODOSSIER  ))';
   end;
   Result:= result +' OR (EXISTS (SELECT 1 FROM grpdonnees '
   + ' left join liendosgrp on grp_nom=ldo_nom and grp_id=ldo_grpid'
   + ' left join liendonnees on grp_nom=lnd_nom and grp_id=lnd_grpid '
   +' WHERE grp_nom = "GROUPECONF"'
   + 'AND LDO_NODOSSIER = DOS_NODOSSIER AND lnd_USERid="'+V_PGI.user+'" ))))';
  end;
end;
{$endif}
end;

function RTDroitModifTiers (CodeTiers : string): Boolean;
var Sqlconf : string;
begin
Result:=true;
if (GetParamsocSecur('SO_RTCONFIDENTIALITE',False) = False) then exit;
Sqlconf := RTXXWhereConfident('MOD');
Result :=  VH_RT.DroitModifTiers ;
if VH_RT.DroitModifTiers  and (trim(Sqlconf) <> '') then result:=ExisteSQL('SELECT T_TIERS from RTTIERS WHERE T_TIERS ="'+CodeTiers+'"'+ Sqlconf );
end;

function RTDroitModifFou (CodeTiers : string): Boolean;
var Sqlconf : string;
begin
Result:=true;
if (GetParamsocSecur('SO_RFCONFIDENTIALITE',False) = False) then exit;
Sqlconf := RTXXWhereConfident('MODF');
Result :=  VH_RT.RFDroitModifTiers ;
if VH_RT.RFDroitModifTiers  and (trim(Sqlconf) <> '') then result:=ExisteSQL('SELECT T_TIERS from RFFOURNISSEURS WHERE T_TIERS ="'+CodeTiers+'"'+ Sqlconf );
end;


function RTDroitModifActions (CodeTiers,typeAction,Responsable :string): Boolean;
var    QQ :tquery;
       TobAcces : TOB;
       ListeActions,Operateur1,Operateur2 : string;
       Trouve, Autorise, Createur ,vide:boolean;
begin
  result := true;
  if (GetParamsocSecur('SO_RTCONFIDENTIALITE',False) = False) then exit;
{  if ((CodeTiers <>'') and not RTDroitModifTiers(CodeTiers)) then
  begin
    result := False; exit;
  end;    }
  TobAcces:=TOB.create ('PROSPECTCONF',NIL,-1);
  QQ := OpenSQL('Select * from PROSPECTCONF Where RTC_INTERVENANT="'+V_PGI.User+'" and RTC_TYPECONF="ACT"  and RTC_PRODUITPGI="GRC" ',True) ;
  if Not QQ.EOF then
  begin
    TobAcces.SelectDB('',QQ);
    ListeActions := TobAcces.GetValue('RTC_SQLCONF');
    Operateur1 :=  TobAcces.GetValue('RTC_OPER1');
    Operateur2 :=  TobAcces.GetValue('RTC_OPER2');
    if (ListeActions = '') then Vide := true else Vide := False;
    if (pos(typeAction,ListeActions ) > 0 ) then Trouve := true else  Trouve := False;
    If (operateur1 = '=') then Autorise:= true else  Autorise:=False;
    If  (operateur2 = '=') then Createur := true else Createur := False;
    if ((Responsable = '') or (VH_RT.RTResponsable = '')) then Createur := False;
    if ( Createur and (VH_RT.RTResponsable=Responsable) ) or (not Createur) then Createur := true else Createur :=False;
    if (( (Vide) or (Trouve and Autorise) or (not Trouve and not Autorise)) and Createur) then Result := true else Result := False ;
  end else
  begin
    if ((Responsable <> '') and  (VH_RT.RTResponsable <> '') and (VH_RT.RTResponsable=Responsable))
      or ( VH_RT.RTCreatActions ) then
      Result := true else Result := False;;
  end;
  Ferme(QQ);
  TobAcces.free;
end;

function RTDroitModifActionsF (CodeTiers,typeAction,Responsable :string): Boolean;
var    QQ :tquery;
       TobAcces : TOB;
       ListeActions,Operateur1,Operateur2 : string;
       Trouve, Autorise, Createur ,vide:boolean;
begin
  result := true;
  if (GetParamsocSecur('SO_RFCONFIDENTIALITE',False) = False) then exit;
  TobAcces:=TOB.create ('PROSPECTCONF',NIL,-1);
  QQ := OpenSQL('Select * from PROSPECTCONF Where RTC_INTERVENANT="'+V_PGI.User+'" and RTC_TYPECONF="ACT"  and RTC_PRODUITPGI="GRF" ',True) ;
  if Not QQ.EOF then
  begin
    TobAcces.SelectDB('',QQ);
    ListeActions := TobAcces.GetValue('RTC_SQLCONF');
    Operateur1 :=  TobAcces.GetValue('RTC_OPER1');
    Operateur2 :=  TobAcces.GetValue('RTC_OPER2');
    if (ListeActions = '') then Vide := true else Vide := False;
    if (pos(typeAction,ListeActions ) > 0 ) then Trouve := true else  Trouve := False;
    If (operateur1 = '=') then Autorise:= true else  Autorise:=False;
    If  (operateur2 = '=') then Createur := true else Createur := False;
    if ((Responsable = '') or (VH_RT.RTResponsable = '')) then Createur := False;
    if ( Createur and (VH_RT.RTResponsable=Responsable) ) or (not Createur) then Createur := true else Createur :=False;
    if (( (Vide) or (Trouve and Autorise) or (not Trouve and not Autorise)) and Createur) then Result := true else Result := False ;
  end else
  begin
    if ((Responsable <> '') and  (VH_RT.RTResponsable <> '') and (VH_RT.RTResponsable=Responsable)) then
      Result := true else Result := False;;
  end;
  Ferme(QQ);
  TobAcces.free;
end;

function RTControleModifTiers (FF : Tform ;TobProspect: TOB; ActionModif :Boolean): Boolean;
var TobModAcces,TobTiers,TobZonelibre : TOB;
    QQ :TQUERY;
    InfoAcces : string;
    ind :integer;

    Function CompareInfosFiche(TobAcces :TOB):Boolean;
    var Resultat : Array[0..2] of boolean;
        ValeurConf,ValeurConf2,LibelleChamp,LibelleValeur,LibelleValeur2,Operateur,LibOperateur,Lien,Prefixe : Array[0..2] of String;
        NomChamp,ValeurFiche,ValeurConf0 :string;
        ListeChamp: TStrings;
        ResulCompare,i,ResulCompare1,ResulCompare2 :integer;

        Function ResulConf (Resul1, Resul2:Boolean;lien :string):boolean;
        begin
          if (lien = 'Ou') then result := (Resul1 or Resul2)
          else  result := (Resul1 and Resul2);
        end;
    begin
      ResulCompare1:=0; ResulCompare2:=0;
      Result := false;
      if (TobAcces.GetValue('RTC_SQLCONF')='') then exit;
      ListeChamp := TStringList.Create;
      for i:=1 to 3 do
      begin
        NomChamp := TobAcces.getvalue('RTC_CHAMP'+intToStr(i));
        Prefixe [i-1] := Copy(NomChamp,1,Pos('_',NomChamp)-1);
        if (NomChamp <> '') then
        begin
          ListeChamp.add (NomChamp);
          LibelleChamp [i-1] := ChampToLibelle (NomChamp);
          Operateur [i-1] := TobAcces.getvalue('RTC_OPER'+intToStr(i));
          LibOperateur [i-1] := RechDom ('TTCOMPARE', Operateur[i-1],False);
          if (Operateur [i-1]='E') or (Operateur [i-1]='G') then
          begin
            ValeurConf0 := TobAcces.getvalue('RTC_VAL'+intToStr(i));
            ValeurConf    [i-1] := ReadTokenSt (ValeurConf0);
            ValeurConf2   [i-1] := ReadTokenSt (ValeurConf0);
            LibelleValeur [i-1] := RTValeurChampLibre(NomChamp,ValeurConf [i-1]);
            LibelleValeur2[i-1] := RTValeurChampLibre(NomChamp,ValeurConf2 [i-1]);
          end else
          begin
            ValeurConf [i-1] := TobAcces.getvalue('RTC_VAL'+intToStr(i));
            if (Operateur[i-1]='=') or (Operateur[i-1]='<>') or (Operateur[i-1]='>') or (Operateur[i-1]='>=') or (Operateur [i-1]='<') or (Operateur [i-1]='<=') then
              LibelleValeur[i-1]:= RTValeurChampLibre(NomChamp,ValeurConf [i-1])
            else  LibelleValeur[i-1]:= ValeurConf [i-1];
            ValeurConf2 [i-1] := '';   LibelleValeur2 [i-1]:= '';
          end;
        end else
        begin
          ValeurConf [i-1] := '';    Operateur [i-1] := '';
          LibelleChamp [i-1] := '';  LibelleValeur [i-1]:= '';
          ValeurConf2 [i-1] := '';   LibelleValeur2 [i-1]:= '';
        end;
      end;
      if ListeChamp.Count = 0 then Resultat[0] := true;

      Lien[0] := TobAcces.getvalue('RTC_LIEN1');
      Lien[1] := TobAcces.getvalue('RTC_LIEN2');
      Lien[2] := '';

      for i := 0 to ListeChamp.Count - 1 do
      begin
          ValeurFiche := '';
          NomChamp := ListeChamp [i];
          if ( (Operateur[i] = 'E' ) or (Operateur[i] = 'G' ) ) then
          InfoAcces := InfoAcces +'('+LibelleChamp[i]+' '+LibOperateur[i]+' '+LibelleValeur[i]+' et '+LibelleValeur2[i]+') '+Lien[i]+' '
          else
          InfoAcces := InfoAcces + '(' + LibelleChamp[i] + ' ' + LibOperateur[i]+ ' '  + LibelleValeur[i] +') '+Lien[i]+' ';

          if (not ActionModif) and (((copy(FF.name,1,7) = 'GCTIERS') and (Prefixe [i] = 'RPR'))
          or ((FF.name = 'RTPARAMCL') and (Prefixe [i] <> 'RPR'))) then
          begin  //En création controle séparé des Infos Fiche tiers et des Infos complémentaires
            Resultat[i] := true;
            Continue;
          end;
          if (Prefixe [i] = 'T')   then ValeurFiche := TobTiers.Getvalue(NomChamp)
          else if (Prefixe [i] = 'YTC') then  ValeurFiche := TobZoneLibre.Getvalue(NomChamp)
          else if (Prefixe [i] = 'RPR') then  ValeurFiche := TobProspect.Getvalue(NomChamp);
          ResulCompare := AnsiCompareText(ValeurFiche,ValeurConf[i]);
          if ( (Operateur[i] = 'E' ) or (Operateur[i] = 'G' ) ) then
          begin
             ResulCompare1 := AnsiCompareText(ValeurFiche,ValeurConf[i]);
             ResulCompare2 := AnsiCompareText(ValeurFiche,ValeurConf2[i]);
          end;
          if ( (Operateur[i] = '=' ) and (ResulCompare  = 0 ) )
          or ( (Operateur[i] = '<>') and (ResulCompare <> 0 ) )
          or ( (Operateur[i] = '>' ) and (ResulCompare  > 0 ) )
          or ( (Operateur[i] = '>=') and (ResulCompare >= 0 ) )
          or ( (Operateur[i] = '<' ) and (ResulCompare  < 0 ) )
          or ( (Operateur[i] = '<=') and (ResulCompare <= 0 ) )
          or ( (Operateur[i] = 'C' ) and (pos (ValeurConf[i],ValeurFiche )  =  1  ) ) // commence par
          or ( (Operateur[i] = 'D' ) and (pos (ValeurConf[i],ValeurFiche ) <>  1  ) )// ne commence pas par
          or ( (Operateur[i] = 'L' ) and (pos (ValeurConf[i],ValeurFiche )  >  0  ) ) // contient
          or ( (Operateur[i] = 'M' ) and (pos (ValeurConf[i],ValeurFiche )  =  0  ) ) // ne contient pas
          or ( (Operateur[i] = 'I' ) and (pos (ValeurFiche,ValeurConf[i] )  >  0  ) ) // Est dans
          or ( (Operateur[i] = 'J' ) and (pos (ValeurFiche,ValeurConf[i] )  =  0  ) ) // N'est pas dans
          or ( (Operateur[i] = 'E' ) and (ResulCompare1 >= 0 ) and (ResulCompare2 <= 0 ) ) // Est entre
          or ( (Operateur[i] = 'G' ) and ((ResulCompare1 < 0 ) or (ResulCompare2 > 0 ))  ) // N'est pas entre
          then Resultat[i] := true;
      end;
      Result := Resultat[0];
      for i := 1 to ListeChamp.Count - 1 do
      begin
         Result := ResulConf (Result,Resultat[i],Lien[i-1]);
      end;
      ListeChamp.free;
    end;

begin
  Result := true;  InfoAcces := '';
  TobModAcces:=Nil;
  if (GetParamsocSecur('SO_RTCONFIDENTIALITE',False) = False) or ((copy(FF.name,1,7) <> 'GCTIERS') and (FF.name <> 'RTPARAMCL'))
  or (VH_RT.RTConfWhereModif = '') then exit;
  Result := true;
  try
  TobTiers:=TOB.Create ('TIERS', Nil, -1);
  TobZonelibre:=TOB.Create ('TIERSCOMPL', Nil, -1);
  if (copy(FF.name,1,7) = 'GCTIERS') then
  begin
    TobTiers.GetEcran (FF,Nil);
    TobZoneLibre.GetEcran (FF,Nil);
    Tobprospect :=TOB.Create ('PROSPECTS', Nil, -1);
    Tobprospect.SelectDB('"'+TOBTiers.GetValue('T_AUXILIAIRE')+'"',Nil );
  end
  else if (FF.name = 'RTPARAMCL') then
  begin
    TobTiers.SelectDB('"'+Tobprospect.GetValue('RPR_AUXILIAIRE')+'"',Nil );
    TobZoneLibre.SelectDB('"'+Tobprospect.GetValue('RPR_AUXILIAIRE')+'"',Nil );
  end;
  TobModAcces:=TOB.create ('GRC Confidentialité',NIL,-1);
  QQ := OpenSQL('Select * from PROSPECTCONF Where RTC_INTERVENANT="'+V_PGI.User+'" and RTC_TYPECONF LIKE "MO%" and RTC_PRODUITPGI="GRC" ORDER BY RTC_TYPECONF DESC',True) ;
  if Not QQ.EOF then
  begin
    TobModAcces.LoadDetailDB('PROSPECTCONF','','',QQ, False) ;
    Result := False;
    for ind:=0 to TobModAcces.detail.count-1 do
    begin
     if (TobModAcces.detail[ind].GetValue('RTC_SQLCONF')<>'') then
       InfoAcces := InfoAcces + '#13 Critère '+IntToStr(ind+1)+' : ';
     Result := Result or CompareInfosFiche (TobModAcces.detail[ind]);
    end;
    if not Result then
    begin
     if ActionModif then  PGIBox(InfoAcces,'Accès en modification non autorisé')
     else PGIBox(InfoAcces,'Accès en création non autorisé');
    end;
  end;
  Ferme(QQ);

  Finally
    TobModAcces.Free;
    TobTiers.Free;
    TobZoneLibre.Free;
    if (copy(FF.name,1,7) = 'GCTIERS') then Tobprospect.Free;
  end;
end;

function RTControleModifFou (FF : Tform ;TobProspect: TOB; ActionModif :Boolean): Boolean;
var TobModAcces,TobTiers,TobZonelibre : TOB;
    QQ :TQUERY;
    InfoAcces : string;
    ind :integer;

    Function CompareInfosFicheFou(TobAcces :TOB):Boolean;
    var Resultat : Array[0..2] of boolean;
        ValeurConf,ValeurConf2,LibelleChamp,LibelleValeur,LibelleValeur2,Operateur,LibOperateur,Lien,Prefixe : Array[0..2] of String;
        NomChamp,ValeurFiche,ValeurConf0 :string;
        ListeChamp: TStrings;
        ResulCompare,i,ResulCompare1,ResulCompare2 :integer;

        Function ResulConf (Resul1, Resul2:Boolean;lien :string):boolean;
        begin
          if (lien = 'Ou') then result := (Resul1 or Resul2)
          else  result := (Resul1 and Resul2);
        end;
    begin
      ResulCompare1:=0; ResulCompare2:=0;
      Result := false;
      if (TobAcces.GetValue('RTC_SQLCONF')='') then exit;
      ListeChamp := TStringList.Create;
      for i:=1 to 3 do
      begin
        NomChamp := TobAcces.getvalue('RTC_CHAMP'+intToStr(i));
        Prefixe [i-1] := Copy(NomChamp,1,Pos('_',NomChamp)-1);
        if (NomChamp <> '') then
        begin
          ListeChamp.add (NomChamp);
          LibelleChamp [i-1] := ChampToLibelle (NomChamp);
          Operateur [i-1] := TobAcces.getvalue('RTC_OPER'+intToStr(i));
          LibOperateur [i-1] := RechDom ('TTCOMPARE', Operateur[i-1],False);
          if (Operateur [i-1]='E') or (Operateur [i-1]='G') then
          begin
            ValeurConf0 := TobAcces.getvalue('RTC_VAL'+intToStr(i));
            ValeurConf    [i-1] := ReadTokenSt (ValeurConf0);
            ValeurConf2   [i-1] := ReadTokenSt (ValeurConf0);
            LibelleValeur [i-1] := RTValeurChampLibre(NomChamp,ValeurConf [i-1]);
            LibelleValeur2[i-1] := RTValeurChampLibre(NomChamp,ValeurConf2 [i-1]);
          end else
          begin
            ValeurConf [i-1] := TobAcces.getvalue('RTC_VAL'+intToStr(i));
            if (Operateur[i-1]='=') or (Operateur[i-1]='<>') or (Operateur[i-1]='>') or (Operateur[i-1]='>=') or (Operateur [i-1]='<') or (Operateur [i-1]='<=') then
              LibelleValeur[i-1]:= RTValeurChampLibre(NomChamp,ValeurConf [i-1])
            else  LibelleValeur[i-1]:= ValeurConf [i-1];
            ValeurConf2 [i-1] := '';   LibelleValeur2 [i-1]:= '';
          end;
        end else
        begin
          ValeurConf [i-1] := '';    Operateur [i-1] := '';
          LibelleChamp [i-1] := '';  LibelleValeur [i-1]:= '';
          ValeurConf2 [i-1] := '';   LibelleValeur2 [i-1]:= '';
        end;
      end;
      if ListeChamp.Count = 0 then Resultat[0] := true;

      Lien[0] := TobAcces.getvalue('RTC_LIEN1');
      Lien[1] := TobAcces.getvalue('RTC_LIEN2');
      Lien[2] := '';

      for i := 0 to ListeChamp.Count - 1 do
      begin
          ValeurFiche := '';
          NomChamp := ListeChamp [i];
          if ( (Operateur[i] = 'E' ) or (Operateur[i] = 'G' ) ) then
          InfoAcces := InfoAcces +'('+LibelleChamp[i]+' '+LibOperateur[i]+' '+LibelleValeur[i]+' et '+LibelleValeur2[i]+') '+Lien[i]+' '
          else
          InfoAcces := InfoAcces + '(' + LibelleChamp[i] + ' ' + LibOperateur[i]+ ' '  + LibelleValeur[i] +') '+Lien[i]+' ';

          if (not ActionModif) and (((copy(FF.name,1,13) = 'GCFOURNISSEUR') and (Prefixe [i] = 'RD3'))
          or ((copy(FF.name,1,9) = 'RTPARAMCL') and (Prefixe [i] <> 'RD3'))) then
          begin  //En création controle séparé des Infos Fiche tiers et des Infos complémentaires
            Resultat[i] := true;
            Continue;
          end;
          if (Prefixe [i] = 'T')   then ValeurFiche := TobTiers.Getvalue(NomChamp)
          else if (Prefixe [i] = 'YTC') then  ValeurFiche := TobZoneLibre.Getvalue(NomChamp)
          else if (Prefixe [i] = 'RD3') then  ValeurFiche := TobProspect.Getvalue(NomChamp);
          ResulCompare := AnsiCompareText(ValeurFiche,ValeurConf[i]);
          if ( (Operateur[i] = 'E' ) or (Operateur[i] = 'G' ) ) then
          begin
             ResulCompare1 := AnsiCompareText(ValeurFiche,ValeurConf[i]);
             ResulCompare2 := AnsiCompareText(ValeurFiche,ValeurConf2[i]);
          end;
          if ( (Operateur[i] = '=' ) and (ResulCompare  = 0 ) )
          or ( (Operateur[i] = '<>') and (ResulCompare <> 0 ) )
          or ( (Operateur[i] = '>' ) and (ResulCompare  > 0 ) )
          or ( (Operateur[i] = '>=') and (ResulCompare >= 0 ) )
          or ( (Operateur[i] = '<' ) and (ResulCompare  < 0 ) )
          or ( (Operateur[i] = '<=') and (ResulCompare <= 0 ) )
          or ( (Operateur[i] = 'C' ) and (pos (ValeurConf[i],ValeurFiche )  =  1  ) ) // commence par
          or ( (Operateur[i] = 'D' ) and (pos (ValeurConf[i],ValeurFiche ) <>  1  ) )// ne commence pas par
          or ( (Operateur[i] = 'L' ) and (pos (ValeurConf[i],ValeurFiche )  >  0  ) ) // contient
          or ( (Operateur[i] = 'M' ) and (pos (ValeurConf[i],ValeurFiche )  =  0  ) ) // ne contient pas
          or ( (Operateur[i] = 'I' ) and (pos (ValeurFiche,ValeurConf[i] )  >  0  ) ) // Est dans
          or ( (Operateur[i] = 'J' ) and (pos (ValeurFiche,ValeurConf[i] )  =  0  ) ) // N'est pas dans
          or ( (Operateur[i] = 'E' ) and (ResulCompare1 >= 0 ) and (ResulCompare2 <= 0 ) ) // Est entre
          or ( (Operateur[i] = 'G' ) and ((ResulCompare1 < 0 ) or (ResulCompare2 > 0 ))  ) // N'est pas entre
          then Resultat[i] := true;
      end;
      Result := Resultat[0];
      for i := 1 to ListeChamp.Count - 1 do
      begin
         Result := ResulConf (Result,Resultat[i],Lien[i-1]);
      end;
      ListeChamp.free;
    end;

begin
  Result := true;  InfoAcces := '';
  TobModAcces:=Nil;
  if (GetParamsocSecur('SO_RFCONFIDENTIALITE',False) = False) or ((copy(FF.name,1,13) <> 'GCFOURNISSEUR') and (copy(FF.name,1,9) <> 'RTPARAMCL'))
  or (VH_RT.RFConfWhereModif = '') then exit;
  Result := true;
  try
  TobTiers:=TOB.Create ('TIERS', Nil, -1);
  TobZonelibre:=TOB.Create ('TIERSCOMPL', Nil, -1);
  if (copy(FF.name,1,13) = 'GCFOURNISSEUR') then
  begin
    TobTiers.GetEcran (FF,Nil);
    TobZoneLibre.GetEcran (FF,Nil);
    Tobprospect :=TOB.Create ('RTINFOS003', Nil, -1);
    Tobprospect.SelectDB('"'+TOBTiers.GetValue('T_AUXILIAIRE')+'"',Nil );
  end
  else if (copy(FF.name,1,9) = 'RTPARAMCL') then
  begin
    TobTiers.SelectDB('"'+Tobprospect.GetValue('RD3_CLEDATA')+'"',Nil );
    TobZoneLibre.SelectDB('"'+Tobprospect.GetValue('RD3_CLEDATA')+'"',Nil );
  end;
  TobModAcces:=TOB.create ('GRF Confidentialité',NIL,-1);
  QQ := OpenSQL('Select * from PROSPECTCONF Where RTC_INTERVENANT="'+V_PGI.User+'" and RTC_TYPECONF LIKE "MO%" and RTC_PRODUITPGI="GRF" ORDER BY RTC_TYPECONF DESC',True) ;
  if Not QQ.EOF then
  begin
    TobModAcces.LoadDetailDB('PROSPECTCONF','','',QQ, False) ;
    Result := False;
    for ind:=0 to TobModAcces.detail.count-1 do
    begin
     if (TobModAcces.detail[ind].GetValue('RTC_SQLCONF')<>'') then
       InfoAcces := InfoAcces + '#13 Critère '+IntToStr(ind+1)+' : ';
     Result := Result or CompareInfosFicheFou (TobModAcces.detail[ind]);
    end;
    if not Result then
    begin
     if ActionModif then  PGIBox(InfoAcces,'Accès en modification non autorisé')
     else PGIBox(InfoAcces,'Accès en création non autorisé');
    end;
  end;
  Ferme(QQ);

  Finally
    TobModAcces.Free;
    TobTiers.Free;
    TobZoneLibre.Free;
    if ( copy(FF.name,1,13) = 'GCFOURNISSEUR') then Tobprospect.Free;
  end;
end;

function RTValeurChampLibre(NomChamp,ValeurChamp: string):string;
var datatype :string;
begin
   result := ValeurChamp;
   if copy(NomChamp,1,19)='YTC_TABLELIBRETIERS' then    //YTC_TABLELIBRETIERS0 -> GCLIBRETIERS0
      datatype := FindEtReplace(NomChamp,'YTC_TABLE','GC',false)
   else if copy(NomChamp,1,15)='RPR_RPRLIBTABLE' then   //RPR_RPRLIBTABLE0 -> RT RPRLIBTABLE0
      datatype := FindEtReplace(NomChamp,'RPR_','RT',false)
   else
      datatype := Get_Join(NomChamp );
   if datatype <> '' then
     Result :=RechDom(datatype,ValeurChamp,false)
end;

{***********A.G.L.***********************************************
Auteur  ...... : Garnier Marie-Noëlle
Créé le ...... : 21/09/2001
Modifié le ... :   /  /
Description .. : Suppression d'un ligne dans un PopUpMenu
Mots clefs ... : POPUPMENU
*****************************************************************}
procedure AGLRTMajPopMenu( parms: array of variant; nb: integer ) ;
begin
  RTMajPopMenu(TForm(Longint(Parms[0])), String(Parms[1]), Integer(Parms[2]), Boolean(Parms[3])) ;
end;

Procedure RTMajPopMenu( FF : Tform; MenuPop : String; NoLigne : Integer; AConserver : boolean ) ;
var pop : TPopupMenu ;
begin
    pop :=TPopupMenu(TomTofGetControl(FF,MenuPop) ) ;
    if pop <> nil then pop.items[NoLigne].visible:=AConserver;
end;

Procedure CommercialToRessource ;
var T,TR,TC : TOB;
    Q : TQuery;
    i : integer;
begin
T:=Tob.create('Les ressources',Nil,-1);
TC:=Tob.create('Les commerciaux',Nil,-1);
Q := OpenSQL( 'SELECT * FROM COMMERCIAL', True );
if Not Q.EOF then TC.loadDetailDB('COMMERCIAL','','',Q,False );
Ferme(Q);
if TC.detail.count>0 then
   begin
   For i:=0 to TC.detail.count-1 do
       begin
       TR:=Tob.create('RESSOURCE',T,-1);
       TR.initValeurs (False);
       TR.putValue('ARS_TYPERESSOURCE','SAL');
       TR.putValue('ARS_RESSOURCE',TC.detail[i].getvalue('GCL_COMMERCIAL'));
       TR.putValue('ARS_LIBELLE',TC.detail[i].getvalue('GCL_LIBELLE'));
       end;
   T.InsertDBByNivel(True);
   end;
T.free; TC.Free;
end;


Function RTExisteConfidentF : string ;
begin
  result := 'MOD';
  if (GetParamsocSecur('SO_RFCONFIDENTIALITE',False) = False) then exit;
  if VH_RT.RFExisteConfident = false then
     result := '';
end;

Function AGLRTExisteConfidentF ( parms: array of variant; nb: integer ): variant ;
begin
  result:=RTExisteConfidentF;
end;

Function AGLRTExisteConfident ( parms: array of variant; nb: integer ): variant ;
begin
  result := 'MOD';
  if (GetParamsocSecur('SO_RTCONFIDENTIALITE',False) = False) then exit;
  if VH_RT.RTExisteConfident = false then
     result := '';
end;

Function AGLRTEagl ( parms: array of variant; nb: integer ): Variant ;
begin
result:=false;
{$IFDEF EAGLCLIENT}
result:=true;
{$ENDIF}
end;

Procedure RTCorrespondSuspectProspect(CleChoixCode : string; FromTOB, ToTOB : TOB);
var TobCorrespond: TOB;
    PrefixeTo,ChampTo,FieldNameFrom,FieldNameTo :string;
    ind : integer;
begin
PrefixeTo := TableToPrefixe (ToTOB.NomTable);
if (PrefixeTo='YTC') then ChampTo := 'RSP_CHTIERS'
else if (PrefixeTo='RPR') then ChampTo := 'RSP_CHCOMPL';
TobCorrespond:=tob.create('Correspondance',Nil,-1) ;
TobCorrespond.LoadDetailDB('PARSUSPECTCOR','','',nil,false,true);
//RSP_CHSUSPECT:champ libre suspect,RSP_CHTIERS:champ libre tiers,RSP_CHCOMPL:champ libre prospect
if TobCorrespond.detail.count>0 then
begin  
  For ind:=0 to TobCorrespond.detail.count-1 do
    begin
      FieldNameFrom := TobCorrespond.detail[ind].getvalue('RSP_CHSUSPECT');
      FieldNameTo := TobCorrespond.detail[ind].getvalue(ChampTo);
      if (trim (FieldNameFrom)<>'') and (copy (FieldNameTo ,1,3) = PrefixeTo) then
        ToTOB.PutValue (FieldNameTo, FromTOB.GetValue (FieldNameFrom));
    end;
end;
//FQ 10467
if PrefixeTo = 'YTC' then
begin
  ToTOB.PutValue('YTC_REPRESENTANT2', FromTOB.GetValue('RSC_REPRESENTANT2'));
  ToTOB.PutValue('YTC_REPRESENTANT3', FromTOB.GetValue('RSC_REPRESENTANT3'));
  ToTOB.PutValue('YTC_TAUXREPR1', FromTOB.GetValue('RSC_TAUXREPR1'));
  ToTOB.PutValue('YTC_TAUXREPR2', FromTOB.GetValue('RSC_TAUXREPR2'));
  ToTOB.PutValue('YTC_TAUXREPR3', FromTOB.GetValue('RSC_TAUXREPR3'));
end;
// Fin FQ 10467
TobCorrespond.cleardetail;
TobCorrespond.free;
end;

Function RTRechResponsable (Tiers: string) : string;
var Q : TQuery;
    Select,Utilisateur,Commercial : string;
begin
result:='';
Select := 'SELECT T_REPRESENTANT FROM TIERS WHERE T_TIERS = "'+ tiers +'"';
Q:=OpenSQL(Select, True);
if not Q.Eof then
   Commercial := Q.Fields[0].AsString;
Ferme(Q) ;
if commercial <> '' then
    begin
    Select := 'SELECT GCL_UTILASSOCIE FROM COMMERCIAL WHERE GCL_COMMERCIAL = "'+ commercial+'"';
    Q:=OpenSQL(Select, True);
    if not Q.Eof then
       Utilisateur := Q.Fields[0].AsString;
    Ferme(Q) ;
    if Utilisateur <> '' then
        begin
        Select := 'SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_UTILASSOCIE = "'+ Utilisateur+'"';
        Q:=OpenSQL(Select, True);
        if not Q.Eof then
            Result := Q.Fields[0].AsString;
        Ferme(Q) ;
        end;
    end;
end;
Function RTFormateDateHeure (Date: TDateTime; Heure : String) : String;
var hh : TDateTime;
begin
//hh:=strToDateTime(Heure);
hh:=Date+strToDateTime(Heure);
//{$IFNDEF EAGLCLIENT}
result:=UsTime(hh);
//{$ENDIF}
end;

function RTAttribNewCodeSuspect (DernierChrono : string):string;
var ChronoCode :string;
    i_Chrono{,i_Reste}: Int64;
begin
  ChronoCode := DernierChrono ;
  if not isnumeric(ChronoCode) then ChronoCode := '0';
  repeat
    i_Chrono := StrToInt64(ChronoCode);
    i_Chrono := i_Chrono+1;
//    i_Reste:=i_Chrono Mod 10 ;
//    if i_Reste = 0 then i_Chrono := i_Chrono+1; Fiche 10346 (pr ne pas exclure les no 10,20,...)
    ChronoCode := IntToStr(i_Chrono);
    if (Length(ChronoCode) > 17) then ChronoCode := '1';
  until (not ExisteSQL('SELECT RSU_SUSPECT FROM SUSPECTS WHERE RSU_SUSPECT = "'+ChronoCode+'"'));
  Result := ChronoCode;
end;
// déplacement des Concurrents des propositions dans un fichier liens
Procedure RTCreerLiensConcurrents  ;
var Q: TQuery;
    StListe,StConc,Sql : String;
    i : integer;
begin
{if GetParamsoc('SO_RTDEPLACECON') = true then
  PGIBox('Le déplacement des concurrents a déjà été effectué','Déplacement des concurrents')
else
   begin}
   if PGIAsk('Voulez-vous déplacer les concurrents ?', 'Déplacement des concurrents') <> mrYes then exit
   else
      begin

Q := OpenSQL ('SELECT RPE_CONCURRENTS,rpe_perspective,rpe_tiers FROM PERSPECTIVES where rpe_concurrents <> ""',True) ;
while Not Q.EOF do
   begin
   StListe := Q.FindField('RPE_CONCURRENTS').AsString;
   if StListe <> '' then
      Repeat
      StConc:=Trim(ReadTokenSt(StListe)) ;
      // recherche du code à traiter dans les codes restants pour éliminer les doublons
      if StListe <> '' then
      begin
        i:=1;
        While i<>0 do
        begin
          i:=Pos(StConc, StListe);
          if i <> 0 then
          begin
            Delete(StListe,i,length(StConc));
            if Copy(StListe,i,1)=';' then
               Delete(StListe,i,1);                
          end;
        end;
      end;
      // fin recherche du code à traiter dans les codes restants pour éliminer les doublons
      if (StConc<>'') and (length(StConc) < 18) then
          begin
              Sql :='INSERT INTO PERSPECTIVESTIERS (RPT_TIERS,RPT_PERSPECTIVE) VALUES ("';
              Sql := Sql+StConc+'",'+IntToStr(Q.FindField('rpe_perspective').AsInteger)+')';
              ExecuteSQL(Sql);
          end;
      until  StConc='' ;
   Q.Next;
   end;
Ferme(Q) ;

{      ExecuteSql ('Update perspectives set rpe_concurrents="" where rpe_concurrents <> ""');
      SetParamSoc ('SO_RTDEPLACECON',True)
      end;}
   end;
end;

Function  AGLRTAffichePlusMesActions( parms: array of variant; nb: integer ): variant ;
begin
  // parms 1 = 'X' : lire, = '-' : Ecrire
  if parms[1] = 'X' then
  begin
     if GetSynRegKey('RTMesActions','X',TRUE)='X' then
        Result:=False
     else
        Result:=True;
  end
  else
     SaveSynRegKey('RTMesActions',parms[0],TRUE);
end;

Function  AGLRTAfficheParamRappel ( parms: array of variant; nb: integer ): variant ;
var Duree,Critere : String;
begin
  // parms 1 = 'X' : lire, = '-' : Ecrire
  if parms[1] = 'X' then
  begin
    if GetSynRegKey('RTRappelAuto','X',TRUE)='X' then
      Result:='X;'
    else
      Result:='-;';

    Duree:=GetSynRegKey('RTRappelMinutes','',true);
    if Duree='' then
      Result:=Result+IntToStr(GetParamSocSecur('SO_RTACTDUREE',10))
    else
      Result:=Result+Duree;
  end
else
  begin
  Critere:=parms[0];
  SaveSynRegKey('RTRappelAuto',ReadToKenSt(Critere),TRUE);
  SaveSynRegKey('RTRappelMinutes',ReadToKenSt(Critere),TRUE);
  end;
end;

function RTDroitModifTypeAction( ProduitPgi : String) : Boolean;
begin
  //result := true;
  //if (GetParamsoc('SO_RTCONFIDENTIALITE') = False) then exit;
  result:=ExisteSQL('Select RTC_VAL1 from PROSPECTCONF Where RTC_INTERVENANT="'+V_PGI.User+'" and RTC_TYPECONF="ACT"  and RTC_PRODUITPGI="'+ProduitPgi+'" and RTC_VAL1="="');
end;

Function RTAccesMenu(NoMenu,NoTag,NoGrp:integer) : Boolean;
var Q : TQuery;
begin
  result:=true;
  Q:=OpenSql('SELECT MN_ACCESGRP FROM MENU WHERE MN_1='+IntToStr(NoMenu)+' AND MN_TAG='+IntToStr(NoTag),true);
  if not Q.Eof then
    if Copy(Q.FindField('MN_ACCESGRP').AsString,NoGrp,1) = '-' then result:=false;
  ferme(Q);
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Function AGLRTAccesMenu( parms: array of variant; nb: integer ) : variant;
begin
  result:=RTAccesMenu(Integer(Parms[0]),Integer(parms[1]),Integer(parms[2])) ;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

Function RTCalculEch (DateAct : TDateTime;Nbj : Integer; WE : String) : TDateTime;
var DateEch : TDateTime;
begin
DateEch:=PlusDate(DateAct, Nbj ,'J');
if (WE = 'X' ) then
   begin
   if DayOfWeek(DateEch) = 7 then
          DateEch:=PlusDate(DateEch,2,'J')
      else
          if DayOfWeek(DateEch) = 1 then
              DateEch:=PlusDate(DateEch,1,'J');
   end;
Result := DateEch;
end;

{$IFNDEF GIGI} // $$$ JP 09/08/07: pas de CTI en GI (PCL), car fait dans le Bureau exclusivement
procedure RTAGLFonctionCTI (parms:array of variant; nb:integer);
var
   strRef  :string;
begin
     if Length (Parms) > 2 then
         strRef := string (Parms [2])
     else
         strRef := '';

     FonctionCTI (string (Parms [0]), string (Parms [1]), strRef);
end;

// $$$ JP 22/08/07: strRef: CHAMPCLE=VALEURCLE pour retrouver/sélectionner le tiers dans TFCtiAlerte.OnLoadContacts
procedure FonctionCTI (fonction, notel : string; strRef:string);
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
var
   strTelephone   :string;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
begin
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
  if fonction = 'MAKECALL' then
  begin
    strTelephone := Trim (notel);
    if strTelephone <> '' then
      if PgiAsk ('Appeler le ' + strTelephone + ' ?') = mrYes then
      begin
        // $$$ voir si vraiment nécessaire VH_RT.ctiAlerte.Close;
        VH_RT.ctiAlerte.MakeCall (strTelephone, strRef); // $$$ JP 22/08/07
      end;
  end else if fonction = 'CALLWAIT' then
  begin
  	VH_RT.ctiAlerte.BRepondreClick(VH_RT.ctiAlerte);
  end else if fonction = 'CALLBYE' then
  begin
  	VH_RT.ctiAlerte.BTerminerClick(VH_RT.ctiAlerte);
  end else if fonction = 'GETCALL' then
  begin
  	VH_RT.ctiAlerte.BRepondreClick(VH_RT.ctiAlerte);
  end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
end;
{$ENDIF GIGI}

Function AGLRTGetSynRegKey( parms: array of variant; nb: integer ): variant ;
begin
 result:=GetSynRegKey(parms[0],parms[1],parms[2]);
end;

procedure AGLRTSaveSynRegKey( parms: array of variant; nb: integer ) ;
begin
SaveSynRegKey(parms[0],parms[1],parms[2]);
end;

Function AGLRTFinAnnee( parms: array of variant; nb: integer ) : variant ;
var
  Date_Fin: TDateTime;
begin
  Date_Fin := FinAnnee(StrToDateTime(parms[0]));
  result := DateToStr(Date_Fin);
end;

Function AGLRTSAVSeria( parms: array of variant; nb: integer ) : variant ;
begin
  result := VH_GC.SAVSeria;
end;

Function AGLRTGRCSeria( parms: array of variant; nb: integer ) : variant ;
begin
  result := VH_GC.GRCSeria;
end;

Function AGLRTctxGcAff( parms: array of variant; nb: integer ) : variant ;
begin
  result := ( ctxGCAFF in V_PGI.PGIContexte ) or ( ctxAffaire in V_PGI.PGIContexte );
end;

Initialization
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
  RegisterAglFunc( 'RTeAGL', FALSE , 0, AGLRTEagl);
  RegisterAglProc( 'RTMajChampsLibres', TRUE , 0, AGLRTMajChampsLibres);
  RegisterAglProc( 'RTMajChampsLibresFou', TRUE , 0, AGLRTMajChampsLibresFou);
  RegisterAglProc( 'RTMajChampsLibresContact', TRUE , 0, AGLRTMajChampsLibresContact);
  RegisterAglProc( 'RTMajChampsLibresPersp', TRUE , 0, AGLRTMajChampsLibresPersp);
  RegisterAglProc( 'RTMajChampsLibresAction', TRUE , 0, AGLRTMajChampsLibresAction);
  RegisterAglProc( 'RTMajChampsLibresActionF', TRUE , 0, AGLRTMajChampsLibresActionF);
  RegisterAglProc( 'RTMajChampsLibresChainage', TRUE , 0, AGLRTMajChampsLibresChainage);
  RegisterAglProc( 'RTMajChampsLibresChainageF', TRUE , 0, AGLRTMajChampsLibresChainageF);
  RegisterAglProc( 'RTMajChampsTypeAction', TRUE , 0, AGLRTMajChampsTypeAction);
  RegisterAglProc( 'RTMajChampsTypeActionF', TRUE , 0, AGLRTMajChampsTypeActionF);
  RegisterAglProc( 'RTMajChampsTypeChainage', TRUE , 0, AGLRTMajChampsTypeChainage);
  RegisterAglProc( 'RTMajChampsTypeChainageF', TRUE , 0, AGLRTMajChampsTypeChainageF);
  RegisterAglProc( 'RTMajPopMenu', TRUE , 2, AGLRTMajPopMenu);
  RegisterAglFunc( 'RTXXWhereConfident', FALSE , 1, AGLRTXXWhereConfident);
  RegisterAglFunc( 'RTDroitModifTiers', FALSE , 1, AGLRTDroitModifTiers);
  RegisterAglFunc( 'RTDroitModifFou', FALSE , 1, AGLRTDroitModifFou);
  RegisterAglFunc( 'RTDroitModifActions', FALSE , 3, AGLRTDroitModifActions);
  RegisterAglFunc( 'RTDroitModifActionsF', FALSE , 3, AGLRTDroitModifActionsF);
  RegisterAglFunc( 'RTExisteConfidentF', FALSE , 0, AGLRTExisteConfidentF);
  RegisterAglFunc( 'RTExisteConfident', FALSE , 0, AGLRTExisteConfident);
  RegisterAglProc( 'RTMajChampsLibresSuspect', TRUE , 0, AGLRTMajChampsLibresSuspect);
  RegisterAglFunc( 'RTRechResponsable', FALSE , 1, AGLRTRechResponsable);
  RegisterAglProc( 'RTMasqueChampsLibresTiers', TRUE , 0, AGLRTMasqueChampsLibresTiers);
  RegisterAglProc( 'RTMasqueChampsLibresFou', TRUE , 0, AGLRTMasqueChampsLibresFou);
  RegisterAglProc( 'RTMajChampsLibresContactSSelect', TRUE , 0, AGLRTMajChampsLibresContactSSelect);
  RegisterAglFunc( 'RTFormateDateHeure', FALSE , 2, AGLRTFormateDateHeure);
  RegisterAglFunc( 'RTRecupTime', FALSE , 2, AGLRTRecupTime);
  RegisterAglFunc( 'RTRecupDateJour', FALSE , 0, AGLRTRecupDateJour);
  RegisterAglFunc( 'RTRecupResponsable', FALSE , 1, AGLRTRecupResponsable);
  RegisterAglFunc( 'RTAffichePlusMesActions', FALSE , 2, AGLRTAffichePlusMesActions);
  RegisterAglFunc( 'RTAccesMenu', false , 3, AGLRTAccesMenu);
  RegisterAglFunc( 'RTAfficheParamRappel', FALSE , 2, AGLRTAfficheParamRappel);

{$IFNDEF GIGI} // $$$ JP 13/08/07
  RegisterAglProc( 'FonctionCTI', FALSE, 2, RTAGLFonctionCTI);
{$ENDIF}

  RegisterAglFunc( 'RTGetSynRegKey', FALSE , 3, AGLRTGetSynRegKey);
  RegisterAglProc( 'RTSaveSynRegKey', FALSE, 3, AGLRTSaveSynRegKey);
  RegisterAglFunc( 'RTFinAnnee', FALSE, 3, AGLRTFinAnnee);
  RegisterAglFunc( 'RTSAVSeria', FALSE, 0, AGLRTSAVSeria);
  RegisterAglFunc( 'RTGRCSeria', FALSE, 0, AGLRTGRCSeria);
  RegisterAglFunc( 'RTctxGcAff', FALSE, 0, AGLRTctxGcAff);
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
end.



