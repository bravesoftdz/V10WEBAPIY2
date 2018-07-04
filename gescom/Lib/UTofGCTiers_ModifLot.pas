unit UTofGCTiers_ModifLot;

interface
uses UTOF,
{$IFDEF EAGLCLIENT}
     eMul,Maineagl,
{$ELSE}
     Mul, FE_Main,
{$ENDIF}
     GCMZSUtil,Forms,Classes,M3FP,HMsgBox,Controls,HEnt1,HCtrls,EntGc,UtilGc;

Type
     TOF_GCTiers_ModifLot = Class(TOF)
       private
       procedure ModificationParLotDesTiers;
       public
       Argument : string;
       procedure OnClose ; override ;
       procedure OnLoad ; override ;
       procedure OnArgument (stArgument : string); override;
     end;

procedure AGLModifParLotDesTiers(parms : array of variant; nb : integer);
Procedure AFLanceFiche_ModifLot_TiersAff(Argument:string);

implementation
procedure TOF_GCTiers_ModifLot.OnClose;
begin
V_PGI.ExtendedFieldSelection:='' ;  //mcd 06/10/2003 10912
end;

procedure TOF_GCTiers_ModifLot.OnArgument (stArgument : string);
begin
if stArgument = 'CLI' then
  begin   //modif MCD 03/04/2003 pour si Affaire, ne proposer que CLI .. car accès modif lot foun dans menu fourn
   if (ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then SetControlText ('T_NATUREAUXI', 'CLI')
   else SetControlText ('T_NATUREAUXI', 'CLI;PRO'); // mcd ancienne ligne
   end
else
    SetControlText ('T_NATUREAUXI', stArgument);
Argument := stArgument;

  // ajout MCD 31/01/01 pour traduction stat si existe
GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'YTC_TABLELIBRETIERS', 10, '');
GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_VALLIBRE', 3, '');
GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_DATELIBRE', 3, '');
GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'YTC_BOOLLIBRE', 3, '');
GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_TEXTELIBRE', 3, '');
GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '');
If not(ctxScot in V_PGI.PGIContexte)  then
    begin
    SetControlVisible ('TT_MOISCLOTURE' , False);
    SetControlVisible ('T_MOISCLOTURE' , False);
    SetControlVisible ('TT_MOISCLOTURE_' , False);
    SetControlVisible ('T_MOISCLOTURE_' , False);
    end;
end;

procedure TOF_GCTiers_ModifLot.OnLoad;
begin
if Argument = 'CLI' then Ecran.Caption:='Modification en série des clients'
                    else Ecran.Caption:='Modification en série des fournisseurs';
UpdateCaption(Ecran);
end;

/////////////// ModificationParLotDesTiers //////////////
procedure TOF_GCTiers_ModifLot.ModificationParLotDesTiers;
Var F : TFMul ;
    Parametrages : String;
    TheModifLot : TO_ModifParLot;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  begin
  MessageAlerte('Aucun élément sélectionné');
  V_PGI.ExtendedFieldSelection:='' ;   //mcd 06/10/2003 10912
  exit;
  end;
//if HShowMessage('0;Confirmation;Confirmez vous la modification des clients?;Q;YN;N;N;','','')<>mrYes then exit;
TheModifLot := TO_ModifParLot.Create;
TheModifLot.F := F.FListe; TheModifLot.Q := F.Q;
TheModifLot.NatureTiers := Argument;
TheModifLot.Titre := Ecran.Caption;
TheModifLot.TableName:='TIERS;TIERSCOMPL';
TheModifLot.FCode := 'T_AUXILIAIRE';
    // debut affaire  MCD 27/06/00
if ctxAffaire in V_PGI.PGIContexte  then
    begin
    //TheModifLot.Titre := 'Modification en série des Clients';
      // fait de façon différente, car avant fiche à part de AFF. Laisser au cas où ...
    TheModifLot.Nature := 'GC';
    if Argument='CLI' then TheModifLot.FicheAOuvrir := 'GCTIERS'
    else if Argument='FOU' then TheModifLot.FicheAOuvrir := 'GCFOURNISSEUR';
    end else
    begin
    // fin Affaire
    TheModifLot.Nature := 'GC';
    if Argument='CLI' then TheModifLot.FicheAOuvrir := 'GCTIERS'
    else if Argument='FOU' then TheModifLot.FicheAOuvrir := 'GCFOURNISSEUR';
    end;    //affaire
if ctxAffaire  in V_PGI.PGIContexte  then begin
   if TheModifLot.NatureTiers='CLI' then V_PGI.ExtendedFieldSelection:='3';
   if TheModifLot.NatureTiers='FOU' then V_PGI.ExtendedFieldSelection:='4';
   end
else begin
   if TheModifLot.NatureTiers='CLI' then V_PGI.ExtendedFieldSelection:='2';
   if TheModifLot.NatureTiers='FOU' then V_PGI.ExtendedFieldSelection:='1';
   end;
ModifieEnSerie(TheModifLot, Parametrages) ;
V_PGI.ExtendedFieldSelection:='' ;   //mcd 06/10/2003 10912
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLModifParLotDesTiers(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_GCTiers_ModifLot) then TOF_GCTiers_ModifLot(TOTOF).ModificationParLotDesTiers else exit;
end;

Procedure AFLanceFiche_ModifLot_TiersAff(Argument:string);
begin
AGLLanceFiche('AFF','AFTIERS_MODIFLOT','','',argument);
end;


Initialization
registerclasses([TOF_GCTiers_ModifLot]) ;
RegisterAglProc('ModifParLotDesTiers',TRUE,0,AGLModifParLotDesTiers);
end.
 