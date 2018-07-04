unit UtilTableauBord;

interface

uses
  Classes,
  Hctrls,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,
  {$ELSE}
  MaineAGL,
  {$ENDIF}
  sysutils,
  HEnt1,
  HMsgBox,
  UTOB,
  Entgc,
  AGLInit,
  Paramsoc,
  Ulog,
  AGLInitBTP;

type
	ToptPrevuAvanc = (OptDetailDev,OptDetailPBT,OptDetailETU,OptDetailBCE,OptGlobal,OptDetailFAC);
  TChoixPrevuAvanc = set of ToptPrevuAvanc;


procedure AjouteChampSupPrevu(TOBTmp : TOB; Nature : string);
procedure SetprevuAvance (TOBTMP : TOB; NaturePiece : string ; OptionChoixPrevuAvanc : TChoixPrevuAvanc; WherePiece : string='');
procedure AjouteChampSupRealiseDetail (TOBTmp : TOB);
procedure AjouteChampSupFactureDetail (TOBTmp : TOB);
procedure AjouteChampSupRestADep(TOBTMP : TOB);

Procedure MAJ_PremiereLigne(NomChampSup : String; TOBTMP : Tob);
Procedure Repartition_Eclatement(TypeMontant, NaturePres, TypeMt : String; Montant : Double; TOBTMP : Tob);
Procedure ChargeCumulPrevuFacture(TypeRes, NaturePiece : String; TOBTMP : TOB; MontantPa,MontantPr,MontantPV,TpsPrevu : Double);
Procedure Charge_Repartition_Eclatement(NaturePres : String; MontantPA, MontantPR, MontantPV : Double; TOBTMP : TOB);
Procedure ChargePrevu(TOBL, TOBTMP : TOB; Prefixe : String; Tps_Prevu : Double =0);

procedure TraiteOuvrageTBPlat (TOBTMP,TOBL,TOBOuvragePlat : TOB);

//Modif FV1 :
procedure TraiteDetailOuvrageTBPlat (TOBTMP,TOBOUV : TOB);
procedure TraiteDetailOuvrageTB (TOBTMP,TOBOUV : TOB; Qte,QteDuDetail : double);

  //Modif FV : Dev. prioritaire DSL le 05/06/2012
var CoefFG_Param  : Double;
    TauxHoraire   : Double;
    //
implementation
uses UtilsTOB;


procedure CumuleprevufactureAutre (TOBTMP: TOB; NaturePiece : string; MontantPa,MontantPr,MontantPV : double);
var LocNaturePiece : string;
begin

  if (Pos(NaturePiece ,'FBT;B00;ABT;FAC;AVC;FBC;FBP;ABC;')>0) Then
  begin
     TOBTMP.PutValue('FACTURE_AUTRE', TOBTMP.GetValue('FACTURE_AUTRE')+MontantPV);
  end
  else
  begin
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_AUTPA', TOBTMP.GetValue('PREVU_'+NaturePiece+'_AUTPA')+MontantPA);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_AUTPR', TOBTMP.GetValue('PREVU_'+NaturePiece+'_AUTPR')+MontantPR);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_AUTPV', TOBTMP.GetValue('PREVU_'+NaturePiece+'_AUTPV')+MontantPV);
  end;

end;

procedure CumuleprevufactureSalarie (TOBTMP : TOB ; NaturePiece : string; MontantPa,MontantPr,MontantPV,TpsPrevu : double);
var LocNaturePiece  : string;
    TotalMoSalPA    : Double;
    TotalMoSalPR    : Double;
    TotalMoSalPV    : Double;
    TotalMoSal      : Double;
begin

  if (Pos(NaturePiece,'FBT;B00;FBP;FBC')>0) or (NaturePiece = 'FAC') Then
     TOBTMP.PutValue('FACTURE_SALARIE', TOBTMP.GetValue('FACTURE_SALARIE')+MontantPV)
  else if (Pos(NaturePiece,'ABT;AVC;ABC')>0) or (NaturePiece = 'ABP')  then
     TOBTMP.PutValue('FACTURE_SALARIE', TOBTMP.GetValue('FACTURE_SALARIE')+MontantPV)
  else
  begin
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_MOSALPA',TOBTMP.GetValue('PREVU_'+NaturePiece+'_MOSALPA') + MontantPA);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_MOSALPR',TOBTMP.GetValue('PREVU_'+NaturePiece+'_MOSALPR') + MontantPR);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_MOSALPV',TOBTMP.GetValue('PREVU_'+NaturePiece+'_MOSALPV') + MontantPV);
     TOBTMP.PutValue('TPS_PREVU_'+NaturePiece+'_MOSAL',TOBTMP.GetValue('TPS_PREVU_'+NaturePiece+'_MOSAL') + TpsPRevu);
  end;

end;

Procedure Repartition_Eclatement(TypeMontant, NaturePres, TypeMt : String; Montant : Double; TOBTMP : Tob);
var MtAvant, MtApres : Double;
Begin

  if Not TobTMP.FieldExists(TypeMontant + NaturePres + TypeMt) then
  begin
    TOBTMP.AddChampSupValeur(TypeMontant + NaturePres + TypeMt, Montant);
    MAJ_PremiereLigne(TypeMontant + NaturePres + TypeMt, TOBTMP);
  end else
  begin
    MtAvant := TOBTMP.GetValue(TypeMontant + NaturePres + TypeMt);
    MtAvant := Arrondi(MtAvant, V_PGI.OkDecV);
    MtApres := MtAvant + Montant;
    MtApres := Arrondi(MtApres, V_PGI.OkDecV);
    TOBTMP.PutValue(TypeMontant + NaturePres  + TypeMt,  MtApres);
  end;
end;

Procedure MAJ_PremiereLigne(NomChampSup : String; TOBTMP : TOB);
Begin
  // ATTENTION, on crée le champ sup dans la première ligne de la TOB principale pour que le champ soit présent dans les champs disponibles du tobviewer
  if TOBTMP.Parent.Detail.count = 0 Then Exit;
  if Not TOBTMP.Parent.Detail[0].FieldExists(NomChampSup) then
  begin
     TOBTMP.Parent.Detail[0].AddChampSupValeur(NomChampSup, 0.0);
  end;
end;


procedure CumuleprevufactureInterimaire (TOBTMP : TOB ; NaturePiece : string; MontantPa,MontantPr,MontantPV,TpsPrevu : double);
var LocNaturePiece : string;
begin

  if (Pos(NaturePiece,'FBT;B00;FBP;FBC')>0) or (Pos(NaturePiece,'ABT;ABP;ABC')>0) or (NaturePiece = 'FAC') or (NaturePiece = 'AVC') Then
  begin
     TOBTMP.PutValue('FACTURE_INTERIM', TOBTMP.GetValue('FACTURE_INTERIM')+MontantPV);
  end
  else
  begin
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_MOINTPA',TOBTMP.GetValue('PREVU_'+NaturePiece+'_MOINTPA') + MontantPA);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_MOINTPR',TOBTMP.GetValue('PREVU_'+NaturePiece+'_MOINTPR') + MontantPR);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_MOINTPV',TOBTMP.GetValue('PREVU_'+NaturePiece+'_MOINTPV') + MontantPV);
     TOBTMP.PutValue('TPS_PREVU_'+NaturePiece+'_MOINT',TOBTMP.GetValue('TPS_PREVU_'+NaturePiece+'_MOINT') + TpsPRevu);
  end;
end;

procedure CumuleprevufactureLocation (TOBTMP : TOB ; NaturePiece : string; MontantPa,MontantPr,MontantPV : double);
var LocNaturePiece : string;
begin

  if (Pos(NaturePiece ,'FBT;B00;FBP;FBC')>0) or (Pos(NaturePiece,'ABT;ABP;ABC')>0) or (NaturePiece = 'FAC') or (NaturePiece = 'AVC') Then
  begin
     TOBTMP.PutValue('FACTURE_LOCATION', TOBTMP.GetValue('FACTURE_LOCATION')+MontantPV);
  end
  else
  begin
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_LOCPA',TOBTMP.GetValue('PREVU_'+NaturePiece+'_LOCPA') + MontantPA);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_LOCPR',TOBTMP.GetValue('PREVU_'+NaturePiece+'_LOCPR') + MontantPR);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_LOCPV',TOBTMP.GetValue('PREVU_'+NaturePiece+'_LOCPV') + MontantPV);
  end;
end;

procedure CumuleprevufactureMateriel (TOBTMP : TOB; NaturePiece : string; MontantPa,MontantPr,MontantPV : double);
var LocNaturePiece : string;
begin
  if (Pos(NaturePiece,'FBT;B00;FBP;FBC')>0) or (Pos(NaturePiece,'ABT;ABP;ABC')>0) or (NaturePiece = 'FAC') or (NaturePiece = 'AVC') Then
  begin
     TOBTMP.PutValue('FACTURE_MATERIEL', TOBTMP.GetValue('FACTURE_MATERIEL')+MontantPV);
  end else
  begin
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_MATPA',TOBTMP.GetValue('PREVU_'+NaturePiece+'_MATPA') + MontantPA);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_MATPR',TOBTMP.GetValue('PREVU_'+NaturePiece+'_MATPR') + MontantPR);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_MATPV',TOBTMP.GetValue('PREVU_'+NaturePiece+'_MATPV') + MontantPV);
  end;
end;

procedure CumuleprevufactureOutillage (TOBTMP : TOB; NaturePiece : string; MontantPa,MontantPr,MontantPV : double);
var LocNaturePiece : string;
begin
  if (Pos(NaturePiece,'FBT;B00;FBP;FBC')>0) or (Pos(NaturePiece,'ABT;ABP;ABC')>0) or (NaturePiece = 'FAC') or (NaturePiece = 'AVC') Then
  begin
     TOBTMP.PutValue('FACTURE_OUTIL', TOBTMP.GetValue('FACTURE_OUTIL')+MontantPV);
  end else
  begin
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_OUTPA',TOBTMP.GetValue('PREVU_'+NaturePiece+'_OUTPA') + MontantPA);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_OUTPR',TOBTMP.GetValue('PREVU_'+NaturePiece+'_OUTPR') + MontantPR);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_OUTPV',TOBTMP.GetValue('PREVU_'+NaturePiece+'_OUTPV') + MontantPV);
  end;
end;

procedure CumuleprevufactureSousTraitance (TOBTMP : TOB; NaturePiece : string ;MontantPa,MontantPr,MontantPV,TpsPrevu : double);
var LocNaturePiece : string;
begin
  if (Pos(NaturePiece,'FBT;B00;FBP;FBC')>0) or (Pos(NaturePiece,'ABT;ABP;ABC')>0) or (NaturePiece = 'FAC') or (NaturePiece = 'AVC') Then
  begin
     TOBTMP.PutValue('FACTURE_ST', TOBTMP.GetValue('FACTURE_ST')+MontantPV);
  end
  else
  begin
    //FV1 - 23/10/2017 - FS#2704 - AVENEL - En tableau de bord message impossible de convertir type String en type Double
    If TOBTMP.FieldExists('PREVU_'+NaturePiece+'_STPA') then
      TOBTMP.PutValue('PREVU_'+NaturePiece+'_STPA',TOBTMP.GetValue('PREVU_'+NaturePiece+'_STPA') + MontantPA)
    else
      TOBTMP.AddChampSupValeur('PREVU_'+NaturePiece+'_STPA', MontantPA);
    //
    If TOBTMP.FieldExists('PREVU_'+NaturePiece+'_STPR') then
      TOBTMP.PutValue('PREVU_'+NaturePiece+'_STPR',TOBTMP.GetValue('PREVU_'+NaturePiece+'_STPR') + MontantPR)
    else
     TOBTMP.AddChampSupValeur('PREVU_'+NaturePiece+'_STPR', MontantPR);
    //
    If TOBTMP.FieldExists('PREVU_'+NaturePiece+'_STPV') then
      TOBTMP.PutValue('PREVU_'+NaturePiece+'_STPV',TOBTMP.GetValue('PREVU_'+NaturePiece+'_STPV') + MontantPV)
    else
      TOBTMP.AddChampSupValeur('PREVU_'+NaturePiece+'_STPV', MontantPV);
    //
    If TOBTMP.FieldExists('PREVU_'+NaturePiece+'_ST') then
      TOBTMP.PutValue('TPS_PREVU_'+NaturePiece+'_ST',TOBTMP.GetValue('TPS_PREVU_'+NaturePiece+'_ST') + TpsPrevu)
    else
      TOBTMP.AddChampSupValeur('TPS_PREVU_'+NaturePiece+'_ST', TpsPrevu);
    //
  end;
end;

procedure CumuleprevufactureFourniture (TOBTMP : TOB; NaturePiece : string ; MontantPa,MontantPr,MontantPV : double);
var LocNaturePiece : string;
begin
  if (Pos(NaturePiece,'FBT;B00;FBP;FBC')>0) or (Pos(NaturePiece,'ABT;ABP;ABC')>0) or (NaturePiece = 'FAC') or (NaturePiece = 'AVC') Then
  begin
     TOBTMP.PutValue('FACTURE_FOURNITURE', TOBTMP.GetValue('FACTURE_FOURNITURE')+MontantPV);
  end else
  begin
    //FV1 - 23/10/2017 - FS#2704 - AVENEL - En tableau de bord message impossible de convertir type String en type Double
    if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_FOURNPA') then
      TOBTMP.PutValue('PREVU_'+NaturePiece+'_FOURNPA', TOBTMP.GetValue('PREVU_'+NaturePiece+'_FOURNPA') +  MontantPA)
    else
      TOBTMP.AddChampSupValeur('PREVU_'+NaturePiece+'_FOURNPA',MontantPA);
    //
    if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_FOURNPR') then
      TOBTMP.PutValue('PREVU_'+NaturePiece+'_FOURNPR', TOBTMP.GetValue('PREVU_'+NaturePiece+'_FOURNPR') + MontantPR)
    else
      TOBTMP.AddChampSupValeur('PREVU_'+NaturePiece+'_FOURNPR', MontantPR);
    //
    if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_FOURNPV') then
      TOBTMP.PutValue('PREVU_'+NaturePiece+'_FOURNPV', TOBTMP.GetValue('PREVU_'+NaturePiece+'_FOURNPV') + MontantPV)
    else
      TOBTMP.AddChampSupValeur('PREVU_'+NaturePiece+'_FOURNPV', MontantPV);
    //
  end;
end;

//FV1 : 04/02/2014 - FS#863 - SCETEC : Distinguer les frais au niveau des champs de prévisionnel
procedure CumuleprevufactureFrais (TOBTMP : TOB; NaturePiece : string ; MontantPa,MontantPr,MontantPV : double);
var LocNaturePiece : string;
begin
  if (Pos(NaturePiece,'FBT;B00;FBP;FBC')>0) or (Pos(NaturePiece,'ABT;ABP;ABC')>0) or (NaturePiece = 'FAC') or (NaturePiece = 'AVC') Then
  begin
     TOBTMP.PutValue('FACTURE_FRAIS', TOBTMP.GetValue('FACTURE_FRAIS')+MontantPV);
  end
  else
  begin
    //FV1 - 23/10/2017 - FS#2704 - AVENEL - En tableau de bord message impossible de convertir type String en type Double
    if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_FRAISPA') then
      TOBTMP.PutValue('PREVU_'+NaturePiece+'_FRAISPA',TOBTMP.GetValue('PREVU_'+NaturePiece+'_FRAISPA') + MontantPA)
    else
      TOBTMP.AddChampSupValeur('PREVU_'+NaturePiece+'_FRAISPA',  MontantPA);

    if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_FRAISPR') then
      TOBTMP.PutValue('PREVU_'+NaturePiece+'_FRAISPR',TOBTMP.GetValue('PREVU_'+NaturePiece+'_FRAISPR') + MontantPR)
    else
      TOBTMP.AddChampSupValeur('PREVU_'+NaturePiece+'_FRAISPR', MontantPR);
    //
    if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_FRAISPV') then
      TOBTMP.PutValue('PREVU_'+NaturePiece+'_FRAISPV',TOBTMP.GetValue('PREVU_'+NaturePiece+'_FRAISPV') + MontantPV)
    else
      TOBTMP.AddChampSupValeur('PREVU_'+NaturePiece+'_FRAISPV',  MontantPV);
    //
  end;
end;

procedure AjouteChampSupPrevu(TOBTmp : TOB; Nature : string);
begin

  //FV1 : 03/06/2013 - global à l'ensemble des documents du chantier
  TOBTMP.addchampsup('PREVU_'+Nature+'_PA', false); TOBTMP.PutValue('PREVU_'+Nature+'_PA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_PR', false); TOBTMP.PutValue('PREVU_'+Nature+'_PR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_PV', false); TOBTMP.PutValue('PREVU_'+Nature+'_PV', 0.0);

  {Depuis prévision}

  (* Autre *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_AUTPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_AUTPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_AUTPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_AUTPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_AUTPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_AUTPV', 0.0);

  (* fourniture *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_FOURNPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_FOURNPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_FOURNPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_FOURNPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_FOURNPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_FOURNPV', 0.0);

  //FVA : 04/02/2014 - FS#863 - SCETEC : Distinguer les frais au niveau des champs de prévisionnel
  (* frais *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_FRAISPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_FRAISPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_FRAISPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_FRAISPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_FRAISPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_FRAISPV', 0.0);

  (* Main oeuvre interne *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_MOSALPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_MOSALPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MOSALPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_MOSALPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MOSALPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_MOSALPV', 0.0);
  TOBTMP.addchampsup('TPS_PREVU_'+Nature+'_MOSAL', false); TOBTMP.PutValue('TPS_PREVU_'+Nature+'_MOSAL', 0.0);

  (* Main oeuvre interimaire *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_MOINTPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_MOINTPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MOINTPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_MOINTPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MOINTPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_MOINTPV', 0.0);
  TOBTMP.addchampsup('TPS_PREVU_'+Nature+'_MOINT', false); TOBTMP.PutValue('TPS_PREVU_'+Nature+'_MOINT', 0.0);

  (* Location *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_LOCPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_LOCPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_LOCPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_LOCPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_LOCPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_LOCPV', 0.0);

  (* Materiel *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_MATPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_MATPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MATPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_MATPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MATPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_MATPV', 0.0);

  (* Outillage *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_OUTPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_OUTPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_OUTPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_OUTPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_OUTPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_OUTPV', 0.0);

  (* Sous Traitance *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_STPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_STPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_STPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_STPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_STPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_STPV', 0.0);
  TOBTMP.addchampsup('TPS_PREVU_'+Nature+'_ST', false); TOBTMP.PutValue('TPS_PREVU_'+Nature+'_ST', 0.0);

  // Ajout BRL 7/11/2012 : Montants des frais détaillés chantier, généraux et répartis
  TOBTMP.addchampsup('PREVU_'+Nature+'_MONTANTFG', false); TOBTMP.PutValue('PREVU_'+Nature+'_MONTANTFG', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MONTANTFC', false); TOBTMP.PutValue('PREVU_'+Nature+'_MONTANTFC', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MONTANTFR', false); TOBTMP.PutValue('PREVU_'+Nature+'_MONTANTFR', 0.0);
end;

procedure AjouteChampSupRestADep(TOBTmp : TOB);
begin

  TOBTMP.addchampsup('RESTEADEP_FINAFF', false);     TOBTMP.PutValue('RESTEADEP_FINAFF', 0.0);
  TOBTMP.addchampsup('RESTEADEP_MTRESTE', false);    TOBTMP.PutValue('RESTEADEP_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTEADEP_QTRESTE', false);    TOBTMP.PutValue('RESTEADEP_QTRESTE', 0.0);
  TOBTMP.addchampsup('RESTEADEP_DATEARRETEE', false);    TOBTMP.PutValue('RESTEADEP_DATEARRETEE', iDate1900);
  //
  TOBTMP.addchampsup('RESTADEP_AUT_MTRESTE', false); TOBTMP.PutValue('RESTADEP_AUT_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_AUT_FINAFF',  false); TOBTMP.PutValue('RESTADEP_AUT_FINAFF',  0.0);
  // Fourniture deja definie
  TOBTMP.addchampsup('RESTADEP_SAL_MTRESTE', false); TOBTMP.PutValue('RESTADEP_SAL_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_SAL_QTRESTE', false); TOBTMP.PutValue('RESTADEP_SAL_QTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_SAL_FINAFF',  false); TOBTMP.PutValue('RESTADEP_SAL_FINAFF',  0.0);
  //
  TOBTMP.addchampsup('RESTADEP_LOC_MTRESTE', false); TOBTMP.PutValue('RESTADEP_LOC_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_LOC_FINAFF',  false); TOBTMP.PutValue('RESTADEP_LOC_FINAFF',  0.0);
  //
  TOBTMP.addchampsup('RESTADEP_OUT_MTRESTE', false); TOBTMP.PutValue('RESTADEP_OUT_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_OUT_FINAFF',  false); TOBTMP.PutValue('RESTADEP_OUT_FINAFF',  0.0);
  //
  TOBTMP.addchampsup('RESTADEP_ST_MTRESTE',  false); TOBTMP.PutValue('RESTADEP_ST_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_ST_FINAFF',   false); TOBTMP.PutValue('RESTADEP_ST_FINAFF',  0.0);
  //
  TOBTMP.addchampsup('RESTADEP_MAT_MTRESTE', false); TOBTMP.PutValue('RESTADEP_MAT_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_MAT_FINAFF',  false); TOBTMP.PutValue('RESTADEP_MAT_FINAFF',  0.0);
  //
  TOBTMP.addchampsup('RESTADEP_INT_MTRESTE', false); TOBTMP.PutValue('RESTADEP_INT_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_INT_QTRESTE', false); TOBTMP.PutValue('RESTADEP_INT_QTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_INT_FINAFF',  false); TOBTMP.PutValue('RESTADEP_INT_FINAFF',  0.0);
  //
  TOBTMP.addchampsup('RESTADEP_ACH_MTRESTE', false); TOBTMP.PutValue('RESTADEP_ACH_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_ACH_FINAFF',  false); TOBTMP.PutValue('RESTADEP_ACH_FINAFF',  0.0);
  TOBTMP.addchampsup('RESTADEP_STK_MTRESTE', false); TOBTMP.PutValue('RESTADEP_STK_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_STK_FINAFF',  false); TOBTMP.PutValue('RESTADEP_STK_FINAFF',  0.0);
  TOBTMP.addchampsup('RESTADEP_FOU_MTRESTE', false); TOBTMP.PutValue('RESTADEP_FOU_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_FOU_FINAFF',  false); TOBTMP.PutValue('RESTADEP_FOU_FINAFF',  0.0);
  //
  TOBTMP.addchampsup('RESTADEP_FAC_MTRESTE', false); TOBTMP.PutValue('RESTADEP_FAC_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_FAC_FINAFF',  false); TOBTMP.PutValue('RESTADEP_FAC_FINAFF',  0.0);

end;

procedure AjouteChampSupRealiseDetail (TOBTmp : TOB);
begin
  TOBTMP.addchampsup('REALISE_AUTRE_PA', false); TOBTMP.PutValue('REALISE_AUTRE_PA', 0.0);
  TOBTMP.addchampsup('REALISE_AUTRE_PR', false); TOBTMP.PutValue('REALISE_AUTRE_PR', 0.0);
  TOBTMP.addchampsup('REALISE_AUTRE_PV', false); TOBTMP.PutValue('REALISE_AUTRE_PV', 0.0);
  
  // Fourniture deja definie
  TOBTMP.addchampsup('REALISE_SAL_PA', false); TOBTMP.PutValue('REALISE_SAL_PA', 0.0);
  TOBTMP.addchampsup('REALISE_SAL_PR', false); TOBTMP.PutValue('REALISE_SAL_PR', 0.0);
  TOBTMP.addchampsup('REALISE_SAL_PV', false); TOBTMP.PutValue('REALISE_SAL_PV', 0.0);
  //
  TOBTMP.addchampsup('REALISE_LOC_PA', false); TOBTMP.PutValue('REALISE_LOC_PA', 0.0);
  TOBTMP.addchampsup('REALISE_LOC_PR', false); TOBTMP.PutValue('REALISE_LOC_PR', 0.0);
  TOBTMP.addchampsup('REALISE_LOC_PV', false); TOBTMP.PutValue('REALISE_LOC_PV', 0.0);
  //
  TOBTMP.addchampsup('REALISE_OUTIL_PA', false); TOBTMP.PutValue('REALISE_OUTIL_PA', 0.0);
  TOBTMP.addchampsup('REALISE_OUTIL_PR', false); TOBTMP.PutValue('REALISE_OUTIL_PR', 0.0);
  TOBTMP.addchampsup('REALISE_OUTIL_PV', false); TOBTMP.PutValue('REALISE_OUTIL_PV', 0.0);
  //
  TOBTMP.addchampsup('REALISE_ST_PA', false); TOBTMP.PutValue('REALISE_ST_PA', 0.0);
  TOBTMP.addchampsup('REALISE_ST_PR', false); TOBTMP.PutValue('REALISE_ST_PR', 0.0);
  TOBTMP.addchampsup('REALISE_ST_PV', false); TOBTMP.PutValue('REALISE_ST_PV', 0.0);
  //
  TOBTMP.addchampsup('REALISE_MAT_PA', false); TOBTMP.PutValue('REALISE_MAT_PA', 0.0);
  TOBTMP.addchampsup('REALISE_MAT_PR', false); TOBTMP.PutValue('REALISE_MAT_PR', 0.0);
  TOBTMP.addchampsup('REALISE_MAT_PV', false); TOBTMP.PutValue('REALISE_MAT_PV', 0.0);
  //
  TOBTMP.addchampsup('REALISE_INTERIM_PA', false); TOBTMP.PutValue('REALISE_INTERIM_PA', 0.0);
  TOBTMP.addchampsup('REALISE_INTERIM_PR', false); TOBTMP.PutValue('REALISE_INTERIM_PR', 0.0);
  TOBTMP.addchampsup('REALISE_INTERIM_PV', false); TOBTMP.PutValue('REALISE_INTERIM_PV', 0.0);
  //
end;

procedure AjouteChampSupFactureDetail (TOBTmp : TOB);
begin
  TOBTMP.addchampsup('FACTURE_FOURNITURE', false);TOBTMP.PutValue('FACTURE_FOURNITURE', 0.0);
  TOBTMP.addchampsup('FACTURE_AUTRE', false);     TOBTMP.PutValue('FACTURE_AUTRE', 0.0);
  TOBTMP.addchampsup('FACTURE_SALARIE', false);   TOBTMP.PutValue('FACTURE_SALARIE', 0.0);
  TOBTMP.addchampsup('FACTURE_LOCATION', false);  TOBTMP.PutValue('FACTURE_LOCATION', 0.0);
  TOBTMP.addchampsup('FACTURE_OUTIL', false);     TOBTMP.PutValue('FACTURE_OUTIL', 0.0);
  TOBTMP.addchampsup('FACTURE_ST', false);        TOBTMP.PutValue('FACTURE_ST', 0.0);
  TOBTMP.addchampsup('FACTURE_MATERIEL', false);  TOBTMP.PutValue('FACTURE_MATERIEL', 0.0);
  TOBTMP.addchampsup('FACTURE_INTERIM', false);   TOBTMP.PutValue('FACTURE_INTERIM', 0.0);

  //FVA : 04/02/2014 - FS#863 - SCETEC : Distinguer les frais au niveau des champs de prévisionnel
  TOBTMP.addchampsup('FACTURE_FRAIS', false);     TOBTMP.PutValue('FACTURE_FRAIS', 0.0);

end;

procedure SetPrevuAvancePBT (TOBTMP : TOB; OptionChoixPrevuAvanc : TChoixPrevuAvanc);
var req         : string;
		QQ          : TQuery;
    TypeRes     : String;
    NaturePres  : String;
    TypeArticle : String;
    MontantPA, MontantPr,MontantPV, MontantFG, MontantFC, MontantFR, AvancePa,AvancePr,AvancePv,TpsPrevu,TpsAvance : double;
begin

  // Récupération du prévu
  Req := 'SELECT BNP_TYPERESSOURCE,BNP_NATUREPRES, GL_TYPEARTICLE, '+
         'SUM(GL_QTEFACT*GL_DPA) AS ACHAT, '+
  			 'SUM(GL_QTEFACT*GL_DPR) AS REVIENT, '+
         'SUM(GL_TOTALHTDEV) AS VENTE, ' +
         'SUM(GL_MONTANTFG) AS MONTANTFG, ' +
         'SUM(GL_MONTANTFC) AS MONTANTFC, ' +
         'SUM(GL_MONTANTFR) AS MONTANTFR, ' +
         'SUM(GL_QTEFACT) AS TPS_PREVU';
  // .. et de l'avance
  Req := Req + ',SUM(GL_QTEPREVAVANC*GL_DPA) AS AVANCEPA'+
  			   		 ',SUM(GL_QTEPREVAVANC*GL_DPR) AS AVANCEPR'+
               ',SUM(GL_TOTALHTDEV*(GL_POURCENTAVANC/100.0)) AS AVANCEPV'+
               ',SUM(GL_QTEPREVAVANC) AS TPS_AVANCE' ;
  Req := Req + ' FROM LIGNE ' +
  						 'LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE '+
               'LEFT JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES '+
               'WHERE GL_NATUREPIECEG = "PBT" AND GL_TYPELIGNE LIKE "AR%" AND GL_AFFAIRE="' + TOBTMP.GetValue('BCO_AFFAIRE') + '"' +
               'GROUP BY GL_AFFAIRE,GL_TYPEARTICLE,BNP_TYPERESSOURCE,BNP_NATUREPRES';
  QQ := OpenSql (req,True);

  while not QQ.eof do
  begin
    TypeRes   := QQ.findfield('BNP_TYPERESSOURCE').AsString;
    NaturePres:= QQ.findfield('BNP_NATUREPRES').AsString;
    //
    //FV1 : 04/02/2014 - FS#863 - SCETEC : Distinguer les frais au niveau des champs de prévisionnel
    TypeArticle := QQ.findfield('GL_TYPEARTICLE').AsString;
    If TypeArticle <> 'FRA' then
    begin
      //FV1 : Si pas de nature de prestation alors sur Nature 'Fournitures'
      if TypeRes = '' then TypeRes := 'FOU';
      if NaturePres = '' then NaturePres := 'FOURNITURES';
    end
    else
    begin
      //FV1 : pas de nature de prestation alors sur Nature 'Frais'
      if TypeRes = ''    then TypeRes := 'FRA';
      if NaturePres = '' then NaturePres := 'FRAIS';
    end;
    //
    MontantPA := QQ.findfield('ACHAT').AsFloat;
    MontantPR := QQ.findfield('REVIENT').AsFloat;
    MontantPV := QQ.findfield('VENTE').AsFloat;
    MontantFG := QQ.findfield('MONTANTFG').AsFloat;
    MontantFC := QQ.findfield('MONTANTFC').AsFloat;
    MontantFR := QQ.findfield('MONTANTFR').AsFloat;
    //
    AvancePA  := QQ.findfield('AVANCEPA').AsFloat;
    AvancePR  := QQ.findfield('AVANCEPR').AsFloat;
    AvancePV  := QQ.findfield('AVANCEPV').AsFloat;
    //
    TpsPrevu  := QQ.findfield('TPS_PREVU').AsFloat;
    TpsAvance := QQ.findfield('TPS_AVANCE').AsFloat;

    TOBTMP.PutValue('PREVUPA', TOBTMP.GetValue('PREVUPA') + MontantPA);
    TOBTMP.PutValue('PREVUPR', TOBTMP.GetValue('PREVUPR') + MontantPR);
    // BRL 13/08 : ajout cumul montant Frais pour PBT
    TOBTMP.PutValue('PREVU_MONTANTFG', TOBTMP.GetValue('PREVU_MONTANTFG') + MontantFG);
    TOBTMP.PutValue('PREVU_MONTANTFC', TOBTMP.GetValue('PREVU_MONTANTFC') + MontantFC);
    TOBTMP.PutValue('PREVU_MONTANTFR', TOBTMP.GetValue('PREVU_MONTANTFR') + MontantFR);
    TOBTMP.PutValue('PREVUPV', TOBTMP.GetValue('PREVUPV') + MontantPV);

    TOBTMP.PutValue('AVANCEPA', TOBTMP.GetValue('AVANCEPA') + AvancePa);
    TOBTMP.PutValue('AVANCEPR', TOBTMP.GetValue('AVANCEPR') + AvancePr);
    TOBTMP.PutValue('AVANCEPV', TOBTMP.GetValue('AVANCEPV') + AvancePV);

    if TypeRes = 'SAL' then
    begin
      TOBTMP.PutValue('TPS_PREVU', TOBTMP.GetValue('TPS_PREVU')+ TpsPrevu);
      TOBTMP.PutValue('TPS_AVANCE', TOBTMP.GetValue('TPS_AVANCE') + TpsAvance);
    end;

    if OptDetailPBT in OptionChoixPrevuAvanc then
    begin
      // BRL 13/08 : Mise à jour montants de frais par nature et prévus totaux pour PBT
      TOBTMP.PutValue('PREVU_PBT_MONTANTFG', TOBTMP.GetValue('PREVU_MONTANTFG'));
      TOBTMP.PutValue('PREVU_PBT_MONTANTFC', TOBTMP.GetValue('PREVU_MONTANTFC'));
      TOBTMP.PutValue('PREVU_PBT_MONTANTFR', TOBTMP.GetValue('PREVU_MONTANTFR'));
      TOBTMP.PutValue('PREVU_PBT_PA', TOBTMP.GetValue('PREVUPA'));
      TOBTMP.PutValue('PREVU_PBT_PR', TOBTMP.GetValue('PREVUPR'));
      TOBTMP.PutValue('PREVU_PBT_PV', TOBTMP.GetValue('PREVUPV'));
      //
      ChargeCumulPrevuFacture(TypeRes, 'PBT', TOBTMP,MontantPa,MontantPr,MontantPV,TpsPrevu);
    end;
    //
    if (TypeRes <> 'FOU') and (TOBTMP.GetValue('CEclatNatPrest') = 'X') then
    begin
      Charge_Repartition_Eclatement('PBT_'+ TypeRes + '_' + NaturePres,MontantPA, MontantPR, MontantPV, TOBTMP)
    end;


    QQ.next;
  end;
  Ferme(QQ);
end;

//Mise à jour des prevu par type de ressource ou Nature de Prestation.
Procedure  Charge_Repartition_Eclatement(NaturePres : String; MontantPA, MontantPR, MontantPV : Double; TOBTMP : TOB);
Begin

    Repartition_Eclatement('PREVU_',  NaturePres, '_PA', MontantPA, TOBTMP);
    Repartition_Eclatement('PREVU_',  NaturePres, '_PR', MontantPR, TOBTMP);
    Repartition_Eclatement('PREVU_',  NaturePres, '_PV', MontantPV, TOBTMP);
end;

Procedure ChargeCumulPrevuFacture(TypeRes, NaturePiece : String; TOBTMP : TOB; MontantPa,MontantPr,MontantPV,TpsPrevu : Double);
begin

  if TypeRes = 'SAL' then           // prevu salarie
  begin
    //Modif FV : Dev. prioritaire DSL le 05/06/2012
    if (CoefFG_Param <> 0) then
    begin
      MontantPA := Arrondi(TauxHoraire * TpsPrevu,V_PGI.OkDecV);
      MontantPR := Arrondi(MontantPA * CoefFG_Param, V_PGI.OkDecV);
    end;
    CumuleprevufactureSalarie (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end
  else if TypeRes = 'AUT' then      // prevu autre
    CumuleprevufactureAutre (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV)
  else if TypeRes = 'INT' then     // prevu interimaire
    CumuleprevufactureInterimaire (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu)
  else
  if TypeRes = 'LOC' then           // prevu location
    CumuleprevufactureLocation (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV)
  else
  if TypeRes = 'MAT' then           // prevu materiel
    CumuleprevufactureMateriel (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV)
  else if TypeRes = 'OUT' then      // prevu outillage
    CumuleprevufactureOutillage (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV)
  else if TypeRes = 'ST' then       // prevu sous traitance
    CumuleprevufactureSousTraitance (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu)
  else if TypeRes = 'FOU' then       // prevu fourniture
    CumuleprevufactureFourniture (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV)
  else
    CumuleprevufactureFrais (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV);
end;

procedure ChargelesLignesTB (TOBLIgne,TOBPiece : TOB);
var Req : string;
		QQ : TQuery;
begin
	Req := 'SELECT BNP_TYPERESSOURCE,BNP_NATUREPRES,GL_NATUREPIECEG,GL_SOUCHE, GL_NUMERO,'   +
         'GL_INDICEG,GL_NUMLIGNE,GL_TYPEARTICLE,GL_TYPENOMENC,GL_ARTICLE,'  +
         'GL_NUMORDRE,GL_INDICENOMEN,GL_QTEFACT,GL_DPA,GL_DPR,' +
         'GL_MONTANTFC, GL_MONTANTFG, GL_MONTANTFR, GL_TOTALHTDEV,GL_PUHTDEV, GLC_NATURETRAVAIL FROM LIGNE ' +
  			 'LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE ' +
         'LEFT JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES ' +
         'LEFT JOIN LIGNECOMPL ON GLC_NATUREPIECEG=GL_NATUREPIECEG AND ' +
         'GLC_SOUCHE=GL_SOUCHE AND GLC_NUMERO=GL_NUMERO AND ' +
         'GLC_INDICEG=GL_INDICEG AND GLC_NUMORDRE=GL_NUMORDRE ' +
         'WHERE '+
         'GL_TYPELIGNE="ART" ' + // on ne prend que les lignes de type article (les commentaires...pfff)
         'AND GL_NATUREPIECEG = "'+TOBPiece.getValue('GP_NATUREPIECEG')+ '" '+
         'AND GL_SOUCHE="' + TOBPiece.GetValue('GP_SOUCHE') + '" ' +
         'AND GL_NUMERO=' + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ' '+
         'AND GL_INDICEG=' + IntToStr(TOBPiece.GetValue('GP_INDICEG')) + ' '+
         'ORDER BY GL_NUMLIGNE';
  QQ := OpenSql (req,true);
  TOBLigne.loadDetailDb ('LIGNE','','',QQ,false);
  ferme (QQ);
end;

procedure ConstitueOuvragesTB (Prefixe : String; TOBPiece,TOBOuvrage,TOBLocOuvrage : TOB);
var Lig : integer;
		indice : integer;
    TOBLig ,TOBNOuv, TOBPere, TOBnewDet,TOBL : TOB;
    LigneN1,LigneN2,LigneN3,LigneN4,LigneN5 : integer;
    Table : String;
begin
	// Initialisation
	Lig := 0;
  TOBNouv := nil;

  if Prefixe = 'BLO' then
    Table := 'LIGNEOUV'
  else
    Table := 'LIGNEOUVPLAT';
  //
	for Indice := 0 TO TOBLocOuvrage.detail.count -1 do
  begin
    TOBLig := TOBLocOuvrage.detail[Indice];
    if TOBLIg.getValue(Prefixe + '_NUMLIGNE') <> Lig then
    begin
    	// rupture sur N° de ligne --> donc nouvel ouvrage
      TOBNOuv := TOB.create ('NEW OUV',TOBOuvrage,-1);
      Lig := TOBLIg.getValue(Prefixe + '_NUMLIGNE');
      TOBL := TOBPiece.findFirst(['GL_NUMLIGNE'],[Lig],true);
      if TOBL<> nil then TOBL.putValue('GL_INDICENOMEN',TOBOuvrage.detail.count);
    end;
    LigneN1 := TOBLig.GetValue(Prefixe + '_N1');
    LigneN2 := TOBLig.GetValue(Prefixe + '_N2');
    LigneN3 := TOBLig.GetValue(Prefixe + '_N3');
    LigneN4 := TOBLig.GetValue(Prefixe + '_N4');
    LigneN5 := TOBLig.GetValue(Prefixe + '_N5');

    if LigneN5 > 0 then
    begin
			TOBPere:=TOBNOuv.FindFirst([Prefixe + '_NUMLIGNE',Prefixe + '_N1',Prefixe + '_N2',Prefixe + '_N3',Prefixe + '_N4',Prefixe + '_N5'],[Lig,LigneN1,LigneN2,LigneN3,LigneN4,0],True) ;
    end else
    if LigneN4 > 0 then
    begin
    	TOBPere:=TOBNOuv.FindFirst([Prefixe + '_NUMLIGNE',Prefixe + '_N1',Prefixe + '_N2',Prefixe + '_N3',Prefixe + '_N4',Prefixe + '_N5'],[Lig,LigneN1,LigneN2,LigneN3,0,0],True) ;
    end else
    if LigneN3 > 0 then
    begin
    	TOBPere:=TOBNOuv.FindFirst([Prefixe + '_NUMLIGNE',Prefixe + '_N1',Prefixe + '_N2',Prefixe + '_N3',Prefixe + '_N4',Prefixe + '_N5'],[Lig,LigneN1,LigneN2,0,0,0],True) ;
    end else
    if LigneN2 > 0 then
    begin
      TOBPere:=TOBNOuv.FindFirst([Prefixe + '_NUMLIGNE',Prefixe + '_N1',Prefixe + '_N2',Prefixe + '_N3',Prefixe + '_N4',Prefixe + '_N5'],[Lig,LigneN1,0,0,0,0],True) ;
    end else
    begin
    	TOBPere:=TOBNOuv;
    end;

    if TOBPere<>Nil then
    BEGIN
       TOBNewDet:=TOB.Create(Table,TOBPere,-1) ;
       TOBNewDet.Dupliquer(TOBLig,False,True) ;
    END;
  end;
end;

procedure ChargelesOuvragesTB (TOBOuvrage,TOBPiece : TOB);
var Req : string;
		QQ : TQuery;
    TOBLocOuvrage : TOB;
begin
	TOBLocOuvrage := TOB.Create ('OUV LU',nil,-1);
	Req := 'SELECT BNP_TYPERESSOURCE,BNP_NATUREPRES, BLO_TYPEARTICLE,BLO_NATUREPIECEG, BLO_SOUCHE,BLO_ARTICLE,BLO_QTEFACT,BLO_DPA,BLO_REMISEPIED,'+
  			 'BLO_DPR,BLO_PUHTDEV,BLO_QTEDUDETAIL,BLO_NUMLIGNE,BLO_N1,BLO_N2,BLO_N3,BLO_N4,BLO_N5, BLO_NATURETRAVAIL, ' +
         'BLO_MONTANTFC, BLO_MONTANTFG, BLO_MONTANTFR, BLO_NUMERO FROM LIGNEOUV '+
  			 'LEFT JOIN ARTICLE     ON GA_ARTICLE=BLO_ARTICLE '+
         'LEFT JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES '+
         'WHERE '+
         'AND BLO_NATUREPIECEG = "'+TOBPiece.getValue('GP_NATUREPIECEG')+'" '+
         'AND BLO_SOUCHE="' + TOBPiece.GetValue('GP_SOUCHE') + '" ' +
         'AND BLO_NUMERO=' + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ' '+
         'AND BLO_INDICEG=' + IntToStr(TOBPiece.GetValue('GP_INDICEG')) + ' '+
         'ORDER BY BLO_NUMLIGNE,BLO_N1,BLO_N2,BLO_N3,BLO_N4,BLO_N5';
  QQ := OpenSql (req,true);
  TOBLocOuvrage.loadDetailDb ('LIGNEOUV','','',QQ,false);
  ferme (QQ);
  ConstitueOuvragesTB ('BLO', TOBPiece,TOBOuvrage,TOBLocOuvrage);
  TOBLocOuvrage.free;
end;

procedure ChargelesOuvragesTBPlat (TOBOuvragePlat,TOBPiece : TOB);
var Req : string;
		QQ : TQuery;
begin

	Req := 'SELECT BNP_TYPERESSOURCE,BNP_NATUREPRES, BOP_TYPEARTICLE,BOP_NATUREPIECEG, BOP_SOUCHE,BOP_ARTICLE,BOP_QTEFACT,BOP_DPA,BOP_REMISEPIED,'+
  			 'BOP_DPR,BOP_PUHTDEV,BOP_TOTALHTDEV,BOP_MONTANTHTDEV, BOP_NUMLIGNE,BOP_NUMORDRE,BOP_N1,BOP_N2,BOP_N3,BOP_N4,BOP_N5, BOP_NATURETRAVAIL, ' +
         'BOP_MONTANTFC, BOP_MONTANTFG, BOP_MONTANTFR, BOP_NUMERO FROM LIGNEOUVPLAT '+
  			 'LEFT JOIN ARTICLE     ON GA_ARTICLE=BOP_ARTICLE '+
         'LEFT JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES '+
         'WHERE '+
         'AND BOP_NATUREPIECEG = "'+TOBPiece.getValue('GP_NATUREPIECEG')+'" '+
         'AND BOP_SOUCHE="' + TOBPiece.GetValue('GP_SOUCHE') + '" ' +
         'AND BOP_NUMERO=' + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ' '+
         'AND BOP_INDICEG=' + IntToStr(TOBPiece.GetValue('GP_INDICEG')) + ' '+
         'ORDER BY BOP_NUMORDRE';

  QQ := OpenSql (req,true);

  TOBOuvragePlat.loadDetailDb ('LIGNEOUVPLAT','','',QQ,false);

  ferme (QQ);

end;

procedure TraiteLigneTB (TOBTMP,TOBL : TOB);
begin

  //si intervenant sur ligne article pas prise en compte
  if TOBL.GetString('GLC_NATURETRAVAIL') = '001' then exit;

  if (TobL.getString('GLC_NATURETRAVAIL') = '002') Then
  begin
    TOBL.PutValue('BNP_TYPERESSOURCE', 'ST');
    TOBL.PutValue('LIBNATURE', rechDom('AFTTYPERESSOURCE','ST',false));
  end;

  ChargePrevu(TOBL, TOBTMP, 'GL');

end;

Procedure ChargePrevu(TOBL, TOBTMP : TOB; Prefixe : String; Tps_Prevu : Double);
var MontantPa : Double;
    MontantPr : Double;
    MontantPv : Double;
    //
		NaturePiece : string;
    //
    TypeRes   : String;
    NaturePres: String;
    //
    MontantFG : Double;
    MontantFC : Double;
    MontantFR : Double;
    //
    TotalMtPA : Double;
    TotalMtPR : Double;
    TotalMtPV : Double;
    //
    TotalMtFG : Double;
    TotalMtFC : Double;
    TotalMtFR : Double;
    //
    DPA       : Double;
    DPR       : Double;
    PUHT      : Double;
    //
    QteduDetail: Double;
    //
    NumeroPiece: Integer;
    //
    TypeArticle: String;
    //
    LigneLog   : String;
    NoLig      : Integer;
Begin

  LigneLog := '';

	NaturePiece := TOBL.GetValue(Prefixe + '_NATUREPIECEG');
  if pos(NaturePiece,'AFF;DAP')>0  Then NaturePiece := 'PBT';
  //
  TypeRes     := TOBL.GetValue('BNP_TYPERESSOURCE');
  NaturePres  := TOBL.GetValue('BNP_NATUREPRES');

  //FV1 : 04/02/2014 - FS#863 - SCETEC : Distinguer les frais au niveau des champs de prévisionnel
  TypeArticle := TOBL.GetValue(Prefixe + '_TYPEARTICLE');
  If TypeArticle <> 'FRA' then
  begin
    //FV1 : Si pas de nature de prestation alors sur Nature 'Fournitures'
    if TypeRes = '' then TypeRes := 'FOU';
    if NaturePres = '' then NaturePres := 'FOURNITURES';
  end
  else
  begin
    //FV1 : pas de nature de prestation alors sur Nature 'Frais'
    if TypeRes = ''    then TypeRes := 'FRA';
    if NaturePres = '' then NaturePres := 'FRAIS';
  end;
  //
  NumeroPiece := StrToInt(TOBL.GetValue(Prefixe + '_NUMERO'));
  //
  DPA         := TOBL.GetValue(Prefixe + '_DPA');
  DPR         := TOBL.GetValue(Prefixe + '_DPR');
  PUHT        := TOBL.GetValue(Prefixe + '_PUHTDEV');
  //

  if (Tps_Prevu = 0) and (prefixe <> 'BLO') then Tps_Prevu := TOBL.GetValue(Prefixe + '_QTEFACT');
  //
  MontantFG   := TOBL.GetValue(Prefixe + '_MONTANTFG');
  MontantFC   := TOBL.GetValue(Prefixe + '_MONTANTFC');
  MontantFR   := TOBL.GetValue(Prefixe + '_MONTANTFR');
  //
  if prefixe = 'BLO' then
  begin
    //
    MontantPA := (Tps_Prevu*TOBL.GetDouble('BLO_DPA')) ;
    MontantPR := (Tps_Prevu*TOBL.GetDouble('BLO_DPR'));
    MontantPV := (Tps_Prevu*TOBL.GetDouble('BLO_PUHTDEV'));
    //
    if TOBL.GetValue('BLO_REMISEPIED') then
    begin
    	MontantPV := MontantPV - (MontantPV * (TOBL.GetDouble('BLO_REMISEPIED')/100.0));
    end;
    //
    if TOBL.GetValue('BLO_QTEFACT') <> 0 then
      MontantFG := (Tps_Prevu/TOBL.GetDouble('BLO_QTEFACT')) * MontantFG
    else
      MontantFG := 0.0;
    if TOBL.GetValue('BLO_QTEFACT') <> 0 then
      MontantFC := (Tps_Prevu/TOBL.GetDouble('BLO_QTEFACT')) * MontantFC
    else
      MontantFC := 0.0;
    if TOBL.GetValue('BLO_QTEFACT') <> 0 then
      MontantFR := (Tps_Prevu/TOBL.GetDouble('BLO_QTEFACT')) * MontantFR
    else
      MontantFR := 0.0;
    //
  end
  else if prefixe = 'BOP' then
  begin
    MontantPA   := TOBL.GetValue(Prefixe + '_QTEFACT')* DPA;
    MontantPR   := TOBL.GetValue(Prefixe + '_QTEFACT')* DPR;
    MontantPV   := TOBL.GetDouble(Prefixe + '_TOTALHTDEV');
  end
  else
  begin
    MontantPA   := TOBL.GetValue(Prefixe + '_QTEFACT')* DPA;
    MontantPR   := TOBL.GetValue(Prefixe + '_QTEFACT')* DPR;
    MontantPV   := TOBL.GetValue(Prefixe + '_TOTALHTDEV');
    //
{ BRL 9/08 : pas d'avancé par nature de pièce autre que PBT
    AvancePa    := TOBL.GetValue(Prefixe + '_QTEPREVAVANC') * DPA;
    AvancePr    := TOBL.GetValue(Prefixe + '_QTEPREVAVANC') * DPR;
    AvancePV    := TOBL.GetValue(Prefixe + '_TOTALHTDEV')   * (TOBL.GetValue(Prefixe+ '_POURCENTAVANC')/100);
}
  end;


  if TOBL.GetValue(Prefixe + '_TYPEARTICLE')='POU' then
  begin
    // Article de type pourcentage
    MontantPA := 0 ;
    MontantPR := 0;
    MontantPV := (Tps_Prevu * PUHT/100);
  end;

  //Ajout FV1 03/06/2013 : Montant prévu dispatch par nature de pièce
  TotalMtPA := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_PA');
  TotalMtPA := TotalMTPA + MontantPA;

  TotalMtPR := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_PR');
  TotalMtPR := TotalMTPR + MontantPR;

  TotalMtPV := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_PV');
  TotalMtPV := TotalMTPV + MontantPV;

// BRL 13/08 : ces 3 montants ne concernent que la prévision de chantier PBT
  if NaturePiece = 'PBT' then
  begin
    //Ajout FV1 03/06/2013 : Montant prévu dispatch par nature de pièce
    TotalMtFG := TOBTMP.GetValue('PREVU_MONTANTFG');
    TotalMtFG := TotalMTFG + MontantFG;

    TotalMtFC := TOBTMP.GetValue('PREVU_MONTANTFC');
    TotalMtFC := TotalMTFC + MontantFC;

    TotalMtFR := TOBTMP.GetValue('PREVU_MONTANTFR');
    TotalMtFR := TotalMTFR + MontantFR;

    // Ajout BRL 7/11/2012 : Montants des frais détaillés chantier, généraux et répartis
    if TOBTMP.FieldExists('PREVU_MONTANTFG') then TOBTMP.PutValue('PREVU_MONTANTFG', TotalMtFG);
    if TOBTMP.FieldExists('PREVU_MONTANTFC') then TOBTMP.PutValue('PREVU_MONTANTFC', TotalMtFC);
    if TOBTMP.FieldExists('PREVU_MONTANTFR') then TOBTMP.PutValue('PREVU_MONTANTFR', TotalMtFR);
  end;

  if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_MONTANTFG') then
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_MONTANTFG', TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MONTANTFG')+MontantFG);
  if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_MONTANTFC') then
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_MONTANTFC', TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MONTANTFC')+MontantFC);
  if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_MONTANTFR') then
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_MONTANTFR', TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MONTANTFR')+MontantFR);

  TOBTMP.AddChampSupValeur('NUMEROPIECE', NumeroPiece);

  ChargeCumulPrevuFacture(TypeRes, NaturePiece, TOBTMP,MontantPa,MontantPr,MontantPV,Tps_Prevu);

  if not ((Pos(NaturePiece,'FBT;B00;FBP;FBC')>0) or (Pos(NaturePiece,'ABT;ABP;ABC')>0) or (NaturePiece = 'FAC') or (NaturePiece = 'AVC')) Then
  begin
    if (TypeRes <> 'FOU') and (TOBTMP.GetValue('CEclatNatPrest') = 'X') then
    begin
      Charge_Repartition_Eclatement(NaturePiece + '_' + TypeRes + '_' + NaturePres,MontantPA, MontantPR, MontantPV, TOBTMP)
    end;
  end;
end;

procedure TraiteLigneDetailOuvrageTB (Prefixe : String; TOBTMP,TOBL : TOB;Qte,QteDuDetail : double);
var TpsPrevu : double;
begin
  TpsPrevu := 0;

  If prefixe = 'BLO' then TpsPrevu := Qte/QteDudetail;

  ChargePrevu(TOBL, TOBTMP, Prefixe, TpsPrevu);
end;

procedure TraiteDetailOuvrageTBPlat (TOBTMP,TOBOUV : TOB);
begin
    if (TOBOUV.getString('BOP_NATURETRAVAIL') = '002') Then
    begin
      TOBOUV.PutValue('BNP_TYPERESSOURCE','ST');
      TOBOUV.PutValue('LIBNATURE', rechDom('AFTTYPERESSOURCE','ST',false));
    end;

    TraiteLigneDetailOuvrageTB ('BOP',TOBTMP,TOBOUV,0,0);
end;

procedure TraiteOuvrageTB (TOBTMP,TOBL,TOBOuvrage : TOB);
var IndiceOuv : integer;
		TOBOuv : TOB;
begin
  IndiceOuv := TOBL.GetValue('GL_INDICENOMEN');
  if IndiceOuv =0 then exit;
  if IndiceOuv > TOBOuvrage.detail.count then exit;
  TOBOuv := TOBOuvrage.detail[IndiceOuv-1];
  if TOBOUv = nil then exit;
  TraiteDetailOuvrageTB (TOBTMP,TOBOUV,TOBL.GetDouble('GL_QTEFACT'),1);
end;

procedure TraiteDetailOuvrageTB (TOBTMP,TOBOUV : TOB; Qte,QteDuDetail : double);
var QteSui,QteDuDetailSui : double;
		Indice : integer;
    TOBDet : TOB;
begin

  for indice := 0 to TOBOUV.detail.count -1 do
  begin
  	TOBDet := TOBOUV.detail[Indice];
    if (Tobdet.getString('BLO_NATURETRAVAIL') = '001') Then continue;

    if (Tobdet.getString('BLO_NATURETRAVAIL') = '002') Then
    begin
      TOBDet.PutValue('BNP_TYPERESSOURCE','ST');
      TOBDet.PutValue('LIBNATURE', rechDom('AFTTYPERESSOURCE','ST',false));
    end;

    QteSui := Qte * TOBDet.GetDouble('BLO_QTEFACT');
    if TOBDet.GetDouble('BLO_QTEDUDETAIL') <> 0 then
    begin
       QteDuDetailSui := QteDudetail * TOBDet.GetDouble('BLO_QTEDUDETAIL');
    end
    else
    begin
       QteDuDetailSui := QteDudetail;
    end;

    if TOBDet.detail.count > 0 then
    begin
      TraiteDetailOuvrageTB (TOBTMP,TOBDet,QteSui,QteDuDetailSui);
    end else
    begin
    	TraiteLigneDetailOuvrageTB ('BLO',TOBTMP,TOBdet,QteSui,QteDuDetailSui);
    end;
  end;

end;

procedure TraiteOuvrageTBPlat (TOBTMP,TOBL,TOBOuvragePlat : TOB);
var TOBTT : TOB;
    ii,depart : integer;
begin

  TOBTT := TOBOuvragePlat.findfirst(['BOP_NUMORDRE'],[TOBL.GetInteger('GL_NUMORDRE')],false);

  If TOBTT = nil then exit;
  depart := TOBTT.getIndex;
  for II := depart to TOBOuvragePlat.detail.count -1 do
  begin
    TOBTT := TOBOuvragePlat.detail[II];
    if TOBTT.GetValue('BOP_NUMORDRE')<> TOBL.GetInteger('GL_NUMORDRE') then break;
    TraiteDetailOuvrageTBPLAT (TOBTMP,TOBTT);
  end;
end;



procedure TraiteLaPieceTB (TOBTMP,TOBPiece,TOBOuvrage : TOB);
var Indice : integer;
		TOBL : TOB;
begin

	for indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    if (Pos (TOBL.GetValue('GL_TYPEARTICLE'),'OUV;OU1;ARP') > 0) and (TOBL.GetValue('GL_INDICENOMEN')>0) then
    begin
      If (TOBL.GetString('GL_NATUREPIECEG') = 'FBT') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'B00') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'FBP') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'FBC') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'ABC') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'ABT') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'ABP') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'FAC') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'AVC') then
        TraiteOuvrageTBPlat(TOBTMP,TOBL,TOBOuvrage)
      else
       	TraiteOuvrageTB (TOBTMP,TOBL,TOBOuvrage);
    end else
    begin
    	TraiteLigneTB (TOBTMP,TOBL);
    end;
  end;

end;

procedure DefiniPrevuDetail (TOBTMP,TOBPiece: TOB ; NaturePiece : string; OptionChoixPrevuAvanc: TChoixPrevuAvanc);
var TOBLignes       : TOB;
    TOBOuvrages     : TOB;
begin
	TOBLignes := TOB.Create ('LA PIECE',nil,-1);
  TOBLignes.Dupliquer (TOBPiece,false,true);
  TOBOuvrages := TOB.Create ('LES OUVRAGES',nil,-1);

  ChargelesLignesTB (TOBLIgnes,TOBPiece);
  //
  If  (NaturePiece = 'FBT') OR
      (NaturePiece = 'B00') OR
      (NaturePiece = 'FBC') OR
      (NaturePiece = 'ABC') OR
      (NaturePiece = 'FBP') OR
      (NaturePiece = 'ABT') OR
      (NaturePiece = 'ABP') OR
      (NaturePiece = 'FAC') OR
      (NaturePiece = 'AVC') then
    ChargelesOuvragesTBPlat (TOBOuvrages,TOBLignes)
  else
    ChargelesOuvragesTB (TOBOuvrages,TOBLignes);
  //
  TraiteLaPieceTB (TOBTMP,TOBLignes,TOBOuvrages);

  TOBLIgnes.free;
  TOBOuvrages.free;
end;
                                                    
procedure SetPrevuAvanceADetaille (TOBTMP : TOB; NaturePiece : string; OptionChoixPrevuAvanc: TChoixPrevuAvanc; WherePiece : String);
var Req : String;
    TOBPieces   : TOB;
    QQ : TQuery;
    Indice : integer;
    Critere : string;
    RefAffaire : string;
begin

	TOBPieces := TOB.Create ('LES PIECES',nil,-1);

  RefAffaire := 'GP_AFFAIREDEVIS';

  if Naturepiece = 'DBT' then
    Critere := ' AND AFF_ETATAFFAIRE IN ("ACP","TER") '
  else if Naturepiece = 'AFF' then
  BEGIN
    //Critere := ' AND AFF_ETATAFFAIRE IN ("ENC","TER") ';
    //FV1 - 25/06/2015 - FS#1408 - DELABOUDINIERE : en analyse chantier, le prévu contrat n'est pas renseigné.
    Critere := ' AND AFF_ETATAFFAIRE IN ("ACP","ENC","TER") ';
    RefAffaire:='GP_AFFAIRE';
  END
  else if Naturepiece = 'DAP' then
  BEGIN
    Critere := ' AND AFF_ETATAFFAIRE IN ("ACP","TER") ';
  END
  else if NaturePiece = 'ETU' then Critere := ' AND AFF_ETATAFFAIRE="ACP"';

  if WherePiece = '' then
  begin
    if NaturePiece = 'FBT' then
    begin
  	  WherePiece := ' ((GP_NATUREPIECEG="'+NaturePiece+'") OR (GP_NATUREPIECEG="B00") OR ((GP_NATUREPIECEG="FBP") AND (GP_VIVANTE="X"))) '+
      'AND GP_AFFAIRE="'+TOBTMP.GetValue('BCO_AFFAIRE') + '"';
    end else
    begin
  	WherePiece := ' GP_NATUREPIECEG="'+NaturePiece+'" '+'AND GP_AFFAIRE="'+TOBTMP.GetValue('BCO_AFFAIRE') + '"';
    end;
  end;

	Req := 'SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_INDICEG, GP_MONTANTPA, GP_MONTANTPR, GP_TOTALHTDEV FROM PIECE '+
  			 'LEFT JOIN AFFAIRE ON AFF_AFFAIRE='+RefAffaire+' WHERE '+
  			 WherePiece +
         Critere;

  QQ := OPenSql (req,true);
  TOBPieces.LoadDetailDB ('PIECE','','',QQ,false);
  Ferme (QQ);

  for indice := 0 to TOBPieces.detail.count -1 do
  begin
    DefiniPrevuDetail (TOBTMP,TOBPIeces.detail[Indice],NaturePiece,OptionChoixPrevuAvanc);

    if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_PA') then TOBTMP.PutValue('PREVU_' + NaturePiece +'_PA', TOBTMP.GetValue('PREVU_' + NaturePiece +'_PA') + TOBPIeces.detail[Indice].GetValue('GP_MONTANTPA'));
    if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_PR') then TOBTMP.PutValue('PREVU_' + NaturePiece +'_PR', TOBTMP.GetValue('PREVU_' + NaturePiece +'_PR') + TOBPIeces.detail[Indice].GetValue('GP_MONTANTPR'));
    if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_PV') then TOBTMP.PutValue('PREVU_' + NaturePiece +'_PV', TOBTMP.GetValue('PREVU_' + NaturePiece +'_PV') + TOBPIeces.detail[Indice].GetValue('GP_TOTALHTDEV'));

  end;

  TOBPIeces.free;
end;

procedure SetprevuAvance (TOBTMP : TOB; NaturePiece : string ; OptionChoixPrevuAvanc : TChoixPrevuAvanc; WherePiece : string='');
begin

  //Modif FV : Dev. prioritaire DSL le 05/06/2012
  CoefFG_param  := GetParamSocSecur('SO_COEFFG', 0);
  Tauxhoraire   := GetParamSocSecur('SO_TAUXHORAIRE', 0);

	if NaturePiece = 'PBT' then
    SetPrevuAvancePBT (TOBTMP,OptionChoixPrevuAvanc)
  else
    SetPrevuAvanceADetaille (TOBTMP,NaturePiece,OptionChoixPrevuAvanc,WherePiece);

end;

end.
