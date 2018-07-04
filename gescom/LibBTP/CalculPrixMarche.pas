unit CalculPrixMarche;

interface
uses Classes,HEnt1,Ent1,EntGC,UTob, AGLInit, UtilPGI,
{$IFDEF EAGLCLIENT}
     UtileAGL,MaineAGL,
{$ELSE}
     DBTables,EdtREtat,FE_Main,uPDFBatch,
{$ENDIF}
     FactUtil,Voirtob,
     HCtrls,ParamSoc,SysUtils,FactComm,FactOuvrage,
     FactCommBtp;

Type
TQuelCalcul = (TMcPa,TMcPr,TMcPv);
TTypeCalc = (TycProrata,TycMarge);
TEtenduCalc = (TEcAllDoc,TecParag);
//
Function AffecteMontantSurOuvrage (TOBPiece,TOBL,TOBOuvrage : TOB; NewMontant : double ; QuelCalc : TQuelCalcul; TypeCalc : TTypeCalc) : Double;
Function AffecteMontantSurLigne (TOBPiece,TOBL : TOB; NewMontant : double ; QuelCalc : TQuelCalcul; TypeCalc : TTypeCalc) : double;
procedure AffecteMontantSurPiece (TOBPiece: TOB; IndiceDepart,IndiceFin : integer; NewMontant : double ; QuelCalc : TQuelCalcul; TypeCalc : TTypeCalc; Etenducalc : TEtenduCalc);
//
Function AppliqueCoefSurOuvrage (TOBPiece,TOBL,TOBOuvrages : TOB; Coef : double ; QuelCalc : TQuelCalcul; TypeCalc : TTypeCalc;EnHt : boolean ) : Double;
Function AppliqueCoefSurLigne (TOBPiece,TOBL,TOBOuvrages : TOB; Coef : double ; QuelCalc : TQuelCalcul; TypeCalc : TTypeCalc; EnHT : boolean) : double;
//
procedure BeforeCalculMontantOuvrage (TOBPiece,TOBL, TOBOuvrages : TOB; EnHt : boolean);
procedure BeforeCalculMontantLigne (TOBPiece,TOBL, TOBOuvrages : TOB;EnHT : boolean);
procedure BeforeCalculMontantPiece (TOBPiece,TOBOuvrages : TOB; IndiceDepart,IndiceFin : integer);

implementation


Function AffecteMontantSurOuvrage (TOBPiece,TOBL,TOBOuvrage : TOB; NewMontant : double ; QuelCalc : TQuelCalcul; TypeCalc : TTypeCalc ) : Double;
begin
end;

Function AffecteMontantSurLigne (TOBPiece,TOBL : TOB; NewMontant : double ; QuelCalc : TQuelCalcul; TypeCalc : TTypeCalc) : double;
begin
end;

procedure AffecteMontantSurPiece (TOBPiece: TOB; IndiceDepart,IndiceFin : integer; NewMontant : double ; QuelCalc : TQuelCalcul; TypeCalc : TTypeCalc; Etenducalc : TEtenduCalc);
var Indice : integer;
		TheMontant,TheMontantRevient,Coef,TotalInterm : double;
    EnHt : boolean;
    MaxOccurence,Occurence,IndiceChamps,IndiceTypeLigne : integer;
begin
  if TheMontant = 0 then exit; // Une chtite erreur sans doute
	EnHt := TOBPiece.getValue('GP_FACTUREHT')='X';
  ToTalInterm := 0:
  MaxOccurence := 3;
  Occurence := 1;
  repeat
    if QuelCalc = TMcPa then
    begin
      TheMontant := TOBPiece.GetValue ('DPA');
      IndiceChamps := TOBPiece.GetNumChamp ('DPA');
      Coef := NewMontant / TheMontant;
    end else if QuelCalc = TMcPR then
    begin
      TheMontant := TOBPiece.GetValue ('DPR');
      IndiceChamps := TOBPiece.GetNumChamp ('DPR');
      Coef := NewMontant / TheMontant;
    end else if QuelCalc = TMcPV then
    begin
      if Etenducalc = TEcAllDoc then
      begin
        // Sur l'ensemble du document
        if EnHt then
        begin
        	TheMontant := TOBPiece.GetValue ('GP_TOTALHTDEV');
      		IndiceChamps := TOBPiece.GetNumChamp ('GP_TOTALHTDEV');
        end else
        begin
        	TheMontant := TOBPiece.GetValue ('GP_TOTALTTCDEV');
      		IndiceChamps := TOBPiece.GetNumChamp ('GP_TOTALTTCDEV');
        end;
        TheMontantRevient := TOBPiece.GetValue ('DPR');
        if TheMontantRevient = 0 then TheMontantRevient := TOBPiece.GetValue ('DPA');
      end else
      begin
        // Sur un paragraphe ,...voire une ligne
        if EnHt then
        begin
        	TheMontant := TOBPiece.detail[IndiceFin].GetValue ('GL_MONTANTHTDEV')
      		IndiceChamps := TOBPiece.detail[IndiceFin].GetNumChamp ('GL_MONTANTHTDEV');
        end else
        begin
        	TheMontant := TOBPiece.detail[Indicefin].GetValue ('GL_MONTANTTTCDEV');
      		IndiceChamps := TOBPiece.detail[IndiceFin].GetNumChamp ('GL_MONTANTTTCDEV');
        end;
        TheMontantRevient := TOBPiece.detail[IndiceFin].GetValue ('GL_DPR');
        if TheMontantRevient = 0 then TheMontantRevient := TOBPiece.detail[IndiceFin].GetValue ('GL_DPA');
      end;

      if TypeCalc = TycProrata then
      begin
        Coef := NewMontant / TheMontant;
      end else
      begin
        if TheMontantRevient = 0 then exit;
        Coef := (NewMontant - TheMontantRevient) / (TheMontant/TheMontantRevient);
        MaxOccurence := 3;
      end;
    end else exit;
  	ToTalInterm := 0:
    IndiceTypeLigne := TOBPiece.detail[0].GetNumChamp('GL_TYPELIGNE');

    for Indice := IndiceDepart to IndiceFin do
    begin
    	// Application du coef calculé sur le détail le plus fin
      // et recup de la valeur calculée
      TOBL := TOBPiece.detail[Indice];
			IndiceTypeLigne := TOBL.GetNumChamp('GL_TYPELIGNE');
			if OkLigneCalculMarche (TOBL,IndiceTypeLigne) then
      begin
    		TotalInterm := TotalInterm + AppliqueCoefSurLigne (TOBPiece,TOBL,TOBOuvrages,Coef,QuelCalc,TypeCalc,EnHT);
      end
    end;
    //
    if (QuelCalc = TMcPa) or (QuelCalc = TMcPr) or ((QuelCalc=TMcPv) and (EtenduCalc = TecAllDoc)) then
    begin
    	TOBPiece.putValeur (IndiceChamps,TotalInterm);
    end else
    begin
    	TOBPiece.Detail[IndiceFin].PutValeur  (IndiceChamps,TotalInterm);
    end;
    //
    if TotalInterm = newMontant then break;
    //
  until (Occurence >= MaxOccurence);
  // Repartition de l'écart dans le cas ou il en existe un...
  //
  // -----------------------------
end;

Function AppliqueCoefSurOuvrage (TOBPiece,TOBL,TOBOuvrages : TOB; Coef : double ; QuelCalc : TQuelCalcul; TypeCalc : TTypeCalc;EnHt : boolean ) : Double;
begin
end;

Function AppliqueCoefSurLigne (TOBPiece,TOBL,TOBOuvrages : TOB; Coef : double ; QuelCalc : TQuelCalcul; TypeCalc : TTypeCalc; EnHT : boolean) : double;
var TypeLigne : string;
		TheMontant : double;
    IndiceChamps : integer;
begin
	result := 0;
  TypeLigne := TOBL.GetValue('GL_TYPELIGNE');
  if ((TypeLigne = 'ARP') or (TypeLigne = 'OUV')) and (TOBL.GetValue('GL_INDICENOMEN')>0) then
  begin
    result := AppliqueCoefSurOuvrage (TOBPiece,TOBL,TOBOuvrages,Coef,QuelCalc,TypeCalc,EnHT);
  end else
  begin
    if QuelCalc = TMcPa then
    begin
      Result := Arrondi(TOBL.GEtValue('GL_DPA') * Coef,V_PGI.OkdecP);
      TOBL.putValue('GL_DPA',Result);

    end else if QuelCalc = TMcPR then
    begin
      Result := Arrondi(TOBL.GEtValue('GL_DPR') * Coef,V_PGI.OkdecP);
      TOBL.putValue('GL_DPR',Result);

    end else if QuelCalc = TMcPV then
    begin
    	if TypeCalc = TycProrata then
      begin
      	if EnHt then
        begin
          Result := Arrondi(TOBL.GEtValue('GL_PUHTDEV') * * Coef,V_PGI.OkdecP);
          TOBL.PutValue('GL_PUHTDEV',result);
        end else
        begin
          Result := Arrondi(TOBL.GEtValue('GL_PUTTCDEV') * * Coef,V_PGI.OkdecP);
          TOBL.PutValue('GL_PUTTCDEV',result);
        end;
      end else
      begin
        Result := Arrondi(TOBL.GEtValue('GL_DPR') * TOBL.GetValue('__COEFMARG')* Coef,V_PGI.OkdecP);
        TOBL.PutValue('__COEFMARG',TOBL.GetValue('__COEFMARG')*CoefMarg);
      end;
    end;

  end;
end;

procedure BeforeCalculMontantLigne (TOBPiece,TOBL, TOBOuvrages : TOB;EnHT : boolean);
var PRxAchat,PrxRevient,PrxVente : double;
    IndiceMarge : integer;
begin
  if not TOBL.FieldExists ('__COEFMARG') then
  begin
    TOBL.AddChampSup ('__COEFMARG',false);
  end;
  if not TOBL.FieldExists ('__COEFFG') then
  begin
    TOBL.AddChampSup ('__COEFFG',false);
  end;
  PrxAchat := TOBL.GetValue('GL_DPA');
  PrxRevient := TOBL.GetValue('GL_DPR');
  if EnHt then PrxVente := TOBL.GetValue('GL_PUHTDEV')
  				else PrxVente := TOBL.GetValue('GL_PUTTCDEV');
  if PrxRevient = 0 then PrxRevient := PrxAchat;
  if PrxAchat = 0 then TOBL.PutValue('__COEFFG',0)
  								else TOBL.PutValue('__COEFFG',PrxRevient/PrxAchat);
  if PrxRevient = 0 then TOBL.PutValue('__COEFMARG',0)
  									else TOBL.PutValue('__COEFMARG',PrxVente/PrxRevient);

  if TOBL.GetValue('GL_INDICENOMEN')> 0 then BeforeCalculMontantOuvrage (TOBPiece,TOBL,TOBOuvrages,EnHt);
end;

procedure BeforeCalculMontantPiece (TOBPiece,TOBOuvrages : TOB; IndiceDepart,IndiceFin : integer);
var Indice : integer;
		TOBL : TOB;
    EnHt : boolean;
begin
	EnHt := TOBPiece.getValue('GP_FACTUREHT')='X';
	for Indice := IndiceDepart to IndiceFin do
  begin
    TOBL := TOBPiece.detail[Indice] ;
    BeforeCalculMontantLigne (TOBPiece,TOBL,TOBOuvrages,EnHt);
  end;
end;

procedure BeforeCalculMontantDetailOuvrage (TOBPiece,TOBL, TOBOuv : TOB; EnHt : boolean);
var Indice : integer;
		TOBO : TOB;
		PRxAchat,PrxRevient,PrxVente : double;
begin
  if not TOBOuv.FieldExists ('__COEFMARG') then
  begin
    TOBOuv.AddChampSup ('__COEFMARG',false);
  end;
  if not TOBOuv.FieldExists ('__COEFFG') then
  begin
    TOBOuv.AddChampSup ('__COEFFG',false);
  end;
  PrxAchat := TOBOuv.GetValue('BLO_DPA');
  PrxRevient := TOBOuv.GetValue('BLO_DPR');
  if EnHt then PrxVente := TOBOuv.GetValue('BLO_PUHTDEV')
  				else PrxVente := TOBOuv.GetValue('BLO_PUTTCDEV');
  if PrxRevient = 0 then PrxRevient := PrxAchat;
  if PrxAchat = 0 then TOBL.PutValue('__COEFFG',0)
  								else TOBL.PutValue('__COEFFG',PrxRevient/PrxAchat);
  if PrxRevient = 0 then TOBL.PutValue('__COEFMARG',0)
  									else TOBL.PutValue('__COEFMARG',PrxVente/PrxRevient);
  if TOBOUv.detail.count > 0 then
  begin
    for Indice := 0 to TOBOUV.detail.count -1 do
    begin
    	TOBO := TOBOUV.detail[Indice];
      BeforeCalculMontantDetailOuvrage (TOBPiece,TOBL,TOBO,EnHt);
    end;
  end;
end;

procedure BeforeCalculMontantOuvrage (TOBPiece,TOBL, TOBOuvrages : TOB; EnHt : boolean);
var TOBOUVRAGE,TOBOUV : TOB;
		IndiceNomen,Indice : Integer;
begin
	IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  if IndiceNomen = 0 then exit;
  TOBOUVRAGE := TOBOuvrages.detail[IndiceNomen-1]; // commence a 1 ce c.n la !!
  for Indice := 0 to TOBOUVRAGE.detail.count -1 do
  begin
  	TOBOUv := TOBOUVRAGE.detail[Indice];
		BeforeCalculMontantDetailOuvrage (TOBPiece,TOBL, TOBOuv ,EnHt);
  end;
end;

end.
