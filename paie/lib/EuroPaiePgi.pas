{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 01/08/2001
Modifié le ... : 01/08/2001
Description .. : Unit de traitement des différentes tables lors du basculement
Suite ........ : en EURO
Suite ........ : Ne contient que les fonctions de bas niveaux de traitement
Suite ........ : de l'information.
Suite ........ :
Mots clefs ... : PAIE;PGEURO
*****************************************************************}
{
PT1 : 01/08/2001 : V547 : PH
     Modif traitement histobulletin sur les rémunérations dont le montant est de type montant
     Ne prenais pas en compte ce cas, regardais uniquement base,taux,coeff
PT2 : 01/08/2001 : V547 : PH
     Rajout rechargement des paramètres société de la paie en fin de traitement de la bascule
PT3 : 01/08/2001 : V547 : PH ==> Fonction BasculeFrancPaie
     Rajout de la procédure inverse de la bascule .Attention, toutes les tables ne sont pas
     retraduites en inverses notament SALARIES sur les montants car il y aurait des ^problèmes
     d'arrondi. En fait, cela concerne surtout les historiques.
PT4 : 16/08/2001 : V547 : PH
     PRM_BASEMTQTE passe de Boolean en COMBO donc test de valeur à changer
PT5 : 30/10/2001 : V562 : PH Annulation migration ne reconvertit pas les plafonds et tranches
                             et cout patronal

}

unit EuroPaiePGI;

interface


uses windows, hent1, sysutils, UTOB, classes, Forms, ExtCtrls,
{$IFDEF EAGLCLIENT}
{$ELSE}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  HCtrls, Ent1, UtilPGI, EntPaie, ParamSoc;

function BasculePaieEuro(StDev: string; PanelPGI: TPanel; Sens: string): Boolean;

implementation

var
  OldDev: string;
  Tob_Rem, Tob_Cot, Tob_Cum: TOB;
  PgTypSalLib1, PgTypSalLib2, PgTypSalLib3, PgTypSalLib4: Boolean;

  // Fonction de conversion en EURO de chaque champ pour chaque table

procedure ConvertEuroPG(Q: TQuery; NomT: string);
var
  TRub: TOB;
  Nature, Rub, Cum: string;
begin
  if NomT = '' then Exit;
  if NomT = 'ABSENCESALARIE' then
  begin
    if Q.FindField('PCN_TOPCONVERT').AsString <> 'X' then
    begin
      Q.FindField('PCN_BASE').AsFloat := FrancToEuro(Q.FindField('PCN_BASE').AsFloat);
      Q.FindField('PCN_ABSENCE').AsFloat := FrancToEuro(Q.FindField('PCN_ABSENCE').AsFloat);
      Q.FindField('PCN_ABSENCEMANU').AsFloat := FrancToEuro(Q.FindField('PCN_ABSENCEMANU').AsFloat);
      Q.FindField('PCN_VALOX').AsFloat := FrancToEuro(Q.FindField('PCN_VALOX').AsFloat);
      Q.FindField('PCN_VALOMS').AsFloat := FrancToEuro(Q.FindField('PCN_VALOMS').AsFloat);
      Q.FindField('PCN_VALORETENUE').AsFloat := FrancToEuro(Q.FindField('PCN_VALORETENUE').AsFloat);
      Q.FindField('PCN_VALOMANUELLE').AsFloat := FrancToEuro(Q.FindField('PCN_VALOMANUELLE').AsFloat);
      Q.FindField('PCN_TOPCONVERT').AsString := 'X';
    end;
  end else
    if NomT = 'CONTRATTRAVAIL' then
  begin
    if Q.FindField('PCI_TOPCONVERT').AsString <> 'X' then
    begin
      if PgTypSalLib1 then
      begin
        Q.FindField('PCI_SALAIREMOIS1').AsFloat := FrancToEuro(Q.FindField('PCI_SALAIREMOIS1').AsFloat);
        Q.FindField('PCI_SALAIRANN1').AsFloat := FrancToEuro(Q.FindField('PCI_SALAIRANN1').AsFloat);
      end;
      if PgTypSalLib2 then
      begin
        Q.FindField('PCI_SALAIREMOIS2').AsFloat := FrancToEuro(Q.FindField('PCI_SALAIREMOIS2').AsFloat);
        Q.FindField('PCI_SALAIRANN2').AsFloat := FrancToEuro(Q.FindField('PCI_SALAIRANN2').AsFloat);
      end;
      if PgTypSalLib3 then
      begin
        Q.FindField('PCI_SALAIREMOIS3').AsFloat := FrancToEuro(Q.FindField('PCI_SALAIREMOIS3').AsFloat);
        Q.FindField('PCI_SALAIRANN3').AsFloat := FrancToEuro(Q.FindField('PCI_SALAIRANN3').AsFloat);
      end;
      if PgTypSalLib4 then
      begin
        Q.FindField('PCI_SALAIREMOIS4').AsFloat := FrancToEuro(Q.FindField('PCI_SALAIREMOIS4').AsFloat);
        Q.FindField('PCI_SALAIRANN4').AsFloat := FrancToEuro(Q.FindField('PCI_SALAIRANN4').AsFloat);
      end;
      Q.FindField('PCI_TOPCONVERT').AsString := 'X';
    end;
  end else
    if NomT = 'HISTOANALPAIE' then
  begin
    if (Q.FindField('PHA_OBASEREM').AsFloat = 0) and
      (Q.FindField('PHA_OMTREM').AsFloat = 0) and
      (Q.FindField('PHA_OBASECOT').AsFloat = 0) and
      (Q.FindField('PHA_OMTSALARIAL').AsFloat = 0) and
      (Q.FindField('PHA_OMTPATRONAL').AsFloat = 0) then
    begin
      Q.FindField('PHA_OBASEREM').AsFloat := Q.FindField('PHA_BASEREM').AsFloat;
      Q.FindField('PHA_OMTREM').AsFloat := Q.FindField('PHA_MTREM').AsFloat;
      Q.FindField('PHA_OBASECOT').AsFloat := Q.FindField('PHA_BASECOT').AsFloat;
      Q.FindField('PHA_OMTSALARIAL').AsFloat := Q.FindField('PHA_MTSALARIAL').AsFloat;
      Q.FindField('PHA_OMTPATRONAL').AsFloat := Q.FindField('PHA_MTPATRONAL').AsFloat;
      Nature := Q.FindField('PHA_NATURERUB').AsString;
      Rub := Q.FindField('PHA_RUBRIQUE').AsString;
      if Nature = 'AAA' then
      begin // Remuneration
        TRub := TOB_REM.FindFirst(['PRM_RUBRIQUE'], [Rub], FALSE);
        if TRub <> nil then
        begin
          // PT4 : 16/08/2001 : V547  PH
          if TRub.GetValue('PRM_BASEMTQTE') = '001' then Q.FindField('PHA_BASEREM').AsFloat := FrancToEuro(Q.FindField('PHA_BASEREM').AsFloat);
          // PT1 : 01/08/2001 : V547 : PH  Rajout cas PRM_MONTANTMTQTE
          // PT4 : 16/08/2001 : V547  PH
          if ((TRub.GetValue('PRM_BASEMTQTE') = '001') or (TRub.GetValue('PRM_TAUXMTQTE') = 'X')
            or (TRub.GetValue('PRM_COEFFMTQTE') = 'X')) or (TRub.GetValue('PRM_MONTANTMTQTE') = 'X') then Q.FindField('PHA_MTREM').AsFloat := FrancToEuro(Q.FindField('PHA_MTREM').AsFloat);
          // FIN PT1
        end;
      end
      else
      begin // Cotisation
        Q.FindField('PHA_BASECOT').AsFloat := FrancToEuro(Q.FindField('PHA_BASECOT').AsFloat);
        Q.FindField('PHA_MTSALARIAL').AsFloat := FrancToEuro(Q.FindField('PHA_MTSALARIAL').AsFloat);
        Q.FindField('PHA_MTPATRONAL').AsFloat := FrancToEuro(Q.FindField('PHA_MTPATRONAL').AsFloat);
      end;
    end;
  end else
    if NomT = 'HISTOBULLETIN' then
  begin
    if (Q.FindField('PHB_OBASEREM').AsFloat = 0) and
      (Q.FindField('PHB_OTAUXREM').AsFloat = 0) and
      (Q.FindField('PHB_OCOEFFREM').AsFloat = 0) and
      (Q.FindField('PHB_OMTREM').AsFloat = 0) and
      (Q.FindField('PHB_OBASECOT').AsFloat = 0) and
      (Q.FindField('PHB_OTAUXSALARIAL').AsFloat = 0) and
      (Q.FindField('PHB_OMTSALARIAL').AsFloat = 0) and
      (Q.FindField('PHB_OTAUXPATRONAL').AsFloat = 0) and
      (Q.FindField('PHB_OMTPATRONAL').AsFloat = 0) then
    begin
      Nature := Q.FindField('PHB_NATURERUB').AsString;
      Rub := Q.FindField('PHB_RUBRIQUE').AsString;
      if Nature = 'AAA' then TRub := TOB_REM.FindFirst(['PRM_RUBRIQUE'], [Rub], FALSE)
      else TRub := TOB_Cot.FindFirst(['PCT_RUBRIQUE'], [Rub], FALSE);
      if TRub <> nil then
      begin
        Q.FindField('PHB_OBASEREM').AsFloat := Q.FindField('PHB_BASEREM').AsFloat;
        Q.FindField('PHB_OTAUXREM').AsFloat := Q.FindField('PHB_TAUXREM').AsFloat;
        Q.FindField('PHB_OCOEFFREM').AsFloat := Q.FindField('PHB_COEFFREM').AsFloat;
        Q.FindField('PHB_OMTREM').AsFloat := Q.FindField('PHB_MTREM').AsFloat;
        Q.FindField('PHB_OBASECOT').AsFloat := Q.FindField('PHB_BASECOT').AsFloat;
        Q.FindField('PHB_OTAUXSALARIAL').AsFloat := Q.FindField('PHB_TAUXSALARIAL').AsFloat;
        Q.FindField('PHB_OMTSALARIAL').AsFloat := Q.FindField('PHB_MTSALARIAL').AsFloat;
        Q.FindField('PHB_OTAUXPATRONAL').AsFloat := Q.FindField('PHB_TAUXPATRONAL').AsFloat;
        Q.FindField('PHB_OMTPATRONAL').AsFloat := Q.FindField('PHB_MTPATRONAL').AsFloat;
        if Nature = 'AAA' then
        begin // Remuneration
          // PT4 : 16/08/2001 : V547  PH
          if TRub.GetValue('PRM_BASEMTQTE') = '001' then Q.FindField('PHB_BASEREM').AsFloat := FrancToEuro(Q.FindField('PHB_BASEREM').AsFloat);
          if TRub.GetValue('PRM_TAUXMTQTE') = 'X' then Q.FindField('PHB_TAUXREM').AsFloat := FrancToEuro(Q.FindField('PHB_TAUXREM').AsFloat);
          if TRub.GetValue('PRM_COEFFMTQTE') = 'X' then Q.FindField('PHB_COEFFREM').AsFloat := FrancToEuro(Q.FindField('PHB_COEFFREM').AsFloat);
          // PT-1 : 01/08/2001 : V547 : PH  Rajout cas PRM_MONTANTMTQTE
          // PT4 : 16/08/2001 : V547  PH
          if ((TRub.GetValue('PRM_BASEMTQTE') = '001') or (TRub.GetValue('PRM_TAUXMTQTE') = 'X')
            or (TRub.GetValue('PRM_COEFFMTQTE') = 'X')) or (TRub.GetValue('PRM_MONTANTMTQTE') = 'X') then Q.FindField('PHB_MTREM').AsFloat := FrancToEuro(Q.FindField('PHB_MTREM').AsFloat);
          // FIN PT-1
        end
        else
        begin // Cotisation
          Q.FindField('PHB_BASECOT').AsFloat := FrancToEuro(Q.FindField('PHB_BASECOT').AsFloat);
          Q.FindField('PHB_MTSALARIAL').AsFloat := FrancToEuro(Q.FindField('PHB_MTSALARIAL').AsFloat);
          Q.FindField('PHB_MTPATRONAL').AsFloat := FrancToEuro(Q.FindField('PHB_MTPATRONAL').AsFloat);
          Q.FindField('PHB_PLAFOND').AsFloat := FrancToEuro(Q.FindField('PHB_PLAFOND').AsFloat);
          Q.FindField('PHB_PLAFOND1').AsFloat := FrancToEuro(Q.FindField('PHB_PLAFOND1').AsFloat);
          Q.FindField('PHB_PLAFOND2').AsFloat := FrancToEuro(Q.FindField('PHB_PLAFOND2').AsFloat);
          Q.FindField('PHB_PLAFOND3').AsFloat := FrancToEuro(Q.FindField('PHB_PLAFOND3').AsFloat);
          Q.FindField('PHB_TRANCHE1').AsFloat := FrancToEuro(Q.FindField('PHB_TRANCHE1').AsFloat);
          Q.FindField('PHB_TRANCHE2').AsFloat := FrancToEuro(Q.FindField('PHB_TRANCHE2').AsFloat);
          Q.FindField('PHB_TRANCHE3').AsFloat := FrancToEuro(Q.FindField('PHB_TRANCHE3').AsFloat);
        end;
      end;
    end;
  end else
    if NomT = 'HISTOCUMSAL' then
  begin
    Cum := Q.FindField('PHC_CUMULPAIE').AsString;
    TRub := Tob_Cum.FindFirst(['PCL_CUMULPAIE'], [Cum], FALSE);
    if (TRub <> nil) and (Q.FindField('PHC_OMONTANT').AsFloat = 0) then
    begin
      Q.FindField('PHC_OMONTANT').AsFloat := Q.FindField('PHC_MONTANT').AsFloat;
      if TRub.GetValue('PCL_TYPECUMUL') = 'X' then Q.FindField('PHC_MONTANT').AsFloat := FrancToEuro(Q.FindField('PHC_MONTANT').AsFloat);
    end;
  end else
    if NomT = 'HISTOSAISRUB' then
  begin
    if Q.FindField('PSD_TOPCONVERT').AsString <> 'X' then
    begin
      // Cas acompte tjrs un montant
      if Q.FindField('PSD_RUBRIQUE').AsString = 'ACP' then Q.FindField('PSD_MONTANT').AsFloat := FrancToEuro(Q.FindField('PSD_MONTANT').AsFloat)
      else
      begin
        Rub := Q.FindField('PSD_RUBRIQUE').AsString;
        TRub := TOB_REM.FindFirst(['PRM_RUBRIQUE'], [Rub], FALSE);
        if TRub <> nil then
        begin
          // PT4 : 16/08/2001 : V547  PH
          if (Q.FindField('PSD_BASE').AsFloat <> 0) and (TRub.GetValue('PRM_BASEMTQTE') = '001') then Q.FindField('PSD_BASE').AsFloat := FrancToEuro(Q.FindField('PSD_BASE').AsFloat);
          if (Q.FindField('PSD_TAUX').AsFloat <> 0) and (TRub.GetValue('PRM_TAUXMTQTE') = 'X') then Q.FindField('PSD_TAUX').AsFloat := FrancToEuro(Q.FindField('PSD_TAUX').AsFloat);
          if (Q.FindField('PSD_COEFF').AsFloat <> 0) and (TRub.GetValue('PRM_COEFFMTQTE') = 'X') then Q.FindField('PSD_COEFF').AsFloat := FrancToEuro(Q.FindField('PSD_COEFF').AsFloat);
          // PT-1 : 01/08/2001 : V547 : PH  Rajout cas PRM_MONTANTMTQTE
          // PT4 : 16/08/2001 : V547  PH
          if (Q.FindField('PSD_MONTANT').AsFloat <> 0) and (((TRub.GetValue('PRM_BASEMTQTE') = '001')
            or (TRub.GetValue('PRM_TAUXMTQTE') = 'X') or (TRub.GetValue('PRM_COEFFMTQTE') = 'X')) or (TRub.GetValue('PRM_MONTANTMTQTE') = 'X')) then
            Q.FindField('PSD_MONTANT').AsFloat := FrancToEuro(Q.FindField('PSD_MONTANT').AsFloat);
          // FIN PT-1
          Q.FindField('PSD_TOPCONVERT').AsString := 'X';
        end;
      end;
    end;
  end else
    if NomT = 'HISTOSALARIE' then
  begin
    if Q.FindField('PHS_TOPCONVERT').AsString <> 'X' then
    begin
      if PgTypSalLib1 then
      begin
        Q.FindField('PHS_SALAIREMOIS1').AsFloat := FrancToEuro(Q.FindField('PHS_SALAIREMOIS1').AsFloat);
        Q.FindField('PHS_SALAIREANN1').AsFloat := FrancToEuro(Q.FindField('PHS_SALAIREANN1').AsFloat);
      end;
      if PgTypSalLib2 then
      begin
        Q.FindField('PHS_SALAIREMOIS2').AsFloat := FrancToEuro(Q.FindField('PHS_SALAIREMOIS2').AsFloat);
        Q.FindField('PHS_SALAIREANN2').AsFloat := FrancToEuro(Q.FindField('PHS_SALAIREANN2').AsFloat);
      end;
      if PgTypSalLib3 then
      begin
        Q.FindField('PHS_SALAIREMOIS3').AsFloat := FrancToEuro(Q.FindField('PHS_SALAIREMOIS3').AsFloat);
        Q.FindField('PHS_SALAIREANN3').AsFloat := FrancToEuro(Q.FindField('PHS_SALAIREANN3').AsFloat);
      end;
      if PgTypSalLib4 then
      begin
        Q.FindField('PHS_SALAIREMOIS4').AsFloat := FrancToEuro(Q.FindField('PHS_SALAIREMOIS4').AsFloat);
        Q.FindField('PHS_SALAIREANN4').AsFloat := FrancToEuro(Q.FindField('PHS_SALAIREANN4').AsFloat);
      end;
      Q.FindField('PHS_TOPCONVERT').AsString := 'X';
    end;
  end else
    if NomT = 'PAIEENCOURS' then
  begin
    if (Q.FindField('PPU_OCBRUT').AsFloat = 0) and
      (Q.FindField('PPU_OCBRUTFISCAL').AsFloat = 0) and
      (Q.FindField('PPU_OCNETIMPOSAB').AsFloat = 0) and
      (Q.FindField('PPU_OCNETAPAYER').AsFloat = 0) and
      (Q.FindField('PPU_OCCOUTSALARIE').AsFloat = 0) and
      (Q.FindField('PPU_OCCOUTPATRON').AsFloat = 0) and
      (Q.FindField('PPU_OCPLAFONDSS').AsFloat = 0) and
      (Q.FindField('PPU_OCBASESS').AsFloat = 0) then
    begin
      Q.FindField('PPU_OCBRUT').AsFloat := Q.FindField('PPU_CBRUT').AsFloat;
      Q.FindField('PPU_OCBRUTFISCAL').AsFloat := Q.FindField('PPU_CBRUTFISCAL').AsFloat;
      Q.FindField('PPU_OCNETIMPOSAB').AsFloat := Q.FindField('PPU_CNETIMPOSAB').AsFloat;
      Q.FindField('PPU_OCNETAPAYER').AsFloat := Q.FindField('PPU_CNETAPAYER').AsFloat;
      Q.FindField('PPU_OCCOUTSALARIE').AsFloat := Q.FindField('PPU_CCOUTSALARIE').AsFloat;
      Q.FindField('PPU_OCCOUTPATRON').AsFloat := Q.FindField('PPU_CCOUTPATRON').AsFloat;
      Q.FindField('PPU_OCPLAFONDSS').AsFloat := Q.FindField('PPU_CPLAFONDSS').AsFloat;
      Q.FindField('PPU_OCBASESS').AsFloat := Q.FindField('PPU_CBASESS').AsFloat;
      Q.FindField('PPU_CBRUT').AsFloat := FrancToEuro(Q.FindField('PPU_CBRUT').AsFloat);
      Q.FindField('PPU_CBRUTFISCAL').AsFloat := FrancToEuro(Q.FindField('PPU_CBRUTFISCAL').AsFloat);
      Q.FindField('PPU_CNETIMPOSAB').AsFloat := FrancToEuro(Q.FindField('PPU_CNETIMPOSAB').AsFloat);
      Q.FindField('PPU_CNETAPAYER').AsFloat := FrancToEuro(Q.FindField('PPU_CNETAPAYER').AsFloat);
      Q.FindField('PPU_CCOUTSALARIE').AsFloat := FrancToEuro(Q.FindField('PPU_CCOUTSALARIE').AsFloat);
      Q.FindField('PPU_CCOUTPATRON').AsFloat := FrancToEuro(Q.FindField('PPU_CCOUTPATRON').AsFloat);
      Q.FindField('PPU_CPLAFONDSS').AsFloat := FrancToEuro(Q.FindField('PPU_CPLAFONDSS').AsFloat);
      Q.FindField('PPU_CBASESS').AsFloat := FrancToEuro(Q.FindField('PPU_CBASESS').AsFloat);
    end;
  end else
    if NomT = 'SALARIES' then
  begin
    if Q.FindField('PSA_TOPCONVERT').AsString <> 'X' then
    begin
      if PgTypSalLib1 then
      begin
        Q.FindField('PSA_SALAIREMOIS1').AsFloat := FrancToEuro(Q.FindField('PSA_SALAIREMOIS1').AsFloat);
        Q.FindField('PSA_SALAIRANN1').AsFloat := FrancToEuro(Q.FindField('PSA_SALAIRANN1').AsFloat);
      end;
      if PgTypSalLib2 then
      begin
        Q.FindField('PSA_SALAIREMOIS2').AsFloat := FrancToEuro(Q.FindField('PSA_SALAIREMOIS2').AsFloat);
        Q.FindField('PSA_SALAIRANN2').AsFloat := FrancToEuro(Q.FindField('PSA_SALAIRANN2').AsFloat);
      end;
      if PgTypSalLib3 then
      begin
        Q.FindField('PSA_SALAIREMOIS3').AsFloat := FrancToEuro(Q.FindField('PSA_SALAIREMOIS3').AsFloat);
        Q.FindField('PSA_SALAIRANN3').AsFloat := FrancToEuro(Q.FindField('PSA_SALAIRANN3').AsFloat);
      end;
      if PgTypSalLib4 then
      begin
        Q.FindField('PSA_SALAIREMOIS4').AsFloat := FrancToEuro(Q.FindField('PSA_SALAIREMOIS4').AsFloat);
        Q.FindField('PSA_SALAIRANN4').AsFloat := FrancToEuro(Q.FindField('PSA_SALAIRANN4').AsFloat);
      end;
      Q.FindField('PSA_SALAIRETHEO').AsFloat := FrancToEuro(Q.FindField('PSA_SALAIRETHEO').AsFloat);
      Q.FindField('PSA_TAUXHORAIRE').AsFloat := FrancToEuro(Q.FindField('PSA_TAUXHORAIRE').AsFloat);
      Q.FindField('PSA_SALAIREMULTI').AsFloat := FrancToEuro(Q.FindField('PSA_SALAIREMULTI').AsFloat);

      Q.FindField('PSA_TOPCONVERT').AsString := 'X';
    end;
  end;
end;

// Fonction de reconversion en Franc de chaque champ pour chaque table

procedure ConvertFrancPG(Q: TQuery; NomT: string);
var
  TRub: TOB;
  Rub: string;
begin
  if NomT = '' then Exit;
  if NomT = 'ABSENCESALARIE' then
  begin // Attention , Pb arrondi
    if Q.FindField('PCN_TOPCONVERT').AsString = 'X' then
    begin
      Q.FindField('PCN_BASE').AsFloat := EuroToFranc(Q.FindField('PCN_BASE').AsFloat);
      Q.FindField('PCN_ABSENCE').AsFloat := EuroToFranc(Q.FindField('PCN_ABSENCE').AsFloat);
      Q.FindField('PCN_ABSENCEMANU').AsFloat := EuroToFranc(Q.FindField('PCN_ABSENCEMANU').AsFloat);
      Q.FindField('PCN_VALOX').AsFloat := EuroToFranc(Q.FindField('PCN_VALOX').AsFloat);
      Q.FindField('PCN_VALOMS').AsFloat := EuroToFranc(Q.FindField('PCN_VALOMS').AsFloat);
      Q.FindField('PCN_VALORETENUE').AsFloat := EuroToFranc(Q.FindField('PCN_VALORETENUE').AsFloat);
      Q.FindField('PCN_VALOMANUELLE').AsFloat := EuroToFranc(Q.FindField('PCN_VALOMANUELLE').AsFloat);
      Q.FindField('PCN_TOPCONVERT').AsString := '-';
    end;
  end else
    if NomT = 'CONTRATTRAVAIL' then
  begin // Attention , Pb arrondi
    if Q.FindField('PCI_TOPCONVERT').AsString = 'X' then
    begin
      if PgTypSalLib1 then
      begin
        Q.FindField('PCI_SALAIREMOIS1').AsFloat := EuroToFranc(Q.FindField('PCI_SALAIREMOIS1').AsFloat);
        Q.FindField('PCI_SALAIRANN1').AsFloat := EuroToFranc(Q.FindField('PCI_SALAIRANN1').AsFloat);
      end;
      if PgTypSalLib2 then
      begin
        Q.FindField('PCI_SALAIREMOIS2').AsFloat := EuroToFranc(Q.FindField('PCI_SALAIREMOIS2').AsFloat);
        Q.FindField('PCI_SALAIRANN2').AsFloat := EuroToFranc(Q.FindField('PCI_SALAIRANN2').AsFloat);
      end;
      if PgTypSalLib3 then
      begin
        Q.FindField('PCI_SALAIREMOIS3').AsFloat := EuroToFranc(Q.FindField('PCI_SALAIREMOIS3').AsFloat);
        Q.FindField('PCI_SALAIRANN3').AsFloat := EuroToFranc(Q.FindField('PCI_SALAIRANN3').AsFloat);
      end;
      if PgTypSalLib4 then
      begin
        Q.FindField('PCI_SALAIREMOIS4').AsFloat := EuroToFranc(Q.FindField('PCI_SALAIREMOIS4').AsFloat);
        Q.FindField('PCI_SALAIRANN4').AsFloat := EuroToFranc(Q.FindField('PCI_SALAIRANN4').AsFloat);
      end;
      Q.FindField('PCI_TOPCONVERT').AsString := '-';
    end;
  end else
    if NomT = 'HISTOANALPAIE' then
  begin
    if Q.FindField('PHA_OBASEREM').AsFloat <> 0 then
    begin
      Q.FindField('PHA_BASEREM').AsFloat := Q.FindField('PHA_OBASEREM').AsFloat;
      Q.FindField('PHA_OBASEREM').AsFloat := 0;
    end;
    if Q.FindField('PHA_OMTREM').AsFloat <> 0 then
    begin
      Q.FindField('PHA_MTREM').AsFloat := Q.FindField('PHA_OMTREM').AsFloat;
      Q.FindField('PHA_OMTREM').AsFloat := 0;
    end;
    if Q.FindField('PHA_OBASECOT').AsFloat <> 0 then
    begin
      Q.FindField('PHA_BASECOT').AsFloat := Q.FindField('PHA_OBASECOT').AsFloat;
      Q.FindField('PHA_OBASECOT').AsFloat := 0;
    end;
    if Q.FindField('PHA_OMTSALARIAL').AsFloat <> 0 then
    begin
      Q.FindField('PHA_MTSALARIAL').AsFloat := Q.FindField('PHA_OMTSALARIAL').AsFloat;
      Q.FindField('PHA_OMTSALARIAL').AsFloat := 0;
    end;
    if Q.FindField('PHA_OMTPATRONAL').AsFloat <> 0 then
    begin
      Q.FindField('PHA_MTPATRONAL').AsFloat := Q.FindField('PHA_OMTPATRONAL').AsFloat;
      Q.FindField('PHA_OMTPATRONAL').AsFloat := 0;
    end;
  end else
    if NomT = 'HISTOBULLETIN' then
  begin
    if (Q.FindField('PHB_OBASEREM').AsFloat <> 0) then
    begin
      Q.FindField('PHB_BASEREM').AsFloat := Q.FindField('PHB_OBASEREM').AsFloat;
      Q.FindField('PHB_OBASEREM').AsFloat := 0;
    end;
    if (Q.FindField('PHB_OTAUXREM').AsFloat <> 0) then
    begin
      Q.FindField('PHB_TAUXREM').AsFloat := Q.FindField('PHB_OTAUXREM').AsFloat;
      Q.FindField('PHB_OTAUXREM').AsFloat := 0;
    end;
    if (Q.FindField('PHB_OCOEFFREM').AsFloat <> 0) then
    begin
      Q.FindField('PHB_COEFFREM').AsFloat := Q.FindField('PHB_OCOEFFREM').AsFloat;
      Q.FindField('PHB_OCOEFFREM').AsFloat := 0;
    end;
    if (Q.FindField('PHB_OMTREM').AsFloat <> 0) then
    begin
      Q.FindField('PHB_MTREM').AsFloat := Q.FindField('PHB_OMTREM').AsFloat;
      Q.FindField('PHB_OMTREM').AsFloat := 0;
    end;
    if (Q.FindField('PHB_OBASECOT').AsFloat <> 0) then
    begin
      Q.FindField('PHB_BASECOT').AsFloat := Q.FindField('PHB_OBASECOT').AsFloat;
      Q.FindField('PHB_OBASECOT').AsFloat := 0;
    end;
    if (Q.FindField('PHB_OTAUXSALARIAL').AsFloat <> 0) then
    begin
      Q.FindField('PHB_TAUXSALARIAL').AsFloat := Q.FindField('PHB_OTAUXSALARIAL').AsFloat;
      Q.FindField('PHB_OTAUXSALARIAL').AsFloat := 0;
    end;
    if (Q.FindField('PHB_OMTSALARIAL').AsFloat <> 0) then
    begin
      Q.FindField('PHB_MTSALARIAL').AsFloat := Q.FindField('PHB_OMTSALARIAL').AsFloat;
      Q.FindField('PHB_OMTSALARIAL').AsFloat := 0;
    end;
    if (Q.FindField('PHB_OTAUXPATRONAL').AsFloat <> 0) then
    begin
      Q.FindField('PHB_TAUXPATRONAL').AsFloat := Q.FindField('PHB_OTAUXPATRONAL').AsFloat;
      Q.FindField('PHB_OTAUXPATRONAL').AsFloat := 0;
    end;
    if (Q.FindField('PHB_OMTPATRONAL').AsFloat <> 0) then
    begin
      Q.FindField('PHB_MTPATRONAL').AsFloat := Q.FindField('PHB_OMTPATRONAL').AsFloat;
      Q.FindField('PHB_OMTPATRONAL').AsFloat := 0;
    end;
    // PT5 : 30/10/2001 : V562 : PH Annulation migration ne reconvertit pas les plafonds et tranches
    Q.FindField('PHB_PLAFOND').AsFloat := EuroToFranc(Q.FindField('PHB_PLAFOND').AsFloat);
    Q.FindField('PHB_PLAFOND1').AsFloat := EuroToFranc(Q.FindField('PHB_PLAFOND1').AsFloat);
    Q.FindField('PHB_PLAFOND2').AsFloat := EuroToFranc(Q.FindField('PHB_PLAFOND2').AsFloat);
    Q.FindField('PHB_PLAFOND3').AsFloat := EuroToFranc(Q.FindField('PHB_PLAFOND3').AsFloat);
    Q.FindField('PHB_TRANCHE1').AsFloat := EuroToFranc(Q.FindField('PHB_TRANCHE1').AsFloat);
    Q.FindField('PHB_TRANCHE2').AsFloat := EuroToFranc(Q.FindField('PHB_TRANCHE2').AsFloat);
    Q.FindField('PHB_TRANCHE3').AsFloat := EuroToFranc(Q.FindField('PHB_TRANCHE3').AsFloat);

  end else
    if NomT = 'HISTOCUMSAL' then
  begin
    if Q.FindField('PHC_OMONTANT').AsFloat <> 0 then
    begin
      Q.FindField('PHC_MONTANT').AsFloat := Q.FindField('PHC_OMONTANT').AsFloat;
      Q.FindField('PHC_OMONTANT').AsFloat := 0;
    end;
  end else
    if NomT = 'HISTOSAISRUB' then
  begin
    if Q.FindField('PSD_TOPCONVERT').AsString = 'X' then
    begin
      // Cas acompte tjrs un montant Attention, pb arrondi
      if Q.FindField('PSD_RUBRIQUE').AsString = 'ACP' then Q.FindField('PSD_MONTANT').AsFloat := EuroToFranc(Q.FindField('PSD_MONTANT').AsFloat)
      else
      begin
        Rub := Q.FindField('PSD_RUBRIQUE').AsString;
        TRub := TOB_REM.FindFirst(['PRM_RUBRIQUE'], [Rub], FALSE);
        if TRub <> nil then
        begin
          // PT4 : 16/08/2001 : V547  PH
          if (Q.FindField('PSD_BASE').AsFloat <> 0) and (TRub.GetValue('PRM_BASEMTQTE') = '001') then Q.FindField('PSD_BASE').AsFloat := EuroToFranc(Q.FindField('PSD_BASE').AsFloat);
          if (Q.FindField('PSD_TAUX').AsFloat <> 0) and (TRub.GetValue('PRM_TAUXMTQTE') = 'X') then Q.FindField('PSD_TAUX').AsFloat := EuroToFranc(Q.FindField('PSD_TAUX').AsFloat);
          if (Q.FindField('PSD_COEFF').AsFloat <> 0) and (TRub.GetValue('PRM_COEFFMTQTE') = 'X') then Q.FindField('PSD_COEFF').AsFloat := EuroToFranc(Q.FindField('PSD_COEFF').AsFloat);
          // PT4 : 16/08/2001 : V547  PH
          if (Q.FindField('PSD_MONTANT').AsFloat <> 0) and (((TRub.GetValue('PRM_BASEMTQTE') = '001')
            or (TRub.GetValue('PRM_TAUXMTQTE') = 'X') or (TRub.GetValue('PRM_COEFFMTQTE') = 'X')) or (TRub.GetValue('PRM_MONTANTMTQTE') = 'X')) then
            Q.FindField('PSD_MONTANT').AsFloat := EuroToFranc(Q.FindField('PSD_MONTANT').AsFloat);
          Q.FindField('PSD_TOPCONVERT').AsString := '-';
        end;
      end;
    end;
  end else
    if NomT = 'HISTOSALARIE' then
  begin
    if Q.FindField('PHS_TOPCONVERT').AsString <> 'X' then
    begin
      if PgTypSalLib1 then
      begin
        Q.FindField('PHS_SALAIREMOIS1').AsFloat := EuroToFranc(Q.FindField('PHS_SALAIREMOIS1').AsFloat);
        Q.FindField('PHS_SALAIREANN1').AsFloat := EuroToFranc(Q.FindField('PHS_SALAIREANN1').AsFloat);
      end;
      if PgTypSalLib2 then
      begin
        Q.FindField('PHS_SALAIREMOIS2').AsFloat := EuroToFranc(Q.FindField('PHS_SALAIREMOIS2').AsFloat);
        Q.FindField('PHS_SALAIREANN2').AsFloat := EuroToFranc(Q.FindField('PHS_SALAIREANN2').AsFloat);
      end;
      if PgTypSalLib3 then
      begin
        Q.FindField('PHS_SALAIREMOIS3').AsFloat := EuroToFranc(Q.FindField('PHS_SALAIREMOIS3').AsFloat);
        Q.FindField('PHS_SALAIREANN3').AsFloat := EuroToFranc(Q.FindField('PHS_SALAIREANN3').AsFloat);
      end;
      if PgTypSalLib4 then
      begin
        Q.FindField('PHS_SALAIREMOIS4').AsFloat := EuroToFranc(Q.FindField('PHS_SALAIREMOIS4').AsFloat);
        Q.FindField('PHS_SALAIREANN4').AsFloat := EuroToFranc(Q.FindField('PHS_SALAIREANN4').AsFloat);
      end;
      Q.FindField('PHS_TOPCONVERT').AsString := '-';
    end;
  end else
    if NomT = 'PAIEENCOURS' then
  begin
    if Q.FindField('PPU_OCBRUT').AsFloat <> 0 then
    begin
      Q.FindField('PPU_CBRUT').AsFloat := Q.FindField('PPU_OCBRUT').AsFloat;
      Q.FindField('PPU_OCBRUT').AsFloat := 0;
    end;
    if Q.FindField('PPU_OCBRUTFISCAL').AsFloat <> 0 then
    begin
      Q.FindField('PPU_CBRUTFISCAL').AsFloat := Q.FindField('PPU_OCBRUTFISCAL').AsFloat;
      Q.FindField('PPU_OCBRUTFISCAL').AsFloat := 0;
    end;
    if Q.FindField('PPU_OCNETIMPOSAB').AsFloat <> 0 then
    begin
      Q.FindField('PPU_CNETIMPOSAB').AsFloat := Q.FindField('PPU_OCNETIMPOSAB').AsFloat;
      Q.FindField('PPU_OCNETIMPOSAB').AsFloat := 0;
    end;
    if Q.FindField('PPU_OCNETAPAYER').AsFloat <> 0 then
    begin
      Q.FindField('PPU_CNETAPAYER').AsFloat := Q.FindField('PPU_OCNETAPAYER').AsFloat;
      Q.FindField('PPU_OCNETAPAYER').AsFloat := 0;
    end;
    if Q.FindField('PPU_OCCOUTSALARIE').AsFloat <> 0 then
    begin
      Q.FindField('PPU_CCOUTSALARIE').AsFloat := Q.FindField('PPU_OCCOUTSALARIE').AsFloat;
      Q.FindField('PPU_OCCOUTSALARIE').AsFloat := 0;
    end;
    // PT5 : 30/10/2001 : V562 : PH Annulation migration ne reconvertit pas les plafonds et tranches et cout patronal
    if Q.FindField('PPU_OCCOUTPATRON').AsFloat <> 0 then
    begin
      Q.FindField('PPU_OCCOUTPATRON').AsFloat := Q.FindField('PPU_OCCOUTPATRON').AsFloat;
      Q.FindField('PPU_OCCOUTPATRON').AsFloat := 0;
    end;
    if Q.FindField('PPU_OCPLAFONDSS').AsFloat <> 0 then
    begin
      Q.FindField('PPU_CPLAFONDSS').AsFloat := Q.FindField('PPU_OCPLAFONDSS').AsFloat;
      Q.FindField('PPU_OCPLAFONDSS').AsFloat := 0;
    end;
    if Q.FindField('PPU_OCBASESS').AsFloat <> 0 then
    begin
      Q.FindField('PPU_CBASESS').AsFloat := Q.FindField('PPU_OCBASESS').AsFloat;
      Q.FindField('PPU_OCBASESS').AsFloat := 0;
    end;
  end else
    if NomT = 'SALARIES' then
  begin
    if Q.FindField('PSA_TOPCONVERT').AsString = 'X' then
    begin
      if PgTypSalLib1 then
      begin
        Q.FindField('PSA_SALAIREMOIS1').AsFloat := EuroToFranc(Q.FindField('PSA_SALAIREMOIS1').AsFloat);
        Q.FindField('PSA_SALAIRANN1').AsFloat := EuroToFranc(Q.FindField('PSA_SALAIRANN1').AsFloat);
      end;
      if PgTypSalLib2 then
      begin
        Q.FindField('PSA_SALAIREMOIS2').AsFloat := EuroToFranc(Q.FindField('PSA_SALAIREMOIS2').AsFloat);
        Q.FindField('PSA_SALAIRANN2').AsFloat := EuroToFranc(Q.FindField('PSA_SALAIRANN2').AsFloat);
      end;
      if PgTypSalLib3 then
      begin
        Q.FindField('PSA_SALAIREMOIS3').AsFloat := EuroToFranc(Q.FindField('PSA_SALAIREMOIS3').AsFloat);
        Q.FindField('PSA_SALAIRANN3').AsFloat := EuroToFranc(Q.FindField('PSA_SALAIRANN3').AsFloat);
      end;
      if PgTypSalLib4 then
      begin
        Q.FindField('PSA_SALAIREMOIS4').AsFloat := EuroToFranc(Q.FindField('PSA_SALAIREMOIS4').AsFloat);
        Q.FindField('PSA_SALAIRANN4').AsFloat := EuroToFranc(Q.FindField('PSA_SALAIRANN4').AsFloat);
      end;
      Q.FindField('PSA_SALAIRETHEO').AsFloat := EuroToFranc(Q.FindField('PSA_SALAIRETHEO').AsFloat);
      Q.FindField('PSA_TAUXHORAIRE').AsFloat := EuroToFranc(Q.FindField('PSA_TAUXHORAIRE').AsFloat);
      Q.FindField('PSA_SALAIREMULTI').AsFloat := EuroToFranc(Q.FindField('PSA_SALAIREMULTI').AsFloat);

      Q.FindField('PSA_TOPCONVERT').AsString := '-';
    end;
  end;
end;

// Affichage de la ligne indiquant la table en cours de traitement

procedure AfficheTitre(NomTable, Libel: string; PanelPGI: TPanel);
var
  NumT: integer;
begin
  if PanelPGI = nil then Exit;
  NumT := TableToNum(NomTable);
  if NumT <= 0 then Exit;
  PanelPGI.Caption := 'Bascule Paie & GRH : ' + NomTable + ' ' + Libel;
  Application.ProcessMessages;
end;
// Fonction d'encapsulation du traitement d'une table et affichage compteur

procedure ConvertTablePaiePGI(NomTable: string; PanelPGI: TPanel; Total: Integer; Sens: string);
var
  Q: TQuery;
  ii, nbre, r: integer;
begin
  AfficheTitre(NomTable, '', PanelPGI);
  nbre := 0;
  BeginTrans;
  Q := OpenSQL('SELECT * FROM ' + NomTable, False);
  ii := 0;
  while not Q.EOF do
  begin
    Q.Edit;
    if Sens = 'E' then ConvertEuroPG(Q, NomTable)
    else ConvertFrancPG(Q, NomTable);
    Q.Post;
    inc(ii);
    inc(nbre);
    if ii >= 500 then
    begin
      CommitTrans;
      BeginTrans;
      ii := 0;
    end;
    r := (nbre mod 10);
    if r = 0 then AfficheTitre(NomTable, ' Lus : ' + IntToStr(nbre) + ' sur ' + IntToStr(Total) + ' à traiter', PanelPGI);
    Q.Next;
  end;
  Ferme(Q);
  CommitTrans;
end;
// Fonction qui test les contenus des tables et qui lance la conversion

function BasculeEuroPaie(PanelPGI: TPanel): boolean;
var
  st: string;
  Q: TQuery;
  Nbre: Integer;
begin
  Nbre := 0;
  St := 'select count (*) from ABSENCESALARIE where PCN_TOPCONVERT="-"';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('ABSENCESALARIE', PanelPGI, Nbre, 'E');
  Nbre := 0;
  St := 'select count (*) from CONTRATTRAVAIL where PCI_TOPCONVERT="-"';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('CONTRATTRAVAIL', PanelPGI, Nbre, 'E');
  Nbre := 0;
  St := 'select count (*) from HISTOANALPAIE where PHA_OBASEREM=0 AND ' +
    ' PHA_OMTREM =0 AND PHA_OBASECOT=0 AND PHA_OMTSALARIAL=0 AND PHA_OMTPATRONAL=0';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('HISTOANALPAIE', PanelPGI, Nbre, 'E');
  Nbre := 0;
  St := 'select count (*) from HISTOBULLETIN where PHB_OBASEREM=0 AND PHB_OCOEFFREM=0 AND PHB_OMTREM=0 AND ' +
    ' PHB_OBASECOT =0 AND PHB_OTAUXSALARIAL=0 AND PHB_OMTSALARIAL=0 AND PHB_OMTPATRONAL=0';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('HISTOBULLETIN', PanelPGI, Nbre, 'E');
  Nbre := 0;
  St := 'select count (*) from HISTOCUMSAL where PHC_OMONTANT=0';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('HISTOCUMSAL', PanelPGI, Nbre, 'E');
  St := 'select count (*) from HISTOSAISRUB where PSD_TOPCONVERT="-"';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('HISTOSAISRUB', PanelPGI, Nbre, 'E');
  St := 'select count (*) from HISTOSALARIE where PHS_TOPCONVERT="-"';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('HISTOSALARIE', PanelPGI, Nbre, 'E');
  Nbre := 0;
  St := 'select count (*) from PAIEENCOURS where PPU_OCBRUT=0 AND PPU_OCBRUTFISCAL=0 AND PPU_OCNETIMPOSAB=0 AND ' +
    'PPU_OCNETAPAYER=0 AND PPU_OCCOUTSALARIE=0 AND PPU_OCCOUTPATRON=0 AND PPU_OCPLAFONDSS=0 AND PPU_OCBASESS=0';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('PAIEENCOURS', PanelPGI, Nbre, 'E');
  St := 'select count (*) from SALARIES where PSA_TOPCONVERT="-"';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('SALARIES', PanelPGI, Nbre, 'E');
  Result := True;
end;
// Fonction qui test les contenus des tables et qui lance l'annulation de la conversion en EURO

function BasculeFrancPaie(PanelPGI: TPanel): boolean;
var
  st: string;
  Q: TQuery;
  Nbre: Integer;
begin
  Nbre := 0;
  St := 'select count (*) from ABSENCESALARIE where PCN_TOPCONVERT="X"';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('ABSENCESALARIE', PanelPGI, Nbre, 'F');
  Nbre := 0;
  St := 'select count (*) from CONTRATTRAVAIL where PCI_TOPCONVERT="X"';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('CONTRATTRAVAIL', PanelPGI, Nbre, 'F');
  Nbre := 0;
  St := 'select count (*) from HISTOANALPAIE where PHA_OBASEREM<>0 OR ' +
    ' PHA_OMTREM <>0 OR PHA_OBASECOT<>0 OR PHA_OMTSALARIAL<>0 OR PHA_OMTPATRONAL<>0';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('HISTOANALPAIE', PanelPGI, Nbre, 'F');
  Nbre := 0;
  St := 'select count (*) from HISTOBULLETIN where PHB_OBASEREM<>0 OR PHB_OCOEFFREM<>0 AND PHB_OMTREM<>0 OR ' +
    ' PHB_OBASECOT<>0 OR PHB_OTAUXSALARIAL<>0 OR PHB_OMTSALARIAL<>0 AND PHB_OMTPATRONAL<>0';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('HISTOBULLETIN', PanelPGI, Nbre, 'F');
  Nbre := 0;
  St := 'select count (*) from HISTOCUMSAL where PHC_OMONTANT<>0';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('HISTOCUMSAL', PanelPGI, Nbre, 'F');
  St := 'select count (*) from HISTOSAISRUB where PSD_TOPCONVERT="X"';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('HISTOSAISRUB', PanelPGI, Nbre, 'F');
  St := 'select count (*) from HISTOSALARIE where PHS_TOPCONVERT="X"';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('HISTOSALARIE', PanelPGI, Nbre, 'F');
  Nbre := 0;
  St := 'select count (*) from PAIEENCOURS where PPU_OCBRUT<>0 OR PPU_OCBRUTFISCAL<>0 OR PPU_OCNETIMPOSAB<>0 OR ' +
    'PPU_OCNETAPAYER<>0 OR PPU_OCCOUTSALARIE<>0 OR PPU_OCCOUTPATRON<>0 OR PPU_OCPLAFONDSS<>0 OR PPU_OCBASESS<>0';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('PAIEENCOURS', PanelPGI, Nbre, 'F');
  St := 'select count (*) from SALARIES where PSA_TOPCONVERT="X"';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  if Nbre > 0 then ConvertTablePaiePGI('SALARIES', PanelPGI, Nbre, 'F');
  Result := True;
end;

// Lancement de la bascule et mémorisation des infos nécessaires au traitement

function BasculePaieEuro(StDev: string; PanelPGI: TPanel; Sens: string): Boolean;
var
  st: string;
  Q: TQuery;
  SystemTime0: TSystemTime;
begin
  PgTypSalLib1 := (GetParamSoc('SO_PGTYPSALLIB1')); // Si elemnts variables fiches salaries sont des montants = X
  PgTypSalLib2 := (GetParamSoc('SO_PGTYPSALLIB2')); // Si elemnts variables fiches salaries sont des montants = X
  PgTypSalLib3 := (GetParamSoc('SO_PGTYPSALLIB3')); // Si elemnts variables fiches salaries sont des montants = X
  PgTypSalLib4 := (GetParamSoc('SO_PGTYPSALLIB4')); // Si elemnts variables fiches salaries sont des montants = X

  Tob_Cum := TOB.Create('Les Cumuls', nil, -1);
  Tob_Cum.LoadDetailDB('CUMULPAIE', '', '', nil, FALSE, False);
  Tob_Cot := TOB.Create('Les Cotisations', nil, -1);
  St := 'SELECT * FROM COTISATION WHERE ##PCT_PREDEFINI## PCT_NATURERUB="COT" OR PCT_NATURERUB="BAS"';
  Q := OpenSql(st, TRUE);
  TOB_Cot.LoadDetailDB('COTISATION', '', '', Q, FALSE, FALSE);
  Ferme(Q);
  TOB_Rem := TOB.Create('Les Remunerations', nil, -1);
  TOB_Rem.LoadDetailDB('REMUNERATION', '', '', nil, FALSE, False);

  OldDev := StDev;
  if PanelPGI <> nil then
  begin
    PanelPGI.Visible := True;
    if Sens = 'E' then PanelPGI.Caption := 'Basculement en EURO Paie & GRH'
    else PanelPGI.Caption := 'Annulation du basculement en EURO Paie & GRH';
    PanelPGI.BringToFront;
    Application.ProcessMessages;
  end;
  if Sens = 'E' then
  begin
    Result := BasculeEuroPaie(PanelPGI);
    if result then
    begin
      GetLocalTime(SystemTime0);
      SetParamSoc('SO_PGTENUEEURO', 'X');
      SetParamSoc('SO_PGDATEBASCULEURO', SystemTimeToDateTime(SystemTime0));
      //    PT-2 : 01/08/2001 : V547 Rechargement des paramsoc
      ChargeParamsPaie;
    end;
  end
  else
  begin
    Result := BasculeFrancPaie(PanelPGI);
    if result then
    begin
      SetParamSoc('SO_PGTENUEEURO', '-');
      SetParamSoc('SO_PGDATEBASCULEURO', Idate1900);
      ChargeParamsPaie;
    end;
  end;
  if PanelPGI <> nil then PanelPGI.Visible := False;
  if Tob_Cum <> nil then
  begin
    Tob_Cum.Free;
    Tob_Cum := nil;
  end;
  if TOB_Rem <> nil then
  begin
    TOB_Rem.Free;
    TOB_Rem := nil;
  end;
  if Tob_Cot <> nil then
  begin
    Tob_Cot.Free;
    Tob_Cot := nil;
  end;
end;

end.

