{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Unit d'intgration des données provenant soit de la saisie par
Suite ........ : rubrique,des acomptes, des lignes de commentaires
Mots clefs ... : PAIE;PGBULLETIN
*****************************************************************}
{
PT1 : 20/12/2001 PH V571 Recupèration des infos de la saisie par rubrique par rapport
      au début et de fin mois et non par rapport à la date de paie qui peut contenir la
      date d'entrée ou date de sortie
PT2 : 27/03/2002 PH V571 Recupèration de la saisie par rubrique si au moins une valeur renseignée
PT3 : 19/06/2002 PH V585 fonctions d'intégration de la saisie par rubrique dans un bulletin complémentaire
PT4 : 01/12/2003 PH FQ 10996 V620 Prise en compte en saisie par rubrique des lignes de type MLB provenant fichier
                         import
PT5 : 05/12/2003 PH FQ 10996 V620 Prise en compte du libellé de la rubrique provenant du fichier d'import
PT6 : 05/12/2003 PH V620 Modif pour traitement des lignes de la saisie par rubrique impliquant un calcul de la paie à l'envers
PT7 : 07/01/2004 PH V_50 Prise en compte des rémunérations de type elt permanent dans la saisie par rubrique
PT8 : 22/03/2004 PH V_50 Prise en compte des lignes saisie des primes/participation
PT9 : 13/12/2004 PH V_60 Prise en compte des lignes de commentaires issues fichier import
}
unit P5RecupInfos;

interface

uses Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls,
{$IFNDEF EAGLCLIENT}
  DBTables, M3Code, Fe_Main,
{$ELSE}
  MaineAGL,
{$ENDIF}
  uPaieRemunerations,
  Hent1, HCtrls, ComCtrls, HRichEdt, HRichOLE, HMsgBox, HStatus,
  ExtCtrls, Grids, ImgList, Mask, ClipBrd, richedit, UTOB, P5Def, M3VM,
  ParamSoc, EntPaie;


procedure RecupereHistoSaisRub(Tob_PSD: tob; DateD, Datef: tdatetime; Salarie: string);
procedure IntegreHistoSaisrub(Tob_Rub, Tob_PSD: tob; DateD, DateF: tdatetime; Salarie, Etab, Action: string);
procedure IntegreHistoSaisrubACP(Tob_Rub, Tob_PSD: tob; DateD, DateF: tdatetime; Salarie, Etab, Action: string);
// PT5 : 05/12/2003 PH V620 Prise en compte du libellé de la rubrique provenant du fichier d'import
procedure IntegreRubHistoSaisRub(tob_rub, TRC: tob; OrigineMvt, ARub: string; DateD, DateF: tdatetime; SBase, Stx, SCoef, SMt: double; Libelle: string = '');
procedure IntegreRubHistoSaisRubACP(tob_rub, TRC: tob; OrigineMvt, ARub: string; DateD, DateF: tdatetime; SMt: double);
procedure ChargeProfilSPR(Salarie, TPE: TOB; Rub, Natrub: string);
procedure IntegreRubCommentaire(tob_rub, T, T_sal: tob; OrigineMvt, Arub, Salarie, Etab: string; DateD, DateF: tdatetime; Libelle: string = '');
function RecupereNouvellesHistoSaisRub(Tob_PSD: tob; DateD, Datef: tdatetime; Salarie: string): boolean;
procedure SupprRubCommentaire(tob_rub, T, T_sal: tob; OrigineMvt, Arub, Salarie, Etab: string; DateD, DateF: tdatetime);

function RecupereHistoSaisprim(Tob_PSP: tob; DateD, Datef: tdatetime; Salarie: string; ActionB: TActionBulletin): boolean;
procedure IntegreHistoSaisPrim(Tob_Rub, Tob_PSP: tob; DateD, DateF: tdatetime; Salarie, Etab, Action: string);

implementation

uses P5Util;
// Fonction intrégration de la saisie par rubrique

procedure IntegreHistoSaisrub(Tob_Rub, Tob_PSD: tob; DateD, DateF: tdatetime; Salarie, Etab, Action: string);
var
  t, TRC: tob;
  OrigineMvt, OrigineMvtP, ARub, ARubP: string;
  SBase, Stx, SCoef, SMt: double;
  AncLib: string;
begin
  if TOB_PSD = nil then exit;
  t := Tob_PSD.findfirst([''], [''], false);
  if t = nil then exit;

  OrigineMvt := t.getvalue('PSD_ORIGINEMVT');
  OrigineMvtP := OrigineMvt;
  SBase := 0;
  Stx := 0;
  SCoef := 0;
  SMt := 0;
  Arub := T.getvalue('PSD_RUBRIQUE');
  ArubP := ARub;
  SupprRubCommentaire(tob_rub, T, Tob_Salarie, ORIGINEMVT, Arub, Salarie, Etab, DateD, DateF);
  if (OrigineMvt <> 'MLB') and (OrigineMvt <> 'ACP') then // ne doit jamais arriver !
    ChargeProfilSPR(Tob_Salarie, Tob_rub, ARub, 'AAA');
  while t <> nil do
  begin
    Arub := T.getvalue('PSD_RUBRIQUE');
    OrigineMvt := t.getvalue('PSD_ORIGINEMVT');
    if OrigineMvt = 'ACP' then
    begin
      t := Tob_PSD.findnext([''], [''], false);
      continue;
    end;
    if OrigineMvt = 'MLB' then
    begin // Cas d'une ligne de commentaire
      if (OrigineMvt <> OrigineMvtP) or (Arub <> ArubP) then
      begin
        // DEB PT9
        TRC := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [ARubP], FALSE);
        if OrigineMvtP <> 'MLB' then IntegreRubHistoSaisRub(tob_rub, TRC, OrigineMvtP, ArubP, DateD, DateF, SBase, Stx, SCoef, SMt, AncLib); //@@@@
        OrigineMvtP := OrigineMvt;
        ArubP := ARub;
        SBase := 0;
        Stx := 0;
        SCoef := 0;
        SMt := 0;
        // FIN PT9
        SupprRubCommentaire(tob_rub, T, Tob_Salarie, t.getvalue('PSD_ORIGINEMVT'), t.getvalue('PSD_RUBRIQUE'), Salarie, Etab, DateD, DateF);
      end;
      AncLib := '';
      IntegreRubCommentaire(tob_rub, T, Tob_salarie, OrigineMvt, Arub, Salarie, Etab, DateD, DateF);
      t.putvalue('PSD_DATEINTEGRAT', DateF);
      t := Tob_PSD.findnext([''], [''], false);
      continue;
    end;
    if (OrigineMvt <> OrigineMvtP) or (Arub <> ArubP) then
    begin
      TRC := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [ARubP], FALSE); // $$$$
      if TRC <> nil then
        if OrigineMvtP <> 'MLB' then
          // PT5 : 05/12/2003 PH V620 Prise en compte du libellé de la rubrique provenant du fichier d'import
          IntegreRubHistoSaisRub(tob_rub, TRC, OrigineMvtP, ArubP, DateD, DateF, SBase, Stx, SCoef, SMt, AncLib);
      if (AncLib <> '') and (OrigineMvt = 'SRB') then // ######
      begin
        IntegreRubCommentaire(tob_rub, T, Tob_salarie, OrigineMvtP, ArubP, Salarie, Etab, DateD, DateF, Anclib);
      end;
      AncLib := '';
      if OrigineMvt <> 'MLB' then
        ChargeProfilSPR(Tob_Salarie, Tob_rub, ARub, 'AAA');

      OrigineMvtP := OrigineMvt;
      ArubP := ARub;
      SBase := 0;
      Stx := 0;
      SCoef := 0;
      SMt := 0;
      SupprRubCommentaire(tob_rub, T, Tob_Salarie, t.getvalue('PSD_ORIGINEMVT'), t.getvalue('PSD_RUBRIQUE'), Salarie, Etab, DateD, DateF);
    end;
    SBase := SBase + T.getvalue('PSD_BASE');
    Stx := STx + T.getvalue('PSD_TAUX');
    SCoef := SCoef + T.getvalue('PSD_COEFF');
    SMt := SMt + T.getvalue('PSD_MONTANT');
    // PT5 : 05/12/2003 PH V620 Prise en compte du libellé de la rubrique provenant du fichier d'import
    if AncLib = '' then AncLib := T.GetValue('PSD_LIBELLE');
    t.putvalue('PSD_DATEINTEGRAT', DateF);
    t := Tob_PSD.findnext([''], [''], false);
  end;
  // Ecriture du dernier
  TRC := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [ARubP], FALSE); // $$$$
  if TRC <> nil then
    if OrigineMvtP <> 'MLB' then
      if OrigineMvtP <> 'ACP' then
      begin
        // PT5 : 05/12/2003 PH V620 Prise en compte du libellé de la rubrique provenant du fichier d'import
        IntegreRubHistoSaisRub(tob_rub, TRC, OrigineMvtP, ArubP, DateD, DateF, SBase, Stx, SCoef, SMt, AncLib);
        if (AncLib <> '') and (OrigineMvt = 'SRB') then // ######
        begin
          IntegreRubCommentaire(tob_rub, T, Tob_salarie, OrigineMvtP, ArubP, Salarie, Etab, DateD, DateF, Anclib);
        end;

      end;
end;

procedure ChargeProfilSPR(Salarie, TPE: TOB; Rub, Natrub: string);
var
  TRC, THB: TOB;
begin
  if Rub = '' then exit;
  TRC := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rub], FALSE); // $$$$
  if TRC = nil then exit; // Integration rubrique de remunératoin inexistante
  if Salarie = nil then exit;
  if TPE.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', Rub], FALSE) = nil then // $$$$
  begin
    THB := TOB.Create('HISTOBULLETIN', TPE, -1);
    RemplirHistoBulletin(THB, Salarie, TRC, nil, TPE.GetValue('PPU_DATEDEBUT'), TPE.GetValue('PPU_DATEFIN'));
  end;
  //  TRC := nil;
end;



//Alimentation des valeurs dans les rubriques d'absence
// PT5 : 05/12/2003 PH V620 Prise en compte du libellé de la rubrique provenant du fichier d'import

procedure IntegreRubHistoSaisRub(tob_rub, TRC: tob; OrigineMvt, ARub: string; DateD, DateF: tdatetime; SBase, Stx, SCoef, SMt: double; Libelle: string = '');
var
  THB: tob;
begin
  THB := Tob_rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', ARub], FALSE); // $$$$
  if THB <> nil then
  begin // PT7
    if (TRC.GetValue('PRM_TYPEBASE') = '00') or (TRC.GetValue('PRM_TYPEBASE') = '01') then
      THB.PutValue('PHB_BASEREM', SBase);
    if (TRC.GetValue('PRM_TYPETAUX') = '00') or (TRC.GetValue('PRM_TYPETAUX') = '01') then
      THB.PutValue('PHB_TAUXREM', STx);
    if (TRC.GetValue('PRM_TYPECOEFF') = '00') or (TRC.GetValue('PRM_TYPECOEFF') = '01') then
      THB.PutValue('PHB_COEFFREM', SCoef);
    if (TRC.GetValue('PRM_TYPEMONTANT') = '00') or (TRC.GetValue('PRM_TYPEMONTANT') = '01') then
      THB.PutValue('PHB_MTREM', SMt);
    // FIN PT7
    {Positionne indicateur pour indiquer la provenance de la ligne dans le bulletin
    ET indiquer que la saisie n'est pas possible car les champs sont précalculés}
    THB.PutValue('PHB_ORIGINELIGNE', 'SAL'); // Calculé par les absences
    THB.putValue('PHB_ORIGINEINFO', OrigineMvt); // Calculé par les absences
    if (THB.GetValue('PHB_LIBELLE') <> Libelle) and (Libelle <> '') and (OrigineMvt <> 'SRB') then // ######
      THB.PutValue('PHB_LIBELLE', Libelle); // Libellé pouvant être saisi
  end;
end;
// FIN PT5
//Rubriques d'acompte

procedure IntegreRubHistoSaisRubACP(tob_rub, TRC: tob; OrigineMvt, ARub: string; DateD, DateF: tdatetime; SMt: double);
var
  THB: tob;
begin
  THB := Tob_rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', ARub], FALSE); // $$$$
  if THB <> nil then
  begin
    THB.PutValue('PHB_MTREM', SMt);
    {Positionne indicateur pour indiquer la provenance de la ligne dans le bulletin
    ET indiquer que la saisie n'est pas possible car les champs sont précalculés}
    THB.PutValue('PHB_ORIGINELIGNE', 'SAL'); // Calculé par les absences
    THB.putValue('PHB_ORIGINEINFO', OrigineMvt); // Calculé par les absences
  end;
end;

procedure RecupereHistoSaisRub(Tob_PSD: tob; DateD, Datef: tdatetime; Salarie: string);
var
  st: string;
  Q: tQuery;
begin
  // PT1 : 20/12/2001 V571 PH Recupèration des infos par rapport au debut de mois ie dates saisie par rubrique
  st := 'SELECT * FROM HISTOSAISRUB LEFT JOIN REMUNERATION ON PRM_RUBRIQUE = PSD_RUBRIQUE AND ##PRM_PREDEFINI## WHERE PRM_THEMEREM <> "ZZZ" AND' +
    ' PSD_SALARIE = "' + Salarie + '"' +
    ' AND (PSD_DATEDEBUT >="' + usdatetime(DebutdeMois(Dated)) +
    '" AND PSD_DATEFIN <= "' + usdatetime(FindeMois(Datef)) + '")' +
    ' AND PSD_DATEINTEGRAT <="' + usdatetime(10) + '"' +
    // PT2 : 27/03/2002 PH V571 Recupèration de la saisie par rubrique si au moins une valeur renseignée
// PT4 : 01/12/2003 PH V620 Prise en compte en saisie par rubrique des lignes de type MLB provenant fichier import
  ' AND ((PSD_BASE <> 0 OR PSD_TAUX <> 0 OR PSD_COEFF <> 0 OR PSD_MONTANT <> 0) OR (PSD_ORIGINEMVT = "MLB"))' +
    ' ORDER BY PSD_ORIGINEMVT,PSD_RUBRIQUE';
  Q := OpenSql(st, TRUE);
  if not (Q.eof) then
    Tob_PSD.LoadDetailDB('HISTOSAISRUB', '', '', Q, False);
  Ferme(Q);
end;

function RecupereNouvellesHistoSaisRub(Tob_PSD: tob; DateD, Datef: tdatetime; Salarie: string): boolean;
var
  st: string;
  Q: tQuery;
  t: tob;
begin
  //PT6 : 05/12/2003 PH V620 on exclut les rubriques de themes net à payer
  st := 'SELECT * FROM HISTOSAISRUB LEFT JOIN REMUNERATION ON PRM_RUBRIQUE = PSD_RUBRIQUE AND ##PRM_PREDEFINI## WHERE PRM_THEMEREM <> "ZZZ" AND' +
    ' PSD_SALARIE = "' + Salarie + '"' +
    ' AND (PSD_DATEDEBUT >="' + usdatetime(Dated) +
    '" AND PSD_DATEFIN <= "' + usdatetime(Datef) + '")' +
    ' AND PSD_DATEINTEGRAT <="' + usdatetime(10) + '"' +
    // PT2 : 27/03/2002 PH V571 Recupèration de la saisie par rubrique si au moins une valeur renseignée
// PT4 : 01/12/2003 PH V620 Prise en compte en saisie par rubrique des lignes de type MLB provenant fichier import
  ' AND ((PSD_BASE <> 0 OR PSD_TAUX <> 0 OR PSD_COEFF <> 0 OR PSD_MONTANT <> 0)  OR (PSD_ORIGINEMVT = "MLB"))' +
    ' ORDER BY PSD_RUBRIQUE,PSD_ORIGINEMVT';
  Q := OpenSql(st, TRUE);
  if (Q.eof) then
    result := false else
  begin
    Ferme(Q);
    //PT6 : 05/12/2003 PH V620 on exclut les rubriques de themes net à payer
    st := 'SELECT * FROM HISTOSAISRUB LEFT JOIN REMUNERATION ON PRM_RUBRIQUE = PSD_RUBRIQUE AND ##PRM_PREDEFINI## WHERE PRM_THEMEREM <> "ZZZ" AND PSD_SALARIE = "' + Salarie + '"' +
      ' AND (PSD_DATEDEBUT >="' + usdatetime(Dated) + '" AND PSD_DATEFIN <= "' + usdatetime(Datef) + '")' +
      ' ORDER BY PSD_ORIGINEMVT,PSD_RUBRIQUE';
    Q := OpenSql(st, TRUE);
    Tob_PSD.LoadDetailDB('HISTOSAISRUB', '', '', Q, False);
    if notVide(Tob_PSD, true) then
    begin
      result := true;
      t := Tob_PSD.findfirst([''], [''], true);
      while t <> nil do
      begin
        t.putvalue('PSD_DATEINTEGRAT', 0);
        t := Tob_PSD.findnext([''], [''], true);
      end;
    end
    else
      result := false;
  end;
  Ferme(Q);
end;
{ Nouvelles fonctions intégration des lignes de la saisie des primes dans le bulletins
Attention, ces lignes seront intégrées uniquement dans un bulletin complémentaire
PT3 : 19/06/2002 PH V585 fonctions d'intégration de la saisie par rubrique dans un bulletin complémentaire
}

function RecupereHistoSaisprim(Tob_PSP: tob; DateD, Datef: tdatetime; Salarie: string; ActionB: TActionBulletin): boolean;
var
  st: string;
  Q: tQuery;
  t: tob;
begin
  result := false;
  if not VH_Paie.PgPrimIntegCompl then exit;
  if (ActionB <> taCreation) and (ActionB <> PremCreation) and (ActionB <> taModification) then exit;
  //PT6 : 05/12/2003 PH V620 on exclut les rubriques de themes net à payer
  // PT8 : 22/03/2004 PH V_50 Prise en compte des lignes saisie des primes/participation
  st := 'SELECT * FROM HISTOSAISPRIM LEFT JOIN REMUNERATION ON PRM_RUBRIQUE = PSP_RUBRIQUE AND ##PRM_PREDEFINI## WHERE PRM_THEMEREM <> "ZZZ" AND' +
    ' PSP_SALARIE = "' + Salarie + '" AND PSP_AREPORTER <> "NON"' + // FIN PT8
  ' AND (PSP_DATEDEBUT >="' + usdatetime(Dated) +
    '" AND PSP_DATEFIN <= "' + usdatetime(Datef) + '")' +
    ' AND PSP_DATEINTEGRAT <="' + usdatetime(10) + '"' +
    ' AND (PSP_BASE <> 0 OR PSP_TAUX <> 0 OR PSP_COEFF <> 0 OR PSP_MONTANT <> 0) ' +
    ' ORDER BY PSP_ORIGINEMVT,PSP_RUBRIQUE';
  Q := OpenSql(st, TRUE);
  if ActionB = taCreation then
  begin
    if not (Q.eof) then
      Tob_PSP.LoadDetailDB('HISTOSAISPRIM', '', '', Q, False);
    Ferme(Q);
  end
  else
  begin
    if (Q.eof) then
      result := false else
    begin
      Ferme(Q);
      //PT6 : 05/12/2003 PH V620 on exclut les rubriques de themes net à payer
      st := 'SELECT * FROM HISTOSAISPRIM LEFT JOIN REMUNERATION ON PRM_RUBRIQUE = PSP_RUBRIQUE AND ##PRM_PREDEFINI## WHERE PRM_THEMEREM <> "ZZZ" AND PSD_SALARIE = "' + Salarie + '"' +
        ' AND (PSP_DATEDEBUT >="' + usdatetime(Dated) + '" AND PSP_DATEFIN <= "' + usdatetime(Datef) + '")' +
        ' ORDER BY PSP_ORIGINEMVT,PSP_RUBRIQUE';
      Q := OpenSql(st, TRUE);
      Tob_PSD.LoadDetailDB('HISTOSAISPRIM', '', '', Q, False);
      if notVide(Tob_PSP, true) then
      begin
        result := true;
        t := Tob_PSP.findfirst([''], [''], true);
        while t <> nil do
        begin
          t.putvalue('PSP_DATEINTEGRAT', 0);
          t := Tob_PSP.findnext([''], [''], true);
        end;
      end
      else
        result := false;
    end;
    Ferme(Q);
  end;
end;
// Fonction intrégration de la saisie des primes dans un bulletin complémentaire

procedure IntegreHistoSaisPrim(Tob_Rub, Tob_PSP: tob; DateD, DateF: tdatetime; Salarie, Etab, Action: string);
var
  t, TRC: tob;
  OrigineMvt, OrigineMvtP, ARub, ARubP: string;
  SBase, Stx, SCoef, SMt: double;
begin
  if not VH_Paie.PgPrimIntegCompl then exit;
  if TOB_PSP = nil then exit;
  t := Tob_PSP.findfirst([''], [''], false);
  if t = nil then exit;

  OrigineMvt := t.getvalue('PSP_ORIGINEMVT');
  OrigineMvtP := OrigineMvt;
  SBase := 0;
  Stx := 0;
  SCoef := 0;
  SMt := 0;
  Arub := T.getvalue('PSP_RUBRIQUE');
  ArubP := ARub;
  // On recherche si la rubrique est presente dans le bulletin
  ChargeProfilSPR(Tob_Salarie, Tob_rub, ARub, 'AAA');
  while t <> nil do
  begin
    Arub := T.getvalue('PSP_RUBRIQUE');
    OrigineMvt := t.getvalue('PSP_ORIGINEMVT');
    if (OrigineMvt <> OrigineMvtP) or (Arub <> ArubP) then
    begin
      TRC := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [ARubP], FALSE); // $$$$
      if TRC <> nil then
        if OrigineMvtP <> 'MLB' then
          IntegreRubHistoSaisRub(tob_rub, TRC, OrigineMvtP, ArubP, DateD, DateF, SBase, Stx, SCoef, SMt);
      ChargeProfilSPR(Tob_Salarie, Tob_rub, ARub, 'AAA');
      OrigineMvtP := OrigineMvt;
      ArubP := ARub;
      SBase := 0;
      Stx := 0;
      SCoef := 0;
      SMt := 0;
    end;
    SBase := SBase + T.getvalue('PSP_BASE');
    Stx := STx + T.getvalue('PSP_TAUX');
    SCoef := SCoef + T.getvalue('PSP_COEFF');
    SMt := SMt + T.getvalue('PSP_MONTANT');
    t.putvalue('PSP_DATEINTEGRAT', DateF);
    t := Tob_PSP.findnext([''], [''], false);
  end;
  // Ecriture du dernier
  TRC := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [ARubP], FALSE); // $$$$
  if TRC <> nil then
    IntegreRubHistoSaisRub(tob_rub, TRC, OrigineMvtP, ArubP, DateD, DateF, SBase, Stx, SCoef, SMt);
end;
// FIN PT3 : 19/06/2002 PH V585 fonctions d'intégration de la saisie

procedure SupprRubCommentaire(tob_rub, T, T_sal: tob; OrigineMvt, Arub, Salarie, Etab: string; DateD, DateF: tdatetime);
var
  TW: tob;
begin
  // on supprime d'abord les lignes commenataires crées
  TW := Tob_Rub.FindFirst(['PHB_NATURERUB', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
    ['AAA', DateD, DateF, Salarie], TRUE);
  while TW <> nil do
  begin
    if ((copy(TW.GetValue('PHB_RUBRIQUE'), 1, length(ARub) + 1) = ARub + '.') and
      (TW.GetValue('PHB_ORIGINEINFO') = OrigineMvt)) then
      TW.free;
    TW := Tob_Rub.FindNext(['PHB_NATURERUB', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
      ['AAA', DateD, DateF, Salarie], TRUE);
  end;
end;
// Fonction intgration des lignes de commentaires

procedure IntegreRubCommentaire(tob_rub, T, T_sal: tob; OrigineMvt, Arub, Salarie, Etab: string; DateD, DateF: tdatetime; Libelle: string = '');
var
  Thh, TW, TR: tob;
  i, j: integer;
begin

  TW := Tob_Rub.FindFirst(['PHB_NATURERUB', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_ETABLISSEMENT', 'PHB_SALARIE'],
    ['AAA', DateD, DateF, Etab, Salarie], TRUE);
  i := 0;
  while TW <> nil do
  begin
    if copy(TW.GetValue('PHB_RUBRIQUE'), 1, length(ARub) + 1) = ARub + '.' then
    begin
      j := strtoint(copy(TW.GetValue('PHB_RUBRIQUE'), (length(ARub) + 2), 1));
      if j > i then i := j;
    end;

    TW := Tob_Rub.FindNext(['PHB_NATURERUB', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_ETABLISSEMENT', 'PHB_SALARIE'],
      ['AAA', DateD, DateF, Etab, CodSal], TRUE);
  end;
  if i < 9 then
    i := i + 1
  else
    exit;
  THH := TOB.create('HISTOBULLETIN', Tob_Rub, -1);
  THH.PutValue('PHB_RUBRIQUE', ARub + '.' + IntToStr(i));
  THH.PutValue('PHB_NATURERUB', 'AAA');
  THH.PutValue('PHB_DATEDEBUT', DateD);
  THH.PutValue('PHB_DATEFIN', DateF);
  THH.PutValue('PHB_ETABLISSEMENT', Etab);
  THH.PutValue('PHB_SALARIE', Salarie);
  THH.PutValue('PHB_BASEREM', 0);
  THH.PutValue('PHB_TAUXREM', 0);
  THH.PutValue('PHB_COEFFREM', 0);
  THH.PutValue('PHB_MTREM', 0);
  THH.PutValue('PHB_CONSERVATION', 'BUL');
  if Libelle = '' then THH.PutValue('PHB_LIBELLE', T.GetValue('PSD_LIBELLE'))
  else THH.PutValue('PHB_LIBELLE', Libelle); // ######
  THH.PutValue('PHB_ORIGINELIGNE', 'SAL'); // Calculé par les absences
  THH.putValue('PHB_ORIGINEINFO', OrigineMvt); // Calculé par les absences

  THH.PutValue('PHB_IMPRIMABLE', 'X');
  THH.PutValue('PHB_TRAVAILN2', t_Sal.GetValue('PSA_TRAVAILN2'));
  THH.PutValue('PHB_TRAVAILN3', t_Sal.GetValue('PSA_TRAVAILN3'));
  THH.PutValue('PHB_TRAVAILN4', t_Sal.GetValue('PSA_TRAVAILN4'));
  THH.PutValue('PHB_TRAVAILN1', t_Sal.GetValue('PSA_TRAVAILN1'));
  THH.PutValue('PHB_CODESTAT', t_Sal.GetValue('PSA_CODESTAT'));
  THH.PutValue('PHB_LIBREPCMB1', t_Sal.GetValue('PSA_LIBREPCMB1'));
  THH.PutValue('PHB_LIBREPCMB2', t_Sal.GetValue('PSA_LIBREPCMB2'));
  THH.PutValue('PHB_LIBREPCMB3', t_Sal.GetValue('PSA_LIBREPCMB3'));
  THH.PutValue('PHB_LIBREPCMB4', t_Sal.GetValue('PSA_LIBREPCMB4'));
  THH.PutValue('PHB_CONFIDENTIEL', t_Sal.GetValue('PSA_CONFIDENTIEL'));
  TR := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [ARub], FALSE);
  if TR <> nil then THH.PutValue('PHB_ORDREETAT', TR.GetValue('PRM_ORDREETAT'));
end;


// integre rubrique acompte

procedure IntegreHistoSaisrubACP(Tob_Rub, Tob_PSD: tob; DateD, DateF: tdatetime; Salarie, Etab, Action: string);
var
  t, TRC: tob;
  OrigineMvt, ARub, ARubP: string;
  SMt: double;
  DatePaiement: tdatetime;
begin
  if TOB_PSD = nil then exit;
  t := Tob_PSD.findfirst(['PSD_ORIGINEMVT'], ['ACP'], false);
  if t = nil then exit;
  while t <> nil do
  begin
    DatePaiement := T.GetValue('PSD_DATEPAIEMENT');
    if DatePaiement >= DateD then break;
    t := Tob_PSD.findnext(['PSD_ORIGINEMVT'], ['ACP'], false);
  end;
  if t = nil then exit;
  SMt := 0;
  Arub := T.getvalue('PSD_RUBRIQUE');
  ArubP := ARub;
  SupprRubCommentaire(tob_rub, T, Tob_Salarie, t.getvalue('PSD_ORIGINEMVT'), Arub, Salarie, Etab, DateD, DateF);
  ChargeProfilSPR(Tob_Salarie, Tob_rub, ARub, 'AAA');
  while t <> nil do
  begin
    DatePaiement := T.GetValue('PSD_DATEPAIEMENT');
    if DatePaiement < DateD then
    begin
      t := Tob_PSD.findnext(['PSD_ORIGINEMVT'], ['ACP'], false);
      continue;
    end;
    Arub := T.getvalue('PSD_RUBRIQUE');
    OrigineMvt := t.getvalue('PSD_ORIGINEMVT');
    IntegreRubCommentaire(tob_rub, T, Tob_Salarie, OrigineMvt, Arub, Salarie, Etab, DateD, DateF);
    t.putvalue('PSD_DATEINTEGRAT', DateF);
    if Arub = Arubp then
    begin
      sMt := smt + t.getvalue('PSD_MONTANT');
      t := Tob_PSD.findnext(['PSD_ORIGINEMVT'], ['ACP'], false);
      continue;
    end;

    TRC := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [ARubP], FALSE); // $$$$
    if TRC <> nil then
      IntegreRubHistoSaisRubACP(tob_rub, TRC, 'ACP', ArubP, DateD, DateF, SMt);
    ChargeProfilSPR(Tob_Salarie, Tob_rub, ARub, 'AAA');
    SMt := 0;
    ArubP := ARub;
    t := Tob_PSD.findnext(['PSD_ORIGINEMVT'], ['ACP'], false);
    SupprRubCommentaire(tob_rub, T, Tob_Salarie, t.getvalue('PSD_ORIGINEMVT'), T.getvalue('PSD_RUBRIQUE'), Salarie, Etab, DateD, DateF);
  end;
  TRC := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [ARubP], FALSE); // $$$$
  if TRC <> nil then
    IntegreRubHistoSaisRubACP(tob_rub, TRC, 'ACP', ArubP, DateD, DateF, SMt);

end;

end.

