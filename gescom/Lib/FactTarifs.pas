unit FactTarifs;

interface

uses
  uTob,
  {$IFNDEF EAGLCLIENT}
    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  {$IFDEF BTP}
  UtilSuggestion,
  {$ENDIF}
  hCtrls,
  hEnt1,
  SaisUtil,
  EntGC,
  FactUtil
  ,uEntCommun
  ;

function TrouveTarif(TOBA, TOBTiers, TOBL, TOBPiece, TOBTarif: TOB; ForceTous, EnHT: boolean): boolean; //JD
function TarifVersLigne(TobA, TobTiers, TobL, TobLigneTarif, TobPiece, TobTarif: Tob; ForceTous, Totale: boolean; DEV: RDevise; CtrlCatFrs : Boolean = True): T_ActionTarifArt;

//Traitement sur TobLigneTarif
procedure LoadLesLignesTarifs(TobLigneTarif: Tob; CleDoc: R_CleDoc);
procedure LoadLigneTarif(TobLigneTarif, TobL: Tob);
procedure SupprimeLigneTarif(TobPiece, TobLigneTarif: Tob; ARow: Integer);
function CreerLigneTarifMere(TobLigneTarif: Tob; ARow: Integer; CleDoc: R_CleDoc; NumLigne: Integer): TOB; overload;
function CreerLigneTarifMere(TobLigneTarif: Tob; ARow: Integer; TobL: Tob; NumLigne: Integer): Tob; overload;
function GetLigneTarif(TobLigneTarif, TobLigne: Tob): Tob;

procedure InsertTobLigneTarif(TOBLigneTarif: TOB; ARow: integer; CleDoc: R_CleDoc);
procedure NumeroteLigneTarif(TOBLigneTarif: TOB);
procedure DeleteTobLigneTarif(TobPiece: Tob; NumLigne: Integer); overload;
procedure DeleteTobLigneTarif(TobPiece: Tob; Tobl: Tob); overload;

implementation

uses
  ParamSoc,
  FactTob,
  UtilTarif,
  UtilPGI,
  Tarifs;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 19/01/2000
Modifié le ... : 19/01/2000
Description .. : Recherche et alimentation d'un tarif depuis saisie
Mots clefs ... : ARTICLE;SAISIE;RECHERCHE;
*****************************************************************}
function TrouveTarif(TOBA, TOBTiers, TOBL, TOBPiece, TOBTarif: TOB; ForceTous, EnHT: boolean): boolean;
begin
  Result := False;
  if TOBA = nil then Exit;
  if ((VH_GC.GCOuvreTarifQte) and (not ForceTous) and (False)) then
  begin
    Result := ChercheTarif(TOBA, TOBTiers, TOBL, TOBPiece, TOBTarif, True, EnHT);
    if Result then ; // Affichage du choix
  end else
  begin
    Result := ChercheTarif(TOBA, TOBTiers, TOBL, TOBPiece, TOBTarif, False, EnHT);
    TOBL.PutValue('RECALCULTARIF', 'X');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 15/09/2003
Modifié le ... : 15/09/2003
Description .. :
Suite ........ : Initilaisation des lignes pièce à partir du système tarifaire
Mots clefs ... : LIGNES/TARIFS
*****************************************************************}
function TarifVersLigne(TobA, TobTiers, TobL, TobLigneTarif, TobPiece, TobTarif: Tob; ForceTous, Totale: boolean; DEV: RDEVISE; CTRLCatFrs : Boolean = True): T_ActionTarifArt;
var
  Rem, Px: Double;
  nFixe, nRemiseEnMontant,CoefUaUs: Double;
  CodeDT, CodeDP: string;
  EnHT: Boolean;
  //Ua : string;
  {$IFDEF BTP}
  // PxAch: Double;
  // InfoTarFou : TGinfostarifFour;
  {$ENDIF}
begin
	CoefUaUs := 0;
  Result := ataOk;
  CodeDP := TOBPiece.GetValue('GP_DEVISE');
  EnHT := (TOBPiece.GetValue('GP_FACTUREHT') = 'X');
  
  if GetInfoParPiece(TOBL.GetValue('GL_NATUREPIECEG'), 'GPP_CONDITIONTARIF') = 'X' then
  begin
    if GetParamSoc('SO_PREFSYSTTARIF') then
    begin
      nFixe := 0;
      Px := 0;
      Rem := 0;
      nRemiseEnMontant := 0;
      { Recherche du tarif marchandise }
      if (RechercheEnregistreTarifs(RechercheFonctionnalite(TobL.GetValue('GL_NATUREPIECEG')), 'LIGNE', '1', TobTiers, TobA, TobPiece, TobL, TobLigneTarif, nFixe, Px, Rem, nRemiseEnMontant, CodeDT)) then
      begin
        TobL.PutValue('RECALCULTARIF', 'X');
      end
      else
      begin
        { Si pas de tarif particulier alors on prend les prix dans la ligne de pièce }
        CodeDT := V_PGI.DevisePivot;
        Rem := 0;
        if EnHT then Px := TobL.GetValue('GL_PUHTDEV') else Px := TobL.GetValue('GL_PUTTCDEV');
      end;
    end
    else if TrouveTarif(TOBA, TOBTiers, TOBL, TOBPiece, TOBTarif, ForceTous, EnHT) then
    begin
      CodeDT := TOBTarif.GetValue('GF_DEVISE');
      Px := TOBTarif.GetValue('GF_PRIXUNITAIRE');
      {$IFDEF BTP}
      if TOBPiece.GetValue('GP_VENTEACHAT') = 'ACH' then
      begin
(*
				InfoTarFou := GetInfosTarifAch (TOBTiers.getValue('T_TIERS'),'',TurAchat,true,true,TOBA,TOBL.GetValue('GL_QTEFACT'),TOBL.GetValue('GL_DATEPIECE'));
        if InfoTarFou.Tarif <> 0 then
        begin
        	Px := InfoTarFou.Tarif;
          TOBL.PutValue ('FROMTARIF','X');
          TOBL.PutValue('GL_DPA',Px);
        end else
        begin
        	Px := TOBL.GetValue('GL_PUHTDEV');
        end;
*)
        //
        //fv1 : 22/05/2014 - FS#1008 - DELABOUDINIERE : Ajout contrôle référencement fournisseur en saisie de pièces d'achats
        if CTRLCatFrs then
          Px := RechTarifArticle(TOBA, TOBTarif, TOBTiers, TOBPIECE, TOBL)
        else
          Px := RechPrixArticle(TOBA, TobTarif, Tobtiers, TobPiece, TOBL,TOBPiece.GetValue('GP_VENTEACHAT'));
        //
      end
      else
      BEGIN
      {$ENDIF}
        Px := TOBTarif.GetValue('GF_PRIXUNITAIRE');
      {$IFDEF BTP}
      	TOBL.PutValue ('FROMTARIF','X');
      	END;
      {$ENDIF}
      // --
      Rem := TOBTarif.GetValue('GF_REMISE');
      if Px = 0 then
      begin
        if EnHT then Px := TOBL.GetValue('GL_PUHTDEV') else Px := TOBL.GetValue('GL_PUTTCDEV');
      end;
    end else
    begin
      Rem := 0;
      CodeDT := V_PGI.DevisePivot;
      {$IFDEF BTP}
      if TOBPiece.GetValue('GP_VENTEACHAT') = 'ACH' then
      begin
        //fv1 : 22/05/2014 - FS#1008 - DELABOUDINIERE : Ajout contrôle référencement fournisseur en saisie de pièces d'achats
        if CTRLCatFrs then
          Px := RechTarifArticle(TOBA, TOBTarif, TOBTiers, TOBPIECE, TOBL)
        else
          Px := RechPrixArticle(TOBA, TobTarif, Tobtiers, TobPiece, TOBL,TOBPiece.GetValue('GP_VENTEACHAT'));
        //
        {*
      	Ua := '';
        PxAch := RecupTarifAch(TOBA, TOBTarif, TOBTiers, Ua, CoefUaUs, TurAchat, false,true,TOBL.GetValue('GL_QTEFACT'),TOBL.GetValue('GL_DATEPIECE'));
        if PxAch <> 0 then
        BEGIN
        	Px := PxAch;
          TOBL.PutValue ('FROMTARIF','X');
          TOBL.PutValue('GL_DPA',Px);
          if Ua <> '' then TOBL.putValue('GL_QUALIFQTEVTE',UA);
        END else
          if UA = 'PASCATALOGUE' then
            PX := 0
          Else
            Px := TOBL.GetValue('GL_PUHTDEV');
        *}
        Rem := TOBTarif.GetValue('GF_REMISE');
      end else
        {$ENDIF}
        if EnHT then Px := TOBL.GetValue('GL_PUHTDEV') else Px := TOBL.GetValue('GL_PUTTCDEV');
    end;
  end else
  begin
    Rem := 0;
    CodeDT := V_PGI.DevisePivot;

    {$IFDEF BTP}
    //fv1 : 22/05/2014 - FS#1008 - DELABOUDINIERE : Ajout contrôle référencement fournisseur en saisie de pièces d'achats
    if CTRLCatFrs then
      Px := RechTarifArticle(TOBA, TOBTarif, TOBTiers,TOBPIECE, TOBL)
    else
      Px := RechPrixArticle(TOBA, TobTarif, Tobtiers, TobPiece, TOBL, TOBPiece.GetValue('GP_VENTEACHAT'));
    {$ENDIF}

    if EnHT then Px := TOBL.GetValue('GL_PUHTDEV') else Px := TOBL.GetValue('GL_PUTTCDEV');

  end;
  if CodeDT <> V_PGI.DevisePivot then // Tarif et saisie en devise
  begin
    if EnHT then
    begin
      TOBL.PutValue('GL_PUHTDEV', Px);
      TOBL.PutValue('GL_PUHT', DeviseToPivotEx(Px, DEV.Taux, DEV.Quotite,V_PGI.okdecP));
      TOBL.PutValue ('GL_COEFMARG',0);
    end else
    begin
      TOBL.PutValue('GL_PUTTCDEV', Px);
      TOBL.PutValue('GL_PUTTC', DeviseToPivotEx(Px, DEV.Taux, DEV.Quotite,V_PGI.okdecP));
      TOBL.PutValue ('GL_COEFMARG',0);
    end;
  end else
  begin
    if CodeDT = CodeDP then // Tarif et saisie en pivot
    begin
      if EnHT then
      begin
        if (TOBL.GetDouble('GL_PUHTDEV')<> Px) then
        begin
          TOBL.PutValue('GL_PUHTDEV', Px);
          TOBL.PutValue('GL_PUHT', Px);
          TOBL.PutValue ('GL_COEFMARG',0);
        end;
      end else
      begin
        if (TOBL.GetDouble('GL_PUTTCDEV')<> Px) then
        begin
          TOBL.PutValue('GL_PUTTCDEV', Px);
          TOBL.PutValue('GL_PUTTC', Px);
          TOBL.PutValue ('GL_COEFMARG',0);
        end;
      end;
    end else // Tarif en franc et saisie devise
    begin
      if VH_GC.GCComportePrixDev = 'ALE' then
      begin
        Result := ataAlerte;
        if EnHT then
        begin
          TOBL.PutValue('GL_PUHT', 0);
          TOBL.PutValue('GL_PUHTDEV', 0);
        end else
        begin
          TOBL.PutValue('GL_PUTTC', 0);
          TOBL.PutValue('GL_PUTTCDEV', 0);
        end;
      end else if VH_GC.GCComportePrixDev = 'BLO' then
      begin
        Result := ataCancel;
        if EnHT then
        begin
          TOBL.PutValue('GL_PUHT', 0);
          TOBL.PutValue('GL_PUHTDEV', 0);
        end else
        begin
          TOBL.PutValue('GL_PUTTC', 0);
          TOBL.PutValue('GL_PUTTCDEV', 0);
        end;
      end else
      begin
        Result := ataOk;
        if EnHT then
        begin
          TOBL.PutValue('GL_PUHT', Px);
          // MODIF MODE LM
          Px := Px * V_PGI.TauxEuro;
          TOBL.PutValue('GL_PUHTDEV', PivotToDevise(Px, DEV.Taux, DEV.Quotite,V_PGI.okdecP));
      	  TOBL.PutValue ('GL_COEFMARG',0);
        end else
        begin
          TOBL.PutValue('GL_PUTTC', Px);
          // MODIF MODE LM
          Px := Px * V_PGI.TauxEuro;
          TOBL.PutValue('GL_PUTTCDEV', PivotToDevise(Px, DEV.Taux, DEV.Quotite, V_PGI.okdecP));
      	  TOBL.PutValue ('GL_COEFMARG',0);
        end;
        // MODIF MODE LM
        TOBL.PutValue('RECALCULTARIF', '-');
      end;
    end;
  end;
  // MODIF MODE LM
  //TOBL.PutValue('RECALCULTARIF','-') ;

  { Remarque en recherche de tarif : il faut remettre à 0 si on ne trouve pas }
  if ((Result = ataOk) and (nRemiseEnMontant<>0)) then
  begin
    if Totale then
      TOBL.PutValue('GL_VALEURREMDEV', nRemiseEnMontant);
  end;

  if ((Result = ataOk) and (nFixe<>0)) then
  begin
    if Totale then
      TOBL.PutValue('GL_VALEURFIXEDEV', nFixe);
  end;

  // reinitialisation des remises pour le cas ou le fournisseur n'en a pas (de remise bien sur)
  TOBL.PutValue('GL_REMISELIGNE', 0);
  TOBL.PutValue('GL_REMISECASCADE', '');

  if ((Result = ataOk) and (Rem <> 0) and (TOBL.GetValue('GL_REMISABLELIGNE')='X') ) then
  begin
    if (Totale) then
    begin
      TOBL.PutValue('GL_REMISELIGNE',   Rem);
      TOBL.PutValue('GL_REMISECASCADE', TOBTarif.GetValue('GF_CASCADEREMISE'));
    end;
  end;

  if Result = ataOk then TOBL.PutValue('GL_TARIF', TOBTarif.GetValue('GF_TARIF'))
  else TOBL.PutValue('GL_TARIF', 0);

  if (Result = ataOk) and (ctxMode in V_PGI.PGIContexte) then
  begin
    if (ctxFO in V_PGI.PGIContexte) then
    begin
      if Totale then TOBL.PutValue('GL_CODEARRONDI', TOBTarif.GetValue('GF_ARRONDI'));
      if TOBL.GetValue('GL_NATUREPIECEG') = 'FFO' then
      begin
        if TOBTarif.FieldExists('PRIXBASE') then
        begin
          if TOBTarif.GetValue('PRIXBASE') = 0 then
            if EnHT then TOBTarif.PutValue('PRIXBASE', TOBA.GetValue('GA_PVHT')) else
              TOBTarif.PutValue('PRIXBASE', TOBA.GetValue('GA_PVTTC')); // Prix de base pour le front
          if EnHT then TOBL.PutValue('GL_PUHTBASE', TOBTarif.GetValue('PRIXBASE')) else TOBL.PutValue('GL_PUTTCBASE', TOBTarif.GetValue('PRIXBASE'));
        end;
      end;
    end;
    if Totale then TOBL.PutValue('GL_TYPEREMISE', TOBTarif.GetValue('GF_DEMARQUE'));
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 02/10/2002
Description .. : Charge les lignes de LIGNETARIF pour une pièce.
Suite ........ : TobLigneTarif est structurée :
Suite ........ : Mere
Suite ........ :     LIGNE1               (Table LIGNE)
Suite ........ :          LIGNETARIF1     (Table LIGNETARIF)
Suite ........ :          LIGNETARIF2
Suite ........ :          LIGNETARIF3
Suite ........ :     LIGNE2
Suite ........ :          LIGNETARIF1
*****************************************************************}
procedure LoadLesLignesTarifs(TobLigneTarif: TOB; CleDoc: R_CleDoc);
{
var
  Q: tQuery;
  TheTob, TheTobMere: tob;
  iLig: integer;
  TobPieceLocal: Tob;
}
begin
(* SUPPRIME -- SUITE A DROPTABLE DANS PGIMAJVER
  TheTob := Tob.Create('_LIGNETARIF_', nil, -1);
  try
    Q := OpenSQL('SELECT * FROM LIGNETARIF WHERE ' + WherePiece(CleDoc, ttdLigneTarif, False), True);
    try
      TheTob.LoadDetailDB('LIGNETARIF', '', '', Q, False, True);
      if (TheTob.Detail.count <> 0) then
      begin
        TheTob.Detail.Sort('GLT_NUMLIGNE; GLT_FONCTIONNALITE; GLT_RANG');
        TobLigneTarif.ClearDetail;
        iLig := -1;
        TheTobMere := nil;
        while (TheTob.Detail.Count > 0) do
        begin
          if (iLig <> TheTob.Detail[0].GetValue('GLT_NUMLIGNE')) then
          begin
            iLig := TheTob.Detail[0].GetValue('GLT_NUMLIGNE');
            TheTobMere := CreerLigneTarifMere(TobLigneTarif, -1, CleDoc, iLig);
          end;
          if (TheTobMere <> nil) then
            TheTob.Detail[0].ChangeParent(TheTobMere, -1);
        end;
      end;
    finally
      Ferme(Q);
    end;
  finally
    TheTob.Free;
  end;
  { Synchronise la TobPiece et TobLigneTarif }
  TobPieceLocal := GetTobPieceInTobLigneTarif(TobLigneTarif);
  for iLig := 0 to TobPieceLocal.Detail.Count - 1 do
  begin
    TheTob := GetLigneTarif(TobLigneTarif, TobPieceLocal.Detail[iLig]);
    if TheTob = nil then
      CreerLigneTarifMere(TobLigneTarif, iLig, TobPieceLocal.Detail[iLig], iLig + 1);
  end;
*)
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 15/09/2003
Modifié le ... :   /  /
Description .. : Chargement de LIGNETARIF pour une ligne de pièce
Mots clefs ... : LIGNES/TARIFS/LIGNETARIF
*****************************************************************}
procedure LoadLigneTarif(TobLigneTarif, TobL: Tob);
{
var
  T: Tob;
  NumL: Integer;
  Q: TQuery;
  CleDoc: R_CleDoc;
}
begin
(* SUPPRIME -- SUITE A DROPTABLE DANS PGIMAJVER
  if (TobLigneTarif <> nil) and (TobL <> nil) and (Tobl.GetValue('GL_TYPELIGNE') = 'ART') then
  begin
    NumL := TobL.GetValue('GL_NUMLIGNE');
    CleDoc := TOB2CleDoc(TobL);
    CleDoc.NumLigne := TobL.GetValue('GL_NUMLIGNE');
    { Recherche la ligne }
    T := GetLigneTarif(TobLigneTarif, TobL);
    if T = nil then T := CreerLigneTarifMere(TobLigneTarif, -1, CleDoc, NumL);
    T.ClearDetail;
    { Recharge les Lignes de Tarifs }
    Q := OpenSQL('SELECT * FROM LIGNETARIF WHERE ' + WherePiece(CleDoc, ttdLigneTarif, True), True);
    try
      T.LoadDetailDB('LIGNETARIF', '', '', Q, False, True);
    finally
      Ferme(Q);
    end;
  end;
*)
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 15/09/2003
Modifié le ... :   /  /
Description .. : Suppression de LIGNETARIF d'une ligne de pièce
Mots clefs ... : LIGNES/TARIFS/LIGNETARIF
*****************************************************************}
procedure SupprimeLigneTarif(TobPiece, TobLigneTarif: Tob; ARow: Integer);
var
  Tobl, TobLTarif: Tob;
begin
  Tobl := GetTobLigne(TobPiece, ARow);
  if Tobl = nil then Exit;
  TobLTarif := GetLigneTarif(TobLigneTarif, Tobl);
  if ToblTarif <> nil then ToblTarif.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 15/09/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : Création de la Tob mère à  TobLigneTarif
Mots clefs ... : LIGNES/TARIFS/LIGNETARIF
*****************************************************************}
function CreerLigneTarifMere(TobLigneTarif: Tob; ARow: Integer; CleDoc: R_CleDoc; NumLigne: Integer): TOB;
begin
	result := nil;
(* SUPPRIME -- SUITE A DROPTABLE DANS PGIMAJVER
 	if TobLigneTarif = nil then exit;
  Result := Tob.Create('_LIGNE_', TobLigneTarif, ARow);
  Result.AddChampSupValeur('GL_NATUREPIECEG', CleDoc.NaturePiece);
  Result.AddChampSupValeur('GL_SOUCHE', CleDoc.Souche);
  Result.AddChampSupValeur('GL_NUMERO', CleDoc.NumeroPiece);
  Result.AddChampSupValeur('GL_INDICEG', CleDoc.Indice);
  Result.AddChampSupValeur('GL_NUMLIGNE', NumLigne);
*)
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 15/09/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : Création de la Tob mère à  TobLigneTarif
Mots clefs ... : LIGNES/TARIFS/LIGNETARIF
*****************************************************************}
function CreerLigneTarifMere(TobLigneTarif: Tob; ARow: Integer; TobL: Tob; NumLigne: Integer): TOB;
var
  CleDoc: R_CleDoc;
begin
  CleDoc.NaturePiece := TobL.GetValue('GL_NATUREPIECEG');
  CleDoc.Souche := TobL.GetValue('GL_SOUCHE');
  CleDoc.NumeroPiece := TobL.GetValue('GL_NUMERO');
  CleDoc.Indice := TobL.GetValue('GL_INDICEG');
  CleDoc.NumLigne := TobL.GetValue('GL_NUMLIGNE');
  Result := CreerLigneTarifMere(TobLigneTarif, ARow, CleDoc, NumLigne);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 15/09/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : Chargement de la table LIGNETARIF pour une ligne de
Suite ........ : pièce
Mots clefs ... : LIGNES/TARIFS/LIGNETARIF
*****************************************************************}
function GetLigneTarif(TobLigneTarif, TobLigne: Tob): Tob;
begin
  Result := nil;
  if (TobLigneTarif <> nil) and (TobLigne <> nil) and (TobLigne.GetValue('GL_TYPEREF') = 'ART') then
  begin
    Result := TobLigneTarif.FindFirst(['GL_NATUREPIECEG', 'GL_SOUCHE', 'GL_NUMERO', 'GL_INDICEG', 'GL_NUMLIGNE'], [TobLigne.GetValue('GL_NATUREPIECEG'),
      TobLigne.GetValue('GL_SOUCHE'), TobLigne.GetValue('GL_NUMERO'), TobLigne.GetValue('GL_INDICEG'), TobLigne.GetValue('GL_NUMLIGNE')], True);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 16/09/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : Numérotation des LIGNETARIFS
Mots clefs ... : LIGNES/TARIFS/LIGNETARIF
*****************************************************************}
procedure NumeroteLigneTarif(TOBLigneTarif: TOB);
var
  i, j : integer;
  TobLTarif: TOB;
begin
  if Assigned(TobLigneTarif) then
  begin
    { Scan les lignes }
    for i := 0 to TOBLigneTarif.Detail.Count - 1 do
    begin
      TobLTarif := TobLigneTarif.Detail[i];
      TobLTarif.PutValue('GL_NUMLIGNE', i + 1);
      { Scan les sous-lignes }
      for j := 0 to TobLTarif.Detail.Count - 1 do
      begin
        TobLTarif.Detail[j].PutValue('GLT_NUMLIGNE', i + 1);
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 16/09/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : Création de la Tob mère de la tob TobLigneTarif
Mots clefs ... : LIGNES/TARIFS/LIGNETARIF
*****************************************************************}
procedure InsertTobLigneTarif(TOBLigneTarif: TOB; ARow: integer; CleDoc: R_CleDoc);
begin
  CreerLigneTarifMere(TobLigneTarif, ARow - 1, CleDoc, ARow);
end;

procedure DeleteTobLigneTarif(TobPiece: Tob; NumLigne: Integer);
var
  TobLigneTarifLocal: Tob;
begin
  TobLigneTarifLocal := GetTobLigneTarifInTobPiece(TobPiece);
  if Assigned(TobLigneTarifLocal) then
  begin
     if (NumLigne >= 0) and (NumLigne <= TobLigneTarifLocal.Detail.Count - 1) then
       TobLigneTarifLocal.Detail[NumLigne].Free;
     TobLigneTarifLocal.GetIndex
  end;
end;
procedure DeleteTobLigneTarif(TobPiece: Tob; Tobl: Tob); overload;
begin
  if TobL.Parent = TobPiece then
    DeleteTobLigneTarif(TobPiece, Tobl.GetIndex);
end;
end.
