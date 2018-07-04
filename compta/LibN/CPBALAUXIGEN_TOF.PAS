{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 30/07/2003
Modifié le ... : 31/07/2003
Description .. : Source TOF de la FICHE : CPBALAUXIGEN
Suite ........ : 
Suite ........ : Gestion du QRS1 de la balance auxiliaire par générale
Suite ........ : 
Suite ........ : Attention : les traitements communs aux balances 
Suite ........ : combinatoires sont dans une tof intermédiaire :
Suite ........ : TOF_METH --> TOF_CPBALANCECOMBI --> 
Suite ........ : TOF_CPBALAUXIGEN
Suite ........ : 
Suite ........ : JP 01/07/05 : Gestion des caractères Joker : fonctions de
Suite ........ : base définies dans TofMeth
Suite ........ :
Mots clefs ... : TOF;CPBALAUXIGEN
*****************************************************************}
Unit CPBALAUXIGEN_TOF ;

Interface

Uses StdCtrls, Classes,
{$IFDEF EAGLCLIENT}
     MainEAgl, utob, eQRS1,
{$ELSE}
     db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$ENDIF}
     sysutils, Spin, Ent1,
     HCtrls, HEnt1, UTOF,
     TofMeth,
     uLibExercice,  // CRelatifVersExercice
     CPBALANCECOMBI_TOF
  ,UentCommun     
      ;

procedure CPLanceFiche_BalanceAuxiGen ( Args : String = '' ) ;

Type
  TOF_CPBALAUXIGEN = Class (TOF_CPBALANCECOMBI)
    AuxiFermes             : THEdit;
    // EVT TOF
    procedure OnNew                    ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    // EVT FICHE
    procedure NatureCptChanged(Sender : TObject);   override ;
    procedure TableLibreChanged(Sender : TObject);  override ;
    procedure CompteOnExit(Sender: TObject) ;       override ;
    procedure Compte2OnExit(Sender: TObject) ;      override ;
    // TRAITEMENT
    // -> Procédure d'insertion des comptes cibles
    function  GenererRequeteComptes       : String ;      override ;
    function  GenererRequeteComptesCompar : String ;      override ;
    function  GenererRequeteCombis ( vStCode, vStLib, vStRup : String )       : String ; override ;
    function  GenererRequeteCombisCompar ( vStCode, vStLib, vStRup : String ) : String ; override ;
    // -> Procedure de calcul des totaux / soldes pour alimentation table temporaire
    procedure CalculDebitCredit( vStCompte, vStCompte2: String ; vBoCompar : Boolean ; var vDebit, vCredit: Double );  override ;
    // -> Construit la requête de l'état
    function  GetNomChampCPT : String ;  override ;
    function  GetNomChampCPTCombi : String ;  override ;
    procedure RemplirEDTBALANCE                ; override ;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  uLibWindows; // TestJoker

//==============================================================================
procedure CPLanceFiche_BalanceAuxiGen ( Args : String ) ;
begin
	AGLLanceFiche('CP','CPBALAUXIGEN','','',Args);
end;
//==============================================================================

{TOF_CPBALAUXIGEN}

//==============================================================================
//====================         EVT TOF        ==================================
//==============================================================================

procedure TOF_CPBALAUXIGEN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPBALAUXIGEN.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPBALAUXIGEN.OnLoad ;
begin

  Inherited ;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if not TestJoker(CompteDe.Text) then begin
    // Auto-remplissage si comptes auxiliaires non renseignés
    if (Trim(CompteDe.Text) = '') then
      CompteDe.Text := GetMinCompte('TIERS', 'T_AUXILIAIRE', NatureCpt.Value);
    if (Trim(CompteA.Text) = '') then
      CompteA.Text := GetMaxCompte('TIERS', 'T_AUXILIAIRE', NatureCpt.Value);
  end;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if not TestJoker(Compte2De.Text) then begin
    // Auto-remplissage si comptes généraux non renseignés
    if (Trim(Compte2De.Text) = '') then
      Compte2De.Text := GetMinCompte('GENERAUX', 'G_GENERAL', '' );
    if (Trim(Compte2A.Text) = '') then
      Compte2A.Text := GetMaxCompte('GENERAUX', 'G_GENERAL', '' );
  end;

  // Les ruptures...
  if Rupture.Value <> 'SANS' then
    begin
    // rupture sur n° compte
    if (RuptureType.Value = 'RUPGROUPES') and (NivoRupture.value > 0) then
      AvecNivoRupture.Text := IntToStr(NivoRupture.value);
    // rupture sur champ libre tiers
    if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
      XXRupture.Text := 'T_TABLE' + GetNumTableLibre ;
    // rupture sur plan de correspondance
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
      XXRupture.Text := 'T_CORRESP' + IntToStr(Corresp.ItemIndex + 1) ;
    end;

end ;

procedure TOF_CPBALAUXIGEN.OnArgument (S : String ) ;
begin

  // CONTROLES SPECIFIQUES
  // Contrôles critères et options - Page standard
  Exercice    	:= THValComboBox(Getcontrol('E_EXERCICE', true));
  // Contrôles critères et options - Page avancés
  Devise      	:= THValComboBox(Getcontrol('E_DEVISE', true));
  AuxiFermes := THEdit(GetControl('AUXIFERMES', true));
  // Controles communs
  Inherited ;

  // Rubrique d'aide
  Ecran.HelpContext := 7460000;

  // Type de plan comptable :
  Case V_PGI.LaSerie of
    S7 :  begin
          Corresp.plus := 'AND (CO_CODE = "AU1" OR CO_CODE = "AU2")';
          end;
    else  begin
          Corresp.plus := 'AND CO_CODE = "AU1"';
          end;
    end;

  // Longueur max des comptes
  CompteDe.MaxLength  := VH^.CPta[fbAux].Lg;
  CompteA.MaxLength   := VH^.CPta[fbAux].Lg;
  Compte2De.MaxLength := VH^.CPta[fbGene].Lg;
  Compte2A.MaxLength  := VH^.CPta[fbGene].Lg;

end ;

//==============================================================================
//====================        EVT QRS1        ==================================
//==============================================================================

procedure TOF_CPBALAUXIGEN.NatureCptChanged(Sender: TObject);
begin
  CompteDe.Text := '';
  CompteA.Text := '';
  if NatureCpt.ItemIndex > 0 then
    begin
    CompteDe.Plus := ' AND T_NATUREAUXI="'+NatureCPT.Value+'"' ;
    CompteA.Plus := ' AND T_NATUREAUXI="'+NatureCPT.Value+'"' ;
    end
  else
    begin
    CompteDe.Plus := '' ;
    CompteA.Plus := '';
    end ;
end;

procedure TOF_CPBALAUXIGEN.TableLibreChanged(Sender: TObject);
begin
  if TableLibre.ItemIndex < 0 then
    begin
    LibreDe.DataType  := '';
    LibreA.DataType   := '';
    LibreDe.Text      := '';
    LibreA.Text       := '';
    end
  else
    begin
    LibreDe.DataType  := 'TZNATTIERS' + GetNumTableLibre;
    LibreA.DataType   := 'TZNATTIERS' + GetNumTableLibre;
    LibreDe.Text      := '';
    LibreA.Text       := '';
    end;
end;

//==============================================================================
//====================   TRAITEMENTS DES DONNEES  ==============================
//==============================================================================
procedure TOF_CPBALAUXIGEN.CalculDebitCredit(vStCompte, vStCompte2: String; vBoCompar: Boolean; var vDebit, vCredit: Double);
Var QSum   : TQuery ;
    lStReq : String ;
begin
  vDebit  := 0 ;
  vCredit := 0 ;

  // **** CALCUL SIMPLE DES TOTAUX ****
  // ----------
  // Construction de la requête de récupération des montants
  // ----------
  // --> Champs
  if Affichage.Value = 'EURO'
    then lStReq := 'SELECT SUM(E_DEBIT) D, SUM(E_CREDIT) C'
    else lStReq := 'SELECT SUM(E_DEBITDEV) D, SUM(E_CREDITDEV) C';
  // --> Table
  lStReq := lStReq + ' FROM ' + GetTablePourBase('ECRITURE') ;
  // --> Conditions
  lStReq := lStReq + ' WHERE E_GENERAL = "' + vStCompte + '"' ;
  lStReq := lStReq + ' AND E_AUXILIAIRE = "' + vStCompte2 + '"' ;
  // --> Condition ecritures : sur période ciblée ou comparée ?
  if vBoCompar then
    lStReq := lStReq + ' AND ' + RecupWhereBalanceCompar
  else
    lStReq := lStReq + ' AND ' + RecupWhereQualifPiece
                     + FindEtReplace(RecupWhereBalance,'WHERE',' AND', False) ;

  // ----------
  // Récupération des montants
  // ----------
  QSum := OpenSQL( lStReq, True);
  if not QSum.Eof then
    begin
    vDebit  := QSum.FindField('D').AsFloat ;
    vCredit := QSum.FindField('C').AsFloat ;
    end ;

  Ferme(QSum) ;

end;

function TOF_CPBALAUXIGEN.GenererRequeteCombis(vStCode, vStLib, vStRup: String): String;
var	listeCptExcept  : String ;
    stCpt           : String ;
begin
// INSERT des comptes
  result := 'INSERT INTO ' + GetTablePourBase('CEDTBALANCE') + ' (CED_USER, CED_COMPTE, CED_LIBELLE';
  result := result + ',CED_CREDIT1, CED_DEBIT1, CED_CREDIT2, CED_DEBIT2';
  result := result + ', CED_RUPTURE, CED_COLLECTIF, CED_COMPTE2, CED_LIBELLE2) ';
// SELECT des comptes
  // Champs
  result := result + 'SELECT "' + V_PGI.User + '", E_GENERAL, G_LIBELLE';
  result := result + ',0, 0, 0, 0';
  // Rupture
  result := result + ', "' + vStRup + '"' ;
  // Collectif, compte2, libelle2
  result := result + ', "-", "' + vStCode + '", "' + CheckdblQuote(vStLib) + '"' ;
  // Table Temp + Compte
  result := result + ' FROM ' + GetTablePourBase('ECRITURE');
  // Table Ecriture si besoin
  result := result + ' LEFT JOIN ' + GetTablePourBase('GENERAUX') + ' ON G_GENERAL = E_GENERAL';
  // Conditions du WHERE (sur les comptes uniquement donc)
    // Comptes de ... à
//  result := result + ' WHERE E_GENERAL>="' + Compte2De.Text + '"';
  //result := result + ' AND E_GENERAL<="' + Compte2A.Text + '"';
  {JP 01/07/05 : Gestion des caractères jokers}
  Result := Result + ' WHERE ' + ConvertitCaractereJokers(Compte2De, Compte2A, 'E_GENERAL');

  result := result + ' AND E_AUXILIAIRE="' + vStCode + '"';
   // Conditions sur les écritures :
  result := result + ' AND ' + RecupWhereQualifPiece ;
  
  { FQ 19715 BVE 07.06.07 }
  if ModeSelect.value = 'NSL' then
     result := result + ' AND ((SELECT SUM(E_CREDIT)-SUM(E_DEBIT)'
                      + ' FROM ' + GetTablePourBase('ECRITURE')
      	              + ' WHERE E_GENERAL = G_GENERAL) <> 0 )';
  { END FQ 19715 }

  result := result + FindEtReplace(RecupWhereBalance,'WHERE',' AND', False) ;
  // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
  if CptExcept2.Text <> '' then
    begin
    listeCptExcept := FindEtReplace(CptExcept2.Text,',',';', True);
    while (listeCptExcept <> '') do
      begin
      stCpt := Trim(ReadTokenSt(listeCptExcept)) ;
      if stCpt <> '' then
	result := result + ' AND E_GENERAL NOT LIKE "' + stCpt + '%"';
      end ;
    end;

  // SI COMPARATIF, on se contente d'ajouter les comptes non présent dans CEDTBALANCE
  if AvecComparatif.Checked then
    Result := result + ' AND (E_GENERAL not in (SELECT CED_COMPTE FROM ' + GetTablePourBase('CEDTBALANCE')
                                            + ' WHERE CED_USER = "' + V_PGI.User + '"'
                                            + ' AND CED_COMPTE2 = "' + vStCode + '") )' ;

  // Group by (??)
  result := result + ' GROUP BY E_GENERAL, G_LIBELLE';

end;

function TOF_CPBALAUXIGEN.GenererRequeteCombisCompar(vStCode, vStLib, vStRup: String): String;
var	listeCptExcept  : String ;
    stCpt           : String ;
begin
// INSERT des comptes
  result := 'INSERT INTO ' + GetTablePourBase('CEDTBALANCE') + ' (CED_USER, CED_COMPTE, CED_LIBELLE';
  result := result + ',CED_CREDIT1, CED_DEBIT1, CED_CREDIT2, CED_DEBIT2';
  result := result + ', CED_RUPTURE, CED_COLLECTIF, CED_COMPTE2, CED_LIBELLE2) ';
// SELECT des comptes
  // Champs
  result := result + 'SELECT "' + V_PGI.User + '", E_GENERAL, G_LIBELLE';
  result := result + ',0, 0, 0, 0';
  // Rupture
  result := result + ', "' + vStRup + '"' ;
  // Collectif, compte2, libelle2
  result := result + ', "-", "' + vStCode + '", "' + CheckdblQuote(vStLib) + '"' ;
  // Table Temp + Compte
  result := result + ' FROM ' + GetTablePourBase('ECRITURE');
  // Table Ecriture si besoin
  result := result + ' LEFT JOIN ' + GetTablePourBase('GENERAUX') + ' ON E_GENERAL = G_GENERAL';
  // Conditions du WHERE (sur les comptes uniquement donc)
    // Comptes de ... à
//  result := result + ' WHERE E_GENERAL>="' + Compte2De.Text + '"';
  //result := result + ' AND E_GENERAL<="' + Compte2A.Text + '"';
  {JP 01/07/05 : Gestion des caractères jokers}
  Result := Result + ' WHERE ' + ConvertitCaractereJokers(Compte2De, Compte2A, 'E_GENERAL');

  result := result + ' AND E_AUXILIAIRE="' + vStCode + '"';
   // Conditions sur les écritures :
  result := result + ' AND ' + RecupWhereBalanceCompar ;

  { FQ 19715 BVE 07.06.07 }
  if ModeSelect.value = 'NSL' then
     result := result + ' AND ((SELECT SUM(E_CREDIT)-SUM(E_DEBIT)'
                      + ' FROM ' + GetTablePourBase('ECRITURE')
      	              + ' WHERE E_GENERAL = G_GENERAL) <> 0 )';
  { END FQ 19715 }

  // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
  if CptExcept2.Text <> '' then
    begin
    listeCptExcept := FindEtReplace(CptExcept2.Text,',',';', True);
    while (listeCptExcept <> '') do
      begin
      stCpt := Trim(ReadTokenSt(listeCptExcept)) ;
      if stCpt <> '' then
	result := result + ' AND E_GENERAL NOT LIKE "' + stCpt + '%"';
      end ;
    end;

  // Group by (??)
  result := result + ' GROUP BY E_GENERAL, G_LIBELLE';

end;

function TOF_CPBALAUXIGEN.GenererRequeteComptes: String;
var	Q								: TQuery ;
    DateDebExo      : TDatetime ;
    DateFinExo      : TDatetime ;
    listeCptExcept  : String ;
    stCpt           : String ;
begin
  result := 'SELECT T_AUXILIAIRE CODE, T_LIBELLE LIBELLE';
  // Rupture ?
  if XXRupture.Text <> ''
    then result := result + ', ' + XXRupture.Text + ' RUPTURE'
    else result := result + ', "" RUPTURE';
  // Table Temp + Compte
  result := result + ' FROM TIERS';
  // Table Ecriture si besoin
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then
    result := result + ' LEFT JOIN ' + GetTablePourBase('ECRITURE') + ' ON E_AUXILIAIRE = T_AUXILIAIRE' ;
  // Conditions du WHERE (sur les comptes uniquement donc)
    // Comptes de ... à
//  result := result + ' WHERE T_AUXILIAIRE>="' + CompteDe.Text + '"';
  //result := result + ' AND T_AUXILIAIRE<="' + CompteA.Text + '"';
  {JP 01/07/05 : Gestion des caractères jokers}
  Result := Result + ' WHERE ' + ConvertitCaractereJokers(CompteDe, CompteA, 'T_AUXILIAIRE');

    // Table Ecriture si besoin
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then
    result := result + ' AND ' + RecupWhereQualifPiece
                     + FindEtReplace(RecupWhereBalance(False),'WHERE',' AND', False) ;

    // Uniquement des comptes non fermés
//  result := result + ' AND T_FERME="-"' ;
  // Nature de comptes
  if NatureCPT.Value <> '' then
    result := result + ' AND T_NATUREAUXI="' + NatureCPT.Value + '"';
    // Mode de sélection des comptes
  if ModeSelect.value = 'EXO' then
    begin
    Q:=OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="' + CRelatifVersExercice(Exercice.value) + '"' ,TRUE) ;
    if Not Q.EOF then
      begin
      DateDebExo := Q.FindField('EX_DATEDEBUT').asDateTime ;
      DateFinExo := Q.FindField('EX_DATEFIN').asDateTime ;
      result     := result + ' AND E_DATECOMPTABLE>="' + USDateTime(DateDebExo)
      	      	          + '" AND E_DATECOMPTABLE<="' + USDateTime(DateFinExo) + '"';
      end;
    Ferme(Q) ;
    end
  else if ModeSelect.value = 'PER' then
    result := result + ' AND E_DATECOMPTABLE>="' + USDateTime(DateDebutPer)
                     + '" AND E_DATECOMPTABLE<="' + USDateTime(DateFinPer) + '"'
  else if ModeSelect.value = 'NSL' then
    result := result + ' AND ( (SELECT SUM(E_CREDIT)-SUM(E_DEBIT) FROM ' + GetTablePourBase('ECRITURE')
  			    + ' WHERE E_AUXILIAIRE = T_AUXILIAIRE AND ' + RecupWhereQualifPiece
                            + FindEtReplace(RecupWhereBalance,'WHERE',' AND', False) + ') <> 0 )';
  // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
  if CptExcept.Text <> '' then
    begin
    listeCptExcept := FindEtReplace(CptExcept.Text,',',';', True);
    while (listeCptExcept <> '') do
      begin
      stCpt := Trim(ReadTokenSt(listeCptExcept)) ;
      if stCpt <> '' then
       	result := result + ' AND T_AUXILIAIRE NOT LIKE "' + stCpt + '%"';
      end ;
    end;

  // Conditions sur champ de rupture
  result := result + GetConditionRupture ;

  // Group by
  result := result + ' GROUP BY T_AUXILIAIRE, T_LIBELLE';
  if XXRupture.Text <> '' then
    result := result + ', ' + XXRupture.Text ;

end;

function TOF_CPBALAUXIGEN.GenererRequeteComptesCompar: String;
var	Q								: TQuery ;
    DateDebExo      : TDatetime ;
    DateFinExo      : TDatetime ;
    listeCptExcept  : String ;
    stCpt           : String ;
begin
  result := 'SELECT T_AUXILIAIRE CODE, T_LIBELLE LIBELLE';
  // Rupture ?
  if XXRupture.Text <> ''
    then result := result + ', ' + XXRupture.Text + ' RUPTURE'
    else result := result + ', "" RUPTURE';
  // Table Temp + Compte
  result := result + ' FROM ' + GetTablePourBase('TIERS');
  // Table Ecriture si besoin
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then
    result := result + ' LEFT JOIN ' + GetTablePourBase('ECRITURE') + ' ON E_AUXILIAIRE = T_AUXILIAIRE' ;
  // Conditions du WHERE (sur les comptes uniquement donc)
    // Comptes de ... à
//  result := result + ' WHERE T_AUXILIAIRE>="' + CompteDe.Text + '"';
  //result := result + ' AND T_AUXILIAIRE<="' + CompteA.Text + '"';
  {JP 01/07/05 : Gestion des caractères jokers}
  Result := Result + ' WHERE ' + ConvertitCaractereJokers(CompteDe, CompteA, 'T_AUXILIAIRE');

    // Table Ecriture si besoin
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then
    result := result + ' AND ' + RecupWhereBalanceCompar( False ) ;
    // Uniquement des comptes collectifs et non fermés
  result := result + ' AND T_FERME="-"' ;
    // Nature de comptes
  if NatureCPT.Value <> '' then
    result := result + ' AND T_NATUREAUXI="' + NatureCPT.Value + '"';
    // Mode de sélection des comptes
  if ModeSelect.value = 'EXO' then
    begin
    Q:=OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+CRelatifVersExercice(ComparExo.value)+'"' ,TRUE) ;
    if Not Q.EOF then
      begin
      DateDebExo := Q.FindField('EX_DATEDEBUT').asDateTime ;
      DateFinExo := Q.FindField('EX_DATEFIN').asDateTime ;
      result     := result + ' AND E_DATECOMPTABLE>="' + USDateTime(DateDebExo)
      			  + '" AND E_DATECOMPTABLE<="' + USDateTime(DateFinExo) + '"';
      end;
    Ferme(Q) ;
    end
  else if ModeSelect.value = 'PER' then
    result := result + ' AND E_DATECOMPTABLE>="' + USDateTime(DateDebutCompar)
                    + '" AND E_DATECOMPTABLE<="' + USDateTime(DateFinCompar) + '"'
  else if ModeSelect.value = 'NSL' then
    result := result + ' AND ( (SELECT SUM(E_CREDIT)-SUM(E_DEBIT) FROM ' + GetTablePourBase('ECRITURE')
  			    + ' WHERE E_AUXILIAIRE = T_AUXILIAIRE AND ' + RecupWhereBalanceCompar + ') <> 0 )';
  // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
  if CptExcept.Text <> '' then
    begin
    listeCptExcept := FindEtReplace(CptExcept.Text,',',';', True);
    while (listeCptExcept <> '') do
      begin
      stCpt := Trim(ReadTokenSt(listeCptExcept)) ;
      if stCpt <> '' then
	result := result + ' AND T_AUXILIAIRE NOT LIKE "' + stCpt + '%"';
      end ;
    end;
  // Conditions sur champ de rupture
  result := result + GetConditionRupture ;

  // Group by
  result := result + ' GROUP BY T_AUXILIAIRE, T_LIBELLE';
  if XXRupture.Text <> '' then
    result := result + ', ' + XXRupture.Text ;

end;

function TOF_CPBALAUXIGEN.GetNomChampCPT: String;
begin
  Result := 'CED_GENERAL' ;
end;

function TOF_CPBALAUXIGEN.GetNomChampCPTCombi: String;
begin
  Result := 'CED_AUXILIAIRE' ;
end;

procedure TOF_CPBALAUXIGEN.CompteOnExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  // Complétion auto du numéro de compte si possible
  if not CompleteAuto( Sender, fbAux ) then
    THEdit(Sender).ElipsisClick(nil);
end;

procedure TOF_CPBALAUXIGEN.Compte2OnExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  // Complétion auto du numéro de compte si possible
  if not CompleteAuto( Sender, fbGene ) then
    THEdit(Sender).ElipsisClick(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 31/01/2006
Modifié le ... : 12/09/2006
Description .. : Ne pas prendre les comptes auxi fermés FQ 17207
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALAUXIGEN.RemplirEDTBALANCE;
begin
inherited   ;
  //YMO 12/09/2006 Ne pas éliminer les enregs avec un débit ou un crédit <> 0
  //YMO 22/09/2006 FQ18810 correction+prise en compte de l'utilisateur
  //YMO 29/09/2006 FQ18810 Suppression aussi des cptes generaux fermés ET non mouvementés
  IF ExecuteSQL('DELETE FROM CEDTBALANCE WHERE CED_USER = "' + V_PGI.User + '"'
  +' AND (CED_DEBIT1=0 AND CED_CREDIT1=0)'
  +' AND ((CED_COMPTE2 =(SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE = CED_COMPTE2 AND T_FERME = "X"))'
  +' OR  (CED_COMPTE=(SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = CED_COMPTE AND G_FERME = "X"))'
  +' AND CED_COMPTE2="")')>0 then
      AuxiFermes.Text := 'X'
  else
      AuxiFermes.Text := '-';

end;

Initialization
  registerclasses ( [ TOF_CPBALAUXIGEN ] ) ;
end.
