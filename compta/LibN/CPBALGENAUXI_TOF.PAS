{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 30/07/2003
Modifié le ... : 31/07/2003
Description .. : Source TOF de la FICHE : CPBALGENAUXI
Suite ........ : 
Suite ........ : Gestion du QRS1 de la balance générale par auxiliaire
Suite ........ : 
Suite ........ : Attention : les traitements communs aux balances 
Suite ........ : combinatoires sont dans une tof intermédiaire :
Suite ........ : TOF_METH --> TOF_CPBALANCECOMBI --> 
Suite ........ : TOF_CPBALGENAUXI
Suite ........ : 
Suite ........ : JP 01/07/05 : Gestion des caractères Joker : fonctions de
Suite ........ : base définies dans TofMeth
Mots clefs ... : TOF;CPBALGENAUXI
*****************************************************************}
Unit CPBALGENAUXI_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     MainEAgl, utob, eQRS1,
{$ELSE}
     db, Fe_Main, QRS1, MajTable,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms, sysutils, ComCtrls, Spin, Ent1,
     Filtre,
     HCtrls, HEnt1, HMsgBox, UTOF, hqry,
     TofMeth,
     uLibExercice,  // CRelatifVersExercice
     uLibWindows,   // TraductionTHMultiValComboBox
     CPBALANCECOMBI_TOF,
     Dialogs ;

procedure CPLanceFiche_BalanceGenAuxi ( Args : String = '' ) ;

Type
  TOF_CPBALGENAUXI = Class (TOF_CPBALANCECOMBI)
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
    function  GenererRequeteCombisBalsit ( vStCode, vStLib, vStRup : String ) : String ; override ;
    // -> Procedure de calcul des totaux / soldes pour alimentation table temporaire
    procedure CalculDebitCredit( vStCompte, vStCompte2: String ; vBoCompar : Boolean ; var vDebit, vCredit: Double );  override ;
    // -> Construit la requête de l'état
    function  GetNomChampCPT : String ;  override ;
    function  GetNomChampCPTCombi : String ;  override ;
    procedure RemplirEDTBALANCE      ; override ;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Paramsoc;

//==============================================================================
procedure CPLanceFiche_BalanceGenAuxi ( Args : String ) ;
begin
	AGLLanceFiche('CP','CPBALGENAUXI','','',Args);
end;
//==============================================================================

{TOF_CPBALGENAUXI}

//==============================================================================
//====================         EVT TOF        ==================================
//==============================================================================

procedure TOF_CPBALGENAUXI.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPBALGENAUXI.OnUpdate ;
var
  SQL : string;
begin                
  Inherited ;

  { FQ 20337 BVE 15.06.07 }
  SQL := 'SELECT * FROM ' + GetTablePourBase('CEDTBALANCE') + ' ' +
         'WHERE CED_COMPTE = "" ' +
         'AND CED_LIBELLE = "" ' +
         'AND CED_USER = "' + V_PGI.User + '"';
  if ExisteSQL(SQL) then
  begin
     SQL := 'DELETE FROM ' + GetTablePourBase('CEDTBALANCE') + ' ' +
            'WHERE CED_COMPTE = "" ' +
            'AND CED_LIBELLE = "" ' +
            'AND CED_USER = "' + V_PGI.User + '"';
     ExecuteSQL(SQL);
  end;      
  SQL := 'SELECT * FROM ' + GetTablePourBase('CEDTBALANCE') + ' ' +
         'WHERE CED_DEBIT1 = 0 AND CED_DEBIT2 = 0 ' +
         'AND CED_CREDIT1 = 0 AND CED_CREDIT2 = 0 ' +
         'AND CED_USER = "' + V_PGI.User + '"';
  if ExisteSQL(SQL) then
  begin
     SQL := 'DELETE FROM ' + GetTablePourBase('CEDTBALANCE') + ' ' +
            'WHERE CED_DEBIT1 = 0 AND CED_DEBIT2 = 0 ' +
            'AND CED_CREDIT1 = 0 AND CED_CREDIT2 = 0 ' +
            'AND CED_USER = "' + V_PGI.User + '"';
     ExecuteSQL(SQL);
  end;
  { END FQ 20337 }
end ;

procedure TOF_CPBALGENAUXI.OnLoad ;
begin

  Inherited ;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if not TestJoker(CompteDe.Text) then begin
    // Auto-remplissage si comptes généraux non renseignés
    if (Trim(CompteDe.Text) = '') then
      CompteDe.Text := GetMinCompte('GENERAUX', 'G_GENERAL', NatureCpt.Value );
    if (Trim(CompteA.Text) = '') then
      CompteA.Text := GetMaxCompte('GENERAUX', 'G_GENERAL', NatureCpt.Value );
  end;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if not TestJoker(Compte2De.Text) then begin
    // Auto-remplissage si comptes auxiliaires non renseignés
    if (Trim(Compte2De.Text) = '') then
      Compte2De.Text := GetMinCompte('TIERS', 'T_AUXILIAIRE', '');
    if (Trim(Compte2A.Text) = '') then
      Compte2A.Text := GetMaxCompte('TIERS', 'T_AUXILIAIRE', '');
  end;

  // Les ruptures...
  if Rupture.Value <> 'SANS' then
    begin
    // rupture sur n° compte
    if (RuptureType.Value = 'RUPGROUPES') and (NivoRupture.value > 0) then
      AvecNivoRupture.Text := IntToStr(NivoRupture.value);
    // rupture sur champ libre tiers
    if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
      XXRupture.Text := 'G_TABLE' + GetNumTableLibre ;
    // rupture sur plan de correspondance
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
      XXRupture.Text := 'G_CORRESP' + IntToStr(Corresp.ItemIndex + 1) ;
    end;

end ;

procedure TOF_CPBALGENAUXI.OnArgument (S : String ) ;
begin

  // CONTROLES SPECIFIQUES
  // Contrôles critères et options - Page standard
  Exercice		:= THValComboBox(Getcontrol('E_EXERCICE', true));
  // Contrôles critères et options - Page avancés
  Devise		:= THValComboBox(Getcontrol('E_DEVISE', true));
  // Controles comparatifs
  BalSit		:= THEdit(GetControl('BALSIT', true));
  AuxiFermes := THEdit(GetControl('AUXIFERMES', true));
  // Controles communs
  Inherited ;

  // Evènement sur type de comparatif
  ComparType.OnClick 	:= ComparTypeClick;

  // Rubrique d'aide
  Ecran.HelpContext := 7457000;

  // Type de plan comptable :   
  { FQ 20338 BVE 15.06.07 }
  if GetParamSocSecur('SO_CORSGE1', True, True) then
  begin
     if GetParamSocSecur('SO_CORSGE2', True, True) then
        Corresp.plus := 'AND (CO_CODE = "GE1" OR CO_CODE = "GE2")'
     else
        Corresp.plus := 'AND (CO_CODE = "GE1")';
  end
  else
  begin
     if GetParamSocSecur('SO_CORSGE2', True, True) then
        Corresp.plus := 'AND (CO_CODE = "GE2")'
     else
     begin
        Corresp.Text := '';
        Corresp.Enabled := false;
     end;
  end;
  { END FQ 20338 }
  // Longueur max des comptes
  CompteDe.MaxLength := VH^.CPta[fbGene].Lg;
  CompteA.MaxLength  := VH^.CPta[fbGene].Lg;
  Compte2De.MaxLength := VH^.CPta[fbAux].Lg;
  Compte2A.MaxLength  := VH^.CPta[fbAux].Lg;

  // Type de balance de situation
  BalSit.Plus := ' BSI_TYPECUM = "G/T" ' ;


end ;

//==============================================================================
//====================        EVT QRS1        ==================================
//==============================================================================

procedure TOF_CPBALGENAUXI.NatureCptChanged(Sender: TObject);
begin
  CompteDe.Text := '';
  CompteA.Text  := '';
  if NatureCpt.ItemIndex > 0 then
    begin
    CompteDe.Plus := ' AND G_NATUREGENE="'+NatureCPT.Value+'"' ;
    CompteA.Plus  := ' AND G_NATUREGENE="'+NatureCPT.Value+'"' ;
    end
  else
    begin
    CompteDe.Plus := '' ;
    CompteA.Plus  := '';
    end ;
end;

procedure TOF_CPBALGENAUXI.TableLibreChanged(Sender: TObject);
begin
  if TableLibre.ItemIndex < 0 then
    begin
    LibreDe.DataType := '';
    LibreA.DataType  := '';
    LibreDe.Text     := '';
    LibreA.Text      := '';
    end
  else
    begin
    LibreDe.DataType  := 'TZNATGENE' + GetNumTableLibre ;
    LibreA.DataType   := 'TZNATGENE' + GetNumTableLibre ;
    LibreDe.Text      := '';
    LibreA.Text       := '';
    end;
end;

//==============================================================================
//====================   TRAITEMENTS DES DONNEES  ==============================
//==============================================================================

procedure TOF_CPBALGENAUXI.CalculDebitCredit( vStCompte, vStCompte2: String ; vBoCompar : Boolean ; var vDebit, vCredit: Double );
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
  lStReq := lStReq + ' WHERE E_AUXILIAIRE = "' + vStCompte + '"' ;
  lStReq := lStReq + ' AND E_GENERAL = "' + vStCompte2 + '"' ;
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

function TOF_CPBALGENAUXI.GenererRequeteCombis( vStCode, vStLib, vStRup : String ) : String ;
var	listeCptExcept  : String ;
    stCpt           : String ;
begin
// INSERT des comptes
  result := 'INSERT INTO ' + GetTablePourBase('CEDTBALANCE') + ' (CED_USER, CED_COMPTE, CED_LIBELLE';
  result := result + ',CED_CREDIT1, CED_DEBIT1, CED_CREDIT2, CED_DEBIT2';
  result := result + ', CED_RUPTURE, CED_COLLECTIF, CED_COMPTE2, CED_LIBELLE2) ';
// SELECT des comptes
  // Champs
  result := result + 'SELECT "' + V_PGI.User + '", E_AUXILIAIRE, T_LIBELLE';
  result := result + ',0, 0, 0, 0';
  // Rupture
  result := result + ', "' + vStRup + '"' ;
  // Collectif, compte2, libelle2
  result := result + ', "-", "' + vStCode + '", "' + CheckdblQuote(vStLib) + '"' ;
  // Table Temp + Compte
  result := result + ' FROM ' + GetTablePourBase('ECRITURE');
  // Table Ecriture si besoin
  result := result + ' LEFT JOIN ' + GetTablePourBase('TIERS') + ' ON E_AUXILIAIRE = T_AUXILIAIRE';
  // Conditions du WHERE (sur les comptes uniquement donc)
    // Comptes de ... à
//  result := result + ' WHERE E_AUXILIAIRE>="' + Compte2De.Text + '"';
//  result := result + ' AND E_AUXILIAIRE<="' + Compte2A.Text + '"';
  {JP 01/07/05 : Gestion des caractères jokers}
  Result := Result + ' WHERE ' + ConvertitCaractereJokers(Compte2De, Compte2A, 'E_AUXILIAIRE');

  result := result + ' AND E_GENERAL="' + vStCode + '"';
   // Conditions sur les écritures :
  result := result + ' AND ' + RecupWhereQualifPiece ;
  result := result + FindEtReplace(RecupWhereBalance,'WHERE',' AND', False) ;
  { FQ 19715 BVE 07.06.07 }
  if ModeSelect.value = 'NSL' then
     result := result + ' AND ((SELECT SUM(E_CREDIT)-SUM(E_DEBIT)'
                      + ' FROM ' + GetTablePourBase('ECRITURE')
      	              + ' WHERE E_AUXILIAIRE = T_AUXILIAIRE) <> 0 )';
  { END FQ 19715 }
  // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
  if CptExcept2.Text <> '' then
    begin
    listeCptExcept := FindEtReplace(CptExcept2.Text,',',';', True);
    while (listeCptExcept <> '') do
      begin
      stCpt := Trim(ReadTokenSt(listeCptExcept)) ;
      if stCpt <> '' then
	result := result + ' AND E_AUXILIAIRE NOT LIKE "' + stCpt + '%"';
      end ;
    end;

  // SI COMPARATIF, on se contente d'ajouter les comptes non présent dans CEDTBALANCE
  if AvecComparatif.Checked then
    Result := result + ' AND (T_AUXILIAIRE not in (SELECT CED_COMPTE FROM ' + GetTablePourBase('CEDTBALANCE')
                                            + ' WHERE CED_USER = "' + V_PGI.User + '"'
                                            + ' AND CED_COMPTE2 = "' + vStCode + '") )' ;

  // Group by (??)
  result := result + ' GROUP BY E_AUXILIAIRE, T_LIBELLE';

end;

function TOF_CPBALGENAUXI.GenererRequeteComptes: String;
var Q		    : TQuery ;
    DateDebExo      : TDatetime ;
    DateFinExo      : TDatetime ;
    listeCptExcept  : String ;
    stCpt           : String ;
begin
  result := 'SELECT G_GENERAL CODE, G_LIBELLE LIBELLE';
  // Rupture ?
  if XXRupture.Text <> ''
    then result := result + ', ' + XXRupture.Text + ' RUPTURE'
    else result := result + ', "" RUPTURE';
  // Table Temp + Compte
  result := result + ' FROM ' + GetTablePourBase('GENERAUX');
  // Table Ecriture si besoin
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then
    result := result + ' LEFT JOIN ' + GetTablePourBase('ECRITURE') + ' ON E_GENERAL = G_GENERAL' ;
  // Conditions du WHERE (sur les comptes uniquement donc)
    // Comptes de ... à
//  result := result + ' WHERE G_GENERAL>="' + CompteDe.Text + '"';
  //result := result + ' AND G_GENERAL<="' + CompteA.Text + '"';
  {JP 01/07/05 : Gestion des caractères jokers}
  Result := Result + ' WHERE ' + ConvertitCaractereJokers(CompteDe, CompteA, 'G_GENERAL');

    // Table Ecriture si besoin
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then
    result := result + ' AND ' + RecupWhereQualifPiece
                     + FindEtReplace(RecupWhereBalance(False),'WHERE',' AND', False) ;
    // Uniquement des comptes collectifs et non fermés
  result := result + ' AND G_COLLECTIF="X"' ; // AND G_FERME="-"' ; YMO 29/09/2006 les comptes fermé et non mouvementés sont supprimés dans CEDTBALANCE plus loin
    // Nature de comptes
  if NatureCPT.Value <> '' then
    result := result + ' AND G_NATUREGENE="' + NatureCPT.Value + '"';
    // Mode de sélection des comptes
  if ModeSelect.value = 'EXO' then
    begin
    Q:=OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+CRelatifVersExercice(Exercice.value)+'"' ,TRUE) ;
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
  begin
    result := result + ' AND ( (SELECT SUM(E_CREDIT)-SUM(E_DEBIT) FROM ' + GetTablePourBase('ECRITURE')
    		            + ' WHERE E_GENERAL = G_GENERAL AND ' + RecupWhereQualifPiece
                            + FindEtReplace(RecupWhereBalance,'WHERE',' AND', False) + ') <> 0 )';
  end;
  // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
  if CptExcept.Text <> '' then
    begin
    listeCptExcept := FindEtReplace(CptExcept.Text,',',';', True);
    while (listeCptExcept <> '') do
      begin
      stCpt := Trim(ReadTokenSt(listeCptExcept)) ;
      if stCpt <> '' then
	result := result + ' AND G_GENERAL NOT LIKE "' + stCpt + '%"';
      end ;
    end;
  // Conditions sur champ de rupture
  result := result + GetConditionRupture ;

  // Group by
  result := result + ' GROUP BY G_GENERAL, G_LIBELLE';
  if XXRupture.Text <> '' then
    result := result + ', ' + XXRupture.Text ;

end;

function TOF_CPBALGENAUXI.GenererRequeteCombisCompar(vStCode, vStLib, vStRup: String): String;
var	listeCptExcept  : String ;
    stCpt           : String ;
begin
// INSERT des comptes
  result := 'INSERT INTO ' + GetTablePourBase('CEDTBALANCE') + ' (CED_USER, CED_COMPTE, CED_LIBELLE';
  result := result + ',CED_CREDIT1, CED_DEBIT1, CED_CREDIT2, CED_DEBIT2';
  result := result + ', CED_RUPTURE, CED_COLLECTIF, CED_COMPTE2, CED_LIBELLE2) ';
// SELECT des comptes
  // Champs
  result := result + 'SELECT "' + V_PGI.User + '", E_AUXILIAIRE, T_LIBELLE';
  result := result + ',0, 0, 0, 0';
  // Rupture
  result := result + ', "' + vStRup + '"' ;
  // Collectif, compte2, libelle2
  result := result + ', "-", "' + vStCode + '", "' + CheckdblQuote(vStLib) + '"' ;
  // Table Temp + Compte
  result := result + ' FROM ' + GetTablePourBase('ECRITURE');
  // Table Ecriture si besoin
  result := result + ' LEFT JOIN ' + GetTablePourBase('TIERS') + ' ON E_AUXILIAIRE = T_AUXILIAIRE';
  // Conditions du WHERE (sur les comptes uniquement donc)
    // Comptes de ... à
  //result := result + ' WHERE E_AUXILIAIRE>="' + Compte2De.Text + '"';
  //result := result + ' AND E_AUXILIAIRE<="' + Compte2A.Text + '"';
  {JP 01/07/05 : Gestion des caractères jokers}
  Result := Result + ' WHERE ' + ConvertitCaractereJokers(Compte2De, Compte2A, 'E_AUXILIAIRE');

  result := result + ' AND E_GENERAL="' + vStCode + '"';
   // Conditions sur les écritures :
  result := result + ' AND ' + RecupWhereBalanceCompar ;
  { FQ 19715 BVE 07.06.07 }
  if ModeSelect.value = 'NSL' then
     result := result + ' AND ((SELECT SUM(E_CREDIT)-SUM(E_DEBIT)'
                      + ' FROM ' + GetTablePourBase('ECRITURE')
      	              + ' WHERE E_AUXILIAIRE = T_AUXILIAIRE) <> 0 )';
  { END FQ 19715 }
  // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
  if CptExcept2.Text <> '' then
    begin
    listeCptExcept := FindEtReplace(CptExcept2.Text,',',';', True);
    while (listeCptExcept <> '') do
      begin
      stCpt := Trim(ReadTokenSt(listeCptExcept)) ;
      if stCpt <> '' then
	result := result + ' AND E_AUXILIAIRE NOT LIKE "' + stCpt + '%"';
      end ;
    end;

  // Group by (??)
  result := result + ' GROUP BY E_AUXILIAIRE, T_LIBELLE';

end;

function TOF_CPBALGENAUXI.GenererRequeteComptesCompar: String;
var Q		    : TQuery ;
    DateDebExo      : TDatetime ;
    DateFinExo      : TDatetime ;
    listeCptExcept  : String ;
    stCpt           : String ;
begin
  result := 'SELECT G_GENERAL CODE, G_LIBELLE LIBELLE';
  // Rupture ?
  if XXRupture.Text <> ''
    then result := result + ', ' + XXRupture.Text + ' RUPTURE'
    else result := result + ', "" RUPTURE';
  // Table Temp + Compte
  result := result + ' FROM ' + GetTablePourBase('GENERAUX');
  // Table Ecriture si besoin
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then
    result := result + ' LEFT JOIN ' + GetTablePourBase('ECRITURE') + ' ON E_GENERAL = G_GENERAL' ;
  // Conditions du WHERE (sur les comptes uniquement donc)
    // Comptes de ... à
  //result := result + ' WHERE G_GENERAL>="' + CompteDe.Text + '"';
  //result := result + ' AND G_GENERAL<="' + CompteA.Text + '"';
  {JP 01/07/05 : Gestion des caractères jokers}
  Result := Result + ' WHERE ' + ConvertitCaractereJokers(CompteDe, CompteA, 'G_GENERAL');

  // Table Ecriture si besoin
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then
    result := result + ' AND ' + RecupWhereBalanceCompar( False ) ;
    // Uniquement des comptes collectifs et non fermés
  result := result + ' AND G_COLLECTIF="X" AND G_FERME="-"' ;
    // Nature de comptes
  if NatureCPT.Value <> '' then
    result := result + ' AND G_NATUREGENE="' + NatureCPT.Value + '"';
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
  				 + ' WHERE E_GENERAL = G_GENERAL AND ' + RecupWhereBalanceCompar + ') <> 0 )';
  // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
  if CptExcept.Text <> '' then
    begin
    listeCptExcept := FindEtReplace(CptExcept.Text,',',';', True);
    while (listeCptExcept <> '') do
      begin
      stCpt := Trim(ReadTokenSt(listeCptExcept)) ;
      if stCpt <> '' then
	result := result + ' AND G_GENERAL NOT LIKE "' + stCpt + '%"';
      end ;
    end;

  // Conditions sur champ de rupture
  result := result + GetConditionRupture ;

  // Group by
  result := result + ' GROUP BY G_GENERAL, G_LIBELLE';
  if XXRupture.Text <> '' then
    result := result + ', ' + XXRupture.Text ;

end;

function TOF_CPBALGENAUXI.GetNomChampCPT: String;
begin
  Result := 'CED_AUXILIAIRE' ;
end;

function TOF_CPBALGENAUXI.GetNomChampCPTCombi: String;
begin
  Result := 'CED_GENERAL';
end;

procedure TOF_CPBALGENAUXI.CompteOnExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if Trim(THEdit(Sender).Text) = '' then Exit;
  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  // Complétion auto du numéro de compte si possible
  if not CompleteAuto( Sender, fbGene ) then
    THEdit(Sender).ElipsisClick(nil);
end;

procedure TOF_CPBALGENAUXI.Compte2OnExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if Trim(THEdit(Sender).Text) = '' then Exit;
  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  // Complétion auto du numéro de compte si possible
  if not CompleteAuto( Sender, fbAux ) then
    THEdit(Sender).ElipsisClick(nil);
end;

function TOF_CPBALGENAUXI.GenererRequeteCombisBalsit(vStCode, vStLib, vStRup: String): String;
var	listeCptExcept  : String ;
    stCpt           : String ;
begin
// INSERT des comptes
  result := 'INSERT INTO ' + GetTablePourBase('CEDTBALANCE') + ' (CED_USER, CED_COMPTE, CED_NATURE, CED_LIBELLE';
  result := result + ',CED_CREDIT1, CED_DEBIT1, CED_CREDIT2, CED_DEBIT2';
  result := result + ', CED_RUPTURE, CED_COLLECTIF, CED_COMPTE2, CED_LIBELLE2) ';
// SELECT des comptes
  // --> Alimentation des champs de base
  result := result + 'SELECT "' + V_PGI.User + '", BSE_COMPTE2, "", T_LIBELLE, 0, 0';
  // --> Alimentation des montants
  if Affichage.Value = 'EURO'
    then result := result + ', BSE_Credit, BSE_Debit'         // montants Euro
    else result := result + ', BSE_CreditDev, BSE_DebitDev';  // montants Devise
  // Rupture
  result := result + ', "' + vStRup + '"' ;
  // Collectif, compte2, libelle2
  result := result + ', "-", "' + vStCode + '", "' + CheckdblQuote(vStLib) + '"' ;

  // -->  Tables
  result := result + ' FROM CBALSITECR' ;
  result := result + ' LEFT JOIN ' + GetTablePourBase('TIERS') + ' ON BSE_COMPTE2 = T_AUXILIAIRE';
  // --> conditions
  result := result + ' WHERE BSE_CODEBAL="' + BALSIT.text + '"';
  // Conditions du WHERE (sur les comptes uniquement donc)
  // *> Compte géné ?
  result := result + ' AND BSE_COMPTE1="' + vStCode + '"';
  // *> Compte auxi ?
//  result := result + ' AND BSE_COMPTE2>="' + Compte2De.text + '"';
//  result := result + ' AND BSE_COMPTE2<="' + Compte2A.text + '"';
  {JP 01/07/05 : Gestion des caractères jokers}
  Result := Result + ' AND ' + ConvertitCaractereJokers(Compte2De, Compte2A, 'BSE_COMPTE2');
  { FQ 19715 BVE 07.06.07 }
  if ModeSelect.value = 'NSL' then
     result := result + ' AND ((SELECT SUM(E_CREDIT)-SUM(E_DEBIT)'
                      + ' FROM ' + GetTablePourBase('ECRITURE')
      	              + ' WHERE E_AUXILIAIRE = T_AUXILIAIRE) <> 0 )';
  { END FQ 19715 }
  // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
  if CptExcept2.Text <> '' then
    begin
    listeCptExcept := FindEtReplace(CptExcept2.Text,',',';', True);
    while (listeCptExcept <> '') do
      begin
      stCpt := Trim(ReadTokenSt(listeCptExcept)) ;
      if stCpt <> '' then
	result := result + ' AND BSE_COMPTE2 NOT LIKE "' + stCpt + '%"';
      end ;
    end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 31/01/2006
Modifié le ... : 12/09/2006
Description .. : Ne pas prendre les comptes auxi fermés FQ 17207
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALGENAUXI.RemplirEDTBALANCE;
begin
inherited   ;
  //YMO 12/09/2006 Ne pas éliminer les enregs avec un débit ou un crédit <> 0
  //YMO 22/09/2006 FQ18810 correction+prise en compte de l'utilisateur
  //YMO 29/09/2006 FQ18810 Suppression aussi des cptes generaux fermés ET non mouvementés
  IF ExecuteSQL('DELETE FROM CEDTBALANCE WHERE CED_USER = "' + V_PGI.User + '"'
  +' AND (CED_DEBIT1=0 AND CED_CREDIT1=0)'
  +' AND ((CED_COMPTE =(SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE = CED_COMPTE AND T_FERME = "X"))'
  +' OR  (CED_COMPTE2=(SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = CED_COMPTE2 AND G_FERME = "X"))'
  +' AND CED_COMPTE="")')>0 then
      AuxiFermes.Text := 'X'
  else
      AuxiFermes.Text := '-';
  
end;


Initialization
  registerclasses ( [ TOF_CPBALGENAUXI ] ) ;
end.
