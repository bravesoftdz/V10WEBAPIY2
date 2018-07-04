{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 30/07/2003
Modifié le ... : 31/07/2003
Description .. : Source TOF de la FICHE : CPBALANALGEN
Suite ........ :
Suite ........ : Gestion du QRS1 de la balance analytique par général
Suite ........ :
Suite ........ : Attention : les traitements communs aux balances
Suite ........ : combinatoires sont dans une tof intermédiaire :
Suite ........ : TOF_METH --> TOF_CPBALANCECOMBI -->
Suite ........ : TOF_CPBALANALGEN
Suite ........ :
Suite ........ : JP 01/07/05 : Gestion des caractères Joker : fonctions de
Suite ........ : base définies dans TofMeth
Suite ........ :
Mots clefs ... : TOF;CPBALANALGEN
*****************************************************************}
Unit CPBALANALGEN_TOF ;

Interface

Uses StdCtrls, Classes,Windows,
{$IFDEF EAGLCLIENT}
     MainEAgl, utob, eQRS1,
{$ELSE}
     db,Fe_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     sysutils, Spin, Ent1,
     HCtrls, HEnt1, UTOF, hqry,
     CritEdt,      //ClassCritEdt
     TofMeth,
     uLibExercice,  // CRelatifVersExercice
     uLibWindows,   // TraductionTHMultiValComboBox
     CPBALANCECOMBI_TOF,
     HZoomSp,
     utilPGI  //b FP TSQLAnaCroise
      ;

procedure CPLanceFiche_BalanceAnalGen ( Args : String = '' ) ;

Type
  TOF_CPBALANALGEN = Class (TOF_CPBALANCECOMBI)
    SectionsFermees             : THEdit;
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
    procedure onCompteKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
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
    // -> Construction des conditions du WHERE pour les cas spéciaux
    function  RecupWhereQualifPiece : String ;   override ; // pour éviter les 'like 'xx%' or like...)
    function  RecupWhereBalance( vBoAvecDate : Boolean = True ) : String; override;
    function  RecupWhereBalanceCompar( vBoAvecDate : Boolean = True ) : String ; override ; // Condition sur la période comparée
    // Gestion CritEdt
    procedure ChargementCritEdt             ; override ;
    procedure RemplirEDTBALANCE                ; override ;
    private
        fOnSaveKeyDownCompte : procedure(Sender: TObject; var Key: Word; Shift:
      TShiftState) of object;
        fOnSaveKeyDownCompte_ : procedure(Sender: TObject; var Key: Word; Shift:
      TShiftState) of object;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  AglInit;      // TheData



//==============================================================================
procedure CPLanceFiche_BalanceAnalGen ( Args : String ) ;
begin
	AGLLanceFiche('CP','CPBALANALGEN','','',Args);
end;
//==============================================================================

{TOF_CPBALANALGEN}

//==============================================================================
//====================         EVT TOF        ==================================
//==============================================================================

procedure TOF_CPBALANALGEN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPBALANALGEN.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPBALANALGEN.OnLoad ;
begin

  Inherited ;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if not TestJoker(CompteDe.Text) then begin
    // Auto-remplissage si comptes non renseignés
    if (Trim(CompteDe.Text) = '') then
      CompteDe.Text := GetMinCompte('SECTION', 'S_SECTION', NatureCpt.Value);
    if (Trim(CompteA.Text) = '') then
      CompteA.Text := GetMaxCompte('SECTION', 'S_SECTION', NatureCpt.Value);
  end;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if not TestJoker(Compte2De.Text) then begin
    // Auto-remplissage si comptes généraux non renseignés
    if (Trim(Compte2De.Text) = '') then
      Compte2De.Text := GetMinCompteVentil( '', GetNumAxe );
    if (Trim(Compte2A.Text) = '') then
      Compte2A.Text := GetMaxCompteVentil( '', GetNumAxe );
  end;

  // Les ruptures...
  if Rupture.Value <> 'SANS' then
    begin
    // rupture sur n° compte
    if (RuptureType.Value = 'RUPGROUPES') and (NivoRupture.value > 0) then
      AvecNivoRupture.Text := IntToStr(NivoRupture.value);
    // rupture sur champ libre tiers
    if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
      XXRupture.Text := 'S_TABLE' + GetNumTableLibre ;
    // rupture sur plan de correspondance
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
      XXRupture.Text := 'S_CORRESP' + IntToStr(Corresp.ItemIndex + 1) ;
    end;

end ;

procedure TOF_CPBALANALGEN.OnArgument (S : String ) ;
begin

  // CONTROLES SPECIFIQUES
  // Contrôles critères et options - Page standard
  Exercice     := THValComboBox(Getcontrol('Y_EXERCICE', true));
  // Contrôles critères et options - Page avancés
  Devise       := THValComboBox(Getcontrol('Y_DEVISE', true));
  SectionsFermees := THEdit(GetControl('SECTIONSFERMEES', true));
  // Controles communs
  Inherited ;

  // Rubrique d'aide
  Ecran.HelpContext := 7469000;

  // Type de plan comptable :
  Case V_PGI.LaSerie of
    S7 :  begin
          Corresp.plus := 'AND (CO_CODE = "AU1" OR CO_CODE = "AU2")';
          end;
    else  begin
          Corresp.plus := 'AND CO_CODE = "AU1"';
          end;
    end;

  // Longueur des comptes
  Compte2De.MaxLength := VH^.CPta[fbGene].Lg;
  Compte2A.MaxLength  := VH^.CPta[fbGene].Lg;

  { CA - 06/10/2005 - Pour la saisie des axes structurés }
  if assigned(CompteDe) then
  begin
    fOnSaveKeyDownCompte := CompteDe.OnKeyDown;
    CompteDe.OnKeyDown := onCompteKeyDown;
  end;
  if assigned(CompteA) then
  begin
    fOnSaveKeyDownCompte_ := CompteA.OnKeyDown;
    CompteA.OnKeyDown := onCompteKeyDown;
  end;

end ;

//==============================================================================
//====================        EVT QRS1        ==================================
//==============================================================================

procedure TOF_CPBALANALGEN.NatureCptChanged(Sender: TObject);
var stAxe : String ;
    lTFB  : TFichierBase ;
begin
  // ---------------------
  // Au changement d'axe :
  // ---------------------
  if NatureCpt.ItemIndex < 0 then Exit ;
  stAxe := GetNumAxe ;

  // 1. Type de section
  CompteDe.Text := '';
  CompteA.Text := '';
  if stAxe = '1' then
    begin
    CompteDe.DataType := 'TZSECTION' ;
    CompteA.DataType  := 'TZSECTION' ;
    end
  else
    begin
    CompteDe.DataType := 'TZSECTION' + stAxe ;
    CompteA.DataType := 'TZSECTION' + stAxe ;
    end ;

  // 2. Type de généraux
  Compte2De.Text := '';
  Compte2A.Text := '';
  Compte2De.DataType := 'TZGVentil' + stAxe ;
  Compte2A.DataType  := 'TZGVentil' + stAxe ;

  // 3. Type de Plan de correspondance
  Case V_PGI.LaSerie of
    S7 :  begin
          Corresp.plus := 'AND (CO_CODE = "A'  + stAxe + '1" OR CO_CODE = "A' + stAxe + '2")';
          end;
    else  begin
          Corresp.plus := 'AND CO_CODE = "A' + stAxe + '1"';
          end;
    end;

  // 3. Longueur des sections
  lTFB := AxeToFb( 'A' + stAxe ) ;
  CompteDe.MaxLength  := VH^.Cpta[lTFB].Lg ;
  CompteA.MaxLength   := VH^.Cpta[lTFB].Lg ;

end;

procedure TOF_CPBALANALGEN.TableLibreChanged(Sender: TObject);
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
    LibreDe.DataType  := 'TZNATSECT' + GetNumTableLibre ;
    LibreA.DataType   := 'TZNATSECT' + GetNumTableLibre ;
    LibreDe.Text      := '';
    LibreA.Text       := '';
    end;
end;

//==============================================================================
//====================   TRAITEMENTS DES DONNEES  ==============================
//==============================================================================

procedure TOF_CPBALANALGEN.CalculDebitCredit( vStCompte, vStCompte2: String ; vBoCompar : Boolean ; var vDebit, vCredit: Double );
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
    then lStReq := 'SELECT SUM(Y_DEBIT) D, SUM(Y_CREDIT) C'
    else lStReq := 'SELECT SUM(Y_DEBITDEV) D, SUM(Y_CREDITDEV) C';
  // --> Table
  lStReq := lStReq + ' FROM ' + GetTablePourBase('ANALYTIQ') ;
  // --> Conditions
  lStReq := lStReq + ' WHERE Y_GENERAL = "' + vStCompte + '"' ;
  {b FP}
  //lStReq := lStReq + ' AND Y_SECTION = "' + vStCompte2 + '"' ;
  lStReq := lStReq + ' AND '+TSQLAnaCroise.ChampSection(NatureCpt.Value)+' = "' + vStCompte2 + '"' ;
  {e FP}
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


function TOF_CPBALANALGEN.GenererRequeteCombis(vStCode, vStLib, vStRup: String): String;
var listeCptExcept  : String ;
    stCpt           : String ;
begin
// INSERT des comptes
  result := 'INSERT INTO ' + GetTablePourBase('CEDTBALANCE') + ' ( CED_USER, CED_COMPTE, CED_LIBELLE';
  result := result + ',CED_CREDIT1, CED_DEBIT1, CED_CREDIT2, CED_DEBIT2';
  result := result + ', CED_NATURE'; // GCO - 29/11/2006 - 18927
  result := result + ', CED_RUPTURE, CED_COLLECTIF, CED_COMPTE2, CED_LIBELLE2) ';
// SELECT des comptes
  // Champs
  result := result + 'SELECT "' + V_PGI.User + '", Y_GENERAL, G_LIBELLE';
  result := result + ',0, 0, 0, 0';
  result := result + ', "A' + GetNumAxe + '"'; // GCO - 29/11/2006 - 18927

  // Rupture
  result := result + ', "' + vStRup + '"' ;
  // Collectif, compte2, libelle2
  result := result + ', "-", "' + vStCode + '", "' + CheckdblQuote(vStLib) + '"' ;
  // Table Temp + Compte
  result := result + ' FROM ' + GetTablePourBase('ANALYTIQ');
  result := result + ' LEFT JOIN GENERAUX ON G_GENERAL = Y_GENERAL';
  // Conditions du WHERE (sur les comptes uniquement donc)
    // Comptes de ... à

  {JP 01/07/05 : Gestion des caractères jokers}
  Result := Result + ' WHERE ' + ConvertitCaractereJokers(Compte2De, Compte2A, 'Y_GENERAL');
//  result := result + ' WHERE Y_GENERAL>="' + Compte2De.Text + '"';
  //result := result + ' AND Y_GENERAL<="' + Compte2A.Text + '"';

  {b FP}
  //result := result + ' AND Y_SECTION="' + vStCode + '"';
  result := result + ' AND '+TSQLAnaCroise.ChampSection(NatureCpt.Value)+'="' + vStCode + '"';
  {e FP}
    // Conditions sur les écritures :
  result := result + ' AND ' + RecupWhereQualifPiece ;
  result := result + FindEtReplace(RecupWhereBalance,'WHERE',' AND', False) ;
  // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
  if CptExcept2.Text <> '' then
    begin
    listeCptExcept := FindEtReplace(CptExcept2.Text,',',';', True);
    while (listeCptExcept <> '') do
      begin
      stCpt := Trim(ReadTokenSt(listeCptExcept)) ;
      if stCpt <> '' then
	result := result + ' AND Y_GENERAL NOT LIKE "' + stCpt + '%"';
      end ;
    end;

  // SI COMPARATIF, on se contente d'ajouter les comptes non présent dans CEDTBALANCE
  if AvecComparatif.Checked then
    Result := result + ' AND (Y_GENERAL not in (SELECT CED_COMPTE FROM ' + GetTablePourBase('CEDTBALANCE')
                     + ' WHERE CED_USER = "' + V_PGI.User + '"'
                     + ' AND CED_COMPTE2 = "' + vStCode + '") )' ;

  // Group by (??)
  result := result + ' GROUP BY Y_GENERAL, G_LIBELLE';

end;

function TOF_CPBALANALGEN.GenererRequeteCombisCompar(vStCode, vStLib, vStRup: String): String;
var	listeCptExcept  : String ;
    stCpt           : String ;
begin
// INSERT des comptes
  result := 'INSERT INTO ' + GetTablePourBase('CEDTBALANCE') + ' (CED_USER, CED_COMPTE, CED_LIBELLE';
  result := result + ',CED_CREDIT1, CED_DEBIT1, CED_CREDIT2, CED_DEBIT2';
  result := result + ', CED_NATURE'; // GCO - 29/11/2006 - 18927
  result := result + ', CED_RUPTURE, CED_COLLECTIF, CED_COMPTE2, CED_LIBELLE2) ';
// SELECT des comptes
  // Champs
  result := result + 'SELECT "' + V_PGI.User + '", Y_GENERAL, G_LIBELLE';
  result := result + ',0, 0, 0, 0';
  result := result + ', "A' + GetNumAxe + '"'; // GCO - 29/11/2006 - 18927
  // Rupture
  result := result + ', "' + vStRup + '"' ;
  // Collectif, compte2, libelle2
  result := result + ', "-", "' + vStCode + '", "' + CheckdblQuote(vStLib) + '"' ;
  // Table Temp + Compte
  result := result + ' FROM ' + GetTablePourBase('ANALYTIQ');
  // Table Ecriture si besoin
  result := result + ' LEFT JOIN ' + GetTablePourBase('GENERAUX') + ' ON G_GENERAL = Y_GENERAL';
  // Conditions du WHERE (sur les comptes uniquement donc)
    // Comptes de ... à
  {JP 01/07/05 : Gestion des caractères jokers}
  Result := Result + ' WHERE ' + ConvertitCaractereJokers(Compte2De, Compte2A, 'Y_GENERAL');
//  result := result + ' WHERE Y_GENERAL>="' + Compte2De.Text + '"';
  //result := result + ' AND Y_GENERAL<="' + Compte2A.Text + '"';

  {b FP}
  //result := result + ' AND Y_SECTION="' + vStCode + '"';
  result := result + ' AND '+TSQLAnaCroise.ChampSection(NatureCpt.Value)+'="' + vStCode + '"';
  {e FP}
   // Conditions sur les écritures :
  result := result + ' AND ' + RecupWhereBalanceCompar ;
  // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
  if CptExcept2.Text <> '' then
  	begin
    listeCptExcept := FindEtReplace(CptExcept2.Text,',',';', True);
		while (listeCptExcept <> '') do
    	begin
      stCpt := Trim(ReadTokenSt(listeCptExcept)) ;
      if stCpt <> '' then
	  		result := result + ' AND Y_GENERAL NOT LIKE "' + stCpt + '%"';
      end ;
    end;

  // Group by (??)
	result := result + ' GROUP BY Y_GENERAL, G_LIBELLE';

end;

function TOF_CPBALANALGEN.GenererRequeteComptes: String;
var Q		    : TQuery ;
    DateDebExo      : TDatetime ;
    DateFinExo      : TDatetime ;
    listeCptExcept  : String ;
    stCpt           : String ;
begin
  result := 'SELECT S_SECTION CODE, S_LIBELLE LIBELLE';
  // Rupture ?
  if XXRupture.Text <> ''
    then result := result + ', ' + XXRupture.Text + ' RUPTURE'
    else result := result + ', "" RUPTURE';
  // Table Temp + Compte
  result := result + ' FROM ' + GetTablePourBase('SECTION');
  // Table Ecriture si besoin
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then begin
    {b FP}
    result := result + ' LEFT JOIN ' + GetTablePourBase('ANALYTIQ') +
                //' ON Y_SECTION = S_SECTION' ;
                ' ON '+TSQLAnaCroise.ChampSection(NatureCpt.Value)+' = S_SECTION' ;
    {e FP}
  end;
  // Conditions du WHERE (sur les comptes uniquement donc)
    // Comptes de ... à

  {JP 01/07/05 : Gestion des caractères jokers}
  Result := Result + ' WHERE ' + ConvertitCaractereJokers(CompteDe, CompteA, 'S_SECTION');
//  result := result + ' WHERE S_SECTION>="' + CompteDe.Text + '"';
  //result := result + ' AND S_SECTION<="' + CompteA.Text + '"';

  // Table Ecriture si besoin
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then
    result := result + ' AND ' + RecupWhereQualifPiece
                     + FindEtReplace(RecupWhereBalance(False),'WHERE',' AND', False) ;
    // Uniquement des comptes non fermés
//YMO  result := result + ' AND S_FERME="-"' ;
    // Nature de comptes
  if NatureCPT.Value <> '' then
    result := result + ' AND S_AXE="' + NatureCPT.Value + '"';
  // Mode de sélection des comptes
  if ModeSelect.value = 'EXO' then
    begin
    Q:=OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="' + CRelatifVersExercice(Exercice.value) + '"' ,TRUE) ;
    if Not Q.EOF then
      begin
      DateDebExo := Q.FindField('EX_DATEDEBUT').asDateTime ;
      DateFinExo := Q.FindField('EX_DATEFIN').asDateTime ;
      result     := result + ' AND Y_DATECOMPTABLE>="' + USDateTime(DateDebExo)
      			   + '" AND Y_DATECOMPTABLE<="' + USDateTime(DateFinExo) + '"';
      end;
    Ferme(Q) ;
    end
  else if ModeSelect.value = 'PER' then
    result := result + ' AND Y_DATECOMPTABLE>="' + USDateTime(DateDebutPer)
                     + '" AND Y_DATECOMPTABLE<="' + USDateTime(DateFinPer) + '"'
  else if ModeSelect.value = 'NSL' then
    result := result + ' AND ( (SELECT SUM(Y_CREDIT)-SUM(Y_DEBIT) FROM ' + GetTablePourBase('ANALYTIQ')
                     {b FP}
  		     //+ ' WHERE Y_SECTION = S_SECTION AND ' + RecupWhereQualifPiece
                     + ' WHERE '+TSQLAnaCroise.ChampSection(NatureCpt.Value)+' = S_SECTION AND ' + RecupWhereQualifPiece
                     {e FP}
                     + FindEtReplace(RecupWhereBalance,'WHERE',' AND', False) + ') <> 0 )';
  // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
  if CptExcept.Text <> '' then
    begin
    listeCptExcept := FindEtReplace(CptExcept.Text,',',';', True);
    while (listeCptExcept <> '') do
      begin
      stCpt := Trim(ReadTokenSt(listeCptExcept)) ;
      if stCpt <> '' then
	result := result + ' AND S_SECTION NOT LIKE "' + stCpt + '%"';
      end ;
    end;
  // Conditions sur champ de rupture
  result := result + GetConditionRupture ;

  // Group by
  result := result + ' GROUP BY S_SECTION, S_LIBELLE';
  if XXRupture.Text <> '' then
    result := result + ', ' + XXRupture.Text ;

end;

function TOF_CPBALANALGEN.GenererRequeteComptesCompar: String;
var Q		    : TQuery ;
    DateDebExo      : TDatetime ;
    DateFinExo      : TDatetime ;
    listeCptExcept  : String ;
    stCpt           : String ;
begin
  result := 'SELECT S_SECTION CODE, S_LIBELLE LIBELLE';
  // Rupture ?
  if XXRupture.Text <> ''
    then result := result + ', ' + XXRupture.Text + ' RUPTURE'
    else result := result + ', "" RUPTURE';
  // Table Temp + Compte
  result := result + ' FROM ' + GetTablePourBase('SECTION');
    // Table Ecriture si besoin
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then begin
    {b FP}
    result := result + ' LEFT JOIN ' + GetTablePourBase('ANALYTIQ') +
      //' ON Y_SECTION = S_SECTION' ;
      ' ON '+TSQLAnaCroise.ChampSection(NatureCpt.Value)+' = S_SECTION' ;
    {e FP}
  end;
  // Conditions du WHERE (sur les comptes uniquement donc)
    // Comptes de ... à

  {JP 01/07/05 : Gestion des caractères jokers}
  Result := Result + ' WHERE ' + ConvertitCaractereJokers(CompteDe, CompteA, 'S_SECTION');
  //result := result + ' WHERE S_SECTION>="' + CompteDe.Text + '"';
  //result := result + ' AND S_SECTION<="' + CompteA.Text + '"';

  // Table Ecriture si besoin
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then
    result := result + ' AND ' + RecupWhereBalanceCompar( False ) ;
    // Uniquement des comptes non fermés
  result := result + ' AND S_FERME="-"' ;
    // Nature de comptes
  if NatureCPT.Value <> '' then
    result := result + ' AND S_AXE="' + NatureCPT.Value + '"';
    // Mode de sélection des comptes
  if ModeSelect.value = 'EXO' then
    begin
    Q:=OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+CRelatifVersExercice(ComparExo.value)+'"' ,TRUE) ;
    if Not Q.EOF then
      begin
      DateDebExo := Q.FindField('EX_DATEDEBUT').asDateTime ;
      DateFinExo := Q.FindField('EX_DATEFIN').asDateTime ;
      result     := result + ' AND Y_DATECOMPTABLE>="' + USDateTime(DateDebExo)
      			   + '" AND Y_DATECOMPTABLE<="' + USDateTime(DateFinExo) + '"';
      end;
    Ferme(Q) ;
    end
  else if ModeSelect.value = 'PER' then
    result := result + ' AND Y_DATECOMPTABLE>="' + USDateTime(DateDebutCompar)
                     + '" AND Y_DATECOMPTABLE<="' + USDateTime(DateFinCompar) + '"'
  else if ModeSelect.value = 'NSL' then
    result := result + ' AND ( (SELECT SUM(Y_CREDIT)-SUM(Y_DEBIT) FROM ' + GetTablePourBase('ANALYTIQ')
                            {b FP}
    			    //+ ' WHERE Y_SECTION = S_SECTION AND ' + RecupWhereBalanceCompar + ') <> 0 )';
                            + ' WHERE '+TSQLAnaCroise.ChampSection(NatureCpt.Value)+' = S_SECTION AND ' + RecupWhereBalanceCompar + ') <> 0 )';
                            {e FP}
  // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
  if CptExcept.Text <> '' then
    begin
    listeCptExcept := FindEtReplace(CptExcept.Text,',',';', True);
    while (listeCptExcept <> '') do
      begin
      stCpt := Trim(ReadTokenSt(listeCptExcept)) ;
      if stCpt <> '' then
	result := result + ' AND S_SECTION NOT LIKE "' + stCpt + '%"';
      end ;
    end;
  // Conditions sur champ de rupture
  result := result + GetConditionRupture ;

  // Group by
  result := result + ' GROUP BY S_SECTION, S_LIBELLE';
  if XXRupture.Text <> '' then
    result := result + ', ' + XXRupture.Text ;

end;

function TOF_CPBALANALGEN.RecupWhereBalanceCompar( vBoAvecDate: Boolean ) : String;
begin
  // Exercice
  Result := ' Y_EXERCICE = "' + CRelatifVersExercice(ComparExo.Value) + '" ' ;
  {b FP}
  // Axe
  //Result := result + ' AND Y_AXE = "' + NatureCpt.Value + '" ' ;
  Result := result + ' AND '+TSQLAnaCroise.ConditionAxe(NatureCpt.Value);
  {e FP}
  // Dates comptables
  if vBoAvecDate then
    Result := Result + ' AND Y_DATECOMPTABLE >= "' + UsDateTime(DateDebutCompar) + '" '
                     + ' AND Y_DATECOMPTABLE <= "' + UsDateTime(DateFinCompar) + '" ' ;
  // QualifPiece
  Result := Result + ' AND ' + RecupWhereQualifPiece ;
  // Etablissement
  if ComboEtab.Value<>'' then
    Result := Result + ' AND Y_ETABLISSEMENT = "' + ComboEtab.Value + '" ' ;
  // Devise
  if Devise.Value<>'' then
    Result := Result + ' AND Y_DEVISE = "' + Devise.Value + '" ' ;
end;

function TOF_CPBALANALGEN.RecupWhereQualifPiece: String;
Var lStSql   : String ;
    lstTmp   : String ;
begin
  TraductionTHMultiValComboBox( QualifPiece, lStSql , lstTmp, 'Y_QUALIFPIECE', False );
  Result := lStSql ;
end;

function TOF_CPBALANALGEN.GetNomChampCPT: String;
begin
  Result := 'CED_GENERAL' ;
end;

function TOF_CPBALANALGEN.GetNomChampCPTCombi: String;
begin
  Result := 'CED_SECTION' ;
end;

procedure TOF_CPBALANALGEN.CompteOnExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if Trim(THEdit(Sender).Text) = '' then Exit;
  if NatureCpt.ItemIndex < 0 then Exit ;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  // Complétion auto du numéro de compte si possible
  if not CompleteAuto( Sender, AxeToFb( NatureCpt.value ) ) then
    THEdit(Sender).ElipsisClick(nil);
end;

procedure TOF_CPBALANALGEN.Compte2OnExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  // Complétion auto du numéro de compte si possible
  if not CompleteAuto( Sender, fbGene ) then
    THEdit(Sender).ElipsisClick(nil);
end;

function TOF_CPBALANALGEN.RecupWhereBalance( vBoAvecDate: Boolean ) : String;
begin
  Result := RecupWhereCritere( Pages);
  Result := CMajRequeteExercice ( Exercice.Value, Result);
  {b FP}
  // Axe
  //Result := result + ' AND Y_AXE = "' + NatureCpt.Value + '" ' ; // FQ 15330 SBO 26/11/2005
  Result := result + ' AND '+TSQLAnaCroise.ConditionAxe(NatureCpt.Value);
  {e FP}
  // Dates comptables
  if vBoAvecDate then
    Result := Result + ' AND Y_DATECOMPTABLE >= "' + UsDateTime(StrToDate(DateComptaDe.Text)) + '" '
                     + ' AND Y_DATECOMPTABLE <= "' + UsDateTime(StrToDate(DateComptaA.Text)) + '" ' ;
end;

procedure TOF_CPBALANALGEN.ChargementCritEdt;
begin
  if (TheData <> nil) and (TheData is ClassCritEdt) then
  begin
    if ClassCritEdt(TheData).CritEdt.Bal.Axe <> '' then
      NatureCpt.Value := ClassCritEdt(TheData).CritEdt.Bal.Axe;
  end ;

  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 05/10/2005
Modifié le ... :   /  /    
Description .. : Gestion des accès au choix des sections dans les zones de 
Suite ........ : saisie des sections
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPBALANALGEN.onCompteKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  St : string;
  fb : TFichierBase ;
begin
  St := THCritMaskEdit(Sender).Text;
  fb := AxeToFb(NatureCpt.Value);
  if (Shift = []) and (Key = 187) then
  begin
    Key := 0;
    CompteA.Text := CompteDe.Text;
  end else if ((Shift=[ssCtrl]) And (Key=VK_F5)) then
  begin
    If (fb in [fbAxe1..fbAxe5]) And
       VH^.Cpta[fb].Structure and
       // GCO - 29/11/2006 - FQ 19175
       ExisteSQL('SELECT SS_AXE FROM STRUCRSE WHERE SS_AXE = "' + FBToAxe(fb) + '"')
    Then
    begin
      if ChoisirSousPlan( fb, St , True,taModif) then
      begin
        if ((THCritMaskEdit(Sender) = CompteA) and EstJoker(St)) then CompteDe.Text := St
        else THCritMaskEdit(Sender).Text := St;
      end;
      Key := 0;
    end;
  end;
  if THCritMaskEdit(Sender) = CompteDe then fOnSaveKeyDownCompte (Sender, Key, Shift)
  else fOnSaveKeyDownCompte_ (Sender, Key, Shift);
end;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 30/01/2006
Modifié le ... : 12/09/2006
Description .. : Ne pas prendre les sections fermées FQ 17206
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALANALGEN.RemplirEDTBALANCE;
begin
inherited   ;
  //YMO 12/09/2006 Ne pas éliminer les enregs avec un débit ou un crédit <> 0
  //YMO 22/09/2006 FQ18810 correction+prise en compte de l'utilisateur
  //YMO 29/09/2006 FQ18810 Suppression aussi des cptes generaux fermés ET non mouvementés
  IF ExecuteSQL('DELETE FROM CEDTBALANCE WHERE CED_USER = "' + V_PGI.User + '"'
  +' AND (CED_DEBIT1=0 AND CED_CREDIT1=0)'
  +' AND ((CED_COMPTE2 =(SELECT S_SECTION FROM SECTION WHERE S_SECTION = CED_COMPTE2 AND S_FERME = "X"))'
  +' OR  (CED_COMPTE=(SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = CED_COMPTE AND G_FERME = "X"))'
  +' AND CED_COMPTE2="")')>0 then
      SectionsFermees.Text := 'X'
  else
      SectionsFermees.Text := '-';

end;

Initialization
  registerclasses ( [ TOF_CPBALANALGEN ] ) ;
end.
