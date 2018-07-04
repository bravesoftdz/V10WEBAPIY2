{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 30/07/2003
Modifié le ... : 31/07/2003
Description .. : Source TOF de la FICHE : CPBALGENANAL
Suite ........ :
Suite ........ : Gestion du QRS1 de la balance générale par analytique
Suite ........ :
Suite ........ : Attention : les traitements communs aux balances
Suite ........ : combinatoires sont dans une tof intermédiaire :
Suite ........ : TOF_METH --> TOF_CPBALANCECOMBI -->
Suite ........ : TOF_CPBALGENANAL
Suite ........ :
Suite ........ : JP 01/07/05 : Gestion des caractères Joker : fonctions de
Suite ........ : base définies dans TofMeth
Mots clefs ... : TOF;CPBALGENANAL
*****************************************************************}
Unit CPBALGENANAL_TOF ;

Interface

Uses StdCtrls, Classes, Windows,
{$IFDEF EAGLCLIENT}
     MainEAgl, utob, eQRS1,
{$ELSE}
     db, Fe_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     sysutils, Spin, Ent1,
     HCtrls, HEnt1, UTOF,
     AglInit,      // TheData
     CritEdt,      //ClassCritEdt
     TofMeth,
     uLibExercice,  // CRelatifVersExercice
     HQry,          // RecupWhereCritere
     uLibWindows,   // TraductionTHMultiValComboBox
     CPBALANCECOMBI_TOF,
     HZoomSp
      ;

procedure CPLanceFiche_BalanceGenAnal ( Args : String = '' ) ;

Type
  TOF_CPBALGENANAL = Class (TOF_CPBALANCECOMBI)
    Axe       : THValComboBox ;
    SectionsFermees             : THEdit;
    // EVT TOF
    procedure OnNew                    ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    // EVT FICHE
    procedure NatureCptChanged  (Sender : TObject);   override ;
    procedure TableLibreChanged (Sender : TObject);   override ;
    procedure AxeChanged        (Sender : TObject) ;  override ;
    procedure CompteOnExit      (Sender: TObject) ;   override ;
    procedure Compte2OnExit     (Sender: TObject) ;   override ;
    procedure onCompte2KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    // Filtres
    procedure InitZones ;                     override ;
    procedure ChargementCritEdt             ; override ;

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
    function  GetNumAxe : String ; override ;
    procedure RemplirEDTBALANCE                ; override ;
    private
        fOnSaveKeyDownCompte2 : procedure(Sender: TObject; var Key: Word; Shift:
      TShiftState) of object;
        fOnSaveKeyDownCompte2_ : procedure(Sender: TObject; var Key: Word; Shift:
      TShiftState) of object;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  utilPGI;  // TSQLAnaCroise


//==============================================================================
procedure CPLanceFiche_BalanceGenAnal ( Args : String ) ;
begin
	AGLLanceFiche('CP','CPBALGENANAL','','',Args);
end;
//==============================================================================

{TOF_CPBALGENANAL}

//==============================================================================
//====================         EVT TOF        ==================================
//==============================================================================

procedure TOF_CPBALGENANAL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPBALGENANAL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPBALGENANAL.OnLoad ;
begin

  Inherited ;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if not TestJoker(CompteDe.Text) then begin
    // Auto-remplissage si comptes généraux non renseignés
    if (Trim(CompteDe.Text) = '') then
      CompteDe.Text := GetMinCompteVentil( NatureCpt.Value, GetNumAxe );
    if (Trim(CompteA.Text) = '') then
      CompteA.Text := GetMaxCompteVentil( NatureCpt.Value, GetNumAxe );
  end;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if not TestJoker(Compte2De.Text) then begin
    // Auto-remplissage si sections non renseignées
    if (Trim(Compte2De.Text) = '') then
      Compte2De.Text := GetMinCompte('SECTION', 'S_SECTION', Axe.Value);
    if (Trim(Compte2A.Text) = '') then
      Compte2A.Text := GetMaxCompte('SECTION', 'S_SECTION', Axe.Value);
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

procedure TOF_CPBALGENANAL.OnArgument (S : String ) ;
begin

  // CONTROLES SPECIFIQUES
  // Contrôles critères et options - Page standard
  Exercice    := THValComboBox(Getcontrol('Y_EXERCICE', true));
  {b FP Renommer le controle dans décla?}
  Axe         := THValComboBox(Getcontrol('Y_AXE', true));
  Axe.Name    := 'YAXE';   {La fonction RecupWhereCritere n'ajoutera pas la condition sur l'axe}
  {e FP}
  // Contrôles critères et options - Page avancés
  Devise      := THValComboBox(Getcontrol('Y_DEVISE', true));
  SectionsFermees := THEdit(GetControl('SECTIONSFERMEES', true));
  // Controles communs
  Inherited ;

  // Tablette des natures de comptes généraux : Tout !
  NatureCpt.DataType := 'TTNATGENE' ;

  // Evt spécifique
  Axe.OnChange := AxeChanged ;

  // Rubrique d'aide
  Ecran.HelpContext:=7466000;

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
  CompteDe.MaxLength  := VH^.CPta[fbGene].Lg;
  CompteA.MaxLength   := VH^.CPta[fbGene].Lg;

  { CA - 06/10/2005 - Pour la saisie des axes structurés }
  if assigned(Compte2De) then
  begin
    fOnSaveKeyDownCompte2 := Compte2De.OnKeyDown;
    Compte2De.OnKeyDown := onCompte2KeyDown;
  end;
  if assigned(Compte2A) then
  begin
    fOnSaveKeyDownCompte2_ := Compte2A.OnKeyDown;
    Compte2A.OnKeyDown := onCompte2KeyDown;
  end;
end ;

//==============================================================================
//====================        EVT QRS1        ==================================
//==============================================================================

procedure TOF_CPBALGENANAL.NatureCptChanged(Sender: TObject);
begin
  CompteDe.Text := '';
  CompteA.Text := '';
  if NatureCpt.ItemIndex > 0 then
    begin
    CompteDe.Plus := ' AND G_NATUREGENE="'+NatureCPT.Value+'"' ;
    CompteA.Plus := ' AND G_NATUREGENE="'+NatureCPT.Value+'"' ;
    end
  else
    begin
    CompteDe.Plus := '' ;
    CompteA.Plus := '';
    end ;
end;

procedure TOF_CPBALGENANAL.TableLibreChanged(Sender: TObject);
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
    LibreDe.DataType := 'TZNATGENE' + GetNumTableLibre ;
    LibreA.DataType  := 'TZNATGENE' + GetNumTableLibre ;
    LibreDe.Text     := '';
    LibreA.Text      := '';
    end;
end;

//==============================================================================
//====================   TRAITEMENTS DES DONNEES  ==============================
//==============================================================================


procedure TOF_CPBALGENANAL.CalculDebitCredit( vStCompte, vStCompte2: String ; vBoCompar : Boolean ; var vDebit, vCredit: Double );
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
  {b FP}
  // --> Conditions
  //lStReq := lStReq + ' WHERE Y_SECTION = "' + vStCompte + '"' ;
  lStReq := lStReq + ' WHERE '+TSQLAnaCroise.ChampSection(Axe.Value)+' = "' + vStCompte + '"' ;
  {e FP}
  lStReq := lStReq + ' AND Y_GENERAL = "' + vStCompte2 + '"' ;
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

function TOF_CPBALGENANAL.GenererRequeteCombis(vStCode, vStLib, vStRup: String): String;
var	listeCptExcept  : String ;
    stCpt           : String ;
begin
// INSERT des comptes
  result := 'INSERT INTO ' + GetTablePourBase('CEDTBALANCE') + ' (CED_USER, CED_COMPTE, CED_LIBELLE';
  result := result + ',CED_CREDIT1, CED_DEBIT1, CED_CREDIT2, CED_DEBIT2';
  result := result + ', CED_NATURE'; // GCO - 29/11/2006 - 18927
  result := result + ', CED_RUPTURE, CED_COLLECTIF, CED_COMPTE2, CED_LIBELLE2) ';
// SELECT des comptes
  {b FP}
  // Champs
  //result := result + 'SELECT "' + V_PGI.User + '", Y_SECTION, S_LIBELLE';
  result := result + 'SELECT "' + V_PGI.User + '", '+TSQLAnaCroise.ChampSection(Axe.Value)+
                     ', S_LIBELLE';
  {e FP}
  result := result + ',0, 0, 0, 0';
  result := result + ', "A' + GetNumAxe + '"'; // GCO - 29/11/2006 - 18927
  // Rupture
  result := result + ', "' + vStRup + '"' ;
  // Collectif, compte2, libelle2
  result := result + ', "-", "' + vStCode + '", "' + CheckdblQuote(vStLib) + '"' ;
  // Table Temp + Compte
  result := result + ' FROM ' + GetTablePourBase('ANALYTIQ');
  {b FP}
  result := result + ' LEFT JOIN ' + GetTablePourBase('SECTION') +
    //' ON S_SECTION = Y_SECTION'+
    ' ON S_SECTION = '+TSQLAnaCroise.ChampSection(Axe.Value)+
    ' AND S_AXE="' + Axe.Value + '" '; // FQ 15330 SBO 26/11/2005
  {e FP}
// WHERE (sur les comptes uniquement donc)
  // Comptes de ... à
//  result := result + ' WHERE Y_SECTION>="' + Compte2De.Text + '"';
  //result := result + ' AND Y_SECTION<="' + Compte2A.Text + '"';
  {JP 01/07/05 : Gestion des caractères jokers}
  {b FP}
  Result := Result + ' WHERE ' +
    //ConvertitCaractereJokers(Compte2De, Compte2A, 'Y_SECTION');
    ConvertitCaractereJokers(Compte2De, Compte2A, TSQLAnaCroise.ChampSection(Axe.Value));
  {e FP}

  result := result + ' AND Y_GENERAL="' + vStCode + '"';
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
        begin
        {b FP}
	//result := result + ' AND Y_SECTION NOT LIKE "' + stCpt + '%"';
        result := result + ' AND '+TSQLAnaCroise.ChampSection(Axe.Value)+' NOT LIKE "' + stCpt + '%"';
        {e FP}
        end;
      end ;
    end;
 
  // SI COMPARATIF, on se contente d'ajouter les comptes non présent dans CEDTBALANCE
  if AvecComparatif.Checked then begin
    Result := result
      {b FP}
      //+ ' AND (Y_SECTION not in (SELECT CED_COMPTE FROM ' + GetTablePourBase('CEDTBALANCE')
      + ' AND ('+TSQLAnaCroise.ChampSection(Axe.Value)+' not in (SELECT CED_COMPTE FROM ' + GetTablePourBase('CEDTBALANCE')
      {e FP}
      + ' WHERE CED_USER = "' + V_PGI.User + '"'
      + ' AND CED_COMPTE2 = "' + vStCode + '") )' ;
  end;
  // Group by (??)
  {b FP}
  //result := result + ' GROUP BY Y_SECTION, S_LIBELLE';
  result := result + ' GROUP BY '+TSQLAnaCroise.ChampSection(Axe.Value)+', S_LIBELLE';
  {e FP}
end;

function TOF_CPBALGENANAL.GenererRequeteCombisCompar(vStCode, vStLib, vStRup: String): String;
var listeCptExcept  : String ;
    stCpt           : String ;
begin
// INSERT des comptes
  result := 'INSERT INTO ' + GetTablePourBase('CEDTBALANCE') + ' (CED_USER, CED_COMPTE, CED_LIBELLE';
  result := result + ',CED_CREDIT1, CED_DEBIT1, CED_CREDIT2, CED_DEBIT2';
  result := result + ', CED_NATURE'; // GCO - 29/11/2006 - 18927
  result := result + ', CED_RUPTURE, CED_COLLECTIF, CED_COMPTE2, CED_LIBELLE2) ';
// SELECT des comptes
  {b FP}
  // Champs
  //result := result + 'SELECT "' + V_PGI.User + '", Y_SECTION, S_LIBELLE';
  result := result + 'SELECT "' + V_PGI.User + '", '+TSQLAnaCroise.ChampSection(Axe.Value)+', S_LIBELLE';
  {e FP}
  result := result + ',0, 0, 0, 0';
  result := result + ', "A' + GetNumAxe + '"'; // GCO - 29/11/2006 - 18927
  // Rupture
  result := result + ', "' + vStRup + '"' ;
  // Collectif, compte2, libelle2
  result := result + ', "-", "' + vStCode + '", "' + CheckdblQuote(vStLib) + '"' ;
  // Table Temp + Compte
  result := result + ' FROM ' + GetTablePourBase('ANALYTIQ');
  {b FP}
  result := result + ' LEFT JOIN ' + GetTablePourBase('SECTION') +
    //' ON S_SECTION = Y_SECTION AND S_AXE="' + Axe.Value + '" ';  // FQ 15330 SBO 26/11/2005
    ' ON S_SECTION = '+TSQLAnaCroise.ChampSection(Axe.Value)+' AND S_AXE="' + Axe.Value + '" ';
  {e FP}
  // Conditions du WHERE (sur les comptes uniquement donc)
    // Comptes de ... à
//  result := result + ' WHERE Y_SECTION>="' + Compte2De.Text + '"';
  //result := result + ' AND Y_SECTION<="' + Compte2A.Text + '"';
  {JP 01/07/05 : Gestion des caractères jokers}
  {b FP}
  //Result := Result + ' WHERE ' + ConvertitCaractereJokers(Compte2De, Compte2A, 'Y_SECTION');
  Result := Result + ' WHERE ' + ConvertitCaractereJokers(Compte2De, Compte2A, TSQLAnaCroise.ChampSection(Axe.Value));
  {e FP}

  result := result + ' AND Y_GENERAL="' + vStCode + '"';
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
        begin
        {b FP}
	//result := result + ' AND Y_SECTION NOT LIKE "' + stCpt + '%"';
        result := result + ' AND '+TSQLAnaCroise.ChampSection(Axe.Value)+' NOT LIKE "' + stCpt + '%"';
        {e FP}
        end;
      end ;
    end;

  {b FP}
  // Group by (??)
  //result := result + ' GROUP BY Y_SECTION, S_LIBELLE';
  result := result + ' GROUP BY '+TSQLAnaCroise.ChampSection(Axe.Value)+', S_LIBELLE';
  {e FP}
end;

function TOF_CPBALGENANAL.GenererRequeteComptes: String;
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
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then begin
    result := result + ' LEFT JOIN ' + GetTablePourBase('ANALYTIQ') +
    {b FP}
      // FQ 19763 : Pb DB2 sur Y_AXE apparaîssant 2 fois dans la requête
      ' ON Y_GENERAL = G_GENERAL ' ; //AND '+TSQLAnaCroise.ConditionAxe(Axe.Value);
    {e FP}
  end;

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
  // Uniquement des comptes non fermés
//  YMO 29/09/2006 les comptes fermé et non mouvementés sont supprimés dans CEDTBALANCE plus loin
//  result := result + ' AND G_FERME="-"' ;
  // Uniquement des ventilables sur l'axe choisi
  result := result + ' AND G_VENTILABLE="X" AND G_VENTILABLE' + GetNumAxe + '="X"' ;
  // Nature de comptes
  if NatureCPT.Value <> '' then
    result := result + ' AND G_NATUREGENE="' + NatureCPT.Value + '"';
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
    		     + ' WHERE Y_GENERAL = G_GENERAL AND ' + RecupWhereQualifPiece
                     + FindEtReplace(RecupWhereBalance,'WHERE',' AND', False) + ') <> 0 )';
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

function TOF_CPBALGENANAL.GenererRequeteComptesCompar: String;
var	Q								: TQuery ;
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
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then begin
    result := result + ' LEFT JOIN ' + GetTablePourBase('ANALYTIQ') +
             {b FP}
             // FQ 19763 : Pb DB2 sur Y_AXE apparaîssant 2 fois dans la requête
             ' ON Y_GENERAL = G_GENERAL ' // AND '+TSQLAnaCroise.ConditionAxe(Axe.Value);
             {e FP}
  end;
  // Conditions du WHERE (sur les comptes uniquement donc)
    // Comptes de ... à
//  result := result + ' WHERE G_GENERAL>="' + CompteDe.Text + '"';
  //result := result + ' AND G_GENERAL<="' + CompteA.Text + '"';
  {JP 01/07/05 : Gestion des caractères jokers}
  Result := Result + ' WHERE ' + ConvertitCaractereJokers(CompteDe, CompteA, 'G_GENERAL');

    // Table Ecriture si besoin
  if (ModeSelect.value = 'PER') or (ModeSelect.value = 'EXO') then
    result := result + ' AND ' + RecupWhereBalanceCompar( False ) ;
    // Uniquement des comptes non fermés
  result := result + ' AND G_FERME="-"' ;
    // Uniquement des ventilables sur l'axe choisi
  result := result + ' AND G_VENTILABLE="X" AND G_VENTILABLE' + GetNumAxe + '="X"' ;
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
      result := result + ' AND Y_DATECOMPTABLE>="' + USDateTime(DateDebExo)
      		       + '" AND Y_DATECOMPTABLE<="' + USDateTime(DateFinExo) + '"';
      end;
    Ferme(Q) ;
    end
  else if ModeSelect.value = 'PER' then
    result := result + ' AND Y_DATECOMPTABLE>="' + USDateTime(DateDebutCompar)
                     + '" AND Y_DATECOMPTABLE<="' + USDateTime(DateFinCompar) + '"'
  else if ModeSelect.value = 'NSL' then
    result := result + ' AND ( (SELECT SUM(Y_CREDIT)-SUM(Y_DEBIT) FROM ' + GetTablePourBase('ANALYTIQ')
    		     + ' WHERE Y_GENERAL = G_GENERAL AND ' + RecupWhereBalanceCompar + ') <> 0 )';
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

procedure TOF_CPBALGENANAL.AxeChanged(Sender: TObject);
var stAxe : String ;
    lTFB  : TFichierBase ;
begin
  // ---------------------
  // Au changement d'axe :
  // ---------------------
  if Axe.ItemIndex < 0 then Exit ;
  stAxe := GetNumAxe ;

  // 1. Type de section
  Compte2De.Text := '';
  Compte2A.Text := '';
  if stAxe = '1' then
    begin
    Compte2De.DataType := 'TZSECTION' ;
    Compte2A.DataType  := 'TZSECTION' ;
    end
  else
    begin
    Compte2De.DataType := 'TZSECTION' + stAxe ;
    Compte2A.DataType := 'TZSECTION' + stAxe ;
    end ;
  // 2. Type de généraux
  CompteDe.Text := '';
  CompteA.Text := '';
  CompteDe.DataType := 'tzGVentil' + stAxe ;
  CompteA.DataType  := 'tzGVentil' + stAxe ;
  // 3. Longueur des sections
  lTFB := AxeToFb( 'A' + stAxe ) ;
  Compte2De.MaxLength  := VH^.Cpta[lTFB].Lg ;
  Compte2A.MaxLength   := VH^.Cpta[lTFB].Lg ;

end;

function TOF_CPBALGENANAL.RecupWhereBalanceCompar( vBoAvecDate: Boolean ) : String;
begin
  // Exercice
  Result := ' Y_EXERCICE = "' + CRelatifVersExercice(ComparExo.Value) + '" ' ;
  {b FP}
  // Axe
  //Result := result + ' AND Y_AXE = "' + Axe.Value + '" ' ;
  Result := result + ' AND '+TSQLAnaCroise.ConditionAxe(Axe.Value);
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

function TOF_CPBALGENANAL.RecupWhereQualifPiece: String;
Var lStSql   : String ;
    lstTmp   : String ;
begin
  TraductionTHMultiValComboBox( QualifPiece, lStSql , lstTmp, 'Y_QUALIFPIECE', False );
  Result := lStSql ;
end;

function TOF_CPBALGENANAL.GetNomChampCPT: String;
begin
  Result := 'CED_SECTION';
end;

function TOF_CPBALGENANAL.GetNomChampCPTCombi: String;
begin
  Result := 'CED_GENERAL';
end;

procedure TOF_CPBALGENANAL.CompteOnExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if Trim(THEdit(Sender).Text) = '' then Exit;
  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  // Complétion auto du numéro de compte si possible
  if not CompleteAuto( Sender, fbGene ) then
    THEdit(Sender).ElipsisClick(nil);
end;

procedure TOF_CPBALGENANAL.Compte2OnExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if Trim(THEdit(Sender).Text) = '' then Exit;
  if NatureCpt.ItemIndex < 0 then Exit ;
  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  // Complétion auto du numéro de compte si possible
  if not CompleteAuto( Sender, AxeToFb( Axe.Value ) ) then
    THEdit(Sender).ElipsisClick(nil);
end;

function TOF_CPBALGENANAL.RecupWhereBalance(vBoAvecDate: Boolean): String;
begin
  Result := RecupWhereCritere( Pages);
  Result := CMajRequeteExercice ( Exercice.Value, Result);
  {b FP}
  Result := result + ' AND '+TSQLAnaCroise.ConditionAxe(Axe.Value);
  {e FP}
  // Dates comptables
  if vBoAvecDate then
    Result := Result + ' AND Y_DATECOMPTABLE >= "' + UsDateTime(StrToDate(DateComptaDe.Text)) + '" '
                     + ' AND Y_DATECOMPTABLE <= "' + UsDateTime(StrToDate(DateComptaA.Text)) + '" ' ;
end;


procedure TOF_CPBALGENANAL.InitZones;
begin
  Axe.ItemIndex := 0;
  inherited;
end;

function TOF_CPBALGENANAL.GetNumAxe: String;
begin
  result := '' ;
  if NatureCpt.ItemIndex < 0 then Exit ;
  result := copy( Axe.Value, 2, 1 ) ;
end;

procedure TOF_CPBALGENANAL.ChargementCritEdt;
begin

  if (TheData <> nil) and (TheData is ClassCritEdt) then
  begin
    // GCO - 16/09/2005 - FQ 16644
    if ClassCritEdt(TheData).CritEdt.Bal.Axe <> '' then
      Axe.Value := ClassCritEdt(TheData).CritEdt.Bal.Axe;
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
procedure TOF_CPBALGENANAL.onCompte2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  St : string;
  fb : TFichierBase ;
begin
  St := THCritMaskEdit(Sender).Text;
  fb := AxeToFb(Axe.Value);
  if (Shift = []) and (Key = 187) then
  begin
    Key := 0;
    Compte2A.Text := Compte2De.Text;
  end else if ((Shift=[ssCtrl]) And (Key=VK_F5)) then
  begin
    If (fb in [fbAxe1..fbAxe5]) And
       VH^.Cpta[fb].Structure and
       { FQ 22071 BVE 20.12.07 }
       VH^.SaisieTranche[Axe.ItemIndex+1] And
       { END FQ 22071 }
       // GCO - 29/11/2006 - FQ 19175
       ExisteSQL('SELECT SS_AXE FROM STRUCRSE WHERE SS_AXE = "' + FBToAxe(fb) + '"')
    Then
    begin
      if ChoisirSousPlan( fb, St , True,taModif) then
      begin
        if ((THCritMaskEdit(Sender) = Compte2A) and EstJoker(St)) then Compte2De.Text := St
        else THCritMaskEdit(Sender).Text := St;
      end;
      Key := 0;
    end;
  end;
  if THCritMaskEdit(Sender) = Compte2De then fOnSaveKeyDownCompte2 (Sender, Key, Shift)
  else fOnSaveKeyDownCompte2_ (Sender, Key, Shift);
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 30/01/2006
Modifié le ... : 12/09/2006
Description .. : Ne pas prendre les sections fermées FQ 17206
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALGENANAL.RemplirEDTBALANCE;
begin
inherited   ;
  //YMO 12/09/2006 Ne pas éliminer les enregs avec un débit ou un crédit <> 0
  //YMO 22/09/2006 FQ18810 correction+prise en compte de l'utilisateur
  //YMO 29/09/2006 FQ18810 Suppression aussi des cptes generaux fermés ET non mouvementés
  IF ExecuteSQL('DELETE FROM CEDTBALANCE WHERE CED_USER = "' + V_PGI.User + '"'
  +' AND (CED_DEBIT1=0 AND CED_CREDIT1=0)'
  +' AND ((CED_COMPTE =(SELECT S_SECTION FROM SECTION WHERE S_SECTION = CED_COMPTE AND S_FERME = "X"))'
  +' OR  (CED_COMPTE2=(SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = CED_COMPTE2 AND G_FERME = "X"))'
  +' AND CED_COMPTE="")')>0 then
      SectionsFermees.Text := 'X'
  else
      SectionsFermees.Text := '-';

end;


Initialization
  registerclasses ( [ TOF_CPBALGENANAL ] ) ;
end.
