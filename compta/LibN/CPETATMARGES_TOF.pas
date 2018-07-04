{***********UNITE*************************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 15/03/2006
Modifié le ... : 23/03/2006
Description .. : Source TOF de la FICHE : CPETATMARGES
Suite ........ :
Suite ........ : Gestion du QRS1 de l'état des marges (BAN;EMA)
Suite ........ :
Suite ........ : Attention : les traitements communs aux balances
Suite ........ : sont dans une tof intermédiaire :
Suite ........ : TOF_METH --> TOF_CPBALANCE --> TOF_CPETATMARGES
Suite ........ :
Suite ........ :
Suite ........ :
Mots clefs ... : TOF;CPETATMARGES
*****************************************************************}
unit CPETATMARGES_TOF;

interface

uses StdCtrls, Classes, Windows,
  {$IFDEF EAGLCLIENT}
  MainEAgl, utob, eQRS1,
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,
  {$ENDIF}
  sysutils, Spin, Ent1,
  HCtrls, HEnt1, UTOF, hqry,
  AglInit,      // TheData
  CritEdt,      //ClassCritEdt
  TofMeth,
  uLibExercice, // CRelatifVersExercice
  uLibWindows, // TraductionTHMultiValComboBox
  CPBALANCE_TOF,
  HZoomSp,
  HMsgBox;

procedure CPLanceFiche_EtatDesMarges(Args: string = '');

type
  TOF_CPETATMARGES = class(TOF_CPBALANCE)

    AffaireEnCours : TCheckBox;
    DebChantier : ThEdit;
    FinChantier : ThEdit;
    Exercice : THValCombobox;
    DateComptable : ThEdit;
    DateComptable_ : ThEdit;
    RubProduits : ThEdit;
    RubCharges : ThEdit;

    // EVT TOF
    procedure OnNew; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure SetTypeBalance           ; override ;

    // EVT FICHE
    procedure TableLibreChanged ( Sender: TObject ) ; override ;
    procedure CompteOnExit      ( Sender: TObject ) ; override ;
    procedure onCompteKeyDown   ( Sender: TObject ; var Key: Word ; Shift: TShiftState) ;

    // TRAITEMENT
    function  GetSQLCumul       ( vInPer : Integer ; vInCol : Integer ) : String ; override ;
    function  GenererInsertCPT  (vBoCompar : Boolean = False)           : String ; override ;

    // Paramétrage état
    procedure ParamDivers ; override ;

    // Gestion CritEdt
    procedure ChargementCritEdt             ; override ;
    procedure RemplirEDTBALANCE            ; override ;
    function  GetConditionSQLEcr      : String ; override;

    Function  FaitSelect(Racine : String ; Sens : String) : String ;

    procedure UpdateCumulsCEDTBALANCE       ;
    

    private
        fOnSaveKeyDownCompte : procedure(Sender: TObject; var Key: Word; Shift:
      TShiftState) of object;
        fOnSaveKeyDownCompte_ : procedure(Sender: TObject; var Key: Word; Shift:
      TShiftState) of object;
  end;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  utilPGI;      // TSQLAnaCroise


Var LeSens : String;

//==============================================================================
procedure CPLanceFiche_EtatDesMarges(Args: string);
begin
  AGLLanceFiche('CP', 'CPETATMARGES', '', '', Args);
end;

{TOF_CPETATMARGES}

//==============================================================================
//====================         EVT TOF        ==================================
//==============================================================================

procedure TOF_CPETATMARGES.OnNew;
begin
  inherited;
end;

procedure TOF_CPETATMARGES.OnUpdate;
begin
  inherited;
end;

procedure TOF_CPETATMARGES.OnLoad;
begin

  inherited;

end;

procedure TOF_CPETATMARGES.OnArgument(S: string);
begin

  // Controles communs
  inherited;

  AffaireEncours := TCheckbox(GetControl('AFFAIREENCOURS', true));
  DebChantier    := THEdit(GetControl('DEBCHANTIER', true));
  FinChantier    := THEdit(GetControl('FINCHANTIER', true));
  Exercice       := THValComboBox(GetControl('EXERCICE', true));
  DateComptable  := THEdit(GetControl('DATECOMPTABLE', true));
  Datecomptable_ := THEdit(GetControl('DATECOMPTABLE_', true));
  RubCharges     := THEdit(GetControl('RUBCHARGES', true));
  RubProduits    := THEdit(GetControl('RUBPRODUITS', true));
  
  // Rubrique d'aide   à définir
//  Ecran.HelpContext := ;

  // Type de plan comptable :
  case V_PGI.LaSerie of
    S7:
      begin
        Corresp.plus := 'AND (CO_CODE = "A11" OR CO_CODE = "A12")';
      end;
  else
    begin
      Corresp.plus := 'AND CO_CODE = "A11"';
    end;
  end;

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
end;

//==============================================================================
//====================        EVT QRS1        ==================================
//==============================================================================

procedure TOF_CPETATMARGES.TableLibreChanged(Sender: TObject);
var lstVal : String ;
begin
  if TableLibre.ItemIndex < 0 then
  begin
    LibreDe.DataType := '';
    LibreA.DataType := '';
    LibreDe.Text := '';
    LibreA.Text := '';
  end
  else
  begin
    lstVal := GetNumTableLibre ;
    LibreDe.DataType := 'TZNATSECT' + lstVal ;
    LibreA.DataType := 'TZNATSECT' + lstVal ;
    LibreDe.Text := '';
    LibreA.Text := '';
  end;
end;

//==============================================================================
//====================   TRAITEMENTS DES DONNEES  ==============================
//==============================================================================

procedure TOF_CPETATMARGES.CompteOnExit(Sender: TObject);
begin
  if Trim(THEdit(Sender).Text) = '' then Exit;
  if NatureCpt.ItemIndex < 0 then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  // Complétion auto du numéro de compte si possible
  if not CompleteAuto(Sender, AxeToFb( NatureCpt.Value ) ) then
    THEdit(Sender).ElipsisClick(nil);
end;

procedure TOF_CPETATMARGES.ChargementCritEdt;
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
procedure TOF_CPETATMARGES.onCompteKeyDown(Sender: TObject; var Key: Word;
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
    If (fb in [fbAxe1..fbAxe5]) and
       VH^.Cpta[fb].Structure and
       // GCO - 29/11/2006 - FQ 19175
       ExisteSQL('SELECT SS_AXE FROM STRUCRSE WHERE SS_AXE = "' + FBToAxe(fb) + '"') then
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

procedure TOF_CPETATMARGES.SetTypeBalance;
begin
  TypeBal     :=  balAnal ;
  TableEcr    := 'ANALYTIQ' ;
  PfEcr       := 'Y' ;
  TableCpt    := 'SECTION' ;
  PfCpt       := 'S' ;
  ChampCpt    := 'SECTION' ;
  ChampNatCpt := 'S_AXE' ;
end;

procedure TOF_CPETATMARGES.ParamDivers;
begin
  inherited;
  
end;

function TOF_CPETATMARGES.GetSQLCumul(vInPer, vInCol: Integer): String;
begin
  result := inherited GetSQLCumul( vInPer, vInCol ) ;

  result := StringReplace(result, 'ANALYTIQ',
      'ANALYTIQ LEFT JOIN SECTION ON Y_SECTION=S_SECTION AND Y_AXE=S_AXE',[rfIgnoreCase]);

  // SBO traduction requête pour multi-axe
  TSQLAnaCroise.TraduireRequete( NatureCpt.Value , result ) ;
end;

function TOF_CPETATMARGES.GenererInsertCPT(vBoCompar: Boolean): String;
begin
  result := inherited GenererInsertCPT( vBoCompar ) ;

  // SBO traduction requête pour multi-axe
  TSQLAnaCroise.TraduireRequete( NatureCpt.Value , result ) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 24/03/2006
Modifié le ... :   /  /
Description .. : Utilisation des champs CED_DEBIT2 et CED_CREDIT2
                 pour stocker l'écriture de contrepartie générée
                 antérieurement
Mots clefs ... :
*****************************************************************}
procedure TOF_CPETATMARGES.RemplirEDTBALANCE ;
begin
 // 1. Effacement des enregistrements present pour le user
  DeleteCEDTBALANCE ;

// 2. Insertion de la liste des comptes cibles

  // 2.1 : Comptes du comparatifs si besoin (EN PREMIER)
  
    if AvecComparatif.Checked then
      begin
      if (ComparType.Value = 'COMPARBALSIT') and (BALSIT.text <> '')
        then ExecuteSQL(GenererInsertCPTBalSit) ;
      if (ComparType.Value = 'COMPARPERIODE') and (ComparExo.Value <> '')
        then ExecuteSQL(GenererInsertBalanceCompar) ;
      end ;

  // 2.2 : Comptes des généraux (COMPLETE les Comptes du compratifs au besoin)
  ExecuteSQL(GenererInsertCPT);

// 3. Update des totaux / soldes
  UpdateCumulsCEDTBALANCE ;
 
end;


function TOF_CPETATMARGES.GetConditionSQLEcr : String;
begin

  result := inherited GetConditionSQLEcr;
  If LeSens<>'' then Result := Result + FaitSelect('', LeSens);
  //2e passage, charges
  
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. : // Restrictions imposées par cle choix de la rubrique de
Suite ........ : charges
Suite ........ : // et par la rubrique de produits
Mots clefs ... :
*****************************************************************}
Function TOF_CPETATMARGES.FaitSelect(Racine : String ; Sens : String) : String ;
Var
 St, Cpte1, Cpte2, Exclu1, Exclu2, rubTxt, Where : String ;
 Q   : TQuery ;


BEGIN

St:=St+' AND Y_SECTION LIKE "'+Racine+'%"';

If Sens = 'C' Then RubTxt := RubCharges.Text else RubTxt := RubProduits.Text ;

If RubTxt = '' then
begin
  Result:='';
  exit;
end;

Q:=OpenSQL('SELECT RB_COMPTE1, RB_COMPTE2, RB_EXCLUSION1, RB_EXCLUSION2 FROM RUBRIQUE'
          +' WHERE RB_RUBRIQUE = "'+RubTxt+'"', TRUE) ;

While Not Q.Eof Do
  BEGIN
  Cpte1:=Q.Findfield('RB_COMPTE1').AsString;
  Exclu1:=Q.Findfield('RB_EXCLUSION1').AsString;
  Cpte2:=Q.Findfield('RB_COMPTE2').AsString;
  Exclu2:=Q.Findfield('RB_EXCLUSION2').AsString;
  Q.Next ;
  END ;
Ferme(Q) ;

// Restrictions définies au niveau des rubriques

Where:=AnalyseCompte(Cpte1,fbgene,False,False,False) ;
   if Where<>'' then St:=St+' AND '+Where ;
Where:=AnalyseCompte(Exclu1,fbgene,True,False,False) ;
   if Where<>'' then St:=St+' AND '+Where ;
Where:=AnalyseCompte(Cpte2,fbgene,False,False,False) ;
   if Where<>'' then St:=St+' AND '+Where ;
Where:=AnalyseCompte(Exclu2,fbgene,True,False,False) ;
   if Where<>'' then St:=St+' AND '+Where ;

// YMO 21/06/2006 Activation des sélections sur dates affaires
If GetCheckBoxState('AFFAIREENCOURS')<>cbGrayed then
  St:=St+' AND S_AFFAIREENCOURS="'+GetControlText('AFFAIREENCOURS')+'"';
St:=St+' AND S_DEBCHANTIER>="'+UsDateTime(StrToDateTime(DebChantier.Text))+'"';
St:=St+' AND S_FINCHANTIER<="'+UsDateTime(StrToDateTime(FinChantier.Text))+'"';


// La fonction Analysecompte renvoie les sélections sur la table GENERAUX,
// Je remplace G_ par S_ pour éviter de faire une jointure sur GENERAUX
St:=StringReplace(St,'G_','Y_',[rfReplaceAll, rfIgnoreCase]);

Result:=St ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 21/06/2006
Modifié le ... :   /  /
Description .. : On ne prend pas les ecritures de CUTOFF
Mots clefs ... :
*****************************************************************}
procedure TOF_CPETATMARGES.UpdateCumulsCEDTBALANCE;
Var lStReq     : String ;
begin

  try  // finally
    try  // except

      // -----------------------------------
      // ---- Calcul période principale ----
      // -----------------------------------

      //Charges
      LeSens:='C';

      lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT1 = ( ' + GetSQLCumul( 1, 1 ) +' )'
                                                            + ' WHERE CED_USER="'+V_PGI.User+'"' ;
      ExecuteSQL( lStReq ) ;

      //Le problème de la requête précédente est que les champs sont mis à NULL si le compte n'est pas mouvementé sur la période donnée
      ExecuteSQL('UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT1 = 0 WHERE CED_DEBIT1 IS NULL AND CED_USER="' + V_PGI.User + '"' ) ;


      //Pour les charges, partie Credit
      lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT2 = ( ' + GetSQLCumul( 1, 2 ) +' )'
                                                            + ' WHERE CED_USER="'+V_PGI.User+'"' ;
      ExecuteSQL( lStReq ) ;

      ExecuteSQL('UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT2 = 0 WHERE CED_DEBIT2 IS NULL AND CED_USER="' + V_PGI.User + '"' ) ;

      // Produits
      LeSens:='P';


      // >> Champ Crédit
      lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT1 = ( ' + GetSQLCumul( 1, 2 ) +' )'
                                                            + ' WHERE CED_USER="'+V_PGI.User+'"' ;
      ExecuteSQL( lStReq ) ;

      ExecuteSQL('UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT1 = 0 WHERE CED_CREDIT1 IS NULL AND CED_USER="' + V_PGI.User + '"' ) ;


      //Pour les produits, partie débit
      lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT2 = ( ' + GetSQLCumul( 1, 1 ) +' )'
                                                            + ' WHERE CED_USER="'+V_PGI.User+'"' ;
      ExecuteSQL( lStReq ) ;

      ExecuteSQL('UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT2 = 0 WHERE CED_CREDIT2 IS NULL AND CED_USER="' + V_PGI.User + '"' ) ;

      LeSens:='';


    // Try Except
    except
      on E : Exception do
      begin
        PgiError( E.Message, Ecran.Caption );
      end;
    end;
  // Try Finally
  finally
//    FiniMoveProgressForm ;
  end ;

end;


initialization
  registerclasses([TOF_CPETATMARGES]);
end.


