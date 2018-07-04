{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 29/07/2003
Modifié le ... : 01/07/2005
Description .. : Source TOF de la FICHE : CPBALANAL
Suite ........ :
Suite ........ : Gestion du QRS1 de la balance analytique
Suite ........ :
Suite ........ : Attention : les traitements communs aux balances
Suite ........ : sont dans une tof intermédiaire :
Suite ........ : TOF_METH --> TOF_CPBALANCE --> TOF_CPBALANAL
Suite ........ :
Suite ........ : JP 01/07/05 : Gestion des caractères Joker : fonctions de
Suite ........ : base définies dans TofMeth
Mots clefs ... : TOF;CPBALANAL
*****************************************************************}
unit CPBALANAL_TOF;

interface

uses StdCtrls, Classes, Windows,
  {$IFDEF EAGLCLIENT}
  MainEAgl, utob, eQRS1,
  {$ELSE}
  db, Fe_Main,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  sysutils, Spin, Ent1,
  HCtrls, UTOF, hqry,
  AglInit,      // TheData
  CritEdt,      //ClassCritEdt
  TofMeth,
  uLibExercice, // CRelatifVersExercice
  uLibWindows, // TraductionTHMultiValComboBox
  CPBALANCE_TOF,
  utilPGI,      // TSQLAnaCroise
  HZoomSp
  ;

procedure CPLanceFiche_BalanceAnalytique(Args: string = '');

type
  TOF_CPBALANAL = class(TOF_CPBALANCE)
    AvecANO: TCheckBox;
    AvecODanalytique:TCheckBox; //SG6 08/11/2004 FQ : 14869
    SectionsFermees             : THEdit;
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
    private
        fOnSaveKeyDownCompte : procedure(Sender: TObject; var Key: Word; Shift:
      TShiftState) of object;
        fOnSaveKeyDownCompte_ : procedure(Sender: TObject; var Key: Word; Shift:
      TShiftState) of object;
  end;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  CPProcGen,
  {$ENDIF MODENT1}
  HEnt1;


//==============================================================================
procedure CPLanceFiche_BalanceAnalytique(Args: string);
begin
  AGLLanceFiche('CP', 'CPBALANAL', '', '', Args);
end;

{TOF_CPBALANAL}

//==============================================================================
//====================         EVT TOF        ==================================
//==============================================================================

procedure TOF_CPBALANAL.OnNew;
begin
  inherited;
end;

procedure TOF_CPBALANAL.OnUpdate;
begin
  inherited;
end;

procedure TOF_CPBALANAL.OnLoad;
begin

  inherited;

  // ANO
  if not AvecANO.Checked then
    XXWhere.Text := 'Y_ECRANOUVEAU<>"OAN"';

end;

procedure TOF_CPBALANAL.OnArgument(S: string);
begin

  // CONTROLES SPECIFIQUES
  // Options
  AvecANO := TCheckBox(GetControl('AVECANO', True));
  AvecODanalytique:=TCheckBox(GetControl('TYPEANALYTIQUE',True));  //SG6 08/11/2004 FQ : 14869
  SectionsFermees := THEdit(GetControl('SECTIONSFERMEES', true));
  // Controles communs
  inherited;

  // Rubrique d'aide
  Ecran.HelpContext := 7451000;

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

procedure TOF_CPBALANAL.TableLibreChanged(Sender: TObject);
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

procedure TOF_CPBALANAL.CompteOnExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if Trim(THEdit(Sender).Text) = '' then Exit;
  if NatureCpt.ItemIndex < 0 then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  // Complétion auto du numéro de compte si possible
  if not CompleteAuto(Sender, AxeToFb( NatureCpt.Value ) ) then
    THEdit(Sender).ElipsisClick(nil);
end;

procedure TOF_CPBALANAL.ChargementCritEdt;
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
procedure TOF_CPBALANAL.onCompteKeyDown(Sender: TObject; var Key: Word;
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

procedure TOF_CPBALANAL.SetTypeBalance;
begin
  TypeBal     :=  balAnal ;
  TableEcr    := 'ANALYTIQ' ;
  PfEcr       := 'Y' ;
  TableCpt    := 'SECTION' ;
  PfCpt       := 'S' ;
  ChampCpt    := 'SECTION' ;
  ChampNatCpt := 'S_AXE' ;
end;

procedure TOF_CPBALANAL.ParamDivers;
begin
  inherited;

  // ANO
  if not AvecANO.Checked then
    XXWhere.Text := 'Y_ECRANOUVEAU<>"OAN"';

end;

function TOF_CPBALANAL.GetSQLCumul(vInPer, vInCol: Integer): String;
begin
  result := inherited GetSQLCumul( vInPer, vInCol ) ;

  // SBO traduction requête pour multi-axe
  TSQLAnaCroise.TraduireRequete( NatureCpt.Value , result ) ;
end;

function TOF_CPBALANAL.GenererInsertCPT(vBoCompar: Boolean): String;
begin
  result := inherited GenererInsertCPT( vBoCompar ) ;

  // SBO traduction requête pour multi-axe
  TSQLAnaCroise.TraduireRequete( NatureCpt.Value , result ) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 30/01/2006
Modifié le ... : 12/09/2006
Description .. : Ne pas prendre les sections fermées FQ 17206
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALANAL.RemplirEDTBALANCE ;
begin
  inherited;
  //YMO 12/09/2006 Ne pas éliminer les enregs avec un débit ou un crédit <> 0
  //YMO 22/09/2006 prise en compte de l'utilisateur
  IF ExecuteSQL('DELETE FROM CEDTBALANCE WHERE CED_USER = "' + V_PGI.User + '"'
  +' AND (CED_DEBIT1=0 AND CED_CREDIT1=0)'
  +' AND CED_COMPTE ='
  +'(SELECT S_SECTION FROM SECTION WHERE S_SECTION = CED_COMPTE AND S_FERME = "X")')>0 then 
      SectionsFermees.Text := 'X'
  else
      SectionsFermees.Text := '-';
end;

initialization
  registerclasses([TOF_CPBALANAL]);
end.

