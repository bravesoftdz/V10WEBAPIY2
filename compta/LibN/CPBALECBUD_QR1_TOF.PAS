{***********UNITE*************************************************
Auteur  ...... : Thong hor LIM
Créé le ...... : 29/12/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPBALECBUDG_QR1 ()
Mots clefs ... : TOF;CPBALECBUDG_QR1

Unité : Ancêtre des Ecarts budgétaires.
--------------------------------------------------------------------------------------
    Version    |   Date  | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 8.00.001.024    06/07/07   JP   FQ 20973 : on vide la table temporaire en fermeture de fiche
 8.10.002.002    29/10/07   JP   FQ 20973 : Gestion de l'utilisateur dans la table temporaire
--------------------------------------------------------------------------------------
*****************************************************************}
unit CPBALECBUD_QR1_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
   eQRS1,
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  QRS1,
  {$ENDIF}
  uTob,
  TofMeth,
  uLibWindows, //TraductionTHMultiValComboBox
  Forms, SysUtils, ComCtrls,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ELSE}
  tcalccum,   
  {$ENDIF MODENT1}
  HCtrls, HEnt1, UTOF, Ent1;


type
  TTabDate12 = array[1..12] of TDateTime;
  TabTot12   = array[0..12] of TabDC;
  TMontTotal = array[0..2 ] of TabTot12;
  TTypeBalance = (ecbud_Gene, ecbud_Sect, ecbud_SeGe, ecbud_GeSe);

  TOF_CPBALECBUD_QR1 = class (TOF_METH)
    procedure OnArgument(S : string ); override;
    procedure OnLoad                 ; override;
    procedure OnUpdate               ; override;
    procedure OnClose                ; override;
  private
    {JP 11/08/05 : FQ 16054 : Pour le moment, cela se limite aux tables libres sauf pour les
                   sections par compte qui gère les plan de ruptures}
    function GetRupture : Char;
  protected
    FGroupQuelleRupture : THRadioGroup;
    FGroupChoixRupt     : THRadioGroup;
    FRealise            : TCheckBox;
    FJournal            : THValComboBox;
    FEXERCICE           : THValComboBox;
    FEXERCICE_          : THValComboBox;
    FPeriode1           : THValComboBox;
    FPeriode2           : THValComboBox;
    FSELECTCPTE         : THValComboBox;
    FNatBud             : THMultiValComboBox;   // FQ 16536 SBO 18/10/2005
    FStWhereNat         : String ;
    FStLibNat           : String ;
    procedure FJournalChange  (Sender : TObject);
    procedure FExerciceChange (Sender : TObject);
    procedure FExercice_Change(Sender : TObject);
    procedure FPeriode1Change (Sender : TObject);
    procedure FPeriode2Change (Sender : TObject);
    procedure FLibreDblClick  (Sender : TObject);
    procedure FSansRuptClick  (Sender : TObject);
    procedure FRupturesClick  (Sender : TObject); virtual;  // MD FQ 18077
    procedure FRealiseClick   (Sender : TObject); virtual;
  public

    {Pour connaître la balance courante}
    aTypeBalance   : TTypeBalance;

    szExoDebJal    : string;
    szExoFinJal    : string;
    szAxeJal       : string;
    FDateCompta1   : string;     //période début d'édition
    FDateCompta2   : string;     //période fin d'édition
    FSecdRequete   : string;
    FMvmtRequete   : string;
    dtDatDebJal    : TDateTime;
    dtDatFinJal    : TDateTime;
    bDebitPos      : Boolean;
    bRecap         : Boolean;
    Decimale       : Byte;
    TabDate        : TTabDate12;   // à revoir

    {Récupération des ruptures sur tables libres}
    function UpdateRupture(szLibre1, szLibre2 : string) : string;
    {Affecte un évènement au contrôle passé en paramètre et renvoit le dit contrôle}
    function DoAssign(Control, Evt : string; Event : TNotifyEvent) : TControl;
    {Charge les critères en fonctions du journal budgétaire}
    procedure InfosJournal;
    {Actuellement, les ruptures sont gérées uniquement dans BALBUDTESEG_TOF, mais pas dans
     BALBUDTESEC_TOF. En attendant un éventuel changement pour reprendre l'ergonomie des
     balances et du grand livre}
    procedure MajInfoRupture; virtual;

    procedure InitCompte;
    procedure DebToDate;
    procedure FinToDate;

    procedure CalculBud(i : Byte ; var Tot1 : TMontTotal ; NatB : string ; D,C : Double ; Dec : Byte);
    procedure InitCalcBud(T : Tob; var Tot1 : TMontTotal; Resol : string; Taux : Double);

    function  GetMainRequete : string; virtual;
    function  GetSQLMouvement(Compte, Section : string) : string;
    function  GetSQLSectGene (Valeur : string) : string;
    procedure SetSecdRequete ;
    procedure SetMvmtRequete ;

    procedure PrepCalcBud(SQL : string); virtual; abstract;
    procedure MiseMtsReaDansTable(T3: Tob; D, C : double; Resol : string; Taux : Double; Col: integer);

    property Rupture : Char read GetRupture;
  end;

implementation


uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcMetier,
  {$ENDIF MODENT1}
  UtilEDT,  {ExistBud, WhereLibre, OrderLibre}
  CALCOLE,  {GetCumul}
  TabLiEdt, {ChoixTableLibreSur, ChoixTableLibre}
  TriTabLi; {ChoixTriTableLibre}


{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;

  FJournal   := THValComboBox(DoAssign('FJOURNAL', 'OnChange', FJournalChange));
  FEXERCICE  := THValComboBox(DoAssign('FEXERCICE', 'OnChange', FExerciceChange));
  FEXERCICE_ := THValComboBox(DoAssign('FEXERCICE_', 'OnChange', FExercice_Change));
  FPeriode1  := THValComboBox(DoAssign('FPeriode1', 'OnChange', FPeriode1Change));
  FPeriode2  := THValComboBox(DoAssign('FPeriode2', 'OnChange', FPeriode2Change));
  FNatBud    := THMultiValComboBox( GetControl('FNATBUD', True ) );   // FQ 16536 SBO 18/10/2005

  FGroupChoixRupt     := THRadioGroup(DoAssign('FGROUPCHOIXRUPT'    , 'OnClick', FSansRuptClick));
  FGroupQuelleRupture := THRadioGroup(DoAssign('FGROUPQUELLERUPTURE', 'OnClick', FRupturesClick));

  DoAssign('FLIBRE1', 'OnDblClick', FLibreDblClick);
  DoAssign('FLIBRE2', 'OnDblClick', FLibreDblClick);
  DoAssign('FLIBTRIPAR', 'OnDblClick', FLibreDblClick);
  FRealise := TCheckBox(GetControl('FREALISE', True));
  FRealise.OnClick := FRealiseClick;

  // Initialise les contrôles
  if FJournal.Values.Count>0 then FJournal.Value:=FJournal.Values[0];
  THValComboBox(GetControl('FNatBud', True)).ItemIndex := 0;
  THValComboBox(GetControl('FResol', True)).ItemIndex := 1;

  FSELECTCPTE := THValComboBox(GetControl('FSELECTCPTE', True));
  FSELECTCPTE.Value := 'TOO';
  InitCompte;

  if (aTypeBalance = ecbud_Sect) or (aTypeBalance = ecbud_Gene) then
    SetControlEnabled('FONLYCPTASSOCIE', False);

  if (EstSerie(S3)) then SetControlVisible('FGROUPCHOIXRUPT', False);

  FSansRuptClick(nil);

  {b FP 21/06/2006 FQ16054}
  SetControlVisible('FGROUPCHOIXRUPT', False);
  {e FP 21/06/2006 FQ16054}
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.OnLoad;
{---------------------------------------------------------------------------------------}
var
  szDec : string;
  FResol : THValComboBox;
  StTri, StCod1, StCod2 : string;
begin
  inherited;

  {nitialisation des contrôles}
  InitCompte;

  {Gestion du format}
  FResol := THValComboBox(GetControl('FRESOL', True));
  case FResol.ItemIndex of
    0: SetControlText('FORMAT', '#,##0.00;; ;'); // Avec décimales
    1: SetControlText('FORMAT', '#,##0;; ;'); // Sans décimales
    2: SetControlText('FORMAT', '#,##0;; ;'); // Kilo
    3: SetControlText('FORMAT', '#,##0;; ;'); // Mega
  end;


  {Vérification des exercices}
  if not ((QUELEXODTBud(dtDatDebJal)>=szExoDebJal) and (QUELEXODTBud(dtDatFinJal)<=szExoFinJal)) then begin
    LastError := 1;
    LastErrorMsg := 'La fourchette d''exercice ne correspond pas à celle du budget.';
    exit;
  end;

  {Ruptures sur tables libres}
  if (FGROUPCHOIXRUPT.ItemIndex > 0) and (Trim(GetControlText('FLIBTRIPAR'))= '') then begin
    StTri := GetControlText('FLIBTRIPAR');
    if aTypeBalance = ecbud_Sect then begin
      if GetCheckBoxState('FONLYCPTASSOCIE') = cbChecked then ChoixTableLibreSurInfo(AxeToFbBud(szAxeJal), StTri, StCod1, StCod2)
                                                         else ChoixTableLibreInfo(AxeToFbBud(szAxeJal), StTri, StCod1, StCod2);
    end
    else begin
      if GetCheckBoxState('FONLYCPTASSOCIE') = cbChecked then ChoixTableLibreSurInfo(fbBudgen, StTri, StCod1, StCod2)
                                                         else ChoixTableLibreInfo(fbBudgen, StTri, StCod1, StCod2);
    end;
    SetControlText('FLIBTRIPAR', UpdateRupture(StCod1,StCod2));
  end;

  szDec := GetControlText('FRESOL');

  {Gestion du nombre de décimales}
  if szDec = 'C' then Decimale := V_PGI.OkDecV
                 else Decimale := 0;

  if (aTypeBalance <> ecbud_Sect) and (aTypeBalance <> ecbud_Gene) then
    bRecap := GetCheckBoxState('FAVECCPTSECOND') = cbChecked;

  // FQ 16536 SBO 18/10/2005 Mise en place des options de révision dans les balances budgétaires
  FStWhereNat := '' ;
  FStLibNat   := '<<Tous>>' ;
  TraductionTHMultiValComboBox( FNatBud, FStWhereNat, FStLibNat, 'BE_NATUREBUD' ) ;
  SetControlText( 'NATBUD', FStLibNat ) ;
  // Fin FQ 16536 SBO 18/10/2005 Mise en place des options de révision dans les balances budgétaires

  {On vide la table temporaire}
  ExecuteSQL('DELETE FROM CEDTBALBUD WHERE CEB_UTILISATEUR = "' + V_PGI.User + '"');

  PrepCalcBud(GetMainRequete);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.FJournalChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if FJournal.Value <> '' then
    InfosJournal;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.FExerciceChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if FExercice.Value > FExercice_.Value then FExercice_.Value := FExercice.Value;
  ListePeriode(FExercice.Value, FPeriode1.Items, FPeriode1.Values, True);
  FPeriode1.ItemIndex := 0;
  DebToDate;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.FExercice_Change(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if FExercice_.Value < FExercice.Value then FExercice.Value := FExercice_.Value;
  ListePeriode(FExercice_.Value, FPeriode2.Items, FPeriode2.Values, False);
  FPeriode2.ItemIndex := FPeriode2.Items.Count-1;
  FinToDate;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.FPeriode1Change(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {JP 10/08/05 : FQ 16050 : à l'enregistrement d'un filtre, les zones sont mises à blanc
                 ce qui pose un problème sur le StrToDate}
  if (FPeriode1.Value <> '') and (FPeriode2.Value <> '') then begin
    if StrToDate(FPeriode1.Value) > StrToDate(FPeriode2.Value) then
      FPeriode2.Value := DateToStr(FinDeMois(StrToDate(FPeriode1.Value)));
  end;

  FDateCompta1 := FPeriode1.Value;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.FPeriode2Change(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {JP 10/08/05 : FQ 16050 : à l'enregistrement d'un filtre, les zones sont mises à blanc
                 ce qui pose un problème sur le StrToDate}
  if (FPeriode1.Value <> '') and (FPeriode2.Value <> '') then begin
    if StrToDate(FPeriode2.Value) < StrToDate(FPeriode1.Value) then
      FPeriode1.Value := DateToStr(DebutDeMois(StrToDate(FPeriode2.Value)));
  end;

  FDateCompta2 := FPeriode2.Value;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.FLibreDblClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  StTri, StCod1, StCod2 : String;
  LeFb : TFichierBase;
begin
  StCod1 := GetControlText('FLIBRE1');
  StCod2 := GetControlText('FLIBRE2');
  LeFb := fbBudgen;
  case aTypeBalance of
    ecbud_Sect : LeFb := AxeToFbBud(szAxeJal);
    ecbud_SeGe,
    ecbud_GeSe,
    ecbud_Gene : LeFb := fbBudgen;
  end;

  If TEdit(Sender).Name = 'FLIBTRIPAR' then begin
    StTri := THEdit(Sender).Text;
    ChoixTriTableLibre(LeFb, StTri, True, StCod1);
    end
  else begin
    StTri := GetControlText('FLIBTRIPAR');
    if GetCheckBoxState('FONLYCPTASSOCIE')=cbChecked then ChoixTableLibreSur(LeFb, StTri, StCod1, StCod2)
                                                     else ChoixTableLibre(LeFb, StTri, StCod1, StCod2);
    StTri := UpdateRupture(StCod1, StCod2);
  end;
  SetControlText('FLIBTRIPAR', StTri);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.FSansRuptClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  SetControlVisible('SHTRUPTURE', (FGROUPCHOIXRUPT.ItemIndex > 0));

  if (FGROUPCHOIXRUPT.ItemIndex = 2) then SetControlText('TRI', 'X')
                                     else SetControlText('TRI', '');

  if (aTypeBalance = ecbud_SeGe) or (aTypeBalance = ecbud_GeSe) then
    SetControlEnabled('FAVECCPTSECOND', (FGROUPCHOIXRUPT.ItemIndex = 2));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.FRupturesClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  SetControlVisible('FGROUPRUPTURES',  (FGROUPQUELLERUPTURE.ItemIndex = 0));
  SetControlVisible('FONLYCPTASSOCIE', (FGROUPQUELLERUPTURE.ItemIndex = 1));
  SetControlChecked('FONLYCPTASSOCIE', (FGROUPQUELLERUPTURE.ItemIndex = 1));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.FRealiseClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
//
end;

{---------------------------------------------------------------------------------------}
function TOF_CPBALECBUD_QR1.UpdateRupture(szLibre1, szLibre2 : string) : string;
{---------------------------------------------------------------------------------------}
var
  i       : Integer;
  St      : string;
  StListe : string;
  Prefixe : string;
begin
  case aTypeBalance of
    ecbud_Gene,
    ecbud_GeSe : Prefixe := 'B0';
    ecbud_Sect,
    ecbud_SeGe : Prefixe := 'D0';
  end;

  SetControlText('FLIBRE1', szLibre1);
  SetControlText('FLIBRE2', szLibre2);
  StListe := szLibre1;

  i := 0;
  Result := '';

  while StListe <> '' do begin
    St := ReadTokenSt(StListe);
    if St = '' then begin
      Inc(i);
      continue;
    end;
    if (St[1] <> '#') and (St[1] <> '-') then Result := Result + Prefixe + IntToStr(i) + ';';
    Inc(i);
  end;

end;

{---------------------------------------------------------------------------------------}
function TOF_CPBALECBUD_QR1.DoAssign(Control, Evt : string; Event: TNotifyEvent) : TControl;
{---------------------------------------------------------------------------------------}
var
  szEvt : string;
begin
  szEvt  := UpperCase(Evt);
  Result := GetControl(Control, True);

  if Result = nil then Exit;

       if szEvt = 'ONCHANGE'   then TComboBox(Result)   .OnChange   := Event
  else if szEvt = 'ONDBLCLICK' then THEdit(Result)      .OnDblClick := Event
  else if szEvt = 'ONCLICK'    then THRadioGroup(Result).OnClick    := Event;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.InfosJournal;
{---------------------------------------------------------------------------------------}
var
  QJal  : TQuery;
  szOldAxeJal : string;
begin
  QJal := OpenSQL('SELECT BJ_EXODEB, BJ_EXOFIN, BJ_PERDEB, BJ_PERFIN, BJ_AXE, BJ_NATJAL FROM BUDJAL WHERE BJ_BUDJAL="'+FJournal.Value+'"', True);
  try
    if not QJal.EOF then begin
      FEXERCICE.Value  := QJal.FindField('BJ_EXODEB').AsString;
      FEXERCICE_.Value := QJal.FindField('BJ_EXOFIN').AsString;
      szExoDebJal      := QJal.FindField('BJ_EXODEB').AsString;
      szExoFinJal      := QJal.FindField('BJ_EXOFIN').AsString;
      FPeriode1.Value  := QJal.FindField('BJ_PERDEB').AsString;
      FPeriode2.Value  := QJal.FindField('BJ_PERFIN').AsString;
      FDateCompta1     := QJal.FindField('BJ_PERDEB').AsString;
      FDateCompta2     := QJal.FindField('BJ_PERFIN').AsString;
      dtDatDebJal      := QJal.FindField('BJ_PERDEB').AsDateTime;
      dtDatFinJal      := QJal.FindField('BJ_PERFIN').AsDateTime;
      szAxeJal         := QJal.FindField('BJ_AXE'   ).AsString;
      bDebitPos        := QJal.FindField('BJ_NATJAL').AsString = 'CHA';
      szOldAxeJal      := szAxeJal;

      // Si changement d'axe
      if (szOldAxeJal <> szAxeJal) then begin
        case aTypeBalance of
          ecbud_Gene,
          ecbud_GeSe : begin
                       SetControlProperty('FCPTE1', 'DATATYPE', 'TZBUDSEC' + szAxeJal[2]);
                       SetControlProperty('FCPTE2', 'DATATYPE', 'TZBUDSEC' + szAxeJal[2]);
                       SetControlText('FCPTE1', '');
                       SetControlText('FCPTE2', '');
                     end;
          ecbud_Sect,
          ecbud_SeGe : begin
                       SetControlProperty('FSECTION1', 'DATATYPE', 'TZBUDSEC' + szAxeJal[2]);
                       SetControlProperty('FSECTION2', 'DATATYPE', 'TZBUDSEC' + szAxeJal[2]);
                       SetControlText('FSECTION1', '');
                       SetControlText('FSECTION2', '');
                     end;
        end;
        MajInfoRupture;
      end;
    end;
  finally
    Ferme(QJal);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.MajInfoRupture;
{---------------------------------------------------------------------------------------}
begin
  {Actuellement, les ruptures sont gérées uniquement dans BALBUDTESEG_TOF, mais pas dans
   BALBUDTESEC_TOF. En attendant un éventuel changement pour reprendre l'ergonomie des
   balances et du grand livre}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.InitCompte;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  if aTypeBalance <> ecbud_Sect then begin
    Q := OpenSQL('SELECT MIN(BG_BUDGENE), MAX(BG_BUDGENE) FROM BUDGENE WHERE BG_FERME="-"', True);
    if not Q.EOF then begin
      if GetControlText('FCPTE1') = '' then SetControlText('FCPTE1', Q.Fields[0].AsString);
      if GetControlText('FCPTE2') = '' then SetControlText('FCPTE2', Q.Fields[1].AsString);
    end;
    Ferme(Q);
  end;


  if aTypeBalance <> ecbud_Gene then begin
    Q := OpenSQL('SELECT MIN(BS_BUDSECT), MAX(BS_BUDSECT) FROM BUDSECT WHERE BS_FERME = "-" AND BS_AXE = "' +
                  szAxeJal + '"', True);
    if not Q.EOF then begin
      if aTypeBalance = ecbud_Sect then begin
        if GetControlText('FCPTE1') = '' then SetControlText('FCPTE1', Q.Fields[0].AsString);
        if GetControlText('FCPTE2') = '' then SetControlText('FCPTE2', Q.Fields[1].AsString);
      end
      else begin
        if GetControlText('FSECTION1') = '' then SetControlText('FSECTION1', Q.Fields[0].AsString);
        if GetControlText('FSECTION2') = '' then SetControlText('FSECTION2', Q.Fields[1].AsString);
      end;
    end;
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.DebToDate;
{---------------------------------------------------------------------------------------}
begin
  FDateCompta1 := FPeriode1.Value;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.FinToDate;
{---------------------------------------------------------------------------------------}
begin
  FDateCompta2 := FPeriode2.Value;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.CalculBud(i : Byte; var Tot1 : TMontTotal; NatB : string;
                                        D, C : Double; Dec : Byte);
begin
  If (NatB='INI') or (NatB='ANO') then begin        {FP FQ16059}
    Tot1[0][i] .TotDebit  := Arrondi(Tot1[0][i] .TotDebit  + D, Dec);
    Tot1[0][i] .TotCredit := Arrondi(Tot1[0][i] .TotCredit + C, Dec);
  end
  else if Copy(NatB, 1, 2) = 'DM' then begin
    Tot1[1][i] .TotDebit  := Arrondi(Tot1[1][i] .TotDebit  + D, Dec);
    Tot1[1][i] .TotCredit := Arrondi(Tot1[1][i] .TotCredit + C, Dec);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.InitCalcBud(T : Tob; var Tot1 : TMontTotal; Resol : string; Taux : Double);
var
  DateCompta : TDateTime;
  LaNatBud   : string;
  DEBIT      : Double;
  CREDIT     : Double;
  i          : Integer;
begin
  Fillchar(Tot1, SizeOf(Tot1), #0);

  for i := 0 to T.Detail.Count-1 do begin
    DateCompta := T.Detail[i].GetDateTime('BE_DATECOMPTABLE');
    LaNatBud   := T.Detail[i].GetString('BE_NATUREBUD');
    DEBIT      := T.Detail[i].GetDouble('DEBIT');
    CREDIT     := T.Detail[i].GetDouble('CREDIT');

    {
    la rééval sur le détail fait perdre beaucoup de montants
    Reevaluation(DEBIT, CREDIT, Resol, Taux);  // formatage des montants
    }

    // Calculs Bud Période
    if (DateCompta >= StrToDate(FDateCompta1)) and (DateCompta < StrToDate(FDateCompta2)) then
      CalculBud(0, Tot1, LaNatBud, DEBIT, CREDIT, Decimale);
    // Calculs Bud Cumulé
    if (DateCompta >= dtDatDebJal) and (DateCompta < StrToDate(FDateCompta2)) then
      CalculBud(1, Tot1, LaNatBud, DEBIT, CREDIT, Decimale);
    // Calculs Bud Annuel
    if (DateCompta >= dtDatDebJal) and (DateCompta < dtDatFinJal) then
      CalculBud(2, Tot1, LaNatBud, DEBIT, CREDIT, Decimale);
    end;
   // formatage des montants
  Reevaluation(Tot1[0][0].TotDebit, Tot1[0][0].TotCredit , Resol, Taux); // bud ini. PER
  Reevaluation(Tot1[1][0].TotDebit, Tot1[0][0].TotCredit , Resol, Taux); // bud révi. PER

  Reevaluation(Tot1[0][1].TotDebit, Tot1[0][0].TotCredit , Resol, Taux); // bud ini. CUM
  Reevaluation(Tot1[1][1].TotDebit, Tot1[0][0].TotCredit , Resol, Taux); // bud révi. CUM

  Reevaluation(Tot1[0][2].TotDebit, Tot1[0][0].TotCredit , Resol, Taux); // bud ini. Annuel
  Reevaluation(Tot1[1][2].TotDebit, Tot1[0][0].TotCredit , Resol, Taux); // bud révi. Annuel
end;

{Constitution sur les comptes budgétaires
{---------------------------------------------------------------------------------------}
function TOF_CPBALECBUD_QR1.GetMainRequete : string;
{---------------------------------------------------------------------------------------}
var
  sz      : string;
  szCpte1 : string;
  szCpte2 : string;
  Champ   : string;
  Orderby : string;
begin
  {La requête de base repose sur les généraux}
  if (aTypeBalance = ecbud_GeSe) or (aTypeBalance = ecbud_Gene) then begin
    Champ   := 'BG_BUDGENE';
    OrderBy := 'BG_BUDGENE';

    if Rupture in ['T', 'A'] then
      Result := 'SELECT BG_BUDGENE, ' + OrderLibre(GetControlText('FLIBTRIPAR')) + ' BG_RUB FROM BUDGENE G WHERE '
    else
      Result := 'SELECT BG_BUDGENE, BG_LIBELLE, BG_RUB FROM BUDGENE G WHERE ';

    Result := Result + ExistBud(fbBudgen, GetControlText('FSELECTCPTE'), GetControlText('FJOURNAL'), szAxeJal, True);
  end

  {La requête de base repose sur les sections}
  else begin
    Champ   := 'BS_BUDSECT';
    OrderBy := 'BS_AXE, BS_BUDSECT';

    if Rupture in ['T', 'A'] then
      Result := 'SELECT BS_BUDSECT, ' + OrderLibre(GetControlText('FLIBTRIPAR')) + ' BS_RUB FROM BUDSECT S WHERE '
    else
      Result := 'SELECT * FROM BUDSECT S WHERE ';

    Result := Result + ExistBud(AxeToFbBud(szAxeJal), GetControlText('FSELECTCPTE'), GetControlText('FJOURNAL'), szAxeJal, True);
    Result := Result + ' AND BS_AXE = "' + szAxeJal + '" ';
  end;

  {Filtre sur les comptes / sections}
  szCpte1 := GetControlText('FCPTE1');
  szCpte2 := GetControlText('FCPTE2');
  if szCpte1 <> '' then Result := Result + ' AND ' + Champ + ' >= "' + szCpte1 + '" ';
  if szCpte2 <> '' then Result := Result + ' AND ' + Champ + ' <= "' + szCpte2 + '" ';

  {Gestion des ruptures}
  if Rupture in ['T', 'A'] then begin
    szCpte1 := GetControlText('FLIBRE1');
    szCpte2 := GetControlText('FLIBRE2');

    if (aTypeBalance = ecbud_GeSe) or (aTypeBalance = ecbud_Gene) then
      sz := WhereLibre(szCpte1, szCpte2, fbBudgen, (GetCheckBoxState('FONLYCPTASSOCIE') = cbChecked))
    else
      sz := WhereLibre(szCpte1, szCpte2, AxeToFbBud(szAxeJal), (GetCheckBoxState('FONLYCPTASSOCIE')=cbChecked));

    if sz <> '' then Result := Result + sz;
    Result := Result + ' ORDER BY ' + OrderLibre(GetControlText('FLIBTRIPAR')) + ' ' + Champ;
  end
  else
    Result := Result + ' ORDER BY ' + OrderBy;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.SetSecdRequete;
{---------------------------------------------------------------------------------------}
var
  szCpte1 : string;
  szCpte2 : string;
begin
  if (aTypeBalance = ecbud_GeSe) then begin
    szCpte1 := GetControlText('FSECTION1');
    szCpte2 := GetControlText('FSECTION2');

    FSecdRequete := 'SELECT BS_BUDSECT, BS_LIBELLE, BS_RUB FROM BUDSECT S WHERE ';
    FSecdRequete := FSecdRequete + ExistBud(AxeToFbBud(szAxeJal), GetControlText('FSELECTCPTE'), GetControlText('FJOURNAL'), szAxeJal, False);
    FSecdRequete := FSecdRequete + ' AND BS_AXE = "' + szAxeJal + '" ';
    if szCpte1 <> '' then FSecdRequete := FSecdRequete + ' AND BS_BUDSECT >= "' + szCpte1 + '" ';
    if szCpte2 <> '' then FSecdRequete := FSecdRequete + ' AND BS_BUDSECT <= "' + szCpte2 + '" ';
    FSecdRequete := FSecdRequete + ' ORDER BY BS_AXE, BS_BUDSECT';
  end

  else begin
    szCpte1 := GetControlText('FCPTE1');
    szCpte2 := GetControlText('FCPTE2');

    FSecdRequete := 'SELECT BG_BUDGENE, BG_LIBELLE, BG_RUB FROM BUDGENE G WHERE ';
    FSecdRequete := FSecdRequete + ExistBud(fbBudgen, GetControlText('FSELECTCPTE'), GetControlText('FJOURNAL'), szAxeJal, False);
    if szCpte1 <> '' then FSecdRequete := FSecdRequete + ' AND BG_BUDGENE >= "' + szCpte1 + '" ';
    if szCpte2 <> '' then FSecdRequete := FSecdRequete + ' AND BG_BUDGENE <= "' + szCpte2 + '" ';
    FSecdRequete := FSecdRequete + ' ORDER BY BG_BUDGENE';
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.SetMvmtRequete;
{---------------------------------------------------------------------------------------}
var
  szEtab   : string;
  FValide  : TCheckBox;
begin
  FMvmtRequete := 'SELECT BE_BUDGENE, BE_BUDSECT, BE_EXERCICE, BE_DATECOMPTABLE, BE_NUMEROPIECE,'+
                  'BE_REFINTERNE, BE_ETABLISSEMENT, BE_LIBELLE, BE_VALIDE,'+
                  'BE_BUDJAL,BE_NATUREBUD, BE_QUALIFPIECE, '+
                  'BE_DEBIT DEBIT,BE_CREDIT CREDIT '+
                  'FROM BUDECR ';

  if (aTypeBalance = ecbud_SeGe) then
    FMvmtRequete := FMvmtRequete + 'WHERE BE_BUDSECT=:C1 AND BE_BUDGENE=:C2 '
  else
  if (aTypeBalance = ecbud_GeSe) then
    FMvmtRequete := FMvmtRequete + 'WHERE BE_BUDGENE=:C1 AND BE_BUDSECT=:C2 '
  else
  if (aTypeBalance = ecbud_Gene) then
    FMvmtRequete := FMvmtRequete + 'WHERE BE_BUDGENE=:C1 '
  else
  if (aTypeBalance = ecbud_Sect) then
    FMvmtRequete := FMvmtRequete + 'WHERE BE_BUDSECT=:C1 ';

  FMvmtRequete := FMvmtRequete + 'AND BE_AXE = "' + szAxeJal + '" '+
                                 'AND BE_BUDJAL = "' + GetControlText('FJOURNAL') + '" ';


  // FQ 16536 SBO 18/10/2005 Mise en place des options de révision dans les balances budgétaires
  if FStWhereNat <> '' then
    FMvmtRequete := FMvmtRequete + 'AND ' + FStWhereNat + ' ';
  // Fin FQ 16536 SBO 18/10/2005 Mise en place des options de révision dans les balances budgétaires

  {b Thl FQ 18277 05/06/2006}
   // On ramasse les mvts budgets depuis FDateCompta1 sauf si (FDateCompta1 > dtDatDebJal)
   // au dtDatFinJal sauf si (FDateCompta2 > dtDatFinJal).
  if (StrToDate(FDateCompta1) >  dtDatDebJal) then
    FMvmtRequete := FMvmtRequete + 'AND BE_DATECOMPTABLE>="'+USDateTime(dtDatDebJal) + '" '
  else
    FMvmtRequete := FMvmtRequete + 'AND BE_DATECOMPTABLE>="'+USDateTime(StrToDate(FDateCompta1)) + '" ';

  if (StrToDate(FDateCompta2) >  dtDatFinJal) then
    FMvmtRequete := FMvmtRequete + ' AND BE_DATECOMPTABLE<="'+ usdatetime(StrToDate(FDateCompta2))+'" '
  else
    FMvmtRequete := FMvmtRequete + ' AND BE_DATECOMPTABLE<="'+ usdatetime(dtDatFinJal) + '" ';
    FMvmtRequete := FMvmtRequete + ' AND BE_EXERCICE>="' + GetControlText('FEXERCICE') + '" ' +
                                   ' AND BE_EXERCICE<="' + GetControlText('FEXERCICE_')+'" ';
  {e Thl}
  if Assigned(ComboEtab) then
    szEtab := ComboEtab.Value;
  if szEtab<>'' then FMvmtRequete := FMvmtRequete + ' AND BE_ETABLISSEMENT="'+szEtab+'" ';

  FValide := TCheckBox(GetControl('FVALIDE', True));
       if FValide.Checked             then FMvmtRequete := FMvmtRequete + ' AND BE_VALIDE = "X" '
  else if FValide.State = cbUnchecked then FMvmtRequete := FMvmtRequete + ' AND BE_VALIDE = "-" ';

  if (aTypeBalance = ecbud_SeGe) then begin
    if bRecap then FMvmtRequete := FMvmtRequete + ' ORDER BY BE_BUDGENE, BE_BUDSECT, BE_AXE, BE_EXERCICE, BE_DATECOMPTABLE, BE_BUDJAL, BE_NUMEROPIECE, BE_QUALIFPIECE '
              else FMvmtRequete := FMvmtRequete + ' ORDER BY BE_BUDSECT, BE_AXE, BE_BUDGENE, BE_EXERCICE, BE_DATECOMPTABLE, BE_BUDJAL, BE_NUMEROPIECE, BE_QUALIFPIECE ';
  end
  else begin
    if bRecap then FMvmtRequete := FMvmtRequete + ' ORDER BY BE_BUDSECT, BE_BUDGENE, BE_BUDJAL, BE_NATUREBUD, BE_EXERCICE, BE_DATECOMPTABLE, BE_AXE, BE_NUMEROPIECE, BE_QUALIFPIECE '
              else FMvmtRequete := FMvmtRequete + ' ORDER BY BE_BUDGENE, BE_BUDJAL, BE_NATUREBUD, BE_EXERCICE, BE_DATECOMPTABLE, BE_BUDSECT, BE_AXE, BE_NUMEROPIECE, BE_QUALIFPIECE ';
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPBALECBUD_QR1.GetSQLMouvement(Compte, Section : string) : string;
{---------------------------------------------------------------------------------------}
var
  iPos : Integer;
  Chp1 : string;
  Chp2 : string;
begin
  Result := FMvmtRequete;

  if (aTypeBalance = ecbud_SeGe) then begin
    Chp1 := Section;
    Chp2 := Compte;
  end else begin
    Chp1 := Compte;
    Chp2 := Section;
  end;

  {#34 <=> '"' (double quote)}
  iPos := Pos(':C1', Result);
  System.Delete(Result, iPos, 3);
  System.Insert(#34 + Chp1 + #34, Result, iPos);

  iPos := Pos(':C2', Result);
  if Chp2 = '' then begin
    System.Delete(Result, iPos - 1, 4);
    System.Insert('<> ""', Result, iPos - 1);
    end
  else begin
    System.Delete(Result, iPos, 3);
    System.Insert(#34 + Chp2 + #34, Result, iPos);
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPBALECBUD_QR1.GetSQLSectGene(Valeur : string) : string;
{---------------------------------------------------------------------------------------}
var
  iPos   : Integer;
  ToFind : string;
begin
  if (aTypeBalance = ecbud_SeGe) then ToFind := ':BS_BUDSECT'
                                 else ToFind := ':BG_BUDGENE';
  Result := FSecdRequete;

  iPos := Pos(ToFind, Result);
  while iPos > 0 do begin
    System.Delete(Result, iPos, Length(ToFind));
    System.Insert(#34 + Valeur + #34, Result, iPos);
    iPos := Pos(ToFind, Result);
  end;
end;

{Pour le moement, c'est limité aux tables libres, sauf pour le section / général qui gère
 les plans de ruptures
{---------------------------------------------------------------------------------------}
function TOF_CPBALECBUD_QR1.GetRupture : Char;
{---------------------------------------------------------------------------------------}
begin
  if FGroupChoixRupt.ItemIndex = 0 then
    Result := 'S'
  else case FGroupQuelleRupture.ItemIndex of
    0 : if FGroupChoixRupt.ItemIndex = 1 then Result := 'R'  {Avec Plan de ruptures}
                                         else Result := 'U'; {Sur  Plan de ruptures}
    1 : if FGroupChoixRupt.ItemIndex = 1 then Result := 'T'  {Avec Tables libres}
                                         else Result := 'A'; {Sur  Tables libres}
    2 : if FGroupChoixRupt.ItemIndex = 1 then Result := 'C'  {Avec Plan de correspondance}
                                         else Result := 'O'; {Sur  Plan de correspondance}
    else
      Result := 'S';
  end;
end;

procedure TOF_CPBALECBUD_QR1.OnUpdate;
begin
  inherited;

  // Requête finale exécutée par l'état
  TFQRS1(Ecran).WhereSQL := ' CEB_UTILISATEUR = "' + V_PGI.User + '"'  ;
end;

procedure TOF_CPBALECBUD_QR1.MiseMtsReaDansTable(T3: Tob; D, C: double; Resol: string; Taux: Double; Col: integer);
var   dbl: Double;
begin
  Reevaluation(D, C, Resol, Taux);
  dbl := T3.GetValue('CEB_DEBITREA' + IntToStr(col));
  T3.PutValue('CEB_DEBITREA' + IntToStr(col),  dbl + D);
  dbl := T3.GetValue('CEB_CREDITREA'+ IntToStr(col));
  T3.PutValue('CEB_CREDITREA'+ IntToStr(col), dbl + C);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUD_QR1.OnClose;
{---------------------------------------------------------------------------------------}
begin
  {JP 06/07/07 : FQ 20973 : On vide la table temporaire}
  ExecuteSQL('DELETE FROM CEDTBALBUD WHERE CEB_UTILISATEUR = "' + V_PGI.User + '"');
  inherited;
end;

end.
