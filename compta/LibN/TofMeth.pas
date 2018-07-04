{***********UNITE*************************************************
Auteur  ...... : On s'en fou
Créé le ...... :   /  /    
Modifié le ... : 25/10/2007
Description .. : Source TOF de la FICHE CPCONSECR, CPCONSGENE, 
Suite ........ : CPPOINTAGEECR
Suite ........ : GCO - 17/07/2004 -> Blocage Filtre en Public pour CCSTD
Suite ........ : 
Suite ........ : JP 28/06/06 : FQ 16149 : Gestion des restrictions 
Suite ........ : utilsateurs sur les établissements
Suite ........ : 
Suite ........ : JP 25/10/07 : FQ 19970 : limitation de la FQ 16149 sur les
Suite ........ : établissements
Mots clefs ... : TOF;
*****************************************************************}
unit TofMeth;

interface

uses Classes, StdCtrls, Controls, SysUtils,
{$IFDEF EAGLCLIENT}
   Maineagl, eQRS1,
{$ELSE}
   {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
   QRS1, Fe_Main,HQuickRP,  // AGLLanceFiche
{$ENDIF}
     Dialogs,
     comctrls,
     UTof,
     HCtrls,
     Ent1,
     CritEdt,     // TCritEdtChaine
     AGLInit,     // TheData
     Filtre,      // Filtre
     Extctrls,    // TTimer
     hMsgBox,     // PGIERROR
     Htb97,       // ToolBarButton97
     UtilPGI,     // EstMultiSoc
     ParamSoc,    // GetParamSocSecur
     uTob,
     Stat;        // TFStat

procedure AGLLanceFicheEtatChaine( vCritEdtChaine : TCritEdtChaine );

type
  TOF_Meth = class(TOF)

  private
    FComboEtab          : THValComboBox; // JP 28/06/06 : FQ 16149
    FDateDernCloture    : TDateTime;    // GCO - 18/09/2006

    function GetDateDernCloture : TDateTime;

  public
    FFiltres    : THValComBoBox;
    TABLELIBRE  : THValComboBox;
    Pages       : TPageControl;
    PAVANCES    : TTabSheet;
    FTimer      : TTimer;
    FProvisoire : THEdit; // GCO - 08/09/2006

    cLoadfiltre  : boolean; //SG6 14.03.05 Gestion des filtres FQ 14996

    FDateFinEdition : TDateTime; // GCO - 18/09/2006

    procedure OnArgument(Arguments : string) ; override ;
    procedure OnNew ; override ;
    procedure OnLoad ; override ;
    procedure onUpdate; override;
    procedure OnClose; override;

    {JP}
    function  HasJoker(Sender : TObject) : Boolean;

    function  GetNumTableLibre : String; virtual;// GCO à voir avec SBO

    procedure FTimerTimer(Sender: TObject);
    procedure ChargementFCritEdtchaine ;
    procedure OnAfterFormShow; virtual;

    procedure InitSelectFiltre_TofMeth(T: TOB);
    procedure MySelectFiltre; virtual;
    procedure OnChangeFiltre ( Sender : TObject ); virtual;
    procedure ChargementCritEdt; virtual;

    function  GetComboEtab : THValComboBox; {JP 28/06/06 : FQ 16149}
    {JP 28/06/06 : FQ 16149 : gestion des réstrictions Etablissements et à défaut des ParamSoc}
    procedure GereEtablissement;
    {JP 28/06/06 : FQ 16149 : on s'assure que le filtre coincide avec les restrictions utilisateurs sur l'établissement}
    procedure ControlEtab;

    property  ComboEtab : THValComboBox read GetComboEtab write FComboEtab; {JP 28/06/06 : FQ 16149}

  private
    FBoTableToView : Boolean ;
  protected
    FCritEdtchaine : ClassCritEdtchaine;

    procedure DoJalOnExit(JalFocused, Jal1, Jal2: THEdit) ;
    procedure DoCompteOnExit(CptFocused, Cpt1, Cpt2: THEdit) ;
    procedure DoExoToDateOnChange(Exo: THValComboBox; Date1, Date2 : THEdit) ;
    procedure DoExoToPeriodeOnChange(Exo, Periode1, Periode2: THValComboBox) ;
    procedure DoPeriodeOnChange(PeriodeFocused, Periode1, Periode2: THValComboBox; Date1: THEdit = nil; Date2: THEdit = nil) ;
    procedure DoNumOnExit(NumFocused, Num1, Num2: THEdit) ;
    procedure DoDateOnExit(DateFocused, Date1, Date2: THEdit; DateD: TDatetime=2; DateF: TDatetime=73050) ;
    function  NbMoisDansExo(Exo: string): integer;
    function  ExoPrecedent(Exo: string): string;
    procedure ControlVisible(aControl: TControl; const Visible: boolean);
    procedure SetElementComboEtat(ComboEtat: THValComboBox; TypeEtat, NatEtat, Codes: string);
    procedure FillComboValues(Exo1, Exo2: THValComboBox);
    procedure PageOrder(Pages: TPageControl);
    procedure TypeChampCompte(labCpt: THLabel; Cpt1, Cpt2: THEdit; General: boolean);
    // GCO 02/02/2002

  end ;

function  GetMinCompte( vStTable, vStChamp, vStNature : string ; vStWhere : String = '' ) : string;
function  GetMaxCompte( vStTable, vStChamp, vStNature : string ; vStWhere : String = '' ) : string;

// GCO - 08/09/2006 - Affichage de la mention provisoire dans les états
function  GetModeBOI( vDtMin, vDtMax : TDateTime ) : string;

implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPTypeCons,
  {$ENDIF MODENT1}
  lookup,
  HEnt1,
  LicUtil,
  Printers,    // Printer
  uLibWindows; // TestJoker

{ TOF_Meth }

////////////////////////////////////////////////////////////////////////////////
function  GetMinCompte( vStTable, vStChamp, vStNature : string ; vStWhere : String  ) : String;
var lQuery : TQuery;
    lStwhere : string;
begin
  Result := '';
  lStWhere := '';

  if vstTable = 'GENERAUX' then
  begin
    if vStNature <> '' then
     lStWhere  := ' WHERE G_NATUREGENE="' + vStNature + '"'
  end
  else
    if vStTable = 'TIERS' then
    begin
      if vStNature <> '' then
       lStWhere  := ' WHERE T_NATUREAUXI="' + vStNature + '"';
    end
    else
      if vStTable = 'SECTION' then
      begin
        if vStNature <> '' then
           lStWhere  := ' WHERE S_AXE="' + vStNature + '"';
      end
      else Exit ;

  // Ajout d'une condition where en option
  if vStWhere <> '' then
    begin
    if lStWhere <> ''
      then lStWhere := lStWhere + ' AND ' + vStWhere
      else lStWhere := ' WHERE ' + vStWhere ;
    end ;

  try
    try
      lQuery := OpenSQL('SELECT MIN( ' + vStChamp + ' ) CODE FROM ' + vStTable + lStWhere , True);
      if not lQuery.Eof then
        Result := lQuery.FindField('CODE').AsString;

    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : GetMinCompte');
    end;

  finally
    Ferme( lQuery );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
function  GetMaxCompte( vStTable, vStChamp, vStNature : string ; vStWhere : String ) : String;
var lQuery : TQuery;
    lStwhere : string;
begin
  Result := '';
  lStWhere := '';

  if vstTable = 'GENERAUX' then
  begin
    if vStNature <> '' then
     lStWhere  := ' WHERE G_NATUREGENE="' + vStNature + '"'
  end
  else
    if vStTable = 'TIERS' then
    begin
      if vStNature <> '' then
       lStWhere  := ' WHERE T_NATUREAUXI="' + vStNature + '"';
    end
    else
      if vStTable = 'SECTION' then
      begin
        if vStNature <> '' then
           lStWhere  := ' WHERE S_AXE="' + vStNature + '"';
      end
      else Exit ;

  // Ajout d'une condition where en option
  if vStWhere <> '' then
    begin
    if lStWhere <> ''
      then lStWhere := lStWhere + ' AND ' + vStWhere
      else lStWhere := ' WHERE ' + vStWhere ;
    end ;

  try
    try
      lQuery := OpenSQL('SELECT MAX( ' + vStChamp + ' ) CODE FROM ' + vStTable + lStWhere , True);
      if not lQuery.Eof then
        Result := lQuery.FindField('CODE').AsString;

    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : GetMaxCompte');
    end;

  finally
    Ferme( lQuery );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/09/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function  GetModeBOI( vDtMin, vDtMax : TDateTime ) : string;
begin

end;

////////////////////////////////////////////////////////////////////////////////
// GCO
////////////////////////////////////////////////////////////////////////////////
procedure AGLLanceFicheEtatChaine( vCritEdtChaine: TCritEdtChaine );
var lCritEdtChaine : ClassCritEdtChaine;
begin
  lCritEdtChaine := ClassCritEdtChaine.Create;
  lCritEdtchaine.CritEdtChaine := vCritEdtChaine;
  TheData := lCritEdtChaine;
  AGLLanceFiche('CP', vCritEdtChaine.NomFiche, '', vCritEdtChaine.CodeEtat, vCritEdtChaine.NatureEtat );
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_Meth.OnArgument(Arguments : string) ;
begin
  inherited ;

  // Gestion MultiSoc SBO 17/08/2005
  if EstMultiSoc then
  begin
    FBoTableToView := V_PGI.EnableTableToView ;
    V_PGI.EnableTableToView := True ;
  end;

  //SG6 14.03.05 Gestion des filtres FQ 14996
  cLoadFiltre := False;

  FDateFinEdition := idate1900;
  FDateDernCloture := GetDateDernCloture;

  // GCO - 17/02/2004
  if (TheData <> nil) and (TheData is ClassCritEdtchaine) then
    FCritEdtchaine := ClassCritEdtchaine(TheData);

  if FCritEdtChaine <> nil then
  begin
    if FCritEdtchaine.CritEdtchaine.AuFormatPDF then
    begin
      if FCritEdtchaine.CritEdtChaine.MultiPdf then
        V_PGI.QRPDFQueue := ExtractFilePath(FCritEdtchaine.CritEdtChaine.NomPDF) + '\' + V_PGI.NoDossier + '-' + ExtractFileName(FCritEdtchaine.CritEdtChaine.NomPDF)
      else
        V_PGI.QRPDFQueue := FCritEdtchaine.CritEdtChaine.NomPDF;

      V_PGI.QRPDFMerge := '' ;
      if (V_PGI.QRPDFQueue <> '') and FileExists(V_PGI.QRPDFQueue) then
        V_PGI.QRPDFMerge := V_PGI.QRPDFQueue ;
    end;
  end;

  FFiltres := THValComboBox(GetControl('FFILTRES'));
  FFiltres.OnChange := OnChangeFiltre;
  Pages    := TPageControl(GetControl('PAGES'));

  TABLELIBRE := THValComboBox(Getcontrol('TABLELIBRE', False));

  FTimer := TTimer.Create ( Ecran );
  FTimer.Enabled  := False;
  FTimer.OnTimer  := FTimerTimer;
  FTimer.Interval := 1000;
  // FIN GCO

  // Gestion BOI (norme NF) dans les états de la comptabilité
  // FQ 19385 - CA - 21/12/2006 - Attention, dans certains cas, la panel STANDARDS n'existe pas
  if ( GetControl('STANDARDS', False) <> nil ) then
  begin
    FProvisoire         := THEdit.Create(Ecran);
    FProvisoire.Parent  := TTabSheet(GetControl('STANDARDS', true));
    FProvisoire.Name    := 'PROVISOIRE'; // GCO - 18985
    FProvisoire.Text    := '';
    FProvisoire.Visible := False;
  end;

  // Debut modif SBO : gestion de l'acces à la liste des états et au paramétrage des états
  if not (Ecran is TFQRS1) then Exit ; // vraiment util ??

  {JP 25/10/05 : Ouverture du paramEtat en eAgl => suppression de la directive de compilation
   Paramétrage de l'état pour les utilisateurs superviseurs + ouverture à PCL}
  TFQRS1(Ecran).ParamEtat := (not (ctxPCL in V_PGI.PGIContexte)) and ((V_PGI.PassWord=CryptageSt(DayPass(Date))) or ( V_PGI.FSuperviseur )) ;
  TFQRS1(Ecran).ChoixEtat := True ;  // Choix des états toujours possible
// Debut modif SBO

//RR 14/10/2005 demance SIC
  if not (ctxPCL in V_PGI.PGIContexte) then
  begin
    PAvances := TTabSheet(GetControl('AVANCES', true));
    Pavances.TabVisible := True;
    TFQRS1(Ecran).CritAvancesVisibled := True;
  end ;
//FIN RR

  {JP 28/06/06 : FQ 16149 : gestion des restrictions utilisateurs sur les établissements}
  FComboEtab := nil;
  GetComboEtab;
  GereEtablissement;

  // GCO - 20/10/2006
  TFQRS1(Ecran).OnAfterFormShow := OnAfterFormShow;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_Meth.OnAfterFormShow;
begin
  TFQRS1(Ecran).ListeFiltre.OnSelect := InitSelectFiltre_TofMeth;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_Meth.InitSelectFiltre_TofMeth(T: TOB);
var
  Lines: HTStrings;
  i: integer;
  stChamp: string;
  stVal: string;
begin
  if V_PGI.AGLDesigning then exit;
  if T <> nil then else exit;
  Lines := HTStringList.Create;
  for i := 0 to T.Detail.Count - 1 do
  begin
    stChamp := T.Detail[i].GetValue('N');
    stVal := T.Detail[i].GetValue('V');
    Lines.Add(stChamp + ';' + stVal);
  end;
  VideFiltre(FFiltres, Pages, false);
  ChargeCritMemoire(Lines, Pages);
  Lines.Free;

  // Ajoute ton code ici
  MySelectFiltre;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_Meth.MySelectFiltre;
begin
  {JP 28/06/06 : FQ 16149 : après le chargement du filtre, on s'assure que l'établissement
                 reste en cohérence avec les restrictions utilisateurs}
  ControlEtab;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_Meth.OnNew;
begin
inherited ;
{$IFDEF EAGLCLIENT}
{$ELSE}
  if not (Ecran is TFQRS1) then Exit ;
{$ENDIF}
  if FCritEdtChaine <> nil then
    ChargementFCritEdtchaine;

  if (TheData <> nil) and (TheData is ClassCritEdt) then
  begin
    TToolBarButton97(GetControl('BAGRANDIR',False)).Click;
    FTimer.Enabled := True;
  end;

{$IFDEF CCSTD}
  TFQRS1(Ecran).ListeFiltre.ForceAccessibilite := FaPublic;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/01/2004
Modifié le ... :   /  /    
Description .. : Contrôler la fourchette des Date1 et Date2
Suite ........ : DateD, DateF sont paramètres par défaut:'01/01/1900' et '31/12/2099'
Mots clefs ... :
*****************************************************************}
procedure TOF_Meth.DoDateOnExit(DateFocused, Date1, Date2: THEdit; DateD, DateF: TDatetime);
begin
  if (Date1 = nil) or (Date2 = nil) then Exit;
  if (DateFocused = Date1) and ((StrToDate(Date1.Text) < DateD) or (StrToDate(Date1.Text) > DateF)) then
    Date1.Text := DateToStr(DateD);
  if (DateFocused = Date2) and ((StrToDate(Date2.Text) < DateD) or (StrToDate(Date2.Text) > DateF)) then
    Date2.Text := DateToStr(DateF);
  if (DateFocused = Date1) and (StrToDate(Date1.Text) > StrToDate(Date2.Text)) then
    Date2.Text := Date1.Text
  else
    if (DateFocused = Date2) and (StrToDate(Date1.Text) > StrToDate(Date2.Text)) then
      Date1.Text := Date2.Text;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Calculer Date1 et Date2 à partir de l'Exo
Mots clefs ... :
*****************************************************************}
procedure TOF_Meth.DoExoToDateOnChange(Exo: THValComboBox; Date1, Date2: THEdit);
begin
  if (Exo=nil) or (Date1=nil) or (Date2=nil) then Exit;
  ExoToDates(Exo.Value, Date1, Date2);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Calculer Periode1 et Periode2 à partir de l'Exo
Mots clefs ... :
*****************************************************************}
procedure TOF_Meth.DoExoToPeriodeOnChange(Exo, Periode1, Periode2: THValComboBox);
var Date1, Date2: TEdit;
begin
  if (Exo=nil) or (Periode1=nil) or (Periode2=nil) then Exit;
  Date1 := nil;
  Date2 := nil;
  try
    Date1 := TEdit.Create(nil);
    Date2 := TEdit.Create(nil);
    ExoToDates(Exo.Value, Date1, Date2);
    ListePeriode(Exo.Value, Periode1.Items, Periode1.Values, true);
    ListePeriode(Exo.Value, Periode2.Items, Periode2.Values, false);
    Periode1.ItemIndex := Periode1.Values.IndexOf(Date1.Text);
    Periode2.ItemIndex := Periode2.Values.IndexOf(Date2.Text);
  finally
    Date1.Free;
    Date2.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Contrôler la fourchette des journaux: Jal1 et Jal2
Mots clefs ... :
*****************************************************************}
procedure TOF_Meth.DoJalOnExit(JalFocused, Jal1, Jal2: THEdit);
begin
  if (Jal1 = nil) or (Jal2 = nil) then exit;
  (*
  if RechDomLookupCombo(THCritMaskedit(JalFocused)) = 'Error' then begin
    JalFocused.SetFocus;
    SysUtils.Abort;
    Exit;
  end;
  *)
  if (JalFocused = Jal1) and (Jal1.Text > Jal2.Text) and (Jal2.Text <> '') then
    Jal2.Text := Jal1.Text
  else
    if (JalFocused = Jal2) and (Jal1.Text > Jal2.Text) then
      Jal1.Text := Jal2.Text;
end;

procedure TOF_Meth.DoCompteOnExit(CptFocused, Cpt1, Cpt2: THEdit) ;
BEGIN
if (Cpt1 = nil) or (Cpt2 = nil) then exit;
if (CptFocused = Cpt1) and (Cpt1.Text > Cpt2.Text) and (Cpt2.Text <> '') then
  Cpt2.Text := Cpt1.Text
  else if (CptFocused = Cpt2) and (Cpt1.Text > Cpt2.Text) then
  Cpt1.Text := Cpt2.Text;
END;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Contrôler la fourchette des Numéro: Num1 et Num2
Mots clefs ... :
*****************************************************************}
procedure TOF_Meth.DoNumOnExit(NumFocused, Num1, Num2: THEdit);
BEGIN
if (Num1 = nil) or (Num2 = nil) then exit;
if StrToFloat(Num1.Text) < 0 then Num1.Text := '0';
if StrToFloat(Num2.Text) < 0 then Num2.Text := '0';
if StrToFloat(Num1.Text) > High(integer) then Num1.Text := InttoStr(High(integer));
if StrToFloat(Num2.Text) > High(integer) then Num2.Text := InttoStr(High(integer));
if (NumFocused = Num1) and (StrToFloat(Num1.Text) > StrToFloat(Num2.Text)) then
  Num2.Text := Num1.Text
else if (NumFocused = Num2) and (StrToFloat(Num1.Text) > StrToFloat(Num2.Text)) then
  Num1.Text := Num2.Text;
END;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Contrôler la fourchette des Périodes: Periode11 et Periode2
Suite ........ : Paramètres par défault: Date1 et Date2 pour calculer la
Suite ........ : fourchette date à partir de la fouchette période.
Mots clefs ... :
*****************************************************************}
procedure TOF_Meth.DoPeriodeOnChange(PeriodeFocused, Periode1, Periode2: THValComboBox; Date1, Date2: THEdit);
var sDate1, sDate2: string[10];
BEGIN
if (Periode1 = nil) or (Periode2 = nil) then Exit;
sDate1 := THValComboBox(Periode1).Value;
sDate2 := THValComboBox(Periode2).Value;
if (PeriodeFocused = Periode1) and (StrToDate(sDate1) > StrToDate(sDate2)) then begin
  Periode2.ItemIndex := Periode1.ItemIndex;
  sDate2 := sDate1;
end else if (PeriodeFocused = Periode2) and (StrToDate(sDate1) > StrToDate(sDate2)) then begin
  Periode1.ItemIndex := Periode2.ItemIndex;
  sDate1 := sDate2;
end;
if Date1 <> nil then Date1.Text := sDate1;
if Date2 <> nil then Date2.Text := sDate2;
END;

function TOF_Meth.NbMoisDansExo(Exo: string): integer;
var Date1, Date2: TEdit;
    ExoDate: TExoDate;
    MM,AA,NbMois: word;
BEGIN
Date1 := nil;
Date2 := nil;
try
  Date1 := TEdit.Create(nil);
  Date2 := TEdit.Create(nil);
  ExoToDates(Exo, Date1, Date2);
  ExoDate.Deb := StrToDate(Date1.Text);
  ExoDate.Fin := StrToDate(Date2.Text);
  NOMBREPEREXO(ExoDate,MM,AA,NbMois) ;
  Result := NbMois;
finally
  Date1.Free;
  Date2.Free;
end;
END;

function TOF_Meth.ExoPrecedent(Exo: string): string;
var i, j, im : integer;
BEGIN
if Exo = VH^.Suivant.Code then begin
  result := VH^.EnCours.Code;
  Exit;
end;
im:=0;
if Exo = VH^.EnCours.Code then begin
  result := VH^.Precedent.Code;
  Exit;
end;
for i:=5 downto 1 do
if trim(VH^.ExoClo[i].Code)<>'' then begin
  im := i;
  break;
end;
j := -1;
for i:=1 to im do
if VH^.ExoClo[i].Code=Exo then begin
  j := i;
  break;
end;
if j > 1 then result := VH^.ExoClo[j-1].Code
else result := '';
END;

procedure TOF_Meth.ControlVisible(aControl: TControl; const Visible: boolean);
begin
  aControl.Visible := Visible;
end;

procedure TOF_Meth.SetElementComboEtat(ComboEtat: THValComboBox; TypeEtat, NatEtat, Codes: string);
var s, s1, sDefaut: string; i: integer; slV, slI: HTStrings;
BEGIN
s:=Codes;
sDefaut:=ComboEtat.Values[ComboEtat.ItemIndex];
slI:=nil;
slV:=nil;
try
slI:=HTStringList.Create;
slV:=HTStringList.Create;
slI.Assign(ComboEtat.Items);
slV.Assign(ComboEtat.Values);
ComboEtat.Items.Clear ;
ComboEtat.Values.Clear ;
while s<>'' do begin
  s1:=ReadTokenST(s);
  for i:=0 to slV.Count-1 do
  if slV.Strings[i]=s1 then begin
    ComboEtat.Items.Add(slI.Strings[i]);
    ComboEtat.Values.Add(slV.Strings[i]);
  end;
end;
finally
  slI.Free;
  slV.Free;
end;
ComboEtat.ItemIndex:=ComboEtat.Values.IndexOf(sDefaut);
END;

procedure TOF_Meth.FillComboValues(Exo1, Exo2: THValComboBox);
var i: integer; s: string;
BEGIN
if (Exo1=nil) or (Exo2=nil) then exit;
Exo2.Items.Clear;
Exo2.Values.Clear;
for i:=0 to Exo1.Values.Count-1 do begin
  s:=Exo1.Values[i];
  Exo2.Values.Add(ReadTokenST(s));
  Exo2.Items.Add(Exo1.Items[i]);
end;
END;

procedure TOF_Meth.PageOrder(Pages: TPageControl);
var i: integer;
begin
  for i := 0 to Pages.PageCount - 1 do
    Pages.Pages[i].PageIndex := Pages.Pages[i].Tag;
end;

//Change libellé et datatype pour les champs compte selon general
procedure TOF_Meth.TypeChampCompte(labCpt: THLabel; Cpt1, Cpt2: THEdit;
  General: boolean);
begin
  if General then begin
    labCpt.Caption := 'Comptes généraux de';
    Cpt1.Name := 'E_GENERAL';
    Cpt2.Name := 'E_GENERAL_';
    Cpt1.DataType := 'TZGENERAL';
    Cpt2.DataType := 'TZGENERAL';
  end
  else begin
    labCpt.Caption := 'Comptes auxiliaires de';
    Cpt1.Name := 'E_AUXILIAIRE';
    Cpt2.Name := 'E_AUXILIAIRE_';
    Cpt1.DataType := 'TZTTOUS';
    Cpt2.DataType := 'TZTTOUS';
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/06/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_Meth.OnLoad;
begin
  inherited;
  ChargementCritEdt;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/06/2002
Modifié le ... :   /  /
Description .. : Declencle l' impression de l' état
Mots clefs ... :
*****************************************************************}
procedure TOF_Meth.FTimerTimer(Sender: TObject);
begin
  FTimer.Enabled := False;

  if FCritEdtChaine <> nil then
  begin
    V_Pgi.NoPrintDialog := True;

    // On force le Print ou le Preview dans les états chainés
    if FCritEdtchaine.CritEdtChaine.AuFormatPDF then
      TCheckBox(GetControl('FAPERCU')).Checked := True
    else
    begin
      // GCO - 05/09/2007 - FQ 20489 -> marche pas, doit marcher après correction
      // de la fiche AGL que j'ai faite pour le serveur CWAS
      V_Pgi.NbDocCopies := FCritEdtchaine.CritEdtChaine.NombreExemplaire;
      TCheckBox(GetControl('FAPERCU')).Checked := False;
    end;
  end;

{$IFDEF GIL}
  if not V_Pgi.Sav then
  begin
{$ENDIF}
  TToolBarButton97(GetControl('BVALIDER')).Click;

  if FCritEdtChaine <> nil then
    Ecran.Close;
{$IFDEF GIL}
  end;
{$ENDIF}

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : __/__/____
Modifié le ... : 08/09/2006
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_Meth.OnClose;
begin
  inherited;
  TheData := nil;
  FCritEdtChaine.Free;
  FCritEdtChaine := nil;
  FTimer.Free;
  FTimer := nil;

  if Assigned(FProvisoire) then FreeAndNil(FProvisoire);

  // Gestion MultiSoc SBO 17/08/2005
  if EstMultiSoc then
    V_PGI.EnableTableToView := FBoTableToView ;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/02/2003
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_Meth.ChargementFCritEdtchaine;
begin
  if FCritEdtChaine = nil then Exit;
  if FCritEdtchaine.CritEdtChaine.Utiliser then
  begin
    // Chargement des filtres de l'état dans la combo FFiltres
    if FCritEdtChaine.CritEdtChaine.FiltreUtilise <> '' then
    begin
      // Chargement du Filtre utilisé par les états chainés
      FFiltres.ItemIndex := FFiltres.Items.Indexof( FCritEdtChaine.CritEdtChaine.FiltreUtilise );
      FFiltres.OnChange(FFiltres);
    end;
    FTimer.Enabled := True;
  end;
end;

procedure TOF_Meth.OnChangeFiltre(Sender: TObject);
begin
  //SG6 14.03.05 Gestion des filtres FQ 14996
  cLoadfiltre := True;
  if Ecran is TFQRS1 then
    TFQRS1(Ecran).FFiltresChange(Sender)
  else
    TFStat(Ecran).FFiltresChange(Sender);
  cLoadfiltre := False;
end;

{On teste si le THEdit est de type string, qu'il a une tablette attachée et qu'il est lié
 à un autre THEdit marquant une fin de fourchette : dans ce cas, on teste s'il contient un
 caractère joker
{---------------------------------------------------------------------------------------}
function TOF_Meth.HasJoker(Sender : TObject) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Ed : THEdit;
begin
  Result := False;
  if Sender is THEdit then begin
    Ed := THEdit(Sender);
    if (Ed.OpeType = otString) and (Ed.DataType <> '') then begin
      if TFQRS1(Ecran).FindComponent(Ed.Name + '_') <> nil then
        Result := TestJoker(Ed.Text);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_Meth.ChargementCritEdt;
begin
  //
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 28/03/2006
Modifié le ... :
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_Meth.GetNumTableLibre: String;
begin
  Result := '' ;
  if TableLibre = nil then Exit;
  if TableLibre.ItemIndex < 0 then Exit ;
  Result := Copy(TableLibre.Value, 3, 1);
end;

////////////////////////////////////////////////////////////////////////////////

{JP 28/06/06 : Récupération du combo gérant les établissements
{---------------------------------------------------------------------------------------}
function TOF_Meth.GetComboEtab : THValComboBox;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  if FComboEtab = nil then begin
    for n := 0 to Ecran.ComponentCount - 1 do
      if (Pos('ETAB', Ecran.Components[n].Name) > 0) and
         (Ecran.Components[n] is THValComboBox) then begin
        FComboEtab := THValComboBox(Ecran.Components[n]);
        Break;
      end;
  end;
  Result := FComboEtab;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_Meth.GereEtablissement;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ComboEtab) then begin
    {Si l'on ne gère pas les établissement ...}
    if not VH^.EtablisCpta  then begin
      {... on affiche l'établissement par défaut}
      SetControlText(ComboEtab.Name, VH^.EtablisDefaut);
      {... on désactive la zone}
      SetControlEnabled(ComboEtab.Name, False);
    end

    {On gère l'établisement, donc ...}
    else begin
      {... On commence par regarder les restrictions utilisateur}
      PositionneEtabUser(ComboEtab);
      {... s'il n'y a pas de restrictions, on reprend le paramSoc
       JP 25/10/07 : FQ 19970 : Finalement on oublie l'option de l'établissement par défaut
      if GetControlText(ComboEtab.Name) = '' then begin
        {... on affiche l'établissement par défaut
        SetControlText(ComboEtab.Name, VH^.EtablisDefaut);
        {... on active la zone
        SetControlEnabled(ComboEtab.Name, True);
      end;}
    end;
  end;
end;

{JP 28/06/06 : FQ 16149 : on s'assure que le filtre ne va pas à l'encontre des
               restrictions utlisateurs
{---------------------------------------------------------------------------------------}
procedure TOF_Meth.ControlEtab;
{---------------------------------------------------------------------------------------}
var
  Eta : string;
begin
  if not Assigned(ComboEtab) then Exit;
  {S'il n'y a pas de gestion des établissement, logiquement, on ne force pas l'établissement !!!}
  if not VH^.EtablisCpta then Exit;

  Eta := EtabForce;
  {S'il y a une restriction utilisateur et qu'elle ne correspond pas au contenu de la combo ...}
  if (Eta <> '') and (Eta <> GetControlText(ComboEtab.Name)) then begin
    {... on affiche l'établissement des restrictions}
    SetControlText(ComboEtab.Name, Eta);
    {... on désactive la zone}
    SetControlEnabled(ComboEtab.Name, False);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_Meth.OnUpdate;
begin
  inherited;

  // GCO - 18/09/2006
  if assigned (FProvisoire) then   // FQ 19385
  begin
    if FDateFinEdition <> idate1900 then
    begin
      if FDateFinEdition <= FDateDernCloture then
        FProvisoire.Text := ''
      else
        FProvisoire.Text := TraduireMemoire('(provisoire)');
    end;
  end;

  //if (FCritEdtChaine <> nil) and (not FCritEdtchaine.CritEdtChaine.AuFormatPDF) then
  //begin
    // GCO - 05/09/2007 - FQ 20489 -> marche pas
    //Printer.Copies := FCritEdtchaine.CritEdtChaine.NombreExemplaire;
    //V_Pgi.DefaultDocCopies := FCritEdtchaine.CritEdtChaine.NombreExemplaire;
    //V_Pgi.NbDocCopies := FCritEdtchaine.CritEdtChaine.NombreExemplaire;
  //end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/09/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOF_Meth.GetDateDernCloture : TDateTime;
var lDateTemp : TDateTime;
begin
  Result := idate1900;

  lDateTemp := GetParamSocSecur('SO_DATECLOTUREPER', idate1900);
  if lDateTemp <> idate1900 then
  begin
    Result := lDateTemp;
  end;

  if (VH^.Encours.Deb -1) > Result then
  begin
    Result := VH^.EnCours.Deb - 1;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

initialization
RegisterClasses([TOF_Meth]) ;


end.
