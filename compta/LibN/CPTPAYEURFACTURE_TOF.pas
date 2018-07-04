{***********UNITE*************************************************
Auteur  ...... : Fabrice PAILLARD
Créé le ...... : 12/07/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPTPAYEURFACTURE ()
Mots clefs ... : TOF;CPTPAYEURFACTURE
*****************************************************************}
Unit CPTPAYEURFACTURE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_main,
     EdtREtat,
     QRS1,
{$else}
     Maineagl,
     eMul,
     UtileAGL,
     eQRS1,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97,
     UTOB,
     Ent1,
     TofMeth,
     uLibExercice,
     HQry,
     UTILEDT,
     HCompte,
     Paramsoc,
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     SAISUTIL
     ;

Type
  TOF_CPTPAYEURFACTURE = Class (TOF_Meth)
  private
    FTobEcr: TOB;

    {Critères}
    NatureJournal:  THRadioGroup;
    Exercice:       THValComboBox;
    DateComptaDe:   THEdit;
    DateComptaA:    THEdit;
    JournalDe:      THEdit;
    JournalA:       THEdit;
    NumeroPieceDe:  THEdit;
    NumeroPieceA:   THEdit;
    ComptePayeurDe: THEdit;
    ComptePayeurA:  THEdit;
    {Critères avancés}
    RefInterne:       THEdit;
    EcrValide:        TCheckBox;
    Devise:           THValcomboBox;
    SymboleDevise:    THEdit;
    {Options d'édition}
    ChoixMonnaie:     THRadioGroup;
    ChkSymboleDevise: TCheckBox;
    {Dev}
    AvecNatureJournal:  THEdit;
    AvecChoixMonnaie:   THEdit;
    JnlPayeur:          THEdit;
    NumExercice:        THEdit;

    procedure BValiderClick(Sender: TObject);
    procedure ExoOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
    procedure JournalOnExit(Sender: TObject) ;
    procedure NatureJournalChange(Sender: TObject);
    procedure ChoixMonnaieChange(Sender: TObject);

    procedure RecupCritEdt;
    procedure RemplirListeFacture(QPrinc: TQuery; JalP: String; DateP: TDateTime; NumP: Integer);
    procedure RemplirTobEcr;
    function  SQLPrinc: String ;
    function  SQLEcr(CptPayeur, Devise: String; JalP, ExoP: String; DateP: TDateTime; NumP: Integer): String ;
    procedure AuxiElipsisClick(Sender : TObject);

  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

  procedure CPLanceFiche_CPTPAYEURFACTURE;
  function ComposantToCritere(WinCtrl: TWinControl): String;
  procedure PositionneFourchetteST(TC1,TC2 : THEdit; tt: TZoomTable) ;

Implementation

uses
  {$IFDEF MODENT1}
  //CPTypeCons,
  {$ENDIF MODENT1}
  UTofMulParamGen; {26/04/07 YMO F5 sur Auxiliaire }

procedure CPLanceFiche_CPTPAYEURFACTURE;
begin
  AGLLanceFiche('CP','CPTPAYEURFACTURE','','','');
end;

procedure TOF_CPTPAYEURFACTURE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPTPAYEURFACTURE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPTPAYEURFACTURE.OnUpdate ;
var
  ClauseWhere: String;
begin
  Inherited ;

  ClauseWhere := '';
  if RefInterne.Text<>'' then
    ClauseWhere := ClauseWhere + 'And UPPER(E_REFINTERNE) like "'+TraduitJoker(RefInterne.Text)+'" ';
  if ComboEtab.Value<>'' then
    ClauseWhere := ClauseWhere + ' And E_ETABLISSEMENT="'+ComboEtab.Value+'" ';
  if Devise.Value<>'' then
    ClauseWhere := ClauseWhere + ' And E_DEVISE="'+Devise.Value+'" ';
  if EcrValide.State <> CbGrayed then begin
    if EcrValide.Checked then
      ClauseWhere := ClauseWhere + ' AND E_VALIDE = "X" '
    else
      ClauseWhere := ClauseWhere + ' AND E_VALIDE = "-" ';
  end;

  TFQRS1(Ecran).WhereSQL := ClauseWhere;
end ;

procedure TOF_CPTPAYEURFACTURE.OnLoad ;
var
  RDev: RDevise;
begin
  Inherited ;

  SymboleDevise.Text := '';
  if ChkSymboleDevise.Checked then begin
    if (ChoixMonnaie.ItemIndex = 0) or (Devise.Value=V_PGI.DevisePivot) then begin
      SymboleDevise.Text := ' '+V_PGI.SymbolePivot;
    end
    else begin
      RDev.Code := Devise.Value;
      GetInfosDevise(RDev);
      SymboleDevise.Text := ' '+RDev.Symbole;
    end;
  end;
  RecupCritEdt;
end ;

procedure TOF_CPTPAYEURFACTURE.OnArgument (S : String ) ;
begin
  Inherited ;

  Pages          := TPageControl(Getcontrol('PAGES', true));
  {Critères}
  NatureJournal  := THRadioGroup(Getcontrol('NATUREJAL', true));
  Exercice       := THValComboBox(Getcontrol('E_EXERCICE', true));
  DateComptaDe   := THEdit(GetControl('DATECOMPTABLE', true));
  DateComptaA    := THEdit(GetControl('DATECOMPTABLE_', true));
  JournalDe      := THEdit(GetControl('E_JOURNAL', true));
  JournalA       := THEdit(GetControl('E_JOURNAL_', true));
  NumeroPieceDe  := THEdit(GetControl('E_NUMEROPIECE', true));
  NumeroPieceA   := THEdit(GetControl('E_NUMEROPIECE_', true));
  ComptePayeurDe := THEdit(GetControl('COMPTEPAYEUR', true));
  ComptePayeurA  := THEdit(GetControl('COMPTEPAYEUR_', true));
  {Critères avancés}
  RefInterne     := THEdit(Getcontrol('REFERENCEINTERNE', true));
  EcrValide      := TCheckBox(Getcontrol('ECRITUREVALIDE', true));
  Devise         := THValcomboBox(Getcontrol('E_DEVISE', true));
  {Options d'édition}
  ChoixMonnaie      := THRadioGroup(Getcontrol('CHOIXMONNAIE', true));
  ChkSymboleDevise  := TCheckBox(Getcontrol('CHKSYMBOLEDEVISE', true));
  {Dev}
  AvecNatureJournal := THEdit(Getcontrol('AVECNATUREJAL', true));
  AvecChoixMonnaie  := THEdit(Getcontrol('AVECCHOIXMONNAIE', true));
  JnlPayeur         := THEdit(Getcontrol('JNLPAYEUR', true));
  NumExercice       := THEdit(Getcontrol('NUMEXERCICE', true));
  SymboleDevise     := THEdit(Getcontrol('SYMBOLEDEVISE', true));


  DateComptaDe.OnExit   := DateOnExit;
  DateComptaA.OnExit    := DateOnExit;
  JournalDe.OnExit      := JournalOnExit;
  JournalA.OnExit       := JournalOnExit;
  Exercice.onChange     := ExoOnChange;
  NatureJournal.OnClick := NatureJournalChange;
  ChoixMonnaie.OnClick  := ChoixMonnaieChange;
  TToolbarButton97(GetControl('BVALIDER')).OnClick := BValiderClick;

  CInitComboExercice(Exercice);       // Init de la combo Exercice en Relatif
(********************************************************************************)
(********************************************************************************)
  {Dans l'édition de la balance il n'est pas necessaire d'initialiser la property itemindex.
   Je n'ai pas trouvé où cette property est initialisée}
  if (ComboEtab.ItemIndex = -1) and (ComboEtab.Items.Count > 0) then
    ComboEtab.ItemIndex := 0;
(********************************************************************************)
(********************************************************************************)
  Devise.ItemIndex := 0;

(********************************************************************************)
(********************************************************************************)
  {Dans le générateur d'état je ne peux pas utiliser la valeur d'un THRadioGroup}
  NatureJournalChange(NatureJournal);
  ChoixMonnaieChange(ChoixMonnaie);
(********************************************************************************)
(********************************************************************************)

  if (CtxPCl in V_PGI.PgiContexte) and (VH^.CPExoRef.Code <>'') then
    Exercice.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
  else
    Exercice.Value := CExerciceVersRelatif(VH^.Entree.Code) ;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    ComptePayeurDe.OnElipsisClick:=AuxiElipsisClick;
    ComptePayeurA.OnElipsisClick:=AuxiElipsisClick;
  end;

end ;

procedure TOF_CPTPAYEURFACTURE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPTPAYEURFACTURE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPTPAYEURFACTURE.OnCancel () ;
begin
  Inherited ;
end ;


function ComposantToCritere(WinCtrl: TWinControl): String;
var
  i:    Integer;
  s:    String;
  Ctrl: TControl;
begin
  Result := '';
  for i:=0 to WinCtrl.ControlCount-1 do begin
    Ctrl   := WinCtrl.Controls[i];
    s      := '';
    if Ctrl is THRadioGroup then
      s := IntToStr((Ctrl as THRadioGroup).ItemIndex)
    else if Ctrl is THEdit then
      s := (Ctrl as THEdit).Text
    else if Ctrl is THValComboBox then
      s := (Ctrl as THValComboBox).Text
    else if Ctrl is TCheckBox then
      s := CheckToStr((Ctrl as TCheckBox).State);
    if s <> '' then
      Result := Result + '`'+Ctrl.Name+'=' + s 
    else if Ctrl is TWinControl then
      Result := Result + ComposantToCritere(Ctrl as TWinControl);
    end;
end;

procedure TOF_CPTPAYEURFACTURE.BValiderClick(Sender: TObject);
var
  Sql:         String;
  CritEdition: String;
begin
  RecupCritEdt;

  Sql := '';
  if RefInterne.Text<>'' then
    Sql := Sql + 'And UPPER(E_REFINTERNE) like "'+TraduitJoker(RefInterne.Text)+'" ';
  if ComboEtab.Value<>'' then
    Sql := Sql + ' And E_ETABLISSEMENT="'+ComboEtab.Value+'" ';
  if Devise.Value<>'' then
    Sql := Sql + ' And E_DEVISE="'+Devise.Value+'" ';

  FTobEcr := TOB.Create('', nil, -1);
  try
    RemplirTobEcr;
    {Dans le paramétrage de l'état, il faut forcer 2 passages sinon la première ligne
     s'édite 2 fois}
    {Le paramètre Pages n'est pas utilisé en mode CWAS, il faut donc passer les
     paramètres dans la chaine CritEdition. Voir CPCONTROLERUB_TOF}
    CritEdition := ComposantToCritere(Pages);
    if CritEdition <> '' then
      CritEdition := Copy(CritEdition, 2, Length(CritEdition));
    LanceEtatTob('E', 'TFA', 'TFA', FTobEcr, TCheckBox(GetControl('FAPERCU')).Checked,
      False, TCheckBox(GetControl('FREDUIRE')).Checked, nil, '', Ecran.Caption, False,
      0, CritEdition);
  finally
    FreeAndNil(FTobEcr);
    end;
end;

procedure TOF_CPTPAYEURFACTURE.RemplirTobEcr;
var
  Sql:         string ;
  QPrinc:      TQuery ;

  JalP,St,St1: String ;
  DateP:       TDateTime ;
  NumP,NumE:   Integer ;
  Y,M,D:       Word ;
  Ok:          Boolean;
begin
  Sql := SQLPrinc;
  QPrinc := OpenSql(Sql, True);
  try
    while not QPrinc.Eof do begin
      St := QPrinc.FindField('E_PIECETP').AsString ;
      If St<>'' Then BEGIN
        St1:=ReadTokenSt(St) ; Y:=StrToInt(Copy(St1,1,4)) ; M:=StrToInt(Copy(St1,5,2)) ; D:=StrToInt(Copy(St1,7,2)) ; DateP:=EncodeDate(Y,M,D) ;
        St1:=ReadTokenSt(St) ; NumP:=StrToInt(St1) ;
        St1:=ReadTokenSt(St) ;
        St1:=ReadTokenSt(St) ; NumE:=StrToInt(St1) ;
        St1:=ReadTokenSt(St) ; JalP:=St1 ;

        Ok := (NumE <= 1);
        if NumeroPieceDe.Text<>'' then
          Ok := Ok and (NumP >= StrToInt(NumeroPieceDe.Text));
        if NumeroPieceA.Text<>'' then
          Ok := Ok and (NumP <= StrToInt(NumeroPieceA.Text));
        if JournalDe.Text<>'' then
          Ok := Ok and (CompareStr(JalP, JournalDe.Text) >= 0);
        if JournalA.Text<>'' then
          Ok := Ok and (CompareStr(JalP, JournalA.Text)  <= 0);;

        If Ok Then begin
          RemplirListeFacture(QPrinc, JalP, DateP, NumP);
          end;
        end;
      QPrinc.Next ;
    end ;
  finally
    Ferme(QPrinc) ;
    end;
end;

procedure TOF_CPTPAYEURFACTURE.RemplirListeFacture(QPrinc: TQuery; JalP: String; DateP: TDateTime; NumP: Integer);
var
  Sql:           string;
  ExoP:          string;
  CptPayeur:     string;
  Devise:        string;
begin
  ExoP          := QPrinc.FindField('E_EXERCICE').AsString;
  CptPayeur     := QPrinc.FindField('E_AUXILIAIRE').AsString;
  Devise        := QPrinc.FindField('E_DEVISE').AsString;

  Sql := SQLEcr(CptPayeur, Devise, JalP, ExoP, DateP, NumP);
  FTobEcr.LoadDetailFromSQL(Sql, True);
end;

function TOF_CPTPAYEURFACTURE.SQLPrinc: String;
{Liste des écritures générées sur le tiers payeur}
var
  Sql  : string ;
  JalTP: string;
  DateDeb, DateFin: TDateTime;
begin
  If NatureJournal.Value = 'H' Then
    JalTP := VH^.JalATP         {Journal d'achat}
  else
    JalTP := VH^.JalVTP ;       {Journal de vente}

  DateDeb := StrToDate(DateComptaDe.Text);
  DateFin := StrToDate(DateComptaA.Text);

  Sql := 'Select E_EXERCICE, E_PIECETP, E_AUXILIAIRE, E_DEVISE' +
    ' From Ecriture ' +
    ' Where E_QUALIFPIECE="N" ' +
    ' And E_DATECOMPTABLE>="'+USDateTime(DateDeb)+'" And E_DATECOMPTABLE<="'+USDateTime(DateFin)+'" '+
    ' And E_JOURNAL="'+JalTP+'" '+
    ' And E_NUMLIGNE=2 ' +
    ' And E_ECRANOUVEAU="N" ' +
    ' And E_PIECETP<>"" ' +
    ' And E_QUALIFORIGINE="TP" ' ;

  Sql := Sql + 'And E_EXERCICE = "'+Exercice.Value+'" ';
  Sql := CMajRequeteExercice(Exercice.Value, Sql);
  if RefInterne.Text<>'' then
    Sql := Sql + 'And UPPER(E_REFINTERNE) like "'+TraduitJoker(RefInterne.Text)+'" ';
  if ComboEtab.Value<>'' then
    Sql := Sql + ' And E_ETABLISSEMENT="'+ComboEtab.Value+'" ';
  if Devise.Value<>'' then
    Sql := Sql + ' And E_DEVISE="'+Devise.Value+'" ';
  If ComptePayeurDe.Text<>'' Then
    Sql := Sql + ' And E_AUXILIAIRE>="'+ComptePayeurDe.Text+'" ';
  If ComptePayeurA.Text<>'' Then
    Sql := Sql + ' And E_AUXILIAIRE<="'+ComptePayeurA.Text+'" ';

  Sql := Sql + ' Order By E_AUXILIAIRE, E_JOURNAL, E_NUMEROPIECE ';
  Result := Sql;
end;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 26/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTPAYEURFACTURE.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


function TOF_CPTPAYEURFACTURE.SQLEcr(CptPayeur, Devise: String; JalP, ExoP: String; DateP: TDateTime; NumP: Integer): String;
{Liste des écritures constituant la pièce d'origine}
var
  Sql: String;
begin
  Sql := 'Select '+
    ' E_EXERCICE, E_JOURNAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_GENERAL,' +
    ' E_AUXILIAIRE, E_MODEPAIE, E_DATEECHEANCE, E_ETABLISSEMENT, E_NUMECHE, E_VALIDE, ' +
    ' E_NATUREPIECE, E_ECHE, E_ETATLETTRAGE, E_DEVISE, E_TAUXDEV, E_COTATION, E_DATETAUXDEV, ' +
    ' E_REFINTERNE, E_REFEXTERNE, E_DATEREFEXTERNE, E_LIBELLE, E_REFLIBRE, E_AFFAIRE, ' +
    ' TTIERS.T_LIBELLE T_LIBELLE, ' +
    ' G_LIBELLE, ' +
    ' J_LIBELLE, ' +
    ' TPAYEUR.T_AUXILIAIRE T_CPTPAYEUR, TPAYEUR.T_LIBELLE T_LIBELLEPAYEUR, '+
    '';
  (*if EstMonnaieIn(Devise) then
    Sql := Sql + ' "X" MONNAIEIN, '
  else *)
    Sql := Sql + ' "-" MONNAIEIN, ';

  Sql := Sql + ' E_DEBIT, E_CREDIT, E_DEBITDEV, E_CREDITDEV ';

  Sql := Sql +
    ' From Ecriture '+
      ' LEFT JOIN TIERS TTIERS On TTIERS.T_AUXILIAIRE=E_AUXILIAIRE '+
      ' LEFT JOIN TIERS TPAYEUR on TPAYEUR.T_AUXILIAIRE="'+CptPayeur+'"'+
      ' LEFT JOIN GENERAUX On G_GENERAL=E_GENERAL '+
      ' LEFT JOIN JOURNAL  On J_JOURNAL=E_JOURNAL '+
    ' Where E_JOURNAL="'+JalP+'"'+
    '   And E_EXERCICE="'+ExoP+'"'+
    '   And E_DATECOMPTABLE="'+USDATETIME(DateP)+'"'+
    '   And E_NUMEROPIECE='+IntToStr(NumP)+
    '   And E_QUALIFPIECE="N" ';

  if NumeroPieceDe.Text<>'' then
    Sql := Sql + ' And E_NUMEROPIECE>='+NumeroPieceDe.Text+' ';
  if NumeroPieceA.Text<>'' then
    Sql := Sql + ' And E_NUMEROPIECE<='+NumeroPieceA.Text+' ';

  if EcrValide.State <> CbGrayed then begin
    if EcrValide.Checked then
      Sql := Sql + ' AND E_VALIDE = "X" '
    else
      Sql := Sql + ' AND E_VALIDE = "-" ';
  end;

  if JournalDe.Text<>'' then
    Sql := Sql + ' And E_JOURNAL>="'+JournalDe.Text+'" ';
  if JournalA.Text<>'' then
    Sql := Sql + ' And E_JOURNAL<="'+JournalA.Text+'" ';

  If NatureJournal.Value = 'H' Then
    Sql := Sql + ' And J_NATUREJAL="ACH" '       {Nature du journal: achat}
  else
    Sql := Sql + ' And J_NATUREJAL="VTE" ';      {Nature du journal: vente}

  Sql := Sql + ' Order By T_CPTPAYEUR, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE';
  Result := Sql;
end;

procedure TOF_CPTPAYEURFACTURE.DateOnExit(Sender: TObject);
var
  DateD: TDateTime;
  DateF: TDateTime;
begin
  DateD := StrToDate(DateComptaDe.Text);
  DateF := StrToDate(DateComptaA.Text);

  DoDateOnExit(THEdit(Sender), DateComptaDe, DateComptaA, DateD, DateF);
end;

procedure TOF_CPTPAYEURFACTURE.ExoOnChange(Sender: TObject);
begin
  CExoRelatifToDates(Exercice.Value, DateComptaDe, DateComptaA);
  NumExercice.Text := CRelatifVersExercice(Exercice.Value);
end;

procedure TOF_CPTPAYEURFACTURE.JournalOnExit(Sender: TObject);
begin
  DoJalOnExit(THEdit(Sender), JournalDe, JournalA);
end;

procedure TOF_CPTPAYEURFACTURE.ChoixMonnaieChange(Sender: TObject);
begin
  {Choix monnaie possible uniquement si une devise est sélectionnée}
  if (Devise.ItemIndex=0) or (Devise.Value=V_PGI.DevisePivot) then begin
    if ChoixMonnaie.ItemIndex <> 0 then
      HShowMessage('0;'+Ecran.Caption+';Vous devez d''abord sélectionner une devise particulière.;E;O;O;O;','','');
    ChoixMonnaie.ItemIndex := 0;
    end;
  AvecChoixMonnaie.Text := ChoixMonnaie.Value;

end;

procedure TOF_CPTPAYEURFACTURE.NatureJournalChange(Sender: TObject);
begin
  AvecNatureJournal.Text := NatureJournal.Value;
  if NatureJournal.Value = 'H' then begin
    JnlPayeur.Text     := VH^.JalATP;        {Journal d'achat}
    JournalDe.DataType := 'TTJALACHAT';
    end
  else begin
    JnlPayeur.Text     := VH^.JalVTP;        {Journal de vente}
    JournalDe.DataType := 'TTJALVENOD';
    end;
  JournalA.DataType := JournalDe.DataType;
  JournalA.Text     := '';
  JournalDe.Text    := '';
end;

{Voir unité UTILEDT}
procedure PositionneFourchetteST(TC1,TC2 : THEdit; tt: TZoomTable) ;
var
  St:          String;
  Q:           TQuery;
begin
  if (TC1.Text='') And (TC2.Text='') then begin
    Case CaseFic(tt) Of
      fbGene : St:='SELECT MIN(G_GENERAL), Max(G_GENERAL) FROM GENERAUX WHERE G_FERME="-" ' ;
      fbAux : St:='SELECT MIN(T_AUXILIAIRE), Max(T_AUXILIAIRE) FROM TIERS WHERE T_FERME="-" ' ;
      fbJal : St:='SELECT MIN(J_JOURNAL), Max(J_JOURNAL) FROM JOURNAL WHERE J_FERME="-" ' ;
      fbAxe1..fbAxe5 : St:='SELECT MIN(S_SECTION), Max(S_SECTION) FROM SECTION WHERE S_FERME="-" ' ;
      fbBudGen : St:='SELECT MIN(BG_BUDGENE), Max(BG_BUDGENE) FROM BUDGENE WHERE BG_FERME="-" ' ;
      fbBudJal : St:='SELECT MIN(BJ_BUDJAL), Max(BJ_BUDJAL) FROM BUDJAL WHERE BJ_FERME="-" ' ;
      fbBudSec1..fbBudSec5 : St:='SELECT MIN(BS_BUDSECT), Max(BS_BUDSECT) FROM BUDSECT WHERE BS_FERME="-" ' ;
      fbNatCpt : St:='SELECT MIN(NT_NATURE), Max(NT_NATURE) FROM NATCPTE WHERE NT_SOMMEIL="-" ' ;
      end;
    St:=St+RecupWhere(tt) ;

    {**************************************************************************}
    {Il faut modifier l'unité HCompte, la fonction RecupWhere car la nature
    d'un jnl de vente est VTE et pas VEN}
    {**************************************************************************}
    if tt = tzJvente then
      St := StringReplace(St, '"VEN"', '"VTE"', []);
    {**************************************************************************}
    {**************************************************************************}
    Q:=OpenSQL(St,TRUE) ;
    if not Q.EOF then begin
      TC1.Text := Q.Fields[0].AsString;
      TC2.Text := Q.Fields[1].AsString;
      end;
    Ferme(Q);
    end;
end;

procedure TOF_CPTPAYEURFACTURE.RecupCritEdt;
var
  Min, Max: String;
begin
  PositionneFourchetteST(ComptePayeurDe, ComptePayeurA, tzTPayeur);

  If NatureJournal.Value = 'H' Then
    PositionneFourchetteST(JournalDe, JournalA, tzJachat)  {Journal d'achat}
  else if (JournalDe.Text='') And (JournalA.Text='') then begin
    {La tablette affiche les journaux de vente et d'OD}
    PositionneFourchetteST(JournalDe, JournalA, tzJvente); {Journal de vente}
    Min := Trim(JournalDe.Text);
    Max := Trim(JournalA.Text);
    JournalDe.Text :='';
    JournalA.Text  :='';
    PositionneFourchetteST(JournalDe, JournalA, tzJOD);    {Journal d'OD}
    if (Min < JournalDe.Text) and (Min <> '') then
      JournalDe.Text := Min;
    if (Max > JournalA.Text) and (Max <> '') then
      JournalA.Text := Max;
    end;

  If NumeroPieceDe.Text = '' then
    NumeroPieceDe.Text := '0';
  If NumeroPieceA.Text = '' then
    NumeroPieceA.Text := '999999999';
end;

Initialization
  registerclasses ( [ TOF_CPTPAYEURFACTURE ] ) ;
end.

