{***********UNITE*************************************************
Auteur  ...... : Dumont Muriel
Créé le ...... : 09/02/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPMULTIBUD_QR1 ()
Mots clefs ... : TOF;CPMULTIBUD_QR1
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 7.00.001.00X    05/06/06    MD   FQ 18077 Correction plan de ruptures
 8.00.001.011    20/04/07    JP   FQ 19671 : On s'assure qu'il y a des budgets par catégorie
 8.00.001.024    06/07/07    JP   FQ 20973 : on vide la table temporaire en fermeture de fiche
--------------------------------------------------------------------------------------
*****************************************************************}
Unit CPMULTIBUD_QR1_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_Main,
     QRS1,
{$else}
     eMul,
     MaineAGL,
     eQRS1,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     Ent1,
     HEnt1,
     HMsgBox,
     ULibWindows,
     CPBALANCEBUDGET_TOF,
     uTob,
     UTOF ;

procedure Multibudgets;

Type
  TBudget      = record
                   Code       : string;
                   Libelle    : string;
                   DateDebBud : TDateTime;
                   DateFinBud : TDateTime;
                   end;
  TBudgets     = array[1..12] of TBudget;

  TOF_CPMULTIBUD_QR1 = class (TOF_CPBALANCEBUDGET) //Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnClose                  ; override ;
    procedure OnUpdate                 ; override;
  private
    FStWhereJou         : String;
    FStLibJou           : String;
    MonoExo             : boolean;
    DateDeb             : TDateTime;
    DateFin             : TDateTime;
    PasDeRubrique       : boolean;
    TabBudget,
    Budgets             : TBudgets;

    function GetRupture : Char;

  protected
    FGroupQuelleRupture : THRadioGroup;
    FGroupChoixRupt     : THRadioGroup;
    FRealise            : TCheckBox;
    FAvecRub            : TCheckBox;
    FAnaGene            : TCheckBox;
    FCategorie          : THValComboBox;
    FEXERCICE           : THValComboBox;
    FEXERCICE_          : THValComboBox;
    FPeriode1           : THValComboBox;
    FPeriode2           : THValComboBox;
    FSELECTCPTE         : THValComboBox;
    FJournal            : THMultiValComboBox;
    FPLANRUPTURES       : THValComboBox;
    FCodeRupt1,
    FCodeRupt2          : THValComboBox;
    ListeCodesRupture   : THValComboBox;
    AffMontants         : THValComboBox;
    FResol              : THValComboBox;

    procedure FPlanRupturesChange(Sender: TObject);
    procedure FCategorieChange  (Sender : TObject);
    procedure FExerciceChange (Sender : TObject);
    procedure FExercice_Change(Sender : TObject);
    procedure FPeriode1Change (Sender : TObject);
    procedure FPeriode2Change (Sender : TObject);
    procedure FSansRuptClick  (Sender : TObject);
    procedure FRupturesClick  (Sender : TObject); override;
    procedure FAvecRubClick(Sender : TObject);
    procedure AffmontantsChange (Sender : TObject );
    procedure FresolChange (Sender : TObject );

    { Affecte un évènement au contrôle passé en paramètre et renvoit le dit contrôle }
    function DoAssign(Control, Evt : string; Event : TNotifyEvent) : TControl;
    {Charge les critères en fonctions du journal budgétaire}
    procedure DebToDate;
    procedure FinToDate;
    function CalculNbBud : integer;
    procedure InitCompte;
    function TraiteListeJournal : boolean;
    procedure RecupPeriodesJournal( Journal : string;var DtDebBud : TDateTime;var DtFinBud : TDateTime);
    procedure SetSecdRequete ;
    procedure SetMvmtRequete;
    procedure InitCalcBud(T : Tob; var Tot1 : TMontTotal; Resol : string; Taux : Double; Col : Integer;Journal : string);
    procedure DoInsert(szCompte, szLibelleCompte, szRubCompte, szSection, szLibelleSection, szRubSection, szSectionTrie : String; iTri : Integer; TobFils : Tob);
    property Rupture : Char read GetRupture;
    function GetSQLMouvement(Compte, Section : string; Journal : string) : string;

  public
    function  GetMainRequete : string; override;
    procedure MajInfoRupture; override;
    procedure PrepCalcBud(SQL : string); override;
    procedure FRealiseClick (Sender : TObject); override;
    //function GetSQLMouvement(Compte, Section : string; Journal : string) : string; override;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcMetier,
  CPTypeCons,
  {$ENDIF MODENT1}
  UtilEDT,     {ExistBud, WhereLibre, OrderLibre}
  CALCOLE,     { TabloExt,GetCumul}
  TabLiEdt,    {ChoixTableLibreSur, ChoixTableLibre}
  TriTabLi,    {ChoixTriTableLibre}
  BUDSECT_TOM; {SectionRetrieBud}


{JP 20/04/07 : FQ  19671 : On s'assure qu'il y a des budgets par catégories 
{---------------------------------------------------------------------------------------}
procedure Multibudgets;
{---------------------------------------------------------------------------------------}
begin
  if ExisteSQL('SELECT BJ_CATEGORIE FROM BUDJAL WHERE BJ_CATEGORIE <> "" AND BJ_CATEGORIE IS NOT NULL') then
    AGLLanceFiche('CP', 'CPMULTIBUD_QR1', '', '', '')
  else
    PGIBox(TraduireMemoire('Cette édition ne peut être lancée car il n''existe pas de budget par catégorie'));
end;

{---------------------------------------------------------------------------------------}
function TOF_CPMULTIBUD_QR1.GetRupture : Char;
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

{---------------------------------------------------------------------------------------}
function TOF_CPMULTIBUD_QR1.CalculNbBud : integer;
{---------------------------------------------------------------------------------------}
var
  i                : Integer;
  Nb               : Integer;
  NbBudget         : Word;
  sz,
  lstValueTab,
  lstvalueTabDecoupe,
  lstValue,
  lstvalueDecoupe  : string;
begin
  // Budgets --> contient la liste de l'ensemble des budgets dans la limite de 8
  Fillchar(Budgets,SizeOf(Budgets),#0);
  //NbBudget := 0;
  if FJournal.Values.Count > 8 then
    Nb := 8
  else
    Nb := FJournal.Values.Count;

  for i := 0 to Nb - 1 do begin
    lstvalueDecoupe := FJournal.values[i];
    Budgets[i + 1].Code := lstvalueDecoupe;
    Budgets[i + 1].Libelle := FJournal.items[i];
    // Lecture du budget afin de récupérer ses dates de début période et fin période
    RecupPeriodesJournal(Budgets[i + 1].Code,Budgets[i + 1].DateDebBud,Budgets[i + 1].DateFinBud);
    //inc(NbBudget);
    end;

  if (FJournal.Values.Count > 8) and FJournal.Tous then
    NbBudget := 9
  else begin
    // TabBudget --> liste des budgets sélectionnés
    Fillchar(TabBudget,SizeOf(TabBudget),#0);
    if not FJournal.Tous then begin
      NbBudget := 0;
      lstValue := FJournal.value;
      while lstvalue <> '' do begin
        lstvalueDecoupe := ReadTokenSt(lstValue);
        if lstvalueDecoupe <> '' then inc(NbBudget);
        for i:= 1 to 12 do begin
          lstvaluetab := Budgets[i].code;
          lstvalueTabDecoupe := ReadTokenSt(lstvaluetab);
          if lstvalueDecoupe=lstvalueTabDecoupe then begin
            TabBudget[NbBudget].code := Budgets[i].code;
            TabBudget[NbBudget].libelle := Budgets[i].libelle;
            TabBudget[NbBudget].DateDebBud := Budgets[i].DateDebBud;
            TabBudget[NbBudget].DateFinBud := Budgets[i].DateFinBud;
            end;
          end;
        end;
      end
    else begin
      TabBudget := Budgets;
      NbBudget := FJournal.Values.Count;
      end;

    for i := 1 to 8 do begin
      // Ne pas mettre une chaine vide sinon la colonne n'est pas grisée
      sz := '   ';
      if TabBudget[i].code <> '' then
        sz := TabBudget[i].libelle;
      SetControlText('PER' + IntToStr(i), sz);
      end;
    end;
  result := NbBudget;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPMULTIBUD_QR1.GetSQLMouvement(Compte, Section, journal : string) : string;
{---------------------------------------------------------------------------------------}
var
  iPos : Integer;
  Chp1 : string;
  Chp2 : string;
begin

  Result := FMvmtRequete;

  if (aTypeBalance = bal_SeGe) then begin
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

  {#34 <=> '"' (double quote)}
  iPos := Pos(':C3', Result);
  System.Delete(Result, iPos, 3);
  System.Insert(#34 + Journal + #34, Result, iPos);

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.MajInfoRupture;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  FPlanRuptures.Datatype := 'ttRuptSect' + szAxeJal[2];
  { b md 05/06/06 FQ 18077 }
  (*if (FGROUPQUELLERUPTURE.ItemIndex=0) then FPlanRupturesChange(nil);*)
  if (FGROUPQUELLERUPTURE.ItemIndex = 0) then begin
    THValComboBox(getControl('FPlanRuptures', True)).ItemIndex := 0;
    FPlanRupturesChange(nil);
    end;
  { e md }
end;

Procedure PasserQry2Tob(var T:Tob;NomTab,Requette:string);
{Lek 250106 Charge la Requete dans Tob}
var Q:TQuery;
begin
  Q := OpenSql(Requette,true);
  T := Tob.Create(NomTab,nil,-1);
  T.LoadDetailDB(NomTab,'','',Q,true);
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.InitCalcBud(T : Tob; var Tot1 : TMontTotal; Resol : string; Taux : Double; Col : Integer; Journal : string);
{---------------------------------------------------------------------------------------}
var
  DateCompta : TDateTime;
  Jour,
  LaNatBud   : string;
  DEBIT      : Double;
  CREDIT     : Double;
  i,j        : Integer;
begin
  Fillchar(Tot1,SizeOf(Tot1),#0);
  j := col - 1;
  for i := 0 to T.Detail.Count-1 do begin
    DateCompta := T.Detail[i].GetDateTime('BE_DATECOMPTABLE');
    LaNatBud   := T.Detail[i].GetString('BE_NATUREBUD');
    DEBIT      := T.Detail[i].GetDouble('DEBIT');
    CREDIT     := T.Detail[i].GetDouble('CREDIT');
    Jour       := T.Detail[i].GetString('BE_BUDJAL');
    if (DateCOMPTA >= DateDeb) and (DateCompta <= DateFin) and (Jour = Journal) then
        CalculBud(j, Tot1, LaNatBud, DEBIT, CREDIT, Decimale);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.DoInsert(szCompte, szLibelleCompte, szRubCompte, szSection, szLibelleSection,
                                      szRubSection, szSectionTrie: string; iTri : Integer; TobFils: Tob);
{---------------------------------------------------------------------------------------}
var
  T2,
  T3          : Tob;
  j,
  Decimale    : Integer;
  Realise     : TabTot12;
  TotBud      : TMontTotal;
  Journal,
  szSQL,
  Resol,
  Compte1,
  Compte2,
  LExo        : String;
  TD          : TabloExt;
  Taux,
  dbl         : Double;
  TabBudget_l : TBudgets;  { Liste des budgets sélectionnés avec leurs périodes }
  DateDebRea  : TdateTime; { Date début période du réalisé }
  DateFinRea  : TdateTime; { Date fin période du réalisé }
  ok_creat    : boolean;

begin
  if MonoExo then
    LExo := QUELEXODTBud(StrToDate(FDateCompta1))
  else
    // LExo à vide si multi-exercices
    LExo := '';

  T3 := nil; {Avertissement !!}
    
  ok_creat := false;
  Resol := GetControlText('FRESOL');
  Taux := StrToFloat(GetControlText('FTAUX'));
  Decimale := 0;
  if Resol='C' then Decimale := V_PGI.OkDecV;

  TabBudget_l := TabBudget;
  // Boucle sur tous les journaux budgétaires sélectionnés (limités à 8)
  for j := 1 to 8 do begin
    Journal := ReadtokenSt(TabBudget_l[j].code);
    if Journal <> '' then begin
      // Requête des mouvements par journal
      szSQL := GetSQLMouvement(szCompte, szSection, Journal);
      PasserQry2Tob(T2,'ZC',szSql);

      if PasDeRubrique then begin
        Compte1 := szAxeJal + szSection;
        Compte2 := szCompte;
        end
      else begin
        Compte1 := 'S/G' + Journal + szRubSection + ':' + szRubCompte;
        Compte2 := '';
        end;

      Fillchar(Realise, SizeOf(Realise), #0);
      Fillchar(TotBud, SizeOf(TotBud), #0);
      if PasDeRubrique then begin
        if (szSectionTrie <> '') then
          // renvoie le budget du journal dans "TotBud" pour la période sélectionnée
          InitCalcBud(T2, TotBud, GetControlText('FRESOL'), StrToFloat(GetControlText('FTAUX')),j,Journal)
        else begin
          // renvoie le budget du journal dans "dbl" pour la période sélectionnée
          dbl := GetCumul('BUDGET:A/G',Compte1,Compte2,'SAN', GetControlText('FETAB'), Journal, LExo, DateDeb, DateFin, False, True, Nil, TD, False);
          TD[4] := 0;
          TD[5] := 0;
          if dbl > 0 then TD[5] := dbl
          else TD[4] := Abs(dbl);
          end;
        end
      else
        // Si "rubrique" coché
        GetCumul('RUBBUD',Compte1,Compte2,'SAN', GetControlText('FETAB'), 'EUR', LExo, DateDeb, DateFin, False, True, Nil, TD, False);

      if (szSectionTrie = '') then begin
        TotBud[0][j - 1].TotDebit := TD[5];
        TotBud[0][j - 1].TotCredit := TD[4];
        end;

      if (TobFils = nil) and (ok_creat = false) then begin
        T3 := Tob.Create('CEDTBALBUD', nil, -1);
        ok_creat := true;
        T3.InitValeurs(False);
        T3.PutValue('CEB_UTILISATEUR', V_PGI.User);
        T3.PutValue('CEB_BUDGENE', szCompte);
        T3.PutValue('CEB_LIBELLE', szLibelleCompte);
        T3.PutValue('CEB_BUDSECT', szSection);
        T3.PutValue('CEB_LIBSECT', szLibelleSection);
        T3.PutValue('CEB_TRI', iTri);
        T3.PutValue('CEB_PRINTCPTE', '2');
        // Plans de ruptures
        if (szSectionTrie <> '') then begin
          T3.PutValue('CEB_BUDTRI', szSectionTrie);
          T3.PutValue('CEB_LIBTRI', ListeCodesRupture.Items[ListeCodesRupture.Values.IndexOf(szSectionTrie)]);
          if bRecap then T3.PutValue('CEB_BUDSECT', szSectionTrie)
                    else T3.PutValue('CEB_BUDSECT', '');
          T3.PutValue('CEB_BUDGENE', '');
          T3.PutValue('CEB_PRINTCPTE', '0');
          end;
        end;
      //else T3 := TobFils;

      if ok_creat and Assigned(T3) then begin
        // Reévaluation des montants budgétés en fonction de la résolution et du taux de reévaluation
        Reevaluation(TotBud[0][j-1].TotDebit,TotBud[0][j-1].TotCredit, Resol, Taux);

        // Mise à jour des "Montants du Budget"
        dbl := T3.GetValue('CEB_DEBIT'+IntToStr(j));
        T3.PutValue('CEB_DEBIT'+IntToStr(j), dbl + TotBud[0][j-1].TotDebit);
        dbl := T3.GetValue('CEB_CREDIT'+IntToStr(j));
        T3.PutValue('CEB_CREDIT'+IntToStr(j), dbl + TotBud[0][j-1].TotCredit);

        if (szSectionTrie = '') then begin
          TotBud[0][12].TotDebit := Arrondi(TotBud[0][12].TotDebit + TotBud[0][j-1].TotDebit, Decimale);
          TotBud[0][12].TotCredit := Arrondi(TotBud[0][12].TotCredit + TotBud[0][j-1].TotCredit, Decimale);
          end;

        // Mise à jour des "Totaux Budgets"
        dbl := T3.GetValue('CEB_TOTDEBIT');
        T3.PutValue('CEB_TOTDEBIT', dbl + TotBud[0][12].TotDebit);
        dbl := T3.GetValue('CEB_TOTCREDIT');
        T3.PutValue('CEB_TOTCREDIT', dbl + TotBud[0][12].TotCredit);

        if GetCheckBoxState('FREALISE') = cbChecked then begin
          // Option "Montants du réalisé" coché
          // Remarque : on affiche le réalisé du journal en tenant compte des dates de début et fin du budget;
          // ces deux dates doivent être comprises dans la fourchette de dates sélectionnée à l'écran
          if TabBudget_l[j].DateDebBud < DateDeb then
            DateDebRea := DateDeb
          else
            DateDebRea := TabBudget_l[j].DateDebBud;
          if TabBudget_l[j].DateFinBud > DateFin then
            DateFinRea := DateFin
          else
            DateFinRea := TabBudget_l[j].DateFinBud;

          if PasDeRubrique then begin
            // LExo à vide si multi-exercices
            dbl := GetCumul('BUDGET:A/GREA',Compte1,Compte2,'SAN', GetControlText('FETAB'), Journal, LExo, DateDebRea, DateFinRea, False, False, Nil, TD, False);
            TD[4] := 0;
            TD[5] := 0;
            if dbl > 0 then TD[5] := dbl
            else TD[4] := Abs(dbl);
            end
          else
            GetCumul('RUBREA',Compte1,Compte2,'SAN', GetControlText('FETAB'), 'EUR', LExo, DateDebRea, DateFinRea, False, False, Nil, TD, False);

          realise[j - 1].TotDebit := TD[5];
          realise[j - 1].TotCredit := TD[4];

          // Mise à jour des "Montants réalisés"
          dbl := T3.GetValue('CEB_DEBITREA'+IntToStr(j));
          T3.PutValue('CEB_DEBITREA'+IntToStr(j), dbl + realise[j-1].TotDebit);
          dbl := T3.GetValue('CEB_CREDITREA'+IntToStr(j));
          T3.PutValue('CEB_CREDITREA'+IntToStr(j), dbl + realise[j-1].TotCredit);

          Realise[12].TotDebit := Arrondi(Realise[12].TotDebit + Realise[j-1].TotDebit, Decimale);
          Realise[12].TotCredit := Arrondi(Realise[12].TotCredit + Realise[j-1].TotCredit, Decimale);

          // Mise à jour des "Totaux réalisés"
          dbl := T3.GetValue('CEB_TOTDEBITREA');
          T3.PutValue('CEB_TOTDEBITREA', dbl + Realise[12].TotDebit);
          dbl := T3.GetValue('CEB_TOTCREDITREA');
          T3.PutValue('CEB_TOTCREDITREA', dbl + Realise[12].TotCredit);
          end; { réalisé }
        end;
      end; { Journal <> '' }
    end; {  for j := 1 to 8 do }
  T2.Free;
  if (TobFils = nil) and Assigned(T3) then begin
    T3.InsertDB(nil);
    T3.Free;
    end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPMULTIBUD_QR1.GetMainRequete : string;
{---------------------------------------------------------------------------------------}
var
  sz,
  szCpte1,
  szCpte2 : string;
begin
  if Rupture in ['T', 'A']
    then Result := 'SELECT BS_BUDSECT, ' + OrderLibre(GetControlText('FLIBTRIPAR')) + ' BS_RUB, BS_SECTIONTRIE FROM BUDSECT S WHERE '
    else Result := 'SELECT * FROM BUDSECT S WHERE ';
  Result := Result + ExistBud(AxeToFbBud(szAxeJal), GetControlText('FSELECTCPTE'), FCategorie.Value, szAxeJal, True);
  Result := Result + ' AND BS_CATEGORIE ="' + FCategorie.value + '" AND BS_AXE = "' + szAxeJal + '" ';

  szCpte1 := GetControlText('FSECTION1');
  szCpte2 := GetControlText('FSECTION2');
  if szCpte1 <> '' then Result := Result + ' AND BS_BUDSECT >= "' + szCpte1 + '" ';
  if szCpte2 <> '' then Result := Result + ' AND BS_BUDSECT <= "' + szCpte2 + '" ';

  // SUR plans de ruptures
  if Rupture in ['R', 'U'] then begin
    szCpte1 := THValComboBox(GetControl('FCODERUPT1', True)).Text;
    szCpte2 := THValComboBox(GetControl('FCODERUPT2', True)).Text;
    if szCpte1 <> '' then Result := Result + ' AND BS_SECTIONTRIE >= "' + szCpte1 + '" ';
    if szCpte2 <> '' then Result := Result + ' AND BS_SECTIONTRIE <= "' + szCpte2 + '" ';
    end;

  if (FGROUPCHOIXRUPT.ItemIndex > 0) then begin // AVEC ou SUR
    case FGROUPQUELLERUPTURE.ItemIndex  of
      0 : Result := Result + ' ORDER BY BS_AXE, BS_SECTIONTRIE'; {Sur Plan de rupture}

      1 : begin // Sur tables libres
           szCpte1 := GetControlText('FLIBRE1');
           szCpte2 := GetControlText('FLIBRE2');
           sz := WhereLibre(szCpte1, szCpte2, AxeToFbBud(szAxeJal), (GetCheckBoxState('FONLYCPTASSOCIE') = cbChecked));
           if sz <> '' then Result := Result + sz;
           Result := Result + ' ORDER BY ' + OrderLibre(GetControlText('FLIBTRIPAR'))+'BS_BUDSECT ';
           end;
      end;
    end
  else
    Result := Result + ' ORDER BY BS_AXE, BS_BUDSECT';
end;


{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.SetSecdRequete;
{---------------------------------------------------------------------------------------}
var
  szCpte1 : string;
  szCpte2 : string;
begin
  if (aTypeBalance = bal_GeSe) then begin
    szCpte1 := GetControlText('FSECTION1');
    szCpte2 := GetControlText('FSECTION2');
    FSecdRequete := 'SELECT BS_BUDSECT, BS_LIBELLE, BS_RUB FROM BUDSECT S WHERE ';
    FSecdRequete := FSecdRequete + ExistBud(AxeToFbBud(szAxeJal), GetControlText('FSELECTCPTE'),FCategorie.Value, szAxeJal, False);
    FSecdRequete := FSecdRequete + ' AND BS_CATEGORIE="' + FCategorie.value + '" AND BS_AXE = "' + szAxeJal + '" ';
    if szCpte1 <> '' then FSecdRequete := FSecdRequete + ' AND BS_BUDSECT >= "' + szCpte1 + '" ';
    if szCpte2 <> '' then FSecdRequete := FSecdRequete + ' AND BS_BUDSECT <= "' + szCpte2 + '" ';
    FSecdRequete := FSecdRequete + ' ORDER BY BS_AXE, BS_BUDSECT';
  end
  else begin
    szCpte1 := GetControlText('FCPTE1');
    szCpte2 := GetControlText('FCPTE2');
    FSecdRequete := 'SELECT BG_BUDGENE, BG_LIBELLE, BG_RUB FROM BUDGENE G WHERE ';
    FSecdRequete := FSecdRequete + ExistBud(fbBudgen, GetControlText('FSELECTCPTE'), FCategorie.Value, szAxeJal, False);
    if szCpte1 <> '' then FSecdRequete := FSecdRequete + ' AND BG_BUDGENE >= "' + szCpte1 + '" ';
    if szCpte2 <> '' then FSecdRequete := FSecdRequete + ' AND BG_BUDGENE <= "' + szCpte2 + '" ';
    FSecdRequete := FSecdRequete + ' ORDER BY BG_BUDGENE';
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.SetMvmtRequete;
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

  if (aTypeBalance = bal_SeGe) then
    FMvmtRequete := FMvmtRequete + 'WHERE BE_BUDSECT=:C1 AND BE_BUDGENE=:C2 '
  else
    FMvmtRequete := FMvmtRequete + 'WHERE BE_BUDGENE=:C1 AND BE_BUDSECT=:C2 ';

  FMvmtRequete := FMvmtRequete + 'AND BE_AXE = "' + szAxeJal + '" ';

  (***FMvmtRequete := FMvmtRequete + 'AND BE_BUDJAL="' + FCategorie.value + '"';***)
  FMvmtRequete := FMvmtRequete + 'AND BE_BUDJAL=:C3 ';

  if FStWhereNat <> '' then
    FMvmtRequete := FMvmtRequete + 'AND ' + FStWhereNat + ' ';

  FMvmtRequete := FMvmtRequete + ' AND BE_DATECOMPTABLE>="' + USDateTime(StrToDate(FDateCompta1)) + '"' +
                                 ' AND BE_DATECOMPTABLE<="' + USDateTime(StrToDate(FDateCompta2)) + '"' +
                                 ' AND BE_EXERCICE>="' + GetControlText('FEXERCICE') + '"' +
                                 ' AND BE_EXERCICE<="' + GetControlText('FEXERCICE_') + '"';
  szEtab := GetControlText('FETAB');
  if szEtab <> '' then FMvmtRequete := FMvmtRequete + ' AND BE_ETABLISSEMENT="' + szEtab + '" ';

  FValide := TCheckBox(GetControl('FVALIDE', True));
       if FValide.Checked             then FMvmtRequete := FMvmtRequete + ' AND BE_VALIDE = "X" '
  else if FValide.State = cbUnchecked then FMvmtRequete := FMvmtRequete + ' AND BE_VALIDE = "-" ';

  if (aTypeBalance = bal_SeGe) then begin
    if bRecap then FMvmtRequete := FMvmtRequete + ' ORDER BY BE_BUDGENE, BE_BUDSECT, BE_AXE, BE_EXERCICE, BE_DATECOMPTABLE, BE_BUDJAL, BE_NUMEROPIECE, BE_QUALIFPIECE '
              else FMvmtRequete := FMvmtRequete + ' ORDER BY BE_BUDSECT, BE_AXE, BE_BUDGENE, BE_EXERCICE, BE_DATECOMPTABLE, BE_BUDJAL, BE_NUMEROPIECE, BE_QUALIFPIECE ';
  end
  else begin
    if bRecap then FMvmtRequete := FMvmtRequete + ' ORDER BY BE_BUDSECT, BE_BUDGENE, BE_BUDJAL, BE_NATUREBUD, BE_EXERCICE, BE_DATECOMPTABLE, BE_AXE, BE_NUMEROPIECE, BE_QUALIFPIECE '
              else FMvmtRequete := FMvmtRequete + ' ORDER BY BE_BUDGENE, BE_BUDJAL, BE_NATUREBUD, BE_EXERCICE, BE_DATECOMPTABLE, BE_BUDSECT, BE_AXE, BE_NUMEROPIECE, BE_QUALIFPIECE ';
  end;
end;


{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.PrepCalcBud(SQL : string);
{---------------------------------------------------------------------------------------}
var
  TNat,
  TMain,
  TFinal,
  TFils,
  TFind,
  TLastFind,
  TFindNat,
  T,
  TCompte  : Tob;
  i, j, k,
  iTri     : Integer;
  szTemp,
  szSQL,
  szData,
  szOrder,
  szLib,
  szCompte : String;
  LTri     : TStringList;
  cPrint   : Char;
begin
  // Préparation de la requête des sections
  SetSecdRequete;

  // Préparation de la requête des mouvements
  SetMvmtRequete;

  //T contient la liste des sections budgétaires
  PasserQry2Tob(T,'X',Sql);

  // Pour le tri
  LTri := TStringList.Create;
  szOrder := OrderLibre(GetControlText('FLIBTRIPAR'), True);
  i := Pos(',', szOrder);
  while (i > 0) do begin
    System.Delete(szOrder,i,1);
    System.Insert(#10, szOrder, i);
    i := Pos(',', szOrder);
  end;
  LTri.Text := szOrder;

  // Rupture SUR tables libres ( A implémenter plus tard ? )
  if Rupture in ['T', 'A'] then begin
    // Pour le libellé des tables libres
    PasserQry2Tob(TNat,'Y','SELECT NT_NATURE, NT_LIBELLE FROM NATCPTE WHERE NT_TYPECPTE LIKE "D0%"');
    // Liste des données des tables libres à partir de la requête contenu dans SQL
    TMain := TOB.Create('~Main', nil, -1);
    for i := 0 to 9 do
      TOB.Create('TABLE' + IntToStr(i), TMain, -1);

    for i := 0 to T.detail.Count-1 do begin
      for j := 0 to 9 do begin
        if T.Detail[i].FieldExists('BS_TABLE' + IntToStr(j)) then begin
          // Recherche si la donnée est déjà stockée
          TFind := nil;
          szData := T.Detail[i].GetString('BS_TABLE' + IntToStr(j));

          // Récapitulatif par section
          if bRecap then begin
            szSQL := GetSQLSectGene(T.Detail[i].GetString('BS_BUDSECT'));
            PasserQry2Tob(TCompte,'Z',szSql);

            for k := 0 to TCompte.Detail.Count-1 do begin
              szCompte := TCompte.Detail[k].GetString('BG_BUDGENE');
              szLib := TCompte.Detail[k].GetString('BG_LIBELLE');
              if TMain.Detail[j].Detail.Count > 0 then
                TFind := TMain.Detail[j].FindFirst(['CEB_BUDGENE', 'CEB_BUDSECT'],[szData, szCompte], False);

              // Pas trouvé : Crée la donnée
              if (TFind=nil) and (szData<>'') then begin
                TFils :=  Tob.Create('CEDTBALBUD', TMain.Detail[j], -1);
                TFils.PutValue('CEB_UTILISATEUR', V_PGI.User);
                TFils.PutValue('CEB_TRI', LTri.IndexOf('BS_TABLE'+IntToStr(j)));
                TFils.PutValue('CEB_BUDSECT', szCompte);
                TFils.PutValue('CEB_LIBSECT', szLib);
                if k=0                      then cPrint := '0' else
                if k=TCompte.Detail.Count-1 then cPrint := '2' else
                                                 cPrint := ' ';
                TFils.PutValue('CEB_PRINTCPTE', cPrint);
                // Plan de rupture
                if (FGROUPQUELLERUPTURE.ItemIndex=0) then begin
                  szTemp := T.Detail[i].GetString('BS_SECTIONTRIE')+'x';
                  ListeCodesRupture.Value := szTemp;
                  TFils.PutValue('CEB_BUDGENE', szTemp);
                  TFils.PutValue('CEB_LIBELLE', ListeCodesRupture.Text);
                  TFils.PutValue('CEB_TRI', TFils.GetIndex);
                  end
                // Tables libres
                else begin
                  TFils.PutValue('CEB_BUDGENE', szData);
                  TFindNat := TNat.FindFirst(['NT_NATURE'], [szData], False);
                  if (TFindNat<>nil) then TFils.PutValue('CEB_LIBELLE',  TFindNat.GetString('NT_LIBELLE'));
                  end;
                TFind := TFils;
                end;
              end; { fin boucle }
              TCompte.Free;
            end { récapitulatif par section }
          else begin
            if TMain.Detail[j].Detail.Count > 0 then
              TFind := TMain.Detail[j].FindFirst(['CEB_BUDGENE'],[szData], False);

            // Pas trouvé : Crée la donnée
            if (TFind=nil) and (szData<>'') then begin
              TFils :=  Tob.Create('CEDTBALBUD', TMain.Detail[j], -1);
              TFils.PutValue('CEB_UTILISATEUR', V_PGI.User);
              TFils.PutValue('CEB_TRI', LTri.IndexOf('BS_TABLE'+IntToStr(j)));
              TFils.PutValue('CEB_PRINTCPTE', '2');
              // Plan de rupture
              if (FGROUPQUELLERUPTURE.ItemIndex=0) then begin
                szTemp := T.Detail[i].GetString('BS_SECTIONTRIE')+'x';
                ListeCodesRupture.Value := szTemp;
                TFils.PutValue('CEB_BUDGENE', szTemp);
                TFils.PutValue('CEB_LIBELLE', ListeCodesRupture.Text);
                TFils.PutValue('CEB_TRI', TFils.GetIndex);
                end
              // Tables libres
              else begin
                TFils.PutValue('CEB_BUDGENE', szData);
                TFindNat := TNat.FindFirst(['NT_NATURE'], [szData], False);
                if (TFindNat<>nil) then TFils.PutValue('CEB_LIBELLE',  TFindNat.GetString('NT_LIBELLE'));
                TFind := TFils;
                end;
              end;
            end;
          // Crée le compte uniquement pour les tables libres
          if (FGROUPQUELLERUPTURE.ItemIndex=1) then begin
            TFils := Tob.Create('~COMPTE', TFind, -1);
            TFils.AddChampSup('COMPTE', False);
            TFils.SetString('COMPTE', T.Detail[i].GetString('BS_BUDSECT'));
            end;
          end; { if T.Detail[i].FieldExists('BS_TABLE  }
        end; { for j := 0 to 9 do }
      end;  { for i := 0 to T.detail.Count-1 }

    // Liste des comptes
    for i := 0 to T.detail.Count-1 do begin
      for j := 0 to 9 do begin
        if T.Detail[i].FieldExists('BS_TABLE'+IntToStr(j)) then begin
          if (FGROUPQUELLERUPTURE.ItemIndex=0) then szData := T.Detail[i].GetString('BS_SECTIONTRIE')+'x'
                                               else szData := T.Detail[i].GetString('BS_TABLE'+IntToStr(j));

          // Récapitulatif par section
          if bRecap then begin
            szSQL := GetSQLSectGene(T.Detail[i].GetString('BS_BUDSECT'));
            PasserQry2Tob(TCompte,'ZA',szSql);
            for k := 0 to TCompte.Detail.Count-1 do begin
              szCompte := TCompte.Detail[k].GetString('BG_BUDGENE');
              TFind := TMain.Detail[j].FindFirst(['CEB_BUDGENE', 'CEB_BUDSECT'],[szData, szCompte], False);

              if not(TFind=nil) then begin
                DoInsert(szCompte,
                         '',
                         '',
                         T.Detail[i].GetString('BS_BUDSECT'),
                         '',
                         T.Detail[i].GetString('BS_RUB'),
                         '',
                         0,
                         TFind);
                end;
              end;
              TCompte.Free;
            end  { if bRecap then }
          else begin
            TFind := TMain.Detail[j].FindFirst(['CEB_BUDGENE'],[szData], False);
            if not(TFind=nil) then begin
              DoInsert('',
                       '',
                       '',
                       T.Detail[i].GetString('BS_BUDSECT'),
                       '',
                       T.Detail[i].GetString('BS_RUB'),
                       '',
                       0,
                       TFind);
              end;
            end;

          // Plan de ruptures : Pour le total
          if (FGROUPQUELLERUPTURE.ItemIndex=0) and bRecap then begin
            TFind := TMain.Detail[j].FindFirst(['CEB_BUDGENE'],[szData], False);
            TLastFind := TFind;
            while TFind<>nil do begin
              TFind := TMain.Detail[j].FindNext(['CEB_BUDGENE'],[szData], False);
              if TFind<>nil then TLastFind := TFind;
              end;

            if TLastFind<>nil then begin
              TFils := Tob.Create('CEDTBALBUD', TMain.Detail[j], TLastFind.GetIndex+1);
              TFils.InitValeurs;
              szTemp := T.Detail[i].GetString('BS_SECTIONTRIE')+'x';
              TFils.PutValue('CEB_BUDTRI', szTemp);
              TFils.PutValue('CEB_LIBTRI', ListeCodesRupture.Items[ListeCodesRupture.Values.IndexOf(szTemp)]);
              TFils.PutValue('CEB_TRI', TFils.GetIndex);

              DoInsert('',
                       '',
                       '',
                       T.Detail[i].GetString('BS_BUDSECT'),
                       '',
                       T.Detail[i].GetString('BS_RUB'),
                       szTemp,
                       0,
                       TFils);
              end;
            end;

          end; { if T.Detail[i].FieldExists('BS_TABLE' }
        end;  { for j := 0 to 9 do  }
      end; { for i := 0 to T.detail.Count-1 }

    // Bascule les sections sur le compte général
    for i := 0 to 9 do begin
      for j := 0 to TMain.Detail[i].Detail.Count-1 do begin
        szTemp := TMain.Detail[i].Detail[j].GetString('CEB_BUDGENE');
        TMain.Detail[i].Detail[j].PutValue('CEB_BUDGENE', TMain.Detail[i].Detail[j].GetString('CEB_BUDSECT'));
        TMain.Detail[i].Detail[j].PutValue('CEB_BUDSECT', szTemp);
        szTemp := TMain.Detail[i].Detail[j].GetString('CEB_LIBELLE');
        TMain.Detail[i].Detail[j].PutValue('CEB_LIBELLE', TMain.Detail[i].Detail[j].GetString('CEB_LIBSECT'));
        TMain.Detail[i].Detail[j].PutValue('CEB_LIBSECT', szTemp);
        end;
      end;

    TNat.Free;

    // Tri les Tob avant insertion
    TFinal := TOB.Create('~CEDTBALBUD', nil, -1);
    for j := 0 to 9 do begin
      while TMain.Detail[j].Detail.Count>0 do begin
        // Supprime les enfants (liste des comptes)
        TMain.Detail[j].Detail[0].ClearDetail;

        // Bascule sur le nouveau parent
        TMain.Detail[j].Detail[0].ChangeParent(TFinal, -1);
        end;
      end;
    TMain.Free;

    TFinal.Detail.Sort('CEB_BUDSECT');
    TFinal.InsertDB(nil, True);
    TFinal.Free;
    end

  else begin
    // AVEC ou SANS Rupture
    iTri := 0;
    // Boucle sur les sections budgétaires
    for i := 0 to T.Detail.Count-1 do begin
      szSQL := GetSQLSectGene(T.Detail[i].GetString('BS_BUDSECT'));
      //Tcompte contient la liste des comptes budgétaires
      PasserQry2Tob(TCompte,'ZB',szSql);
      for j := 0 to TCompte.Detail.Count-1 do begin
        // Boucle sur les comptes budgétaires de la section
        DoInsert(TCompte.Detail[j].GetString('BG_BUDGENE'),
                 TCompte.Detail[j].GetString('BG_LIBELLE'),
                 TCompte.Detail[j].GetString('BG_RUB'),
                 T.Detail[i].GetString('BS_BUDSECT'),
                 T.Detail[i].GetString('BS_LIBELLE'),
                 T.Detail[i].GetString('BS_RUB'),
                 '',
                 iTri,
                 nil);
        inc(iTri);
        end;

      // Plan de ruptures : Pour le total
      if Rupture in ['R', 'U'] then begin
        DoInsert('',
                 TCompte.Detail[TCompte.Detail.Count-1].GetString('BG_LIBELLE'),
                 TCompte.Detail[TCompte.Detail.Count-1].GetString('BG_RUB'),
                 T.Detail[i].GetString('BS_BUDSECT'),
                 T.Detail[i].GetString('BS_LIBELLE'),
                 T.Detail[i].GetString('BS_RUB'),
                 T.Detail[i].GetString('BS_SECTIONTRIE') + 'x',
                 iTri,
                 nil);
        inc(iTri);
        end;
      TCompte.Free;
      end;
    end;
    LTri.Free;
    T.Free;
  // Plans de ruptures
  if Rupture in ['R', 'U'] then SetControlText('PLAN', 'X')
                           else SetControlText('PLAN', '-');
end;

procedure TOF_CPMULTIBUD_QR1.RecupPeriodesJournal( Journal : string;var DtDebBud : TDateTime;var DtFinBud : TDateTime);

Var
  QJal   : TQuery;
  sWhere : String;
begin
  sWhere := ' WHERE';
  if (FCategorie.value <> '') then
    sWhere := sWhere + ' BJ_CATEGORIE="' + FCategorie.value + '"';
  sWhere := sWhere + ' AND BJ_BUDJAL="' + Journal + '"';

  QJal := OpenSQL('Select BJ_PERDEB, BJ_PERFIN From BUDJAL ' + sWhere ,True) ;
  if Not QJal.EOF then begin
    DtDebBud := QJal.FindField('BJ_PERDEB').AsDateTime;
    DtFinBud := QJal.FindField('BJ_PERFIN').AsDateTime;
    end;
Ferme(QJal) ;
end;

procedure TOF_CPMULTIBUD_QR1.OnLoad ;
var
  szDec            : string;
  StTri,
  StCod1,
  StCod2           : string;
  NbBud,
  NbEx             : integer;
begin
  //Inherited ;

  {Initialisation des contrôles}
  InitCompte;

  case FResol.ItemIndex of
    // Formatage des décimales est fait dans le générateur d'états
    //0: SetControlText('FORMAT', '#,##0.00;; ;'); // Avec décimales
    1: SetControlText('FORMAT', '#,##0;; ;');    // Sans décimales
    2: SetControlText('FORMAT', '#,##0;; ;');    // Kilo
    3: SetControlText('FORMAT', '#,##0;; ;');    // Mega
  end;
  szDec := GetControlText('FRESOL');

  FStWherejou := '' ;
  TraductionTHMultiValComboBox( FJournal,FStWhereJou, FStLibJou, 'BJ_BUDJAL',False);
  SetControlText( 'Journal', FStLibJou) ;

  // Recherche si on est mono-exercice ou multi-exercices
  NbEx := StrToInt(QUELEXODTBud(StrToDate(FDateCompta2))) - StrToInt(QUELEXODTBud(StrToDate(FDateCompta1))) + 1;
  DateDeb := StrToDate(FDateCompta1);
  DateFin := FinDeMois(StrToDate(FDateCompta2));
  MonoExo := (NbEx = 1);

  NbBud := CalculNbBud;

  { On vide la table temporaire avant de tester les erreurs }
  ExecuteSQL('DELETE FROM CEDTBALBUD WHERE CEB_UTILISATEUR = "' + V_PGI.User + '"');

  // Vérification que l'on n'a pas sélectionné plus de 8 budgets
  if NbBud > 8 then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Vous ne pouvez pas sélectionner plus de 8 budgets');
    exit;
    end;

  if NbBud = 0 then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Vous devez sélectionnez au moins un budget');
    exit;
    end;

  { Ruptures sur tables libres ( A implémenter plus tard ? )}
  if (FGROUPCHOIXRUPT.ItemIndex > 0) and (Trim(GetControlText('FLIBTRIPAR'))= '') then begin
    StTri := GetControlText('FLIBTRIPAR');
    if GetCheckBoxState('FONLYCPTASSOCIE') = cbChecked then
      ChoixTableLibreSurInfo(AxeToFbBud(szAxeJal), StTri, StCod1, StCod2)
    else
      ChoixTableLibreInfo(AxeToFbBud(szAxeJal), StTri, StCod1, StCod2);
    SetControlText('FLIBTRIPAR', UpdateRupture(StCod1,StCod2));
    end;

  { Ruptures sur Plans de ruptures }
  if Rupture in ['R', 'U']  then begin
    case SectionRetrieBud(FPLANRUPTURES.Value, szAxeJal, nil) of
      srOk              : ;
      srNonStruct       : begin
                            LastError := 1;
                            LastErrorMsg := TraduireMemoire('L''axe que vous avez choisi n''est pas structuré pour les ruptures.');
                            end;
      srPasEnchainement : begin
                            LastError := 1;
                            LastErrorMsg := TraduireMemoire('Il n''existe a pas d''enchaînement pour ce plan de ruptures.');
                            end;
      end;
    if (LastErrorMsg <> '') then begin
      FGROUPCHOIXRUPT.ItemIndex := 0;
      Exit;
      end;
    end;

  { Gestion du nombre de décimales }
  if szDec = 'C' then Decimale := V_PGI.OkDecV
                 else Decimale := 0;

  PrepCalcBud(GetMainRequete);
end ;

function TOF_CPMULTIBUD_QR1.TraiteListeJournal : boolean;
var
  Q          : TQuery;
  Titre,
  Mess,
  sLejournal,
  sRequete,
  sWhere,
  sFrom,
  LibAbreg   : String;
begin
  result := false;
  THMultiValComboBox(GetControl('FJournal')).Items.clear;
  THMultiValComboBox(GetControl('FJournal')).Values.clear;
  sRequete := 'SELECT BJ_BUDJAL,BJ_LIBELLE,BJ_AXE';
  sFrom := ' FROM BUDJAL';
  sWhere := ' WHERE';
  if (FCategorie.value <> '') then
    sWhere := sWhere + ' BJ_CATEGORIE="' + FCategorie.value + '"';
  sWhere := sWhere + ' ORDER BY BJ_BUDJAL';
  Q := OpenSQL(sRequete + sFrom + sWhere,true);

  if Q.eof then begin
    THValComboBox(GetControl('FCATEGORIE', True)).ItemIndex := 0;
    Titre := TraduireMemoire('Choix des journaux');
    Mess := TraduireMemoire('Aucun élément disponible');
    HShowMessage ( '0;' + Titre + ';' + Mess + ';W;O;O;O','', '');
    szAxeJal :='';
    FCategorieChange(nil);
    end
  else begin
    result := true;
    while not Q.EOF do
      begin
      LibAbreg := Q.FindField('BJ_LIBELLE').AsString;
      sLejournal := Q.FindField('BJ_BUDJAL').AsString;
      szAxeJal := Q.FindField('BJ_AXE').AsString;
      THMultiValComboBox(GetControl('FJournal')).Items.Add(LibAbreg);
      THMultiValComboBox(GetControl('FJournal')).Values.add(sLejournal);
      Q.Next;
      end;
    { Sélection par défaut de tous les budgets }
    (*THMultiValComboBox(GetControl('FJournal')).Value := '';
    THMultiValComboBox(GetControl('FJournal')).SelectAll;*)
    end;
  Ferme(Q);
end;

procedure TOF_CPMULTIBUD_QR1.OnArgument (S : String ) ;

begin
  { Attention ne pas faire hériter de CPBalanceBudget : dans cet état "FJournal" est une THMultiValComboBox }
  //inherited;
  aTypeBalance := bal_SeGe;
  FJournal   := THMultiValComboBox( GetControl('FJournal',True));
  FCategorie := THValComboBox( GetControl('FCATEGORIE', True));
  FEXERCICE  := THValComboBox( GetControl('FEXERCICE', True));
  FEXERCICE_ := THValComboBox( GetControl('FEXERCICE_', True));
  FPeriode1  := THValComboBox( GetControl('FPeriode1', True));
  FPeriode2  := THValComboBox( GetControl('FPeriode2', True));
  FSELECTCPTE := THValComboBox(GetControl('FSELECTCPTE', True));
  FRealise := TCheckBox(GetControl('FREALISE', True));
  FAvecRub := TCheckBox(GetControl('FAvecRub', False));
  FAnaGene := TCheckBox(GetControl('FAnaGene', False));
  FPLANRUPTURES := THValComboBox(GetControl('FPLANRUPTURES',True));
  FCodeRupt1 := THValComboBox(GetControl('FCodeRupt1', True));
  FCodeRupt2 := THValComboBox(GetControl('FCodeRupt2', True));
  FGroupChoixRupt := THRadioGroup(GetControl('FGROUPCHOIXRUPT',True));
  FGroupQuelleRupture := THRadioGroup(GetControl('FGROUPQUELLERUPTURE',True));

  FResol := THValComboBox(GetControl('FRESOL', True));
  AffMontants   := THValComboBox(GetControl('AFFMONTANTS',True)) ;
  FCodeRupt1 := THValComboBox(GetControl('FCodeRupt1', True));
  FCodeRupt2 := THValComboBox(GetControl('FCodeRupt2', True));

  ListeCodesRupture := THValComboBox.Create(Ecran);
  ListeCodesRupture.Parent := Ecran;
  ListeCodesRupture.Visible := False;

  // Vérification de l'existence des composants
  NotifyErrorComponent(FJournal,'FJournal');
  NotifyErrorComponent(FCATEGORIE,'FCATEGORIE');
  NotifyErrorComponent(FEXERCICE,'FEXERCICE');
  NotifyErrorComponent(FEXERCICE_,'FEXERCICE_');
  NotifyErrorComponent(FPeriode1,'FPeriode1');
  NotifyErrorComponent(FPeriode2,'FPeriode2');
  NotifyErrorComponent(FRealise,'FREALISE');
  NotifyErrorComponent(FAvecRub,'FAvecRub');
  NotifyErrorComponent(FAnaGene,'FAnaGene');
  NotifyErrorComponent(FGroupChoixRupt,'FGROUPCHOIXRUPT');
  NotifyErrorComponent(FGroupQuelleRupture,'FGROUPQUELLERUPTURE');
  NotifyErrorComponent(FPLANRUPTURES,'FPLANRUPTURES');
  NotifyErrorComponent(FCodeRupt1,'FCodeRupt1');
  NotifyErrorComponent(FCodeRupt2,'FCodeRupt2');
  NotifyErrorComponent(Fresol,'Fresol');
  NotifyErrorComponent(AffMontants,'AFFMONTANTS');

  FExercice.OnChange := FExerciceChange;
  FExercice_.OnChange := FExercice_Change;
  FPeriode1.OnChange := FPeriode1Change;
  FPeriode2.OnChange := FPeriode2Change;
  FCategorie.OnChange := FCategorieChange;
  FPLANRUPTURES.OnChange := FPlanRupturesChange;
  FGroupChoixRupt.OnClick := FSansRuptClick;
  FGroupQuelleRupture.OnClick := FRupturesClick;
  FAvecrub.OnClick :=  FAvecRubClick;
  FRealise.OnClick := FRealiseClick;
  Fresol.OnChange := FresolChange;
  AffMontants.OnChange := AffMontantsChange;

  //DoAssign('FLIBRE1', 'OnDblClick', FLibreDblClick);
  //DoAssign('FLIBRE2', 'OnDblClick', FLibreDblClick);
  //DoAssign('FLIBTRIPAR', 'OnDblClick', FLibreDblClick);

  // Initialisation des contrôles
  THValComboBox(GetControl('FCATEGORIE', True)).ItemIndex := 0;
  THValComboBox(GetControl('FJournal', True)).ItemIndex := 0;
  THValComboBox(GetControl('FResol', True)).ItemIndex := 1;
  SetControlText('IMPRESSION','ENTI');
  THValComboBox(GetControl('AffMontants', True)).ItemIndex := 0;
  SetControlText('AFFMONTANT','NORMAL');
  THValComboBox(GetControl('FETAB', True)).ItemIndex := 0;

  FSELECTCPTE.Value := 'TOO';

  FExercice.Value := VH^.Entree.Code;

  InitCompte;

  FresolChange(nil);
  AffMontantsChange(nil);
  FCategorieChange(nil);
  FRupturesClick(nil);
  FSansRuptClick(nil);
  FAvecRubClick(nil);
  SetControlVisible('FGROUPCHOIXRUPT', False);

  {JP 15/10/07 : FQ 16149 : gestion des réstrictions Etablissements}
  GereEtablissement;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  {JP 06/07/07 : FQ 20973 : On vide la table temporaire}
  ExecuteSQL('DELETE FROM CEDTBALBUD WHERE CEB_UTILISATEUR = "' + V_PGI.User + '"');

  ListeCodesRupture.Free;
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {JP 29/10/07 : FQ 20973 : On filtre la table temporaire}
  TFQRS1(Ecran).WhereSQL := ' CEB_UTILISATEUR = "' + V_PGI.User + '"'  ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.InitCompte;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT MIN(BG_BUDGENE), MAX(BG_BUDGENE) FROM BUDGENE WHERE BG_FERME="-"', True);
  if not Q.EOF then begin
    if GetControlText('FCPTE1') = '' then SetControlText('FCPTE1', Q.Fields[0].AsString);
    if GetControlText('FCPTE2') = '' then SetControlText('FCPTE2', Q.Fields[1].AsString);
  end;
  Ferme(Q);

  Q := OpenSQL('SELECT MIN(BS_BUDSECT), MAX(BS_BUDSECT) FROM BUDSECT WHERE BS_FERME = "-" AND BS_AXE = "' +
                szAxeJal + '"', True);
  if not Q.EOF then begin
    if GetControlText('FSECTION1') = '' then SetControlText('FSECTION1', Q.Fields[0].AsString);
    if GetControlText('FSECTION2') = '' then SetControlText('FSECTION2', Q.Fields[1].AsString);
    end;
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.DebToDate;
{---------------------------------------------------------------------------------------}
begin
  FDateCompta1 := FPeriode1.Value;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.FinToDate;
{---------------------------------------------------------------------------------------}
begin
  FDateCompta2 := FPeriode2.Value;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.FRealiseClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if FRealise.Checked then TFQRS1(Ecran).FEtat.Value := 'MBR'
                      else TFQRS1(Ecran).FEtat.Value := 'MBJ';
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.FAvecRubClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  PasDeRubrique := not FAvecRub.Checked;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPMULTIBUD_QR1.DoAssign(Control, Evt : string; Event: TNotifyEvent) : TControl;
{---------------------------------------------------------------------------------------}
var
  szEvt : string;
begin
  szEvt  := UpperCase(Evt);
  Result := GetControl(Control, True);

  if Result = nil then Exit;

  if szEvt = 'ONCHANGE' then TComboBox(Result).OnChange := Event
  else
    if szEvt = 'ONDBLCLICK' then THEdit(Result).OnDblClick := Event
    else
      if szEvt = 'ONCLICK' then THRadioGroup(Result).OnClick := Event;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.FCategorieChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if FCategorie.Value <> '' then begin
    if TraiteListeJournal then ;
    end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.FExerciceChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if FExercice.Value > FExercice_.Value then FExercice_.Value := FExercice.Value;
  ListePeriode(FExercice.Value, FPeriode1.Items, FPeriode1.Values, True);
  FPeriode1.ItemIndex := 0;
  DebToDate;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.FExercice_Change(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if FExercice_.Value < FExercice.Value then FExercice.Value := FExercice_.Value;
  ListePeriode(FExercice_.Value, FPeriode2.Items, FPeriode2.Values, False);
  FPeriode2.ItemIndex := FPeriode2.Items.Count-1;
  FinToDate;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.FPeriode1Change(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if (FPeriode1.Value <> '') and (FPeriode2.Value <> '') then begin
    if StrToDate(FPeriode1.Value) > StrToDate(FPeriode2.Value) then
      FPeriode2.Value := DateToStr(FinDeMois(StrToDate(FPeriode1.Value)));
  end;

  FDateCompta1 := FPeriode1.Value;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.FPeriode2Change(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if (FPeriode1.Value <> '') and (FPeriode2.Value <> '') then begin
    if StrToDate(FPeriode2.Value) < StrToDate(FPeriode1.Value) then
      FPeriode1.Value := DateToStr(DebutDeMois(StrToDate(FPeriode2.Value)));
  end;

  FDateCompta2 := FPeriode2.Value;
end;

{---------------------------------------------------------------------------------------}
PROCEDURE TOF_CPMULTIBUD_QR1.FresolChange (Sender : TObject);
BEGIN
case Fresol.ItemIndex of
  0 : SetControlText('IMPRESSION','DECI');
  1 : SetControlText('IMPRESSION','ENTI');
  2 : SetControlText('IMPRESSION','KILO');
  3 : SetControlText('IMPRESSION','MEGA');
  end;
END;
{---------------------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.AffmontantsChange (Sender : TObject);
{---------------------------------------------------------------------------------------}

begin
  case Affmontants.ItemIndex of
    0 : SetControlText('AFFMONTANT','NORMAL');
    1 : SetControlText('AFFMONTANT','SIGNE');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.FPlanRupturesChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Q1  : TQuery;
  St  : String;
begin
  if szAxeJal = '' then Exit;
  if Length(szAxeJal)< 2 then Exit;

  // Rempli la Fourchette des codes ruptures en 'Sur Rupture'
  if (FGROUPQUELLERUPTURE.ItemIndex = 0) then begin
    St := '';
    FCodeRupt1.Clear;
    FCodeRupt2.clear ;
    if ListeCodesRupture <> nil then ListeCodesRupture.Clear;

    { b md FQ 18077 FQ 18077 }
    // Il existe au moins un plan de rupture : on se positionne dessus sinon on sort
    (*if FPlanRuptures.Items.Count > 0 then FPlanRuptures.ItemIndex := 0
                                     else exit;*)
    if FPlanRuptures.Items.Count <= 0 then exit;                                 
    { e md }

    Q1 := OpenSQL('SELECT RU_CLASSE, RU_LIBELLECLASSE FROM RUPTURE WHERE RU_NATURERUPT="RU'+Copy(szAxeJal,2,1)+'" AND RU_PLANRUPT="'+FPlanRuptures.Value+'" ORDER BY RU_CLASSE ', True);

    while Not Q1.Eof do begin
      FCodeRupt1.Items.Add(Q1.Fields[0].AsString);
      FCodeRupt2.Items.Add(Q1.Fields[0].AsString);
      ListeCodesRupture.Values.Add(Q1.Fields[0].AsString);
      ListeCodesRupture.Items.Add(Q1.Fields[1].AsString);
      Q1.Next;
    end;
    FCodeRupt1.ItemIndex := 0;
    FCodeRupt2.ItemIndex := FCodeRupt2.Items.Count-1;
    Ferme(Q1);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.FSansRuptClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  SetControlVisible('SHTRUPTURE', (FGROUPCHOIXRUPT.ItemIndex > 0));
  if (FGROUPCHOIXRUPT.ItemIndex = 2) then SetControlText('TRI', 'X')
                                     else SetControlText('TRI', '');

  if (aTypeBalance = bal_SeGe) or (aTypeBalance = bal_GeSe) then
    SetControlEnabled('FAVECCPTSECOND', (FGROUPCHOIXRUPT.ItemIndex = 2));
  SetControlEnabled('FANAGENE',(FGROUPCHOIXRUPT.ItemIndex = 0));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTIBUD_QR1.FRupturesClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  MajInfoRupture;
  SetControlVisible('FGROUPRUPTURES',  (FGROUPQUELLERUPTURE.ItemIndex = 0));
  SetControlVisible('FONLYCPTASSOCIE', ((FGROUPQUELLERUPTURE.ItemIndex = 1) and (rupture in ['T','A'])));
  SetControlChecked('FONLYCPTASSOCIE', ((FGROUPQUELLERUPTURE.ItemIndex = 1) and (rupture in ['T','A'])));
  SetControlEnabled('FANAGENE',(FGROUPCHOIXRUPT.ItemIndex = 0));
end;

Initialization
  registerclasses ( [ TOF_CPMULTIBUD_QR1 ] ) ;
end.


