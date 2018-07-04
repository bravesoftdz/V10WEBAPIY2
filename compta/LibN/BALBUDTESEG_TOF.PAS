{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/01/2006
Modifié le ... : 25/01/2006
Description .. : Lek
Suite ........ : Je met en mode 2/3.
Suite ........ : Pour être compilé je modifie
Suite ........ : uTob en commun
Suite ........ : OpenSQL affecte à tQuery puis Tob
Mots clefs ... : 
*****************************************************************}
{ Unité : Source TOF de la FICHE : BALBUDTESEG
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
                 16/03/05    VL   Création de l'unité
 6.50.001.013    09/08/05    JP   Création de l'ancêtre TOF_CPBALANCEBUDGET
 7.00.001.00X    05/06/06    MD   FQ 18077 Correction plan de ruptures
--------------------------------------------------------------------------------------}

unit BALBUDTESEG_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL, eQRS1,
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_Main, QRS1,
  {$ENDIF}
  uTob, SysUtils, ComCtrls,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ELSE}
  TCalcCum,
  {$ENDIF MODENT1}
  HCtrls, HEnt1, CPBALANCEBUDGET_TOF, Ent1;

procedure BalBudteSecGen;

type

  TOF_BALBUDTESEG = class (TOF_CPBALANCEBUDGET)
    procedure OnLoad                ; override;
    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
  protected
    FPLANRUPTURES : THValComboBox;

    FCodeRupt1,
    FCodeRupt2 : THValComboBox;

    ListeCodesRupture : THValComboBox;

    procedure FPlanRupturesChange(Sender: TObject);
    procedure FRealiseClick   (Sender : TObject); override;
    { b md 05/06/06 FQ 18077 }
    procedure FRupturesClick   (Sender : TObject); override;
    { e md }
  public
    function  GetMainRequete : string; override;
    procedure MajInfoRupture; override;
    procedure PrepCalcBud(SQL : string); override;
    procedure DoInsert(szCompte, szLibelleCompte, szRubCompte, szSection, szLibelleSection, szRubSection, szSectionTrie : String; iTri : Integer; TobFils : Tob);
  end ;

implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcMetier,
  {$ENDIF MODENT1}
  UtilEDT,     {ExistBud, WhereLibre, OrderLibre}
  CALCOLE,     {GetCumul}
  TabLiEdt,    {ChoixTableLibreSur, ChoixTableLibre}
  TriTabLi,    {ChoixTriTableLibre}
  BUDSECT_TOM; {SectionRetrieBud}


{---------------------------------------------------------------------------------------}
procedure BalBudteSecGen;
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('CP', 'CPBALBUDTESEG', '', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_BALBUDTESEG.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  aTypeBalance := bal_SeGe;
  ListeCodesRupture := THValComboBox.Create(Ecran);
  ListeCodesRupture.Parent := Ecran;
  ListeCodesRupture.Visible := False;

  FPLANRUPTURES := THValComboBox(DoAssign('FPLANRUPTURES', 'OnChange', FPlanRupturesChange));
  FCodeRupt1 := THValComboBox(GetControl('FCodeRupt1', True));
  FCodeRupt2 := THValComboBox(GetControl('FCodeRupt2', True));

  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_BALBUDTESEG.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  {Ruptures sur Plans de ruptures}
  if Rupture in ['R', 'U']  then begin
    case SectionRetrieBud(FPLANRUPTURES.Value, szAxeJal, nil) of
      srOk              : ;
      srNonStruct       : begin LastError := 1; LastErrorMsg := 'L''axe que vous avez choisi n''est pas structuré pour les ruptures.'; end;
      srPasEnchainement : begin LastError := 1; LastErrorMsg := 'Il n''existe a pas d''enchaînement pour ce plan de ruptures.'; end;
    end;

    if (LastErrorMsg <> '') then begin
      FGROUPCHOIXRUPT.ItemIndex := 0;
      Exit;
    end;
  end;

  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_BALBUDTESEG.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  ListeCodesRupture.Free;
  inherited;
end;

{---------------------------------------------------------------------------------------}
function TOF_BALBUDTESEG.GetMainRequete : string;
{---------------------------------------------------------------------------------------}
var
  sz, szCpte1, szCpte2 : string;
begin
  if Rupture in ['T', 'A'] then Result := 'SELECT BS_BUDSECT, ' + OrderLibre(GetControlText('FLIBTRIPAR')) + ' BS_RUB, BS_SECTIONTRIE FROM BUDSECT S WHERE '
                           else Result := 'SELECT * FROM BUDSECT S WHERE ';
  Result := Result + ExistBud(AxeToFbBud(szAxeJal), GetControlText('FSELECTCPTE'), GetControlText('FJOURNAL'), szAxeJal, True);
  Result := Result + ' AND BS_AXE = "' + szAxeJal + '" ';

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
procedure TOF_BALBUDTESEG.MajInfoRupture;
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
{Lek 250106 Charge la Requette dans Tob}
var Q:TQuery;
begin
  Q:=OpenSql(Requette,true);
  T:=Tob.Create(NomTab,nil,-1);
  T.LoadDetailDB(NomTab,'','',Q,true);
  Ferme(Q);
end;

{CEB_PRINTCPTE = 0 : Imprime l'en-tête
 CEB_PRINTCPTE = 1 : Imprime le pied
 CEB_PRINTCPTE = 2 : Imprime les 2
{---------------------------------------------------------------------------------------}
procedure TOF_BALBUDTESEG.PrepCalcBud(SQL : string);
{---------------------------------------------------------------------------------------}
var
  TNat, TMain, TFinal, TFils, TFind, TLastFind, TFindNat, T, TCompte : Tob;
  i, j, k, iTri : Integer;
  szTemp, szSQL, szData, szOrder, szLib, szCompte : String;
  LTri : TStringList;
  cPrint : Char;
begin
  // Requête des sections
  SetSecdRequete;

  // Requête des mouvements
  SetMvmtRequete;

  // Liste des sections budgétaires
//  T := OpenSQL(SQL, True); Lek 250106
  PasserQry2Tob(T,'X',Sql);

  // Pour le tri
  LTri := TStringList.Create;
  szOrder := OrderLibre(GetControlText('FLIBTRIPAR'), True);
  i := Pos(',', szOrder);
  while (i>0) do begin
    System.Delete(szOrder,i,1);
    System.Insert(#10, szOrder, i);
    i := Pos(',', szOrder);
  end;
  LTri.Text := szOrder;

  // Rupture SUR tables libres
  if Rupture in ['T', 'A'] then begin
    // Pour le libellé des tables libres
//    TNat := OpenSQL('SELECT NT_NATURE, NT_LIBELLE FROM NATCPTE WHERE NT_TYPECPTE LIKE "D0%"', True); Lek 250106
    PasserQry2Tob(TNat,'Y','SELECT NT_NATURE, NT_LIBELLE FROM NATCPTE WHERE NT_TYPECPTE LIKE "D0%"');
    // Liste des données des tables libres à partir de la requête contenu dans SQL
    TMain := TOB.Create('~Main', nil, -1);
    for i := 0 to 9 do
      TOB.Create('TABLE'+IntToStr(i), TMain, -1);

    // T Contient la liste des données des TL et des comptes
    for i := 0 to T.detail.Count-1 do begin
      for j := 0 to 9 do begin
        if T.Detail[i].FieldExists('BS_TABLE'+IntToStr(j)) then begin
          // Recherche si la donnée est déjà stockée
          TFind := nil;
          szData := T.Detail[i].GetString('BS_TABLE'+IntToStr(j));

          // Récapitulatif par section
          if bRecap then begin
            szSQL := GetSQLSectGene(T.Detail[i].GetString('BS_BUDSECT'));
//            TCompte := OpenSQL(szSQL, True); Lek 250106
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
            end;
            TCompte.Free;
            end
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
        end;
      end;
    end;

    // Liste des comptes
    for i := 0 to T.detail.Count-1 do begin
      for j := 0 to 9 do begin
        if T.Detail[i].FieldExists('BS_TABLE'+IntToStr(j)) then begin
          if (FGROUPQUELLERUPTURE.ItemIndex=0) then szData := T.Detail[i].GetString('BS_SECTIONTRIE')+'x'
                                               else szData := T.Detail[i].GetString('BS_TABLE'+IntToStr(j));

          // Récapitulatif par section
          if bRecap then begin
            szSQL := GetSQLSectGene(T.Detail[i].GetString('BS_BUDSECT'));
//            TCompte := OpenSQL(szSQL, True); Lek 250106
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
            end
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
        end;
      end;
    end;

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

  // SANS Rupture et AVEC
  else begin
    iTri := 0;
    for i := 0 to T.Detail.Count-1 do begin
      szSQL := GetSQLSectGene(T.Detail[i].GetString('BS_BUDSECT'));
//      TCompte := OpenSQL(szSQL, True); Lek 250106
      PasserQry2Tob(TCompte,'ZB',szSql);
      for j := 0 to TCompte.Detail.Count-1 do begin
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
                 T.Detail[i].GetString('BS_SECTIONTRIE')+'x',
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

{---------------------------------------------------------------------------------------}
procedure TOF_BALBUDTESEG.DoInsert(szCompte, szLibelleCompte, szRubCompte, szSection, szLibelleSection,
                                   szRubSection, szSectionTrie: string; iTri : Integer; TobFils: Tob);
{---------------------------------------------------------------------------------------}
var
  T2, T3 : Tob;
  j, Decimale : Integer;
  Realise, Ecar : TabTot12;
  TotGen, TotBud : TMontTotal;
  szSQL, Resol, Compte1, Compte2, Lexo : String;
  FinTemp,DateTemp : TDateTime;
  TD : TabloExt;
  Taux, dbl : Double;
begin
  Fillchar(Realise, SizeOf(Realise), #0);
  Fillchar(Ecar, SizeOf(Ecar), #0);
  Fillchar(TotGen, SizeOf(TotGen), #0);
  Fillchar(TotBud, SizeOf(TotBud), #0);

  // Anciennement dans InitCalcBud
  // Requête des mouvements
  szSQL := GetSQLMouvement(szCompte, szSection);
//  T2 := OpenSQL(szSQL, True); Lek 250106
  PasserQry2Tob(T2,'ZC',szSql);

  InitCalcBud(T2, TotBud, GetControlText('FRESOL'), StrToFloat(GetControlText('FTAUX')));
  T2.Free;

  if (TobFils=nil) then begin
    T3 := Tob.Create('CEDTBALBUD', nil, -1);
    T3.InitValeurs(False);
    T3.PutValue('CEB_UTILISATEUR', V_PGI.User);
    T3.PutValue('CEB_BUDGENE', szCompte);
    T3.PutValue('CEB_LIBELLE', szLibelleCompte);
    T3.PutValue('CEB_BUDSECT', szSection);
    T3.PutValue('CEB_LIBSECT', szLibelleSection);
    T3.PutValue('CEB_TRI', iTri);
    T3.PutValue('CEB_PRINTCPTE', '2');
    // Plans de ruptures
    if (szSectionTrie<>'') then begin
      T3.PutValue('CEB_BUDTRI', szSectionTrie);
      T3.PutValue('CEB_LIBTRI', ListeCodesRupture.Items[ListeCodesRupture.Values.IndexOf(szSectionTrie)]);
      if bRecap then T3.PutValue('CEB_BUDSECT', szSectionTrie)
                else T3.PutValue('CEB_BUDSECT', '');
      T3.PutValue('CEB_BUDGENE', '');
      T3.PutValue('CEB_PRINTCPTE', '0');
    end;
    end
  else T3 := TobFils;

  for j := 1 to 12 do begin
    // Initial
    dbl := T3.GetValue('CEB_DEBIT'+IntToStr(j));  T3.PutValue('CEB_DEBIT'+IntToStr(j), dbl+TotBud[0][j-1].TotDebit);
    dbl := T3.GetValue('CEB_CREDIT'+IntToStr(j)); T3.PutValue('CEB_CREDIT'+IntToStr(j), dbl+TotBud[0][j-1].TotCredit);

    // Révisé
    dbl := T3.GetValue('CEB_DEBITREV'+IntToStr(j));  T3.PutValue('CEB_DEBITREV'+IntToStr(j), dbl+TotBud[1][j-1].TotDebit);
    dbl := T3.GetValue('CEB_CREDITREV'+IntToStr(j)); T3.PutValue('CEB_CREDITREV'+IntToStr(j), dbl+TotBud[1][j-1].TotCredit);
  end;
  dbl := T3.GetValue('CEB_TOTDEBIT');     T3.PutValue('CEB_TOTDEBIT', dbl+TotBud[0][12].TotDebit);
  dbl := T3.GetValue('CEB_TOTCREDIT');    T3.PutValue('CEB_TOTCREDIT', dbl+TotBud[0][12].TotCredit);
  dbl := T3.GetValue('CEB_TOTDEBITREV');  T3.PutValue('CEB_TOTDEBITREV', dbl+TotBud[1][12].TotDebit);
  dbl := T3.GetValue('CEB_TOTCREDITREV'); T3.PutValue('CEB_TOTCREDITREV', dbl+TotBud[1][12].TotCredit);

  // Montant du réalisé
  if GetCheckBoxState('FREALISE')= cbChecked then begin
    Compte1 := 'S/G'+GetControlText('FJOURNAL') + szRubSection+':'+szRubCompte;
    Compte2 := '';

    for j := 1 to 12 do begin
      LExo := QUELEXODTBud(TabDate[j]);
      DateTemp := PlusMois(TabDate[j],1);
      if (Lexo<>QUELEXODTBud(DateTemp)) then LExo := '';
      FinTemp := FinDeMois(TabDate[j]);
      GetCumul('RUBREA',Compte1,Compte2,'SAN', GetControlText('FETAB'), 'EUR', LExo, TabDate[j], FinTemp, False, True, Nil, TD, False);
      realise[j-1].TotDebit := TD[3];
      realise[j-1].TotCredit := TD[2];

      dbl := T3.GetValue('CEB_DEBITREA'+IntToStr(j));  T3.PutValue('CEB_DEBITREA'+IntToStr(j), dbl+realise[j-1].TotDebit);
      dbl := T3.GetValue('CEB_CREDITREA'+IntToStr(j)); T3.PutValue('CEB_CREDITREA'+IntToStr(j), dbl+realise[j-1].TotCredit);
    end;

    Resol := GetControlText('FRESOL');
    Taux := StrToFloat(GetControlText('FTAUX'));
    Decimale := 0;
    if Resol='C' then Decimale := V_PGI.OkDecV;

    for j := 0 to 11 do begin
      Reevaluation(Realise[j].TotDebit,Realise[j].TotCredit, Resol, Taux);
      Realise[12].TotDebit := Arrondi(Realise[12].TotDebit + Realise[j].TotDebit, Decimale);
      Realise[12].TotCredit := Arrondi(Realise[12].TotCredit + Realise[j].TotCredit, Decimale);
    end;
    dbl := T3.GetValue('CEB_TOTDEBITREA');  T3.PutValue('CEB_TOTDEBITREA', dbl+Realise[12].TotDebit);
    dbl := T3.GetValue('CEB_TOTCREDITREA'); T3.PutValue('CEB_TOTCREDITREA', dbl+Realise[12].TotCredit);
  end;

  if (TobFils=nil) then begin
    T3.InsertDB(nil);
    T3.Free;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_BALBUDTESEG.FRealiseClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if FRealise.Checked then TFQRS1(Ecran).FEtat.Value := 'BA6'
                      else TFQRS1(Ecran).FEtat.Value := 'BA5';
end;

{ b md 05/06/06 FQ 18077 }
{---------------------------------------------------------------------------------------}
procedure TOF_BALBUDTESEG.FRupturesClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  MajInfoRupture;
end;
{ e md }
{---------------------------------------------------------------------------------------}
procedure TOF_BALBUDTESEG.FPlanRupturesChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Q1  : TQuery;
  St : String;
begin
  if szAxeJal='' then Exit;
  if Length(szAxeJal)<2 then Exit;

  // Rempli la Fourchette des codes ruptures en 'Sur Rupture'
  if (FGROUPQUELLERUPTURE.ItemIndex=0) then begin
    St := '';
    FCodeRupt1.Clear;
    FCodeRupt2.clear ;
    if ListeCodesRupture<>nil then ListeCodesRupture.Clear;

    { b md 05/06/06 FQ 18077 }
    // Il exite au moins un plan de rupture : Se positionne dessus sinon sort
    (*if FPlanRuptures.Items.Count>0 then FPlanRuptures.ItemIndex := 0
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

initialization
  RegisterClasses( [ TOF_BALBUDTESEG ] );

end.
