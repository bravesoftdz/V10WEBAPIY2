{***********UNITE*************************************************
Auteur  ...... : Thong hor LIM
Créé le ...... : 29/12/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPBALECBUDSG_QR1 ()
Mots clefs ... : TOF;CPBALECBUDSG_QR1
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 7.00.001.00X    05/06/06    MD   FQ 18077 Correction plan de ruptures
--------------------------------------------------------------------------------------
*****************************************************************}
Unit CPBALECBUDSG_QR1_TOF ;

Interface

Uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL, eQRS1,
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_Main, QRS1,
  {$ENDIF}
  SysUtils, ComCtrls, HCtrls, HEnt1, Ent1, uTob,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ELSE}
  TCalcCum,   
  {$ENDIF MODENT1}
  CPBALECBUD_QR1_TOF;

procedure BalEcartBudParSectionEtCpte;


Type
  TOF_CPBALECBUDSG_QR1 = class(TOF_CPBALECBUD_QR1)
    procedure OnLoad                ; override;
    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
  protected
    FTotIniRev    : TCheckBox;
    FPLANRUPTURES : THValComboBox;
    FCodeRupt1,
    FCodeRupt2    : THValComboBox;

    ListeCodesRupture : THValComboBox;

    procedure FPlanRupturesChange(Sender: TObject);
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
procedure BalEcartBudParSectionEtCpte;
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('CP', 'CPBALECBUDSG_QR1', '', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUDSG_QR1.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  aTypeBalance := ecbud_SeGe;
  ListeCodesRupture := THValComboBox.Create(Ecran);
  ListeCodesRupture.Parent := Ecran;
  ListeCodesRupture.Visible := False;

  FPLANRUPTURES := THValComboBox(DoAssign('FPLANRUPTURES', 'OnChange', FPlanRupturesChange));
  FCodeRupt1 := THValComboBox(GetControl('FCodeRupt1', True));
  FCodeRupt2 := THValComboBox(GetControl('FCodeRupt2', True));

  inherited;
  FTotIniRev := TCheckBox(GetControl('FTOTINIREV', True));
end ;


Procedure PasserQry2Tob(var T: Tob; NomTab, Requette: string);
{Lek 250106 Charge la Requette dans Tob}
var Q: TQuery;
begin
  Q := OpenSql(Requette, True);
  T := Tob.Create(NomTab, nil, -1);
  T.LoadDetailDB(NomTab, '', '', Q, True);
  Ferme(Q);
end;


procedure TOF_CPBALECBUDSG_QR1.PrepCalcBud(SQL: String);
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

  // Liste des comptes budgétaires
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
//    TNat := OpenSQL('SELECT NT_NATURE, NT_LIBELLE FROM NATCPTE WHERE NT_TYPECPTE LIKE "B0%"', True); Lek 250106
    PasserQry2Tob(TNat,'Y','SELECT NT_NATURE, NT_LIBELLE FROM NATCPTE WHERE NT_TYPECPTE LIKE "B0%"');
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

procedure TOF_CPBALECBUDSG_QR1.DoInsert(szCompte, szLibelleCompte, szRubCompte, szSection, szLibelleSection, szRubSection, szSectionTrie : String; iTri : Integer; TobFils : Tob);
var
  T2, T3:                  Tob;
  szSQL:                   string;
  Realise:                 TabTot12;
  TotGen, TotBud :         TMontTotal;
  Resol, Lexo,
  LeCompte, Compte2:       string;
  DateDeb, DateFin:        TDateTime;
  TD:                      TabloExt;   // Array[1..6] of Double
  Taux, dbl:               Double;
begin
  Fillchar(Realise, SizeOf(Realise), #0);
  Fillchar(TotGen, SizeOf(TotGen), #0);
  Fillchar(TotBud, SizeOf(TotBud), #0);

  // Anciennement dans InitCalcBud
  // Requête des mouvements
  szSQL := GetSQLMouvement(szCompte, szSection);
//  T2 := OpenSQL(szSQL, True); Lek 250106
  PasserQry2Tob(T2,'ZC',szSql);

  Resol := GetControlText('FRESOL');
  Taux := StrToFloat(GetControlText('FTAUX'));
  Decimale := 0;
  if Resol='C' then
    Decimale := V_PGI.OkDecV;

  InitCalcBud(T2, TotBud, Resol, Taux); //cumul par les mts par mois avec reformat des mts.
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

  // Budget Initial Période
  dbl := T3.GetValue('CEB_DEBIT1');  T3.PutValue('CEB_DEBIT1',  dbl + TotBud[0][0].TotDebit);
  dbl := T3.GetValue('CEB_CREDIT1'); T3.PutValue('CEB_CREDIT1', dbl + TotBud[0][0].TotCredit);
  // Budget Révisé Période
  dbl := T3.GetValue('CEB_DEBITREV1');  T3.PutValue('CEB_DEBITREV1',  dbl + TotBud[1][0].TotDebit);
  dbl := T3.GetValue('CEB_CREDITREV1'); T3.PutValue('CEB_CREDITREV1', dbl + TotBud[1][0].TotCredit);

  // Budget Initial Cumulé
  dbl := T3.GetValue('CEB_DEBIT2');  T3.PutValue('CEB_DEBIT2',  dbl + TotBud[0][1].TotDebit);
  dbl := T3.GetValue('CEB_CREDIT2'); T3.PutValue('CEB_CREDIT2', dbl + TotBud[0][1].TotCredit);
  // Budget Révisé Cumulé
  dbl := T3.GetValue('CEB_DEBITREV2');  T3.PutValue('CEB_DEBITREV2',  dbl + TotBud[1][1].TotDebit);
  dbl := T3.GetValue('CEB_CREDITREV2'); T3.PutValue('CEB_CREDITREV2', dbl + TotBud[1][1].TotCredit);

  // Budget Initial Annuel
  dbl := T3.GetValue('CEB_DEBIT3');  T3.PutValue('CEB_DEBIT3',  dbl + TotBud[0][2].TotDebit);
  dbl := T3.GetValue('CEB_CREDIT3'); T3.PutValue('CEB_CREDIT3', dbl + TotBud[0][2].TotCredit);
  // Budget Révisé Annuel
  dbl := T3.GetValue('CEB_DEBITREV3');  T3.PutValue('CEB_DEBITREV3',  dbl + TotBud[1][2].TotDebit);
  dbl := T3.GetValue('CEB_CREDITREV3'); T3.PutValue('CEB_CREDITREV3', dbl + TotBud[1][2].TotCredit);

  // Les montants Réalisés à calculer: RÉA_PER, RÉA_CUM, RÉA_PERN_1, RÉA_EXON_1.
  // Attention on ne prend que les écritures dont E_QUALIFPIECE='N'
  if GetCheckBoxState('FREALISE')= cbChecked then begin
    LeCompte := 'S/G'+GetControlText('FJOURNAL') + szRubSection + ':' + szRubCompte;
    Compte2 := '';
    //- RÉA Période = somme des écritures entre les deux dates choisies;
    DateFin := FinDeMois(StrToDate(FDateCompta2));
    DateDeb := StrToDate(FDateCompta1);
    LExo := QUELEXODTBud(StrToDate(FDateCompta1));
    if (Lexo<>QUELEXODTBud(StrToDate(FDateCompta2))) then
      LExo := '';
    GetCumul('RUBREA', LeCompte, Compte2, 'N-', GetControlText('FETAB'), 'EUR', LExo, DateDeb, DateFin, False, True, Nil, TD, False);
    MiseMtsReaDansTable(T3, TD[3], TD[2], Resol, Taux, 1); // alimente CEB_DEBITREA1 et CEB_CREDITREA1

    //- RÉA Cumulé = somme des écritures entre la date début du budget et la date fin choisie;
    LExo := QUELEXODTBud(dtDatDebJal);
    if (Lexo<>QUELEXODTBud(StrToDate(FDateCompta2))) then
      LExo := '';
    DateFin := FinDeMois(StrToDate(FDateCompta2));
    GetCumul('RUBREA', LeCompte, Compte2, 'N-', GetControlText('FETAB'), 'EUR', LExo, dtDatDebJal, DateFin, False, True, Nil, TD, False);
    MiseMtsReaDansTable(T3, TD[3], TD[2], Resol, Taux, 2); // alimente CEB_DEBITREA2 et CEB_CREDITREA2

    //- RÉA PER_N_1 = somme des écritures entre les deux dates choisies - 12 mois.
    DateFin := StrToDate(FDateCompta2);
    DateFin := FinDeMois(PlusMois(DateFin, -12)) ;
    DateDeb := StrToDate(FDateCompta1);
    DateDeb := PlusMois(DateDeb, -12) ;

    LExo := QUELEXODTBud(DateDeb);
    if (Lexo<>QUELEXODTBud(DateFin)) then
      LExo := '';
    GetCumul('RUBREA', LeCompte, Compte2, 'N-', GetControlText('FETAB'), 'EUR', LExo, DateDeb, DateFin, False, True, Nil, TD, False);
    MiseMtsReaDansTable(T3, TD[3], TD[2], Resol, Taux, 3); // alimente CEB_DEBITREA3 et CEB_CREDITREA3

    //- RÉA RÉA_EXO N-1 = somme des écritures de l'exo précédent par rapport à FDateCompta2.
    LExo := VH^.Precedent.Code;
    DateFin := VH^.Precedent.Fin;
    DateDeb := VH^.Precedent.Deb;
    GetCumul('RUBREA', LeCompte, Compte2, 'N-', GetControlText('FETAB'), 'EUR', LExo, DateDeb, DateFin, False, True, Nil, TD, False);
    MiseMtsReaDansTable(T3, TD[3], TD[2], Resol, Taux, 4); // alimente CEB_DEBITREA4 et CEB_CREDITREA4
    end;

  if (TobFils=nil) then begin
    T3.InsertDB(nil);
    T3.Free;
  end;
end;

{ b md 05/06/06 FQ 18077 }
{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUDSG_QR1.FRupturesClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  MajInfoRupture;
end;
{ e md }

procedure TOF_CPBALECBUDSG_QR1.FPlanRupturesChange(Sender: TObject);
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

function TOF_CPBALECBUDSG_QR1.GetMainRequete: string;
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

procedure TOF_CPBALECBUDSG_QR1.MajInfoRupture;
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

procedure TOF_CPBALECBUDSG_QR1.OnClose;
begin
  ListeCodesRupture.Free;
  inherited;
end;

procedure TOF_CPBALECBUDSG_QR1.OnLoad;
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

Initialization
  registerclasses ( [ TOF_CPBALECBUDSG_QR1 ] ) ;
end.

