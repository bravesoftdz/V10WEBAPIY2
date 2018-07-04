{***********UNITE*************************************************
Auteur  ...... : Thong hor LIM
Créé le ...... : 29/12/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPBALECBUDGS_QR1 ()
Mots clefs ... : TOF;CPBALECBUDGS_QR1
*****************************************************************}
Unit CPBALECBUDGS_QR1_TOF ;

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

procedure BalEcartBudParCpteEtSection;


Type
  TOF_CPBALECBUDGS_QR1 = class(TOF_CPBALECBUD_QR1)
  private

   FTotIniRev              : TCheckBox;

   procedure FTotIniRevClick(Sender: TObject);

  public
    procedure OnArgument(S : string); override;
    procedure PrepCalcBud(SQL : string); override;
    procedure DoInsert(szCompte, szLibelleCompte, szRubCompte, szSection, szLibelleSection, szRubSection : string; TobFils : Tob);
  end ;

implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  UtilEDT,  // ExistBud, WhereLibre, OrderLibre
  CALCOLE,  // GetCumul
  TabLiEdt, // ChoixTableLibreSur, ChoixTableLibre
  TriTabLi; // ChoixTriTableLibre

{---------------------------------------------------------------------------------------}
procedure BalEcartBudParCpteEtSection;
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('CP', 'CPBALECBUDGS_QR1', '', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUDGS_QR1.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  aTypeBalance := ecbud_GeSe;
  inherited;

  FTotIniRev := TCheckBox(GetControl('FTOTINIREV', True));
  FTotIniRev.OnClick := FTotIniRevClick;
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

procedure TOF_CPBALECBUDGS_QR1.PrepCalcBud(SQL: String);
var
  TNat, TMain, TFinal, TFils,
  TFind, TFindNat, T, TSection:              Tob;
  szSQL, szData, szOrder, szLib, szSection:  string;
  i, j, k:    Integer;
  LTri:       TStringList;
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
        if T.Detail[i].FieldExists('BG_TABLE'+IntToStr(j)) then begin
          // Recherche si la donnée est déjà stockée
          TFind := nil;
          szData := T.Detail[i].GetString('BG_TABLE'+IntToStr(j));

          // Récapitulatif par section
          if bRecap then begin
            szSQL := GetSQLSectGene(T.Detail[i].GetString('BG_BUDGENE'));
//            TSection := OpenSQL(szSQL, True); Lek 250106
            PasserQry2Tob(TSection,'Z',szSql);

            for k := 0 to TSection.Detail.Count-1 do begin
              szSection := TSection.Detail[k].GetString('BS_BUDSECT');
              szLib := TSection.Detail[k].GetString('BS_LIBELLE');
              if TMain.Detail[j].Detail.Count > 0 then
                TFind := TMain.Detail[j].FindFirst(['CEB_BUDGENE', 'CEB_BUDSECT'],[szData, szSection], False);

              // Pas trouvé : Crée la donnée
              if (TFind=nil) and (szData<>'') then begin
                TFils :=  Tob.Create('CEDTBALBUD', TMain.Detail[j], -1);
                TFils.PutValue('CEB_UTILISATEUR', V_PGI.User);
                TFils.PutValue('CEB_TRI', LTri.IndexOf('BG_TABLE'+IntToStr(j)));
                TFils.PutValue('CEB_BUDGENE', szData);
                TFils.PutValue('CEB_BUDSECT', szSection);
                TFils.PutValue('CEB_LIBSECT', szLib);
                TFindNat := TNat.FindFirst(['NT_NATURE'], [szData], False);
                if (TFindNat<>nil) then TFils.PutValue('CEB_LIBELLE',  TFindNat.GetString('NT_LIBELLE'));
                TFind := TFils;
              end;
            end;
            TSection.Free;
            end
          else begin
            if TMain.Detail[j].Detail.Count > 0 then
              TFind := TMain.Detail[j].FindFirst(['CEB_BUDGENE'],[szData], False);

            // Pas trouvé : Crée la donnée
            if (TFind = nil) and (szData <> '') then begin
              TFils :=  Tob.Create('CEDTBALBUD', TMain.Detail[j], -1);
              TFils.PutValue('CEB_UTILISATEUR', V_PGI.User);
              TFils.PutValue('CEB_TRI', LTri.IndexOf('BG_TABLE'+IntToStr(j)));
              TFils.PutValue('CEB_BUDGENE', szData);
              TFindNat := TNat.FindFirst(['NT_NATURE'], [szData], False);
              if (TFindNat<>nil) then TFils.PutValue('CEB_LIBELLE',  TFindNat.GetString('NT_LIBELLE'));
              TFind := TFils;
            end;
          end;
          // Crée le compte
          TFils := Tob.Create('~COMPTE', TFind, -1);
          TFils.AddChampSup('COMPTE', False);
          TFils.SetString('COMPTE', T.Detail[i].GetString('BG_BUDGENE'));
        end;
      end;
    end;

    // Liste des comptes
    for i := 0 to T.detail.Count-1 do begin
      for j := 0 to 9 do begin
        if T.Detail[i].FieldExists('BG_TABLE'+IntToStr(j)) then begin
          szData := T.Detail[i].GetString('BG_TABLE'+IntToStr(j));

          // Récapitulatif par section
          if bRecap then begin
            szSQL := GetSQLSectGene(T.Detail[i].GetString('BG_BUDGENE'));
//            TSection := OpenSQL(szSQL, True); Lek 250106
            PasserQry2Tob(TSection,'ZA',szSql);
              
            for k := 0 to TSection.Detail.Count-1 do begin
              szSection := TSection.Detail[k].GetString('BS_BUDSECT');
              TFind := TMain.Detail[j].FindFirst(['CEB_BUDGENE', 'CEB_BUDSECT'],[szData, szSection], False);

              if not(TFind=nil) then begin
                DoInsert(T.Detail[i].GetString('BG_BUDGENE'),
                         '',
                         T.Detail[i].GetString('BG_RUB'),
                         szSection,
                         '',
                         '',
                         TFind);
              end;
            end;
            TSection.Free;
            end
          else begin
            TFind := TMain.Detail[j].FindFirst(['CEB_BUDGENE'],[szData], False);
            if not(TFind=nil) then begin
              DoInsert(T.Detail[i].GetString('BG_BUDGENE'),
                       '',
                       T.Detail[i].GetString('BG_RUB'),
                       '',
                       '',
                       '',
                       TFind);
            end;
          end;
        end;
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

    TFinal.Detail.Sort('CEB_BUDGENE');
    TFinal.InsertDB(nil, True);
    TFinal.Free;
    end

  // SANS Rupture et AVEC
  else begin
    for i := 0 to T.Detail.Count-1 do begin
      szSQL := GetSQLSectGene(T.Detail[i].GetString('BG_BUDGENE'));
//      TSection := OpenSQL(szSQL, True); Lek 250106
      PasserQry2Tob(TSection,'ZB',szSql);
        
      for j := 0 to TSection.Detail.Count-1 do begin
        DoInsert(T.Detail[i].GetString('BG_BUDGENE'),
                 T.Detail[i].GetString('BG_LIBELLE'),
                 T.Detail[i].GetString('BG_RUB'),
                 TSection.Detail[j].GetString('BS_BUDSECT'),
                 TSection.Detail[j].GetString('BS_LIBELLE'),
                 TSection.Detail[j].GetString('BS_RUB'),
                 nil);
      end;
      TSection.Free;
    end;
  end;

  LTri.Free;
  T.Free;
end;

procedure TOF_CPBALECBUDGS_QR1.DoInsert(szCompte, szLibelleCompte, szRubCompte, szSection, szLibelleSection, szRubSection : string; TobFils : Tob);

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
    LeCompte := 'G/S'+GetControlText('FJOURNAL') + szRubCompte + ':' + szRubSection;
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

procedure TOF_CPBALECBUDGS_QR1.FTotIniRevClick(Sender: TObject);
begin
  inherited;
//  FQueTotalBud.Enabled:=not (FTotIniRev.Checked) ;
//  if FTotIniRev.Checked then
//    FQueTotalBud.Checked:=False ;
end;

Initialization
  registerclasses ( [ TOF_CPBALECBUDGS_QR1 ] ) ;
end.


