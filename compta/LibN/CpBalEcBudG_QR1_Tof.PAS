{***********UNITE*************************************************
Auteur  ...... : Thong hor LIM
Créé le ...... : 29/12/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPBALECBUDG_QR1 ()
Mots clefs ... : TOF;CPBALECBUDG_QR1
*****************************************************************}
Unit CPBALECBUDG_QR1_TOF ;

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

procedure BalEcartBudParCpte;


Type
  TOF_CPBALECBUDG_QR1 = class(TOF_CPBALECBUD_QR1)
  private

   FQueTotalBud            : TCheckBox;
   FTotIniRev              : TCheckBox;


  public
    procedure OnArgument(S : string); override;
    procedure PrepCalcBud(SQL : string); override;
    procedure DoInsert(szBudGene, szLibelle, szRub : string; TobFils : Tob);
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
procedure BalEcartBudParCpte;
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('CP', 'CPBALECBUDG_QR1', '', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBALECBUDG_QR1.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  aTypeBalance := ecbud_Gene;
  inherited;

  FQueTotalBud := TCheckBox(GetControl('FQUETOTBUD', True));
  FTotIniRev   := TCheckBox(GetControl('FTOTINIREV', True));
end ;

Procedure PasserQry2Tob(var T: Tob; NomTab, Requette: string);
var Q: TQuery;
begin
  Q := OpenSql(Requette, True);
  T := Tob.Create(NomTab, nil, -1);
  T.LoadDetailDB(NomTab, '', '', Q, True);
  Ferme(Q);
end;

procedure TOF_CPBALECBUDG_QR1.PrepCalcBud(SQL: String);
var
  TNat, TMain, TFinal, TFils, TFind, T : Tob;
  i, j : Integer;
  szData, szOrder : String;
  LTri : TStringList;
begin
//  T := OpenSQL(SQL, True);
  PasserQry2Tob(T, 'X', Sql);

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
//    TNat := OpenSQL('SELECT NT_NATURE, NT_LIBELLE FROM NATCPTE WHERE NT_TYPECPTE LIKE "B0%"', True);
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
          if TMain.Detail[j].Detail.Count > 0 then
            TFind := TMain.Detail[j].FindFirst(['CEB_BUDGENE'],[szData], False);

          // Pas trouvé : Crée la donnée
          if (TFind=nil) and (szData<>'') then begin
            TFils :=  Tob.Create('CEDTBALBUD', TMain.Detail[j], -1);
            TFils.PutValue('CEB_UTILISATEUR', V_PGI.User);
            TFils.PutValue('CEB_TRI', LTri.IndexOf('BG_TABLE'+IntToStr(j)));
            TFils.PutValue('CEB_BUDGENE', szData);
            TFind := TNat.FindFirst(['NT_NATURE'], [szData], False);
            if (TFind<>nil) then TFils.PutValue('CEB_LIBELLE',  TFind.GetString('NT_LIBELLE'));
            TFind := TFils;
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
          TFind := TMain.Detail[j].FindFirst(['CEB_BUDGENE'],[szData], False);
          if not(TFind=nil) then
            DoInsert(T.Detail[i].GetString('BG_BUDGENE'), '', T.Detail[i].GetString('BG_RUB'), TFind);
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
  else begin
    for i := 0 to T.detail.Count-1 do
      DoInsert(T.Detail[i].GetString('BG_BUDGENE'), T.Detail[i].GetString('BG_LIBELLE'), T.Detail[i].GetString('BG_RUB'), nil);
  end;

  LTri.Free;
  T.Free;
end;

procedure TOF_CPBALECBUDG_QR1.DoInsert(szBudGene, szLibelle, szRub: string;  TobFils: Tob);
var
  T2, T3:                  Tob;
  szSQL, szEtab:           String;
  FValide:                 TCheckBox;
  Realise:                 TabTot12;
  TotGen, TotBud :         TMontTotal;
  Resol, LeCompte, Lexo:   String;
  DateDeb, DateFin:        TDateTime;
  TD:                      TabloExt;   // Array[1..6] of Double
  Taux, dbl:               Double;
begin
  Fillchar(Realise, SizeOf(Realise), #0);
  Fillchar(TotGen, SizeOf(TotGen), #0);
  Fillchar(TotBud, SizeOf(TotBud), #0);

  // Anciennement dans InitCalcBud
  // Requête des mouvements
  szSQL := 'SELECT BE_BUDGENE, BE_BUDSECT, BE_EXERCICE, BE_DATECOMPTABLE, BE_NUMEROPIECE, ' +
           'BE_REFINTERNE, BE_ETABLISSEMENT, BE_LIBELLE, BE_VALIDE,' +
           'BE_BUDJAL,BE_NATUREBUD, BE_QUALIFPIECE, BE_DEBIT DEBIT,BE_CREDIT CREDIT ' +
           'FROM BUDECR WHERE BE_BUDGENE="'+szBudGene+'" ' +
           'AND BE_BUDJAL="'+GetControlText('FJOURNAL')+'" ';

  if FStWhereNat <> '' then
    szSQL := szSQL + 'AND ' + FStWhereNat + ' ';

  {b Thl FQ 18277 05/06/2006}
   // On ramasse les mvts budgets depuis FDateCompta1 sauf si (FDateCompta1 > dtDatDebJal)
   // au dtDatFinJal sauf si (FDateCompta2 > dtDatFinJal).
  if (StrToDate(FDateCompta1) >  dtDatDebJal) then
    szSQL := szSQL + 'AND BE_DATECOMPTABLE>="'+USDateTime(dtDatDebJal) + '" '
  else
    szSQL := szSQL + 'AND BE_DATECOMPTABLE>="'+USDateTime(StrToDate(FDateCompta1)) + '" ';

  if (StrToDate(FDateCompta2) > dtDatFinJal) then
    szSQL := szSQL + ' AND BE_DATECOMPTABLE<="'+ usdatetime(StrToDate(FDateCompta2))+'" '
  else
    szSQL := szSQL + ' AND BE_DATECOMPTABLE<="'+ usdatetime(dtDatFinJal)+'" ';
    szSQL := szSQL + ' AND BE_EXERCICE>="'+GetControlText('FEXERCICE')+'" AND BE_EXERCICE<="'+GetControlText('FEXERCICE_')+'" ';
  {e Thl}
  szEtab := GetControlText('FETAB');
  if szEtab<>'' then szSQL := szSQL + ' AND BE_ETABLISSEMENT="'+szEtab+'" ';
  FValide := TCheckBox(GetControl('FVALIDE', True));
  if FValide.Checked then szSQL := szSQL + ' AND BE_VALIDE="X" ' else
  if FValide.State = cbUnchecked then szSQL := szSQL + ' AND BE_VALIDE="-" ';
  szSQL := szSQL + 'ORDER BY BE_BUDGENE,BE_BUDJAL,BE_EXERCICE,BE_DATECOMPTABLE,BE_BUDSECT,BE_AXE,BE_NUMEROPIECE,BE_QUALIFPIECE';
//  T2 := OpenSQL(szSQL, True);
  PasserQry2Tob(T2, 'Z', szSQL);

  Resol := GetControlText('FRESOL');
  Taux := StrToFloat(GetControlText('FTAUX'));
  Decimale := 0;
  if Resol='C' then
    Decimale := V_PGI.OkDecV;

  InitCalcBud(T2, TotBud, Resol, Taux); //cumul les mts par nature et période avec reformat des mts.
  T2.Free;

  if (TobFils=nil) then begin
    T3 := Tob.Create('CEDTBALBUD', nil, -1);
    T3.InitValeurs(False);
    T3.PutValue('CEB_UTILISATEUR', V_PGI.User);
    T3.PutValue('CEB_BUDGENE', szBudGene);
    T3.PutValue('CEB_LIBELLE', szLibelle);
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
    LeCompte := 'CBG'+ GetControlText('FJOURNAL') + szRub;
    //- RÉA Période = somme des écritures entre les deux dates choisies;
    DateFin := FinDeMois(StrToDate(FDateCompta2));
    DateDeb := StrToDate(FDateCompta1);
    LExo := QUELEXODTBud(StrToDate(FDateCompta1));
    if (Lexo<>QUELEXODTBud(StrToDate(FDateCompta2))) then
      LExo := '';
    GetCumul('RUBREA', LeCompte, LeCompte, 'N-', GetControlText('FETAB'), 'EUR', LExo, DateDeb, DateFin, False, True, Nil, TD, False);
    MiseMtsReaDansTable(T3, TD[3], TD[2], Resol, Taux, 1); // alimente CEB_DEBITREA1 et CEB_CREDITREA1

    //- RÉA Cumulé = somme des écritures entre la date début du budget et la date fin choisie;
    LExo := QUELEXODTBud(dtDatDebJal);
    if (Lexo<>QUELEXODTBud(StrToDate(FDateCompta2))) then
      LExo := '';
    DateFin := FinDeMois(StrToDate(FDateCompta2));
    GetCumul('RUBREA', LeCompte, LeCompte, 'N-', GetControlText('FETAB'), 'EUR', LExo, dtDatDebJal, DateFin, False, True, Nil, TD, False);
    MiseMtsReaDansTable(T3, TD[3], TD[2], Resol, Taux, 2); // alimente CEB_DEBITREA2 et CEB_CREDITREA2

    //- RÉA PER_N_1 = somme des écritures entre les deux dates choisies - 12 mois.
    DateFin := StrToDate(FDateCompta2);
    DateFin := FinDeMois(PlusMois(DateFin, -12)) ;
    DateDeb := StrToDate(FDateCompta1);
    DateDeb := PlusMois(DateDeb, -12) ;
    LExo := QUELEXODTBud(DateDeb);
    if (Lexo<>QUELEXODTBud(DateFin)) then
      LExo := '';
    GetCumul('RUBREA', LeCompte, LeCompte, 'N-', GetControlText('FETAB'), 'EUR', LExo, DateDeb, DateFin, False, True, Nil, TD, False);
    MiseMtsReaDansTable(T3, TD[3], TD[2], Resol, Taux, 3); // alimente CEB_DEBITREA3 et CEB_CREDITREA3

    //- RÉA RÉA_EXO N-1 = somme des écritures de l'exo précédent par rapport à FDateCompta2.
    LExo := VH^.Precedent.Code;
    DateFin := VH^.Precedent.Fin;
    DateDeb := VH^.Precedent.Deb;
    GetCumul('RUBREA', LeCompte, LeCompte, 'N-', GetControlText('FETAB'), 'EUR', LExo, DateDeb, DateFin, False, True, Nil, TD, False);
    MiseMtsReaDansTable(T3, TD[3], TD[2], Resol, Taux, 4); // alimente CEB_DEBITREA4 et CEB_CREDITREA4
    end;

  if (TobFils=nil) then begin
    T3.InsertDB(nil);
    T3.Free;
  end;
end;

Initialization
  registerclasses ( [ TOF_CPBALECBUDG_QR1 ] ) ;
end.

