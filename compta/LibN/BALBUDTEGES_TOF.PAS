{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/01/2006
Modifié le ... :   /  /    
Description .. : Lek
Suite ........ : Je mets ce module en 2/3 et pour compiler j'ai modifié:
Suite ........ : uTob en commun
Suite ........ : OpenSQL à remplacer pour affecter à l'objet Tob.
Mots clefs ... : 
*****************************************************************}
{ Unité : Source TOF de la FICHE : BALBUDTEGES
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
                 16/03/05    VL   Création de l'unité
 6.50.001.013    09/08/05    JP   Création de l'ancêtre TOF_CPBALANCEBUDGET

--------------------------------------------------------------------------------------}

unit BALBUDTEGES_TOF ;

interface

uses StdCtrls, Controls, Classes,
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
  HCtrls, HEnt1, Ent1, CPBALANCEBUDGET_TOF;

procedure BalBudteGenSec;

type

  TOF_BALBUDTEGES = class(TOF_CPBALANCEBUDGET)
    procedure OnArgument(S : string); override;
  protected
    procedure FRealiseClick(Sender: TObject); override;
  public

    procedure PrepCalcBud(SQL : string); override;
    procedure DoInsert(szCompte, szLibelleCompte, szRubCompte, szSection, szLibelleSection, szRubSection : string; TobFils : Tob);
  end;

Implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  UtilEDT,  // ExistBud, WhereLibre, OrderLibre
  CALCOLE,  // GetCumul
  TabLiEdt, // ChoixTableLibreSur, ChoixTableLibre
  TriTabLi; // ChoixTriTableLibre


{---------------------------------------------------------------------------------------}
procedure BalBudteGenSec;
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('CP', 'CPBALBUDTEGES', '', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_BALBUDTEGES.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  aTypeBalance := bal_GeSe;
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_BALBUDTEGES.FRealiseClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if FRealise.Checked then TFQRS1(Ecran).FEtat.Value := 'BA4'
                      else TFQRS1(Ecran).FEtat.Value := 'BA3';
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


{---------------------------------------------------------------------------------------}
procedure TOF_BALBUDTEGES.PrepCalcBud(SQL : String);
{---------------------------------------------------------------------------------------}
var
  TNat, TMain, TFinal, TFils, TFind, TFindNat, T, TSection : Tob;
  i, j, k : Integer;
  szSQL, szData, szOrder, szLib, szSection : String;
  LTri : TStringList;
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

{---------------------------------------------------------------------------------------}
procedure TOF_BALBUDTEGES.DoInsert(szCompte, szLibelleCompte, szRubCompte, szSection, szLibelleSection, szRubSection: String; TobFils: Tob);
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
    Compte1 := 'G/S'+GetControlText('FJOURNAL') + szRubCompte+':'+szRubSection;
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


initialization
  registerclasses ( [ TOF_BALBUDTEGES ] ) ;
end.
