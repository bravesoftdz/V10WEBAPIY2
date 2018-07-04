{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 30/03/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BALBUDECGEN ()
Mots clefs ... : TOF;BALBUDECGEN
*****************************************************************}
Unit BALBUDECSEC_TOF ;

Interface

Uses StdCtrls, Controls, Classes, 
{$IFDEF EAGLCLIENT}
  eMul, uTob, MaineAGL, eQRS1,
{$ELSE}
  db, dbtables, mul, FE_Main, QRS1,
{$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, Ent1, TCalcCum, CritEdt;

procedure BalBudecSec;

Type
  TabTot12 = Array[0..12] Of TabDC;
  TMontTot = Array[0..2] Of TabTot12;

  TOF_BALBUDECSEC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    FJournal, FEXERCICE, FEXERCICE_, FPeriode1, FPeriode2, FSELECTCPTE, FPLANRUPTURES : THValComboBox;
    FGROUPCHOIXRUPT, FGROUPQUELLERUPTURE : THRadioGroup;
    FCodeRupt1, FCodeRupt2 : THValComboBox;
    szExoDebJal, szExoFinJal, szAxeJal : String3;
    dtDatDebJal, dtDatFinJal : TDateTime;
    bDebitPos : Boolean;
    FDateCompta1, FDateCompta2, szSoldeFormate : String;
    Decimale : Byte;
    FQueTotBud, FTotIniRev : TCheckBox;
    ListeCodesRupture : THValComboBox;
    procedure FJournalChange(Sender: TObject);
    procedure FExerciceChange(Sender: TObject);
    procedure FExercice_Change(Sender: TObject);
    procedure FPeriode1Change(Sender: TObject);
    procedure FPeriode2Change(Sender: TObject);
    procedure FLibreDblClick(Sender: TObject);
    procedure FSansRuptClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure FPlanRupturesChange(Sender: TObject);
    procedure FQueTotBudClick(Sender: TObject);
    procedure FTotIniRevClick(Sender: TObject);
    function UpdateRupture(szLibre1, szLibre2 : String) : String;

    function DoAssign(Control, Evt : String; Event : TNotifyEvent) : TControl;
    procedure InfosJournal;
    procedure InitCompte;
    procedure DebToDate;
    procedure FinToDate;
    function GenereSQL : String;
    procedure PrepCalcBud(SQL : String);
    procedure CalculBud(i : Byte ; Var Tot1 : TMontTot ; NatB : String ; D,C : Double ; Dec : Byte);
    procedure InitCalcBud(T : Tob; var Tot1 : TMontTot; Resol : String; Taux : Double);
    procedure DoInsert(szBudGene, szLibelle, szRub : String; TobFils : Tob);
    procedure AlimTotEdtEcartBudget(Var LeTotal : TabTot12 ; Var SousTot : TMontTot);
    function GetSolde(TD,TC : DOUBLE ; Decim : Integer ; FormMont : String) : Double;

  end ;

Implementation

uses UtilEDT,  // ExistBud, WhereLibre, OrderLibre
     CALCOLE,  // GetCumul
     TabLiEdt, // ChoixTableLibreSur, ChoixTableLibre
     TriTabLi; // ChoixTriTableLibre

procedure BalBudecSec;
begin
  AGLLanceFiche('CP', 'CPBALBUDECSEC', '', '', '');
end;

procedure TOF_BALBUDECSEC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BALBUDECSEC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BALBUDECSEC.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BALBUDECSEC.OnLoad ;
var
  szDec : String;
  FResol : THValComboBox;
  StTri, StCod1, StCod2 : String;
begin
  Inherited;

  if bDebitPos then szSoldeFormate := 'PD'
               else szSoldeFormate := 'PC';

  // Initialisation des contrôles
  InitCompte;

  // Format
  FResol := THValComboBox(GetControl('FRESOL', True));
  case FResol.ItemIndex of
    0: SetControlText('FORMAT', '#,##0.00;; ;'); // Avec décimales
    1: SetControlText('FORMAT', '#,##0;; ;'); // Sans décimales
    2: SetControlText('FORMAT', '#,##0;; ;'); // Kilo
    3: SetControlText('FORMAT', '#,##0;; ;'); // Mega
  end;

  // FourchetteExoOk
  if not ((QUELEXODTBud(dtDatDebJal)>=szExoDebJal) and (QUELEXODTBud(dtDatFinJal)<=szExoFinJal)) then begin
    LastError := 1;
    LastErrorMsg := 'La fourchette d''exercice ne correspond pas à celle du budget.';
    exit;
  end;

  // Tables libres
  if (FGROUPCHOIXRUPT.ItemIndex>0) and (Trim(GetControlText('FLIBTRIPAR'))= '') then begin
    StTri := GetControlText('FLIBTRIPAR');
    if GetCheckBoxState('FONLYCPTASSOCIE')=cbChecked then ChoixTableLibreSurInfo(fbBudgen, StTri, StCod1, StCod2)
                                                     else ChoixTableLibreInfo(fbBudgen, StTri, StCod1, StCod2);
    SetControlText('FLIBTRIPAR', UpdateRupture(StCod1,StCod2));
  end;

  szDec := GetControlText('FRESOL');
  if szDec='C' then Decimale := V_PGI.OkDecV
               else Decimale := 0;

  ExecuteSQL('DELETE FROM CEDTECABUD WHERE CEC_UTILISATEUR = "' + V_PGI.User + '"');
  PrepCalcBud(GenereSQL);
end ;

procedure TOF_BALBUDECSEC.OnArgument (S : String ) ;
begin
  Inherited ;

  ListeCodesRupture := THValComboBox.Create(Ecran);
  ListeCodesRupture.Parent := Ecran;
  ListeCodesRupture.Visible := False;

  FJournal := THValComboBox(DoAssign('FJOURNAL', 'OnChange', FJournalChange));
  FEXERCICE := THValComboBox(DoAssign('FEXERCICE', 'OnChange', FExerciceChange));
  FEXERCICE_ := THValComboBox(DoAssign('FEXERCICE_', 'OnChange', FExercice_Change));
  FPeriode1 := THValComboBox(DoAssign('FPeriode1', 'OnChange', FPeriode1Change));
  FPeriode2 := THValComboBox(DoAssign('FPeriode2', 'OnChange', FPeriode2Change));
  FGROUPCHOIXRUPT := THRadioGroup(DoAssign('FGROUPCHOIXRUPT', 'OnClick', FSansRuptClick));
  FGROUPQUELLERUPTURE := THRadioGroup(DoAssign('FGROUPQUELLERUPTURE', 'OnClick', FRupturesClick));
  FQueTotBud := TCheckBox(DoAssign('FQUETOTBUD', 'CheckOnClick', FQueTotBudClick));
  FTotIniRev := TCheckBox(DoAssign('FTOTINIREV', 'CheckOnClick', FTotIniRevClick));
  FPLANRUPTURES := THValComboBox(DoAssign('FPLANRUPTURES', 'OnChange', FPlanRupturesChange));
  FCodeRupt1 := THValComboBox(GetControl('FCodeRupt1', True));
  FCodeRupt2 := THValComboBox(GetControl('FCodeRupt2', True));

  DoAssign('FLIBRE1', 'OnDblClick', FLibreDblClick);
  DoAssign('FLIBRE2', 'OnDblClick', FLibreDblClick);
  DoAssign('FLIBTRIPAR', 'OnDblClick', FLibreDblClick);

  // Initialise les contrôles
  if FJournal.Values.Count>0 then FJournal.Value:=FJournal.Values[0];
  THValComboBox(GetControl('FNatBud', True)).ItemIndex := 0;
  THValComboBox(GetControl('FResol', True)).ItemIndex := 1;
  THValComboBox(GetControl('FETAB', True)).ItemIndex := 0;
  FSELECTCPTE := THValComboBox(GetControl('FSELECTCPTE', True));
  FSELECTCPTE.Value := 'TOO';
  InitCompte;

  SetControlEnabled('FONLYCPTASSOCIE', False);
  if (EstSerie(S3)) then SetControlVisible('FGROUPCHOIXRUPT', False);

  FSansRuptClick(nil);  
end ;

procedure TOF_BALBUDECSEC.OnClose ;
begin
  ListeCodesRupture.Free;

  Inherited ;
end ;

procedure TOF_BALBUDECSEC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BALBUDECSEC.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BALBUDECSEC.FJournalChange(Sender: TObject);
begin
  if FJournal.Value='' then Exit;
  InfosJournal;
end;

function TOF_BALBUDECSEC.DoAssign(Control, Evt : String; Event: TNotifyEvent) : TControl;
var
  szEvt : String;
begin
  szEvt := UpperCase(Evt);
  Result := GetControl(Control, True);
  if Result = nil then exit;
  if szEvt = 'ONCHANGE' then TComboBox(Result).OnChange := Event else
  if szEvt = 'ONDBLCLICK' then THEdit(Result).OnDblClick := Event else
  if szEvt = 'ONCLICK' then THRadioGroup(Result).OnClick := Event else
  if szEvt = 'CHECKONCLICK' then TCheckBox(Result).OnClick := Event;
end;

procedure TOF_BALBUDECSEC.FExerciceChange(Sender: TObject);
begin
  if FExercice.Value>FExercice_.Value then FExercice_.Value := FExercice.Value;
  ListePeriode(FExercice.Value, FPeriode1.Items, FPeriode1.Values, True);
  FPeriode1.ItemIndex := 0;
  DebToDate;
end;

procedure TOF_BALBUDECSEC.FExercice_Change(Sender: TObject);
begin
  if FExercice_.Value<FExercice.Value then FExercice.Value := FExercice_.Value;
  ListePeriode(FExercice_.Value, FPeriode2.Items, FPeriode2.Values, False);
  FPeriode2.ItemIndex := FPeriode2.Items.Count-1;
  FinToDate;
end;

procedure TOF_BALBUDECSEC.FPeriode1Change(Sender: TObject);
begin
  if StrToDate(FPeriode1.Value)>StrToDate(FPeriode2.Value) then FPeriode2.Value := DateToStr(FinDeMois(StrToDate(FPeriode1.Value)));
  FDateCompta1 := FPeriode1.Value;
end;

procedure TOF_BALBUDECSEC.FPeriode2Change(Sender: TObject);
begin
  if StrToDate(FPeriode2.Value)<StrToDate(FPeriode1.Value) then FPeriode1.Value := DateToStr(DebutDeMois(StrToDate(FPeriode2.Value)));
  FDateCompta2 := FPeriode2.Value;
end;

procedure TOF_BALBUDECSEC.InfosJournal;
var
  QJal  : TQuery;
  szOldAxeJal : String;
begin
  QJal := OpenSQL('SELECT BJ_EXODEB, BJ_EXOFIN, BJ_PERDEB, BJ_PERFIN, BJ_AXE, BJ_NATJAL FROM BUDJAL WHERE BJ_BUDJAL="'+FJournal.Value+'"', True);
  if Not QJal.EOF then begin
    FEXERCICE.Value := QJal.FindField('BJ_EXODEB').AsString;
    FEXERCICE_.Value := QJal.FindField('BJ_EXOFIN').AsString;
    szExoDebJal := QJal.FindField('BJ_EXODEB').AsString;
    szExoFinJal := QJal.FindField('BJ_EXOFIN').AsString;
    FPeriode1.Value := QJal.FindField('BJ_PERDEB').AsString;
    FPeriode2.Value := QJal.FindField('BJ_PERFIN').AsString;
    FDateCompta1 := QJal.FindField('BJ_PERDEB').AsString;
    FDateCompta2 := QJal.FindField('BJ_PERFIN').AsString;
    dtDatDebJal := QJal.FindField('BJ_PERDEB').AsDateTime;
    dtDatFinJal := QJal.FindField('BJ_PERFIN').AsDateTime;
    szOldAxeJal := szAxeJal;
    szAxeJal := QJal.FindField('BJ_AXE').AsString;
    bDebitPos := QJal.FindField('BJ_NATJAL').AsString='CHA';

    // Si changement d'axe
    if (szOldAxeJal <> szAxeJal) then begin
      SetControlProperty('FCPTE1', 'DATATYPE', 'TZBUDSEC'+szAxeJal[2]);
      SetControlProperty('FCPTE2', 'DATATYPE', 'TZBUDSEC'+szAxeJal[2]);
      SetControlText('FCPTE1', '');
      SetControlText('FCPTE2', '');
    end;
  end;
  Ferme(QJal);
end;

procedure TOF_BALBUDECSEC.InitCompte;
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT MIN(BS_BUDSECT), Max(BS_BUDSECT) FROM BUDSECT WHERE BS_FERME="-" AND BS_AXE="'+szAxeJal+'"', True);
  if Not Q.EOF then begin
    if GetControlText('FCPTE1') = '' then SetControlText('FCPTE1', Q.Fields[0].AsString);
    if GetControlText('FCPTE2') = '' then SetControlText('FCPTE2', Q.Fields[1].AsString);
  end;
  Ferme(Q);
end;

procedure TOF_BALBUDECSEC.DebToDate;
begin
  FDateCompta1 := FPeriode1.Value;
end;

procedure TOF_BALBUDECSEC.FinToDate;
begin
  FDateCompta2 := FPeriode2.Value;
end;

function TOF_BALBUDECSEC.GenereSQL : String;
var
  sz, szCpte1, szCpte2 : String;
begin
  if (FGROUPCHOIXRUPT.ItemIndex=2) then Result := 'SELECT BS_BUDSECT, BS_LIBELLE, '+OrderLibre(GetControlText('FLIBTRIPAR'))+' BS_RUB, BS_SECTIONTRIE FROM BUDSECT S WHERE '
                                   else Result := 'SELECT BS_BUDSECT, BS_LIBELLE, BS_RUB FROM BUDSECT S WHERE ';
  Result := Result + ExistBud(AxeToFbBud(szAxeJal), GetControlText('FSELECTCPTE'), GetControlText('FJOURNAL'), szAxeJal, True);
  Result := Result + ' AND BS_AXE="'+szAxeJal+'" ';

  szCpte1 := GetControlText('FCPTE1');
  szCpte2 := GetControlText('FCPTE2');
  if szCpte1<>'' then Result := Result + ' AND BS_BUDSECT>="'+szCpte1+'" ';
  if szCpte2<>'' then Result := Result + ' AND BS_BUDSECT<="'+szCpte2+'" ';

  // SUR plans de ruptures
  if (FGROUPCHOIXRUPT.ItemIndex=2) and (FGROUPQUELLERUPTURE.ItemIndex=0) then begin
    szCpte1 := THValComboBox(GetControl('FCODERUPT1', True)).Text;
    szCpte2 := THValComboBox(GetControl('FCODERUPT2', True)).Text;
    if szCpte1<>'' then Result := Result + ' AND BS_SECTIONTRIE>="'+szCpte1+'" ';
    if szCpte2<>'' then Result := Result + ' AND BS_SECTIONTRIE<="'+szCpte2+'" ';
  end;

  if (FGROUPCHOIXRUPT.ItemIndex>0) then begin // AVEC ou SUR
    case FGROUPQUELLERUPTURE.ItemIndex  of
      0 : begin // Sur Plan de rupture
        Result := Result + ' ORDER BY BS_AXE, BS_SECTIONTRIE';
      end;
      1 : begin // Sur tables libres
        szCpte1 := GetControlText('FLIBRE1');
        szCpte2 := GetControlText('FLIBRE2');
        sz := WhereLibre(szCpte1, szCpte2, AxeToFbBud(szAxeJal), (GetCheckBoxState('FONLYCPTASSOCIE')=cbChecked));
        if sz<>'' then Result := Result + sz;
        Result := Result + ' ORDER BY ' + OrderLibre(GetControlText('FLIBTRIPAR'))+'BS_BUDSECT ';
      end;
    end;
    end
  else Result := Result + ' ORDER BY BS_AXE, BS_BUDSECT';
end;

procedure TOF_BALBUDECSEC.FLibreDblClick(Sender: TObject);
var
  StTri, StCod1, StCod2 : String;
  LeFb : TFichierBase;
begin
  StCod1 := GetControlText('FLIBRE1');
  StCod2 := GetControlText('FLIBRE2');
  LeFb := fbBudgen;

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

procedure TOF_BALBUDECSEC.FSansRuptClick(Sender: TObject);
begin
  SetControlVisible('SHTRUPTURE', (FGROUPCHOIXRUPT.ItemIndex >0));
  if (FGROUPCHOIXRUPT.ItemIndex = 2) then SetControlText('TRI', '')
                                     else SetControlText('TRI', 'X');
end;

procedure TOF_BALBUDECSEC.FRupturesClick(Sender: TObject);
begin
  SetControlVisible('FGROUPRUPTURES', (FGROUPQUELLERUPTURE.ItemIndex=0));
  SetControlVisible('FONLYCPTASSOCIE', (FGROUPQUELLERUPTURE.ItemIndex=1));
  SetControlChecked('FONLYCPTASSOCIE', (FGROUPQUELLERUPTURE.ItemIndex=1));
end;

procedure TOF_BALBUDECSEC.FPlanRupturesChange(Sender: TObject);
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

    // Il exite au moins un plan de rupture : Se positionne dessus sinon sort
    if FPlanRuptures.Items.Count>0 then FPlanRuptures.ItemIndex := 0
                                   else exit;
                                   
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

procedure TOF_BALBUDECSEC.FQueTotBudClick(Sender: TObject);
begin
  FTotIniRev.Enabled := not FQueTotBud.Checked;
  if FQueTotBud.Checked then FTotIniRev.Checked := False;
end;

procedure TOF_BALBUDECSEC.FTotIniRevClick(Sender: TObject);
begin
  FQueTotBud.Enabled := not FTotIniRev.Checked;
  if FTotIniRev.Checked then FQueTotBud.Checked := False;
end;

procedure TOF_BALBUDECSEC.PrepCalcBud(SQL: String);
var
  TNat, TMain, TFinal, TFils, TFind, TFindNat, T : Tob;
  i, j : Integer;
  szTemp, szData, szOrder : String;
  LTri : TStringList;
begin
  T := OpenSQL(SQL, True);

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
  if FGROUPCHOIXRUPT.ItemIndex=2 then begin
    // Pour le libellé des tables libres
    TNat := OpenSQL('SELECT NT_NATURE, NT_LIBELLE FROM NATCPTE WHERE NT_TYPECPTE LIKE "D0%"', True);

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
          if TMain.Detail[j].Detail.Count > 0 then
            TFind := TMain.Detail[j].FindFirst(['CEC_BUDGENE'],[szData], False);

          // Pas trouvé : Crée la donnée
          if (TFind=nil) and (szData<>'') then begin
            TFils :=  Tob.Create('CEDTECABUD', TMain.Detail[j], -1);
            TFils.PutValue('CEC_UTILISATEUR', V_PGI.User);
            TFils.PutValue('CEC_TRI', LTri.IndexOf('BS_TABLE'+IntToStr(j)));

            // Plan de rupture
            if (FGROUPQUELLERUPTURE.ItemIndex=0) then begin
              szTemp := T.Detail[i].GetString('BS_SECTIONTRIE')+'x';
              ListeCodesRupture.Value := szTemp;
              TFils.PutValue('CEC_BUDGENE', szTemp);
              TFils.PutValue('CEC_LIBELLE', ListeCodesRupture.Text);
              TFils.PutValue('CEC_TRI', TFils.GetIndex);
              end
            // Tables libres
            else begin
              TFils.PutValue('CEC_BUDGENE', szData);
              TFindNat := TNat.FindFirst(['NT_NATURE'], [szData], False);
              if (TFindNat<>nil) then TFils.PutValue('CEC_LIBELLE',  TFindNat.GetString('NT_LIBELLE'));
              TFind := TFils;
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
          TFind := TMain.Detail[j].FindFirst(['CEC_BUDGENE'],[szData], False);
          if not(TFind=nil) then
            DoInsert(T.Detail[i].GetString('BS_BUDSECT'), '', T.Detail[i].GetString('BS_RUB'), TFind);
        end;
      end;
    end;

    TNat.Free;
    // Tri les Tob avant insertion
    TFinal := TOB.Create('~CEDTECABUD', nil, -1);
    for j := 0 to 9 do begin
      while TMain.Detail[j].Detail.Count>0 do begin
        // Supprime les enfants (liste des comptes)
        TMain.Detail[j].Detail[0].ClearDetail;

        // Bascule sur le nouveau parent
        TMain.Detail[j].Detail[0].ChangeParent(TFinal, -1);
      end;
    end;
    TMain.Free;

    TFinal.Detail.Sort('CEC_BUDGENE');
    TFinal.InsertDB(nil, True);
    TFinal.Free;
    end
  else begin
    for i := 0 to T.detail.Count-1 do
      DoInsert(T.Detail[i].GetString('BS_BUDSECT'), T.Detail[i].GetString('BS_LIBELLE'), T.Detail[i].GetString('BS_RUB'), nil);
  end;

  LTri.Free;
  T.Free;
end;

procedure TOF_BALBUDECSEC.DoInsert(szBudGene, szLibelle, szRub: String; TobFils: Tob);
var
  T2, T3 : Tob;
  szSQL, szNature, szEtab : String;
  FValide : TCheckBox;
  Realise, Ecar : TabTot12;
  TotEdt, TotIR : TabTot12; // Total de l'édition, total Initial +Révisé
  TotGen, TotBud : TMontTot;
  Resol, LeCompte, Lexo : String;
  FinTemp, dtDate1 : TDateTime;
  TD : TabloExt;
  Taux : Double;
  Decimale : Integer;
  dbl : Double;
  TotLibre : TabMont77;
  Pourcentage : double;
  RealPer, EcarPer, RealCum, EcarCum : TabTot12;
  MReport : TabTRep;
begin
  Fillchar(Realise, SizeOf(Realise), #0);
  Fillchar(Ecar, SizeOf(Ecar), #0);
  Fillchar(TotGen, SizeOf(TotGen), #0);
  Fillchar(TotBud, SizeOf(TotBud), #0);

  // Anciennement dans InitCalcBud
  // Requête des mouvements
  szSQL := 'SELECT BE_BUDSECT, BE_EXERCICE, BE_DATECOMPTABLE, BE_NUMEROPIECE, ' +
           'BE_REFINTERNE, BE_ETABLISSEMENT, BE_LIBELLE, BE_VALIDE,' +
           'BE_BUDJAL,BE_NATUREBUD, BE_QUALIFPIECE, BE_QTE1, BE_QTE2, BE_DEBIT DEBIT,BE_CREDIT CREDIT ' +
           'FROM BUDECR WHERE BE_BUDSECT="'+szBudGene+'" ' +
           'AND BE_AXE="'+szAxeJal+'" '+
           'AND BE_BUDJAL="'+GetControlText('FJOURNAL')+'" ';
  szNature := GetControlText('FNATBUD');
  if szNature<>'' then szSQL := szSQL + 'AND BE_NATUREBUD="'+szNature+'" ';
  szSQL := szSQL + 'AND BE_DATECOMPTABLE>="'+usdatetime(StrToDate(FDateCompta1))+'" AND BE_DATECOMPTABLE<="'+usdatetime(StrToDate(FDateCompta2))+'" '+
                   'AND BE_EXERCICE>="'+GetControlText('FEXERCICE')+'" AND BE_EXERCICE<="'+GetControlText('FEXERCICE_')+'" ';
  FValide := TCheckBox(GetControl('AVECREVISION', True));
  if FValide.Checked then szSQL := szSQL + ' AND BE_QUALIFPIECE="R" ' else
  if FValide.State = cbUnchecked then szSQL := szSQL + ' AND BE_QUALIFPIECE<>"R" ';
  szEtab := GetControlText('FETAB');
  if szEtab<>'' then szSQL := szSQL + ' AND BE_ETABLISSEMENT="'+szEtab+'" ';
  FValide := TCheckBox(GetControl('FVALIDE', True));
  if FValide.Checked then szSQL := szSQL + ' AND BE_VALIDE="X" ' else
  if FValide.State = cbUnchecked then szSQL := szSQL + ' AND BE_VALIDE="-" ';
  szSQL := szSQL + 'ORDER BY BE_BUDSECT,BE_AXE,BE_EXERCICE,BE_DATECOMPTABLE,BE_BUDJAL,BE_BUDGENE,BE_NUMEROPIECE,BE_QUALIFPIECE';
  T2 := OpenSQL(szSQL, True);
  InitCalcBud(T2, TotBud, GetControlText('FRESOL'), StrToFloat(GetControlText('FTAUX')));
  T2.Free;

  if (TobFils=nil) then begin
    T3 := Tob.Create('CEDTECABUD', nil, -1);
    T3.InitValeurs(False);
    T3.PutValue('CEC_UTILISATEUR', V_PGI.User);
    T3.PutValue('CEC_BUDGENE', szBudGene);
    T3.PutValue('CEC_LIBELLE', szLibelle);
    end
  else T3 := TobFils;

  Resol := GetControlText('FRESOL');
  Taux := StrToFloat(GetControlText('FTAUX'));
  Decimale := 0;
  if Resol='C' then Decimale := V_PGI.OkDecV;

  // Affectation du calcul des montants budgétes pour le total du Compte budgétaire aux cases suivantes du tableau...
  // 0 : Budgété de la période, 4 : Budgété Cumulé, 8 : Budgété annuel
  TotGen := TotBud;
  // Idem pour...
  //...Total Initial + Révisé
  AlimTotEdtEcartBudget(TotIR, TotBud);

  //...Total Edition
  if FGROUPCHOIXRUPT.ItemIndex<>2 then begin
    // Pour Total Edition : En rupture le total est fait sur le PrintGroup
    AlimTotEdtEcartBudget(TotEdt, TotBud);
  end;

  //...Rupture : Nature libre
  // Revise
  TotLibre[0]:= Arrondi(TotLibre[0]+TotBud[0][0].totDebit, Decimale);
  TotLibre[1]:= Arrondi(TotLibre[1]+TotBud[0][0].totCredit, Decimale);
  TotLibre[8]:= Arrondi(TotLibre[8]+TotBud[0][1].totDebit, Decimale);
  TotLibre[9]:= Arrondi(TotLibre[9]+TotBud[0][1].totCredit, Decimale);
  TotLibre[16]:= Arrondi(TotLibre[16]+TotBud[0][2].totDebit, Decimale);
  TotLibre[17]:= Arrondi(TotLibre[17]+TotBud[0][2].totCredit, Decimale);
  // Revise
  TotLibre[20]:= Arrondi(TotLibre[20]+TotBud[1][0].totDebit, Decimale);
  TotLibre[21]:= Arrondi(TotLibre[21]+TotBud[1][0].totCredit, Decimale);
  TotLibre[28]:= Arrondi(TotLibre[28]+TotBud[1][1].totDebit, Decimale);
  TotLibre[29]:= Arrondi(TotLibre[29]+TotBud[1][1].totCredit, Decimale);
  TotLibre[36]:= Arrondi(TotLibre[36]+TotBud[1][2].totDebit, Decimale);
  TotLibre[37]:= Arrondi(TotLibre[37]+TotBud[1][2].totCredit, Decimale);

  // Calcul du Réalisé pour...
  // ...La Période
  // Mvt init
  LeCompte := 'CBG'+ GetControlText('FJOURNAL') + szRub;
  LExo := QUELEXODTBud(StrToDate(FDateCompta1));
  if (Lexo<>QUELEXODTBud(StrToDate(FDateCompta2))) then LExo:='';
  FinTemp := FinDeMois(StrToDate(FDateCompta2));
  dtDate1 := StrToDate(FDateCompta1);
  GetCumul('RUBREA', LeCompte, LeCompte, 'N-', GetControlText('FETAB'), 'EUR', LExo, dtDate1, FinTemp, False, True, Nil, TD, False);

  RealPer[0].TotDebit:=TD[3];
  RealPer[0].TotCredit:=TD[2];
  EcarPer[0].TotDebit := Arrondi((RealPer[0].TotDebit-TotGen[0][0].TotDebit), Decimale);
  EcarPer[0].TotCredit := Arrondi((RealPer[0].TotCredit-TotGen[0][0].TotCredit), Decimale);

  Reevaluation(RealPer[0].TotDebit, RealPer[0].TotCredit, Resol, Taux);
  Reevaluation(EcarPer[0].TotDebit, EcarPer[0].TotCredit, Resol, Taux);

  dbl := T3.GetValue('CEC_INIGENCUM1'); T3.PutValue('CEC_INIGENCUM1', dbl + GetSolde(TotGen[0][0].TotDebit, TotGen[0][0].TotCredit, Decimale, szSoldeFormate));
  dbl := T3.GetValue('CEC_INIGENCUM2'); T3.PutValue('CEC_INIGENCUM2', dbl + GetSolde(RealPer[0].TotDebit, RealPer[0].TotCredit, Decimale, szSoldeFormate));
  dbl := T3.GetValue('CEC_INIGENCUM3'); T3.PutValue('CEC_INIGENCUM3', dbl + GetSolde(EcarPer[0].TotDebit, EcarPer[0].TotCredit, Decimale, szSoldeFormate));
  if (T3.GetValue('CEC_INIGENCUM1')<>0) and (T3.GetValue('CEC_INIGENCUM3')<>0) then begin
    Pourcentage := ((EcarPer[0].TotDebit-EcarPer[0].TotCredit)*100)/(TotGen[0][0].TotDebit-TotGen[0][0].TotCredit);
    Pourcentage := Abs(Pourcentage);
    if T3.GetValue('CEC_INIGENCUM3')<0 then Pourcentage := -1*Pourcentage;
    end
  else Pourcentage := 0;
  T3.PutValue('CEC_INIGENCUM4', AfficheMontant('#,##0', '', Pourcentage , False));
  // Mvt Revi
  RealPer[1].TotDebit := RealPer[0].TotDebit;
  RealPer[1].TotCredit := RealPer[0].TotCredit;
  EcarPer[1].TotDebit := Arrondi((RealPer[1].TotDebit-TotGen[1][0].TotDebit), Decimale);
  EcarPer[1].TotCredit := Arrondi((RealPer[1].TotCredit-TotGen[1][0].TotCredit), Decimale);

  dbl := T3.GetValue('CEC_REVGENCUM1'); T3.PutValue('CEC_REVGENCUM1', dbl + GetSolde(TotGen[1][0].TotDebit, TotGen[1][0].TotCredit, Decimale, szSoldeFormate));
  dbl := T3.GetValue('CEC_REVGENCUM2'); T3.PutValue('CEC_REVGENCUM2', dbl + GetSolde(RealPer[1].TotDebit, RealPer[1].TotCredit, Decimale, szSoldeFormate));
  dbl := T3.GetValue('CEC_REVGENCUM3'); T3.PutValue('CEC_REVGENCUM3', dbl + GetSolde(EcarPer[1].TotDebit,EcarPer[1].TotCredit, Decimale, szSoldeFormate));
  if (T3.GetValue('CEC_REVGENCUM1')<>0) and (T3.GetValue('CEC_REVGENCUM3')<>0) then begin
    Pourcentage := ((EcarPer[1].TotDebit-EcarPer[1].TotCredit)*100)/(TotGen[1][0].TotDebit-TotGen[1][0].TotCredit);
    Pourcentage := Abs(Pourcentage);
    if T3.GetValue('CEC_REVGENCUM3')<0 Then Pourcentage := -1*Pourcentage;
    end
  else Pourcentage := 0;
  T3.PutValue('CEC_REVGENCUM4', Pourcentage);

  // Report : réalisé
  MReport[2].TotDebit := Arrondi(RealPer[0].TotDebit, Decimale);
  MReport[2].TotCredit := Arrondi(RealPer[0].TotCredit, Decimale);

  // ...Le Cumulé
  LExo:=QUELEXODTBud(dtDatDebJal);
  if (Lexo<>QUELEXODTBud(StrToDate(FDateCompta2))) then LExo := '';
  FinTemp := FinDeMois(StrToDate(FDateCompta2));
  GetCumul('RUBREA', LeCompte, LeCompte, 'N-', GetControlText('FETAB'), 'EUR', LExo, dtDatDebJal, FinTemp, False, True, Nil, TD, False);

  RealCum[0].TotDebit := TD[3];
  RealCum[0].TotCredit := TD[2];
  EcarCum[0].TotDebit := Arrondi((RealCum[0].TotDebit-TotGen[0][1].TotDebit), Decimale);
  EcarCum[0].TotCredit := Arrondi((RealCum[0].TotCredit-TotGen[0][1].TotCredit), Decimale);

  Reevaluation(RealCum[0].TotDebit, RealCum[0].TotCredit, Resol, Taux);
  Reevaluation(EcarCum[0].TotDebit, EcarCum[0].TotCredit, Resol, Taux);

  dbl := T3.GetValue('CEC_INIGENCUM5'); T3.PutValue('CEC_INIGENCUM5', dbl + GetSolde(TotGen[0][1].TotDebit, TotGen[0][1].TotCredit, Decimale, szSoldeFormate));
  dbl := T3.GetValue('CEC_INIGENCUM6'); T3.PutValue('CEC_INIGENCUM6', dbl + GetSolde(RealCum[0].TotDebit, RealCum[0].TotCredit, Decimale, szSoldeFormate));
  dbl := T3.GetValue('CEC_INIGENCUM7'); T3.PutValue('CEC_INIGENCUM7', dbl + GetSolde(EcarCum[0].TotDebit, EcarCum[0].TotCredit, Decimale, szSoldeFormate));
  if (T3.GetValue('CEC_INIGENCUM5')<>0) and (T3.GetValue('CEC_INIGENCUM7')<>0) then begin
    Pourcentage := ((EcarCum[0].TotDebit-EcarCum[0].TotCredit)*100)/(TotGen[0][1].TotDebit-TotGen[0][1].TotCredit);
    Pourcentage := Abs(Pourcentage);
    if T3.GetValue('CEC_INIGENCUM7')<0 then Pourcentage := -1*Pourcentage;
    end
  else Pourcentage := 0;
  T3.PutValue('CEC_INIGENCUM8', Pourcentage);
  // Revi
  RealCum[1].TotDebit := RealCum[0].TotDebit;
  RealCum[1].TotCredit := RealCum[0].TotCredit;
  EcarCum[1].TotDebit := Arrondi((RealCum[0].TotDebit-TotGen[1][1].TotDebit), Decimale);
  EcarCum[1].TotCredit := Arrondi((RealCum[0].TotCredit-TotGen[1][1].TotCredit), Decimale);

  dbl := T3.GetValue('CEC_REVGENCUM5'); T3.PutValue('CEC_REVGENCUM5', dbl + GetSolde(TotGen[1][1].TotDebit, TotGen[1][1].TotCredit, Decimale, szSoldeFormate));
  dbl := T3.GetValue('CEC_REVGENCUM6'); T3.PutValue('CEC_REVGENCUM6', dbl + GetSolde(RealCum[1].TotDebit, RealCum[1].TotCredit, Decimale, szSoldeFormate));
  dbl := T3.GetValue('CEC_REVGENCUM7'); T3.PutValue('CEC_REVGENCUM7', dbl + GetSolde(EcarCum[1].TotDebit,EcarCum[1].TotCredit, Decimale, szSoldeFormate));
  if (T3.GetValue('CEC_REVGENCUM5')<>0) and (T3.GetValue('CEC_REVGENCUM7')<>0) then begin
    Pourcentage := ((EcarCum[1].TotDebit-EcarCum[1].TotCredit)*100)/(TotGen[1][1].TotDebit-TotGen[1][1].TotCredit);
    Pourcentage := Abs(Pourcentage);
    if T3.GetValue('CEC_REVGENCUM7')<0 then Pourcentage := -1*Pourcentage;
    end
  else Pourcentage := 0;
  T3.PutValue('CEC_REVGENCUM8', Pourcentage);

  // Report : réalisé
  MReport[6].TotDebit := Arrondi(RealCum[0].TotDebit, Decimale);
  MReport[6].TotCredit := Arrondi(RealCum[0].TotCredit, Decimale);
  //  ... L' Annuel
  dbl := T3.GetValue('CEC_INIGENCUM9'); T3.PutValue('CEC_INIGENCUM9', dbl + GetSolde(TotGen[0][2].TotDebit, TotGen[0][2].TotCredit, Decimale, szSoldeFormate));
  dbl := T3.GetValue('CEC_REVGENCUM9'); T3.PutValue('CEC_REVGENCUM9', dbl + GetSolde(TotGen[1][2].TotDebit, TotGen[1][2].TotCredit, Decimale, szSoldeFormate));

  if (TobFils=nil) then begin
    T3.InsertDB(nil);
    T3.Free;
  end;
end;

procedure TOF_BALBUDECSEC.CalculBud(i: Byte; var Tot1: TMontTot; NatB: String; D, C: Double; Dec: Byte);
begin
  if NatB='INI' then begin
    Tot1[0][i].TotDebit:=  Arrondi(Tot1[0][i].TotDebit+D, Dec) ;
    Tot1[0][i].TotCredit:= Arrondi(Tot1[0][i].TotCredit+C, Dec) ;
    Tot1[0][12].TotDebit:= Arrondi(Tot1[0][12].TotDebit+D, Dec) ;
    Tot1[0][12].TotCredit:=Arrondi(Tot1[0][12].TotCredit+C, Dec) ;
    end
  else if Copy(NatB,1,2)='DM' then begin
    Tot1[1][i].TotDebit:=  Arrondi(Tot1[1][i].TotDebit+D, Dec) ;
    Tot1[1][i].TotCredit:= Arrondi(Tot1[1][i].TotCredit+C, Dec) ;
    Tot1[1][12].TotDebit:= Arrondi(Tot1[1][12].TotDebit+D, Dec) ;
    Tot1[1][12].TotCredit:=Arrondi(Tot1[1][12].TotCredit+C, Dec) ;
  end;
end;

procedure TOF_BALBUDECSEC.InitCalcBud(T: Tob; var Tot1: TMontTot; Resol : String; Taux : Double);
var
  DateCompta : TDateTime;
  LaNatBud : String;
  DEBIT, CREDIT : Double;
  i : Integer;
begin
  Fillchar(Tot1,SizeOf(Tot1),#0);

  for i := 0 to T.Detail.Count-1 do begin
    DateCompta := T.Detail[i].GetDateTime('BE_DATECOMPTABLE');
    LaNatBud := T.Detail[i].GetString('BE_NATUREBUD');
    DEBIT := T.Detail[i].GetDouble('DEBIT');
    CREDIT := T.Detail[i].GetDouble('CREDIT');

    Reevaluation(DEBIT,CREDIT, Resol, Taux);

    if (DateCompta>= StrToDate(FDateCompta1)) and (DateCompta < StrToDate(FDateCompta2)) then begin // Calculs Période
      CalculBud(0,Tot1, LaNatBud, DEBIT, CREDIT, Decimale)
    end;
    if (DateCompta>=dtDatDebJal) and (DateCompta< StrToDate(FDateCompta2)) then begin // Calculs Cumulé
      CalculBud(1,Tot1, LaNatBud, DEBIT, CREDIT, Decimale)
    end;
    if (DateCompta>=dtDatDebJal) and (DateCompta<dtDatFinJal) then begin // Calculs Annuel
      CalculBud(2,Tot1, LaNatBud, DEBIT, CREDIT, Decimale)
    end;
  end;
end;

function TOF_BALBUDECSEC.UpdateRupture(szLibre1, szLibre2 : String) : String;
var
  i : Integer;
  St, StListe : String;
begin
  SetControlText('FLIBRE1', szLibre1);
  SetControlText('FLIBRE2', szLibre2);
  StListe := szLibre1;
  i := 0;
  Result := '';
  while StListe<>'' do begin
    St := ReadTokenSt(StListe);
    if St='' then begin
      inc(i);
      continue;
    end;
    if (St[1]<>'#') and (St[1]<>'-') then Result := Result + 'D0' + IntToStr(i) + ';';
    Inc(i);
  end;

end;

procedure TOF_BALBUDECSEC.AlimTotEdtEcartBudget(var LeTotal: TabTot12; var SousTot: TMontTot);
begin
  LeTotal[0].TotDebit:= Arrondi(LeTotal[0].TotDebit+(SousTot[0][0].totDebit+SousTot[1][0].totDebit), Decimale);
  LeTotal[0].TotCredit:= Arrondi(LeTotal[0].TotCredit+(SousTot[0][0].totCredit+SousTot[1][0].totCredit), Decimale);
  LeTotal[4].TotDebit:= Arrondi(LeTotal[4].TotDebit+(SousTot[0][1].totDebit+SousTot[1][1].totDebit), Decimale);
  LeTotal[4].TotCredit:= Arrondi(LeTotal[4].TotCredit+(SousTot[0][1].totCredit+SousTot[1][1].totCredit), Decimale);
  LeTotal[8].TotDebit:= Arrondi(LeTotal[8].TotDebit+(SousTot[0][2].totDebit+SousTot[1][2].totDebit), Decimale);
  LeTotal[8].TotCredit:= Arrondi(LeTotal[8].TotCredit+(SousTot[0][2].totCredit+SousTot[1][2].totCredit), Decimale);
end;

function TOF_BALBUDECSEC.GetSolde(TD, TC: DOUBLE; Decim: Integer; FormMont: String): Double;
begin
  Result := 0;

  if (TD=TC) then begin
    Result := TD-TC;
    end
  else if Abs(TD)>=Abs(TC) then begin
    Result := TD-TC;
    if (Result<0) And (FormMont='PC') then begin
      Result := Result*(-1);
    end;
    end
  else if Abs(TD)<Abs(TC) then begin
    Result := TC-TD;
    if (Result<0) And (FormMont='PD') Then begin
      Result := Result*(-1);
    end;
  end;
end;

Initialization
  registerclasses ( [ TOF_BALBUDECSEC ] ) ;
end.
