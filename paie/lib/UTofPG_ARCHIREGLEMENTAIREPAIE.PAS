{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 08/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ARCHIREGLEMENTAIREPAIE ()
Mots clefs ... : TOF;ARCHIREGLEMENTAIREPAIE
*****************************************************************}
{
PT1 17/12/2007 GGU V_8 Création de la vue en cas de problème
PT2 19/12/2007 GGU V_80 Compatibilité Delphi 5 / Version 7 pour reprise de l'outils en V7
}
Unit UTofPG_ARCHIREGLEMENTAIREPAIE ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
     FE_Main,  //PT40 (AglLanceFiche)
{$else}
     eMul,
     MainEAGL, //PT40 (AglLanceFiche)
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF, HTB97, Hqry;

Type
  TOF_ARCHIREGLEMENTAIREPAIE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure OnBOuvrirClick(Sender : TObject);
  private
    AnalyseArchitecture : Boolean;
  end ;

  Procedure LoadTablettePGDOSSIER;

Implementation

Procedure LoadTablettePGDOSSIER;
var
  TempTob, TempTobFille : Tob;
  Index : Integer;
begin
  TempTob := Tob.Create('Liste des dossiers', nil, -1);
  TempTobFille := Tob.Create('Dossier Ceg ou Std', TempTob, -1);
  TempTobFille.AddChampSupValeur('DOS_NODOSSIER', '000000');
  TempTobFille.AddChampSupValeur('DOS_LIBELLE', 'Cegid ou standard');
  TempTob.LoadDetailFromSQL('SELECT DOS_NODOSSIER, DOS_LIBELLE FROM DOSSIER', True);
  ExecuteSQL('DELETE FROM COMMUN WHERE CO_TYPE="PGD"');
  for Index := 0 to TempTob.Detail.Count -1 do
  begin
    ExecuteSQL('insert into COMMUN (co_type, co_code, co_libelle, co_abrege, co_libre) '
              +'values ("PGD", "'+IntToStr(Index)+'", "'+TempTob.Detail[index].GetString('DOS_LIBELLE')+'", "", "'+TempTob.Detail[index].GetString('DOS_NODOSSIER')+'")')
  end;
  FreeAndNil(TempTob);
  AvertirTable('PGDOSSIER'); 
end;

procedure TOF_ARCHIREGLEMENTAIREPAIE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_ARCHIREGLEMENTAIREPAIE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_ARCHIREGLEMENTAIREPAIE.OnUpdate ;
begin
  Inherited ;
  if not AnalyseArchitecture then
    SetControlText('XX_WHERE', GetControlText('XX_WHERE') + ' AND ARP_NATUREARCHI NOT IN ("PRO", "CUM", "ELT", "TDY") AND ((ARP_PREDEFINI <> "DOS") or (ARP_NODOSSIER = "'+V_PGI.NoDossier+'"))' ) ;
end ;

procedure TOF_ARCHIREGLEMENTAIREPAIE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_ARCHIREGLEMENTAIREPAIE.OnArgument (S : String ) ;
var
  Q : TQuery;                         //PT1
  TobDevues : Tob;                    //PT1
  PGARCHIREGLEMENTAIRE_SQL : String;  //PT1
begin
  //Debut PT1
  Q := OpenSQL('SELECT count(*) from devues where dv_nomvue = "PGARCHIREGLEMENTAIRE"', True);
  If Q.Fields[0].AsInteger = 0 then
  begin
    { La vue n'existe pas, on doit la créer }
    BEGINTRANS;
    try
      TobDevues := Tob.Create('DEVUES', nil, -1);
      TobDevues.PutValue('DV_NOMVUE', 'PGARCHIREGLEMENTAIRE');
      TobDevues.PutValue('DV_LIBELLE', 'Liste des éléments analysables');
      PGARCHIREGLEMENTAIRE_SQL := 'select "VAR" as ARP_NATUREARCHI, pva_variable as ARP_CODE, pva_libelle as ARP_LIBELLE, pva_abrege as ARP_ABREGE, pva_themevar as ARP_THEME, pva_predefini as ARP_PREDEFINI, pva_nodossier as ARP_NODOSSIER from variablepaie '
{$IFNDEF VER130} //PT2
                                + 'where (pva_typevariable = "PAI") '
{$ENDIF}
                                + 'union '
                                + 'select "REM" as ARP_NATUREARCHI, prm_rubrique as ARP_CODE, prm_libelle as ARP_LIBELLE, prm_abrege as ARP_ABREGE, prm_themerem as ARP_THEME, prm_predefini as ARP_PREDEFINI, prm_nodossier as ARP_NODOSSIER from remuneration '
                                + 'union '
                                + 'select "COT" as ARP_NATUREARCHI, pct_rubrique as ARP_CODE, pct_libelle as ARP_LIBELLE, pct_abrege as ARP_ABREGE, pct_themecot as ARP_THEME, pct_predefini as ARP_PREDEFINI, pct_nodossier as ARP_NODOSSIER from cotisation where (PCT_NATURERUB = "COT") '
                                + 'union '
                                + 'select "CUM" as ARP_NATUREARCHI, pcl_cumulpaie as ARP_CODE, pcl_libelle as ARP_LIBELLE, pcl_abrege as ARP_ABREGE, pcl_themecum as ARP_THEME, pcl_predefini as ARP_PREDEFINI, pcl_nodossier as ARP_NODOSSIER from cumulpaie '
                                + 'union '
                                + 'SELECT "PRO" as ARP_NATUREARCHI, PPI_PROFIL as ARP_CODE, PPI_LIBELLE as ARP_LIBELLE, PPI_TYPEPROFIL as ARP_ABREGE, PPI_THEMEPROFIL as ARP_THEME, PPI_PREDEFINI as ARP_PREDEFINI, PPI_NODOSSIER as ARP_NODOSSIER FROM PROFILPAIE '
{$IFNDEF VER130} //PT2
                                + 'union '
                                + 'SELECT "TDY" as ARP_NATUREARCHI, PTE_CODTABL as ARP_CODE, PTE_LIBELLE as ARP_LIBELLE, PTE_NATURETABLE as ARP_ABREGE, PTE_NIVSAIS as ARP_THEME, PTE_PREDEFINI as ARP_PREDEFINI, PTE_NODOSSIER as ARP_NODOSSIER FROM TABLEDIMENT '
{$ENDIF}
                                + 'union '
                                + 'SELECT distinct "ELT" as ARP_NATUREARCHI, PEL_CODEELT as ARP_CODE, PEL_LIBELLE as ARP_LIBELLE, PEL_ABREGE as ARP_ABREGE, PEL_THEMEELT as ARP_THEME, PEL_PREDEFINI as ARP_PREDEFINI, PEL_NODOSSIER as ARP_NODOSSIER FROM ELTNATIONAUX where (1=1)';
      TobDevues.PutValue('DV_SQL', PGARCHIREGLEMENTAIRE_SQL);
      TobDevues.PutValue('DV_DOMAINE', 'P');
      TobDevues.PutValue('DV_NUMVERSION', '100');
      TobDevues.InsertOrUpdateDB();
      ExecuteSQL('CREATE VIEW PGARCHIREGLEMENTAIRE (ARP_NATUREARCHI, ARP_CODE, ARP_LIBELLE, ARP_ABREGE, ARP_THEME, ARP_PREDEFINI, ARP_NODOSSIER) AS '+PGARCHIREGLEMENTAIRE_SQL);
      COMMITTRANS;
    except
      ROLLBACK;
    end;
    FreeAndNil(TobDevues);
  end;
  //Fin PT1
  Inherited ;
  if READTOKENST(S) = 'ARCHITECTURE' then
    AnalyseArchitecture := True //Mode Analyse de l'architecture ou Diagnostic
  else
    AnalyseArchitecture := False;
  if not AnalyseArchitecture then
  begin
    Ecran.Caption := 'Diagnostic des rubriques de paie - Sélection des données à analyser';
    UpdateCaption(Ecran);
    (GetControl('ARP_NATUREARCHI') as THMultiValComboBox).Plus := ' AND CO_CODE NOT IN ("PRO", "CUM", "ELT", "TDY")';
    SetControlText('XX_WHERE', GetControlText('XX_WHERE') + ' AND ARP_NATUREARCHI NOT IN ("PRO", "CUM", "ELT", "TDY") AND (ARP_NODOSSIER in ("000000", "'+V_PGI.NoDossier+'"))' ) ;
    SetControlVisible('LBLNODOSSIER', False);
    SetControlVisible('ARP_NODOSSIER', False);
  end;
  (GetControl('BOuvrir') as TToolBarButton97).OnClick := OnBOuvrirClick;  
  // La tablette PGDOSSIER ne doit pas être chargé par la fiche avant d'être chargé,
  // donc il ne faut pas qu'elle soit affectée au champs sur la fiche.
  LoadTablettePGDOSSIER;
  if Assigned(GetControl('ARP_NODOSSIER')) then
    (GetControl('ARP_NODOSSIER') as THMultiValComboBox).DataType := 'PGDOSSIER';
end ;

procedure TOF_ARCHIREGLEMENTAIREPAIE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_ARCHIREGLEMENTAIREPAIE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_ARCHIREGLEMENTAIREPAIE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_ARCHIREGLEMENTAIREPAIE.OnBOuvrirClick(Sender: TObject);
var
  Q_Mul: THQuery;
  StSelection : String;
  i : Integer;
  TempNature, TempCode, TempPredefini, TempDossier : String;
begin
  { Lecture des éléments sélectionnés et construction de la chaine de paramètres }
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  if Q_Mul = nil then exit;
  StSelection := '';
  If (TFMul(Ecran).FListe.nbSelected > 0) Then
  Begin
    for i := 0 to TFMul(Ecran).FListe.NbSelected - 1 do
    begin
      {$IFDEF EAGLCLIENT}
      TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
      {$ENDIF}
      TFMul(Ecran).FListe.GotoLeBookmark(i);
      TempNature    := TFmul(Ecran).Q.FindField('ARP_NATUREARCHI').AsString;
      TempCode      := TFmul(Ecran).Q.FindField('ARP_CODE').AsString;
      TempPredefini := TFmul(Ecran).Q.FindField('ARP_PREDEFINI').AsString;
      TempDossier   := TFmul(Ecran).Q.FindField('ARP_NODOSSIER').AsString;
      StSelection := StSelection + TempNature+'|'+TempCode+'|'+TempPredefini+'|'+TempDossier+';';
    end;
  end else if TFMul(Ecran).FListe.AllSelected then
  begin
    TFMul(Ecran).Q.First;
    While not TFMul(Ecran).Q.Eof do
    begin
      TempNature    := TFmul(Ecran).Q.FindField('ARP_NATUREARCHI').AsString;
      TempCode      := TFmul(Ecran).Q.FindField('ARP_CODE').AsString;
      TempPredefini := TFmul(Ecran).Q.FindField('ARP_PREDEFINI').AsString;
      TempDossier   := TFmul(Ecran).Q.FindField('ARP_NODOSSIER').AsString;
      StSelection := StSelection + TempNature+'|'+TempCode+'|'+TempPredefini+'|'+TempDossier+';';
      TFMul(Ecran).Q.Next;
    end;
  end;
  { Lancement de l'analyse }
  if AnalyseArchitecture then
    AGLLanceFiche('PAY', 'PGANARCHIREGLPAIE', '', '', StSelection)
  else
    AGLLanceFiche('PAY', 'PGDIAGNOSTIC', '', '', StSelection);
end;

Initialization
  registerclasses ( [ TOF_ARCHIREGLEMENTAIREPAIE ] ) ;
end.
