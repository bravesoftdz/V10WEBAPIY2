{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 09/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DIAGNOSTIQUE ()
Mots clefs ... : TOF;DIAGNOSTIC
*****************************************************************}
Unit UTofPG_DIAGNOSTIQUE ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     PGTreeTobFrame,
     UTob, HTB97;

Type
  TOF_DIAGNOSTIC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TreeTobFrame : TFFrameTreeTob;
  public
  end ;


  Function ChargeToutesLesRubrique(DateDeb, DateFin : TDateTime) : Tob;
  Procedure RecupTobSalarieVide(DateDebut, DateFin : TDateTime);
  Function Diagnostic(ElementsADiagnostiquer : String): Tob;


Implementation

Uses
  SaisBul, P5Util, uPaieEtabCompl, P5Def, ed_Tools, hpanel;

Function ChargeToutesLesRubrique(DateDeb, DateFin : TDateTime) : Tob;
begin
  result := Tob.Create('Rubriques', nil, -1);
  { Chargement des rémunérations }
  result.LoadDetailFromSQL(' SELECT "" as PHB_ETABLISSEMENT, "" as PHB_SALARIE, "'+USDATETIME(DateDeb)+'" as PHB_DATEDEBUT, '
                          +' "'+USDATETIME(DateFin)+'" as PHB_DATEFIN, PRM_NATURERUB as PHB_NATURERUB, '
                          +' PRM_RUBRIQUE as PHB_RUBRIQUE, PRM_LIBELLE as PHB_LIBELLE, PRM_IMPRIMABLE as PHB_IMPRIMABLE, '
                          +' 0 as PHB_BASEREM, 0 as PHB_TAUXREM, 0 as PHB_COEFFREM, 0 as PHB_MTREM, '
                          +' PRM_BASEIMPRIMABLE as PHB_BASEREMIMPRIM, PRM_TAUXIMPRIMABLE as PHB_TAUXREMIMPRIM, '
                          +' PRM_COEFFIMPRIM as PHB_COEFFREMIMPRIM, PRM_ORDREETAT as PHB_ORDREETAT, '
                          +' PRM_ORDREETAT as PHB_OMTSALARIAL, PRM_SENSBUL as PHB_SENSBUL, "..."  as PHB_COTREGUL, '
                          +' 0 as PHB_BASECOT, 0 as PHB_TAUXSALARIAL, 0 as PHB_TAUXPATRONAL, 0 as PHB_MTSALARIAL, '
                          +' 0 as PHB_MTPATRONAL, "...." as PHB_ORGANISME, "-" as PHB_BASECOTIMPRIM, '
                          +' "-" as PHB_TAUXSALIMPRIM,  "-" as PHB_TAUXPATIMPRIM,  '
                          +' "" as PHB_TRAVAILN1, '
                          +' "" as PHB_TRAVAILN2, '
                          +' "" as PHB_TRAVAILN3, '
                          +' "" as PHB_TRAVAILN4, '
                          +' "" as PHB_CODESTAT, '
                          +' "" as PHB_LIBREPCMB1, '
                          +' "" as PHB_LIBREPCMB2, '
                          +' "" as PHB_LIBREPCMB3, '
                          +' "" as PHB_LIBREPCMB4, '
                          +' "-" as PHB_CONFIDENTIEL, '
                          +' 0 as PHB_TAUXAT, 0 as PHB_PLAFOND, 0 as PHB_PLAFOND1, 0 as PHB_PLAFOND2, '
                          +' 0 as PHB_PLAFOND3, "" as PHB_ORIGINELIGNE, "" as PHB_ORIGINEINFO, 0 as PHB_TRANCHE1, '
                          +' 0 as PHB_TRANCHE2, 0 as PHB_TRANCHE3, '
                          +' "PRO" as PHB_CONSERVATION '
                          +' FROM REMUNERATION');
  { Chargement des cotisations }
  result.LoadDetailFromSQL(' SELECT "" as PHB_ETABLISSEMENT, "" as PHB_SALARIE, "'+USDATETIME(DateDeb)+'" as PHB_DATEDEBUT, '
                          +' "'+USDATETIME(DateFin)+'" as PHB_DATEFIN, PCT_NATURERUB as PHB_NATURERUB, '
                          +' PCT_RUBRIQUE as PHB_RUBRIQUE, PCT_LIBELLE as PHB_LIBELLE, PCT_IMPRIMABLE as PHB_IMPRIMABLE, '
                          +' 0 as PHB_BASEREM, 0 as PHB_TAUXREM, 0 as PHB_COEFFREM, 0 as PHB_MTREM, '
                          +' "-" as PHB_BASEREMIMPRIM, "-" as PHB_TAUXREMIMPRIM, '
                          +' "-" as PHB_COEFFREMIMPRIM, 3 as PHB_ORDREETAT, '
                          +' PCT_ORDREETAT as PHB_OMTSALARIAL, "P" as PHB_SENSBUL, "..."  as PHB_COTREGUL, '
                          +' 0 as PHB_BASECOT, 0 as PHB_TAUXSALARIAL, 0 as PHB_TAUXPATRONAL, 0 as PHB_MTSALARIAL, '
                          +' 0 as PHB_MTPATRONAL, PCT_ORGANISME as PHB_ORGANISME, PCT_BASEIMP as PHB_BASECOTIMPRIM, '
                          +' PCT_TXSALIMP as PHB_TAUXSALIMPRIM,  PCT_TXPATIMP as PHB_TAUXPATIMPRIM,  '
                          +' "" as PHB_TRAVAILN1, '
                          +' "" as PHB_TRAVAILN2, '
                          +' "" as PHB_TRAVAILN3, '
                          +' "" as PHB_TRAVAILN4, '
                          +' "" as PHB_CODESTAT, '
                          +' "" as PHB_LIBREPCMB1, '
                          +' "" as PHB_LIBREPCMB2, '
                          +' "" as PHB_LIBREPCMB3, '
                          +' "" as PHB_LIBREPCMB4, '
                          +' "-" as PHB_CONFIDENTIEL, '
                          +' 0 as PHB_TAUXAT, 0 as PHB_PLAFOND, 0 as PHB_PLAFOND1, 0 as PHB_PLAFOND2, '
                          +' 0 as PHB_PLAFOND3, "" as PHB_ORIGINELIGNE, "" as PHB_ORIGINEINFO, 0 as PHB_TRANCHE1, '
                          +' 0 as PHB_TRANCHE2, 0 as PHB_TRANCHE3, '
                          +' "PRO" as PHB_CONSERVATION '
                          +' FROM COTISATION', True);
  if Assigned(result) then MemorisePhb(result.detail.Items[0]);
end;

Procedure RecupTobSalarieVide(DateDebut, DateFin : TDateTime);
begin
  if Assigned(GblTob_VarCp) then FreeAndNil(GblTob_VarCp);
  if Assigned(TOB_Salarie) then FreeAndNil(TOB_Salarie);
  TOB_Salarie := TOB.Create('SALARIES', nil, -1);
  MemorisePsa(TOB_Salarie);
end;

Function Diagnostic(ElementsADiagnostiquer : String): Tob;
var Rub, lib: string;
  Base, Taux, Coeff, Montant: Double;
  TxSal, TxPat, MtSal, MtPat: Double;
  AT: Boolean;
  TobAnalyse, TempTob : Tob;
  Tob_Rub : Tob;
  IndexTobAAnalyser, IndexTobNature : Integer;
  Param, Parametres : String;
  TempNature, TempCode, TempLibelle : String; //TempPredefini, TempDossier : String;
begin
  Parametres := ElementsADiagnostiquer;
  InitMoveProgressForm(nil, 'Diagnostic en cours', 'Diagnostic', 3, True, True);
  TobAnalyse := Tob.Create('Diagnostic du réglementaire', nil, -1);
  Try
    TobAnalyse.AddChampSupValeur('ARP_CODE', '');
    TobAnalyse.AddChampSupValeur('ARP_NATUREARCHI', '');
    TobAnalyse.AddChampSupValeur('ARP_MESSAGE', 'Diagnostic du réglementaire');
    TempTob := TOB.Create('&Variables', TobAnalyse, -1);
    TempTob.AddChampSupValeur('ARP_CODE', 'VAR');
    TempTob.AddChampSupValeur('ARP_NATUREARCHI', '');
    TempTob.AddChampSupValeur('ARP_MESSAGE', 'Variables');
    TempTob := TOB.Create('&Rémunérations', TobAnalyse, -1);
    TempTob.AddChampSupValeur('ARP_CODE', 'REM');
    TempTob.AddChampSupValeur('ARP_NATUREARCHI', '');
    TempTob.AddChampSupValeur('ARP_MESSAGE', 'Rémunérations');
    TempTob := TOB.Create('&Cotisations', TobAnalyse, -1);
    TempTob.AddChampSupValeur('ARP_CODE', 'COT');
    TempTob.AddChampSupValeur('ARP_NATUREARCHI', '');
    TempTob.AddChampSupValeur('ARP_MESSAGE', 'Cotisations');
    TempTob := TOB.Create('&Cumuls', TobAnalyse, -1);
    TempTob.AddChampSupValeur('ARP_CODE', 'CUM');
    TempTob.AddChampSupValeur('ARP_NATUREARCHI', '');
    TempTob.AddChampSupValeur('ARP_MESSAGE', 'Cumuls');
    MoveCurProgressForm('Lecture des données à analyser');
    Param := readtokenst(Parametres);
    While Param <> '' do
    begin
      TempNature := READTOKENPipe(Param, '|');
      TempCode := READTOKENPipe(Param, '|');
      //TempPredefini :=
      READTOKENPipe(Param, '|');
      //TempDossier :=
      READTOKENPipe(Param, '|');
      TempTob := TobAnalyse.FindFirst(['ARP_CODE'], [TempNature], False);
      if Assigned(TempTob) then
      begin
        TempTob := Tob.create('Analyse d''un élément ', TempTob, -1);
//        TempTob.AddChampSupValeur('ARP_NATURE', TempNature);
        TempTob.AddChampSupValeur('ARP_CODE', TempCode);
        TempTob.AddChampSupValeur('ARP_NATUREARCHI', TempNature);
        if TempNature = 'VAR' then
          TempLibelle := RechDom('PGVARIABLE', TempCode, False);
//          TempTob.AddChampSupValeur('ARP_LIBELLE', RechDom('PGVARIABLE', TempCode, False));
        if TempNature = 'REM' then
          TempLibelle := RechDom('PGREMUNERATION', TempCode, False);
//          TempTob.AddChampSupValeur('ARP_LIBELLE', RechDom('PGREMUNERATION', TempCode, False));
        if TempNature = 'COT' then
          TempLibelle := RechDom('PGCOTISATION', TempCode, False);
//          TempTob.AddChampSupValeur('ARP_LIBELLE', RechDom('PGCOTISATION', TempCode, False));
        if TempNature = 'CUM' then
          TempLibelle := RechDom('PGCUMULPAIE', TempCode, False);
//          TempTob.AddChampSupValeur('ARP_LIBELLE', RechDom('PGCUMULPAIE', TempCode, False));
//        TempTob.AddChampSupValeur('ARP_PREDEFINI', TempPredefini);
//        TempTob.AddChampSupValeur('ARP_NODOSSIER', TempDossier);
        TempTob.AddChampSupValeur('ARP_MESSAGE', '('+TempCode+') '+TempLibelle);
      end;
      Param := readtokenst(Parametres);
    end;
    MoveCurProgressForm('Initialisation des données de la paie');
    InitMemorise();
    InitLesTOBPaie;
    ChargeLesTOBPaie;
    initTOB_EtabCompl();
    if Assigned(TOB_ExerSocial) then FreeAndNil(TOB_ExerSocial);
    ChargeLesExercPaie(DateD, DateF);
    RecupTobSalarieVide(DateD, DateF);
    if Assigned(Tob_Rub) then FreeAndNil(Tob_Rub);
    Tob_Rub := ChargeToutesLesRubrique(DateD, DateF);
    MoveCurProgressForm('Analyse en cours...');
    for IndexTobNature := 0 to TobAnalyse.Detail.Count -1 do
    begin
      TempTob := TobAnalyse.Detail.Items[IndexTobNature];
      TempNature := TempTob.getString('ARP_CODE');
      for IndexTobAAnalyser := 0 to TobAnalyse.Detail.Items[IndexTobNature].Detail.Count -1 do
      begin
        TempTob := TobAnalyse.Detail.Items[IndexTobNature].Detail.Items[IndexTobAAnalyser];
        Rub := TempTob.getString('ARP_CODE');
        if      TempNature = 'VAR' then
        begin
          ValVariable(Rub, DateD, DateF, Tob_Rub, LogGetChildLevel(TempTob));
        end else if TempNature = 'REM' then
        begin
          EvalueRem(TOB_Rub, Rub, Base, Taux, Coeff, Montant, Lib, DateD, DateF, taConsultation, 0, LogGetChildLevel(TempTob))
        end else if TempNature = 'COT' then
        begin
          EvalueCot(Tob_Rub, rub, Base, TxSal, TxPat, MtSal, MtPat, Lib, DateD, DateF, taConsultation, AT, LogGetChildLevel(TempTob));
        end;
      end;
    end;
  finally
    FiniMoveProgressForm;
    VideLesTOBPaie(True);
    if Assigned(Tob_Rub) then FreeAndNil(Tob_Rub);
    if Assigned(TOB_Salarie) then FreeAndNil(TOB_Salarie);
    if Assigned(TOB_ExerSocial) then FreeAndNil(TOB_ExerSocial);
    result := TobAnalyse;
    InitMemorise();
  end;
end;


procedure TOF_DIAGNOSTIC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_DIAGNOSTIC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_DIAGNOSTIC.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_DIAGNOSTIC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_DIAGNOSTIC.OnArgument (S : String ) ;
var
  TobDiagnostic : Tob;
  LinkedCtrls: TTreeTobLinkedCtrls;
begin
  Inherited ;
//  TV := (GetControl('TVANALYSE') as TTreeView);
  TobDiagnostic := Diagnostic(S);
  { FrameTreeTob }
  LinkedCtrls := TTreeTobLinkedCtrls.Create();
  try
    LinkedCtrls.BtExpand     := TToolBarButton97(GetControl('BTEXPAND'));
//    LinkedCtrls.
//    LinkedCtrls.BtLoupe      := TToolBarButton97(GetControl('BTLOUPE'));
//    LinkedCtrls.BtParamListe := TToolBarButton97(GetControl('BPARAMLISTE'));
    LinkedCtrls.BtSearch     := TToolBarButton97(GetControl('BTSEARCH'));
//    LinkedCtrls.PmLoupe      := THPopupMenu(GetControl('PMLOUPE'));
    TreeTobFrame := TFFrameTreeTob.Create(Ecran, THPanel(GetControl('PNFRAMETREETOB')), TobDiagnostic, 'PGDIAGTREE', LinkedCtrls);
  finally
    LinkedCtrls.Free;
    FreeAndNil(TobDiagnostic);
  end;
//  TreeTobFrame.TreeGrid.OnDblClick := FListe_OnDblClick;

end ;

procedure TOF_DIAGNOSTIC.OnClose ;
begin
  Inherited ;
  if Assigned(TreeTobFrame) then FreeAndNil(TreeTobFrame);
end ;

procedure TOF_DIAGNOSTIC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_DIAGNOSTIC.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_DIAGNOSTIC ] ) ;
end.
