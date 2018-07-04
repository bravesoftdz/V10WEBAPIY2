{ Unité : Source TOM de la TABLE : CUMULS
--------------------------------------------------------------------------------------
  Version    |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
04.31.101.001  07/07/03  JP   Création de l'unité
04.31.101.001  25/07/03  JP   Rend invisibles les boutons Prev et Next dans le OnArgument
06.51.001.001  30/11/05  JP   FQ 17081 : CBBUDGET était toujours visible => Axe était toujours sur A1
06.51.001.001  01/12/05  JP   FQ 17079 : Suppression du Where dans la clause Plus
07.10.001.001  22/05/06  JP   FQ 17092 : Problème de gestion des @ dans les filtres
                              02/06/06 En attente de modification de l'AGL après la V7
07.00.001.009  30/05/06  JP   FQ 18005 : Manquait des '"' autour des dates dans FaitRequeteBudget
08.10.001.012  03/10/07  JP   FQ 21560 : gestion de la balance générale revue pour utiliser le standard
--------------------------------------------------------------------------------------}
Unit CPTOTRUB_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL, UtileAGL, eMul,
{$ELSE}
     FE_Main, db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     EdtREtat, Mul,
{$ENDIF}

{$IFDEF VER150}
   Variants,
{$ENDIF}

     Spin,
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ELSE}
     TCalcCum,
     {$ENDIF MODENT1}
     Dialogs, HXLSPAS,
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, Graphics, Ent1, UTob, HTB97,
     UTOF, CpteUtil, Menus, Grids, Windows, CalcOle, SaisUtil, ParamDat, HStatus,
     RUBRIQUE_TOM;

Type
  TCol = class
    LCol    : Integer;
    Libelle : string ;
  end;

  TRes = record
    Cpt : string;
    Deb : string;
    Cdt : string;
    Tot : string;
    Sdb : string;
    Scd : string;
    Mix : string;
  end;

  TOF_CPTOTRUB = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    FListe     : THGrid;
    FListe2    : THGrid;
    Rub1       : THEdit;
    Rub2       : THEdit;
    PD1        : THEdit;
    PD1_        : THEdit;
    Famrub     : THValComboBox;
    TypRub     : THValComboBox;
    Exo        : THValComboBox;
    RubAno     : THValComboBox;
    Dev        : THValComboBox;
    Etab       : THValComboBox;
    QualPie    : THValComboBox;
    CbAxe      : THValComboBox;
    CbBudget   : THValComboBox;
    LibGene    : THValComboBox;
    SupCombo   : THValComboBox;
    BDetail    : TToolbarButton97;
    BFerDetail : TToolbarButton97;
    BStop      : TToolbarButton97;
    SM         : THNumEdit;

    BSec       : TMenuItem;
    BCumP      : TMenuItem;
    BGen       : TMenuItem;
    BRub       : TMenuItem;
    FindDialog : TFindDialog;

    OldBrush : TBrush;
    OldPen   : TPen;
    MsgBox   : THMsgBox;
    TobRub   : TOB;

    FiltreEnCours,
    FirstFind,
    Stop ,
    Budget      : Boolean;

    MemoFam,
    MemoTyp,
    Symb,
    Frub,
    TyRub,
    LExo,
    LRub1,
    LRub2,
    LQuali,
    LDev,
    LEtab,
    LRubAno : string;

    MemoLig,
    Dec,
    CurGrid : Integer ;

    D1, D2  : TDateTime ;
    LeFb    : TFichierBase ;

    FormKeyDown  : TKeyEvent;
    procedure OnKeyDownEcran (Sender : TObject; var Key : Word; Shift : TShiftState); virtual;

    (* 02/06/06 En attente de modification de l'AGL après la V7
    {JP 22/05/06 : FQ 17092 : problème des arobases dans la gestions des filtres}
    procedure FiltreChange         (Sender : TObject);
    function  DecrypteFiltre       (Champ : string) : string;
    procedure AfterSelectFiltre;
    {JP 22/05/06 : FQ 17092 : FIN}
    *)
    procedure DevChange            (Sender : TObject);
    procedure ExoChange            (Sender : TObject);
    procedure TypRubChange         (Sender : TObject);
    procedure FamrubChange         (Sender : TObject);
    procedure CbAxeChange          (Sender : TObject);
    procedure BDetailClick         (Sender : TObject);
    procedure BFerDetailClick      (Sender : TObject);
    procedure FListeDblClick       (Sender : TObject);
    procedure FListeTopLeftChanged (Sender : TObject);
    procedure FListe2TopLeftChanged(Sender : TObject);
    procedure BStopClick           (Sender : TObject);
    procedure FListeEnter          (Sender : TObject);
    procedure FListe2Enter         (Sender : TObject);
    procedure CbBudgetChange       (Sender : TObject);
    procedure FListe2DblClick      (Sender : TObject);
    procedure BCumPClick           (Sender : TObject);
    procedure BRubClick            (Sender : TObject);
    procedure BGenClick            (Sender : TObject);
    procedure BSecClick            (Sender : TObject);
    procedure BChercheClick        (Sender : TObject);
    procedure BImprimerClick       (Sender : TObject);
    procedure bExportClick         (Sender : TObject);
    procedure BRechercherClick     (Sender : TObject);
    procedure FindDialogFind       (Sender : TObject);
    procedure Pd1KeyPress          (Sender : TObject; var Key: Char);
    procedure Pd1_KeyPress         (Sender : TObject; var Key: Char);
    procedure FListeRowEnter       (Sender : TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure FListeDrawCell       (Sender : TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState);
    procedure FListe2DrawCell      (Sender : TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState);

    procedure PositionneZoomTable;
    procedure PositionneLesVariables;
    procedure InitChamp;
    procedure InitPourBudget;
    procedure InitMsg;
    procedure FermeMsg;
    procedure RempliComboFamRub;
    procedure RempliComboTypRub;
    procedure RempliLibGene;
    procedure FaitRequeteRubrique;
    procedure RunCalculRubrique;
    procedure RempliGrilleCompte;

    procedure LitFicheBudget      (SQL, What : string ; var Cpt : string);
    procedure QuelCptBud          (TypRub, CodeRub, JalBud : string ; var Cpt1, Cpt2 : string);
    procedure BloqueControle      (InRun : Boolean);
    procedure RendVisibleListeCpte(Avec  : Boolean);
    procedure SwapFamille         (St : string ; TypeRubri : Boolean);
    procedure SwapHintZoom        (St : string);
    procedure PrepareValeurPourEdi(Ch : string; var Res : TRes);

    function  RetourneWhere      (TT : TZoomTable; Syn : string ) : string;
    function  FaitRequeteBudget  (Jal, CpteS, CpteG : string)     : string;
    function  AfficheMsg         (Num : Integer; Av, Ap : string) : Word ;
    function  ChercheTotalBudgete(St, JalBud : string)            : Double;
    function  FaitSoldeMixte     (Lesm : Double)                  : string;
    function  TestStop : Boolean;
  end ;


procedure TotalRubriques ;
procedure TotalRubriquesV2(Budget : Boolean) ;

Implementation

uses
{$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
{$ENDIF MODENT1}
{$IFDEF EAGLCLIENT}
    MenuOLX,
{$ELSE}
    MenuOLG,
{$ENDIF EAGLCLIENT}
    BUDGENE_TOM,
    CPSECTION_TOM,
    CPGENERAUX_TOM,
    CPBALGEN_TOF, CRITEDT, ULibExercice, AGLInit,
    CPTIERS_TOM;

const
  pfStyle : TFontStyles = [];
  pfName  : string = 'MS Sans Serif';
  pfSize  : Byte   = 8;
  pfColor : TColor = clWindowText;
  CodError = -123456789.321;

var
  EnEntree : Boolean;

{---------------------------------------------------------------------------------------}
Procedure TotalRubriques ;
{---------------------------------------------------------------------------------------}
begin
  EnEntree := True;
  AglLanceFiche('CP', 'CPTOTRUB', '', '', '');
end ;

{---------------------------------------------------------------------------------------}
Procedure TotalRubriquesV2(Budget : Boolean) ;
{---------------------------------------------------------------------------------------}
begin
  EnEntree := True;
  if Budget then AglLanceFiche('CP', 'CPTOTRUB', '', '', 'OUI;')
            else AglLanceFiche('CP', 'CPTOTRUB', '', '', 'NON;');
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  Budget := S = 'OUI;';
  InitMsg;
  if not Budget then Ecran.HelpContext := 7805000
                else Ecran.HelpContext := 15440000;

  {Définition de pointeurs sur les composants de la fiche}
  FListe  := THGrid(GetControl('FLISTE1'));
  FListe2 := THGrid(GetControl('FLISTE2'));

  Rub1    := THEdit(GetControl('RUB1'));
  Rub2    := THEdit(GetControl('RUB2'));
  PD1     := THEdit(GetControl('PD1'));
  PD1_     := THEdit(GetControl('PD1_'));

  Famrub  := THValComboBox(GetControl('FAMRUB'));
  TypRub  := THValComboBox(GetControl('TYPRUB'));
  Exo     := THValComboBox(GetControl('EXO'));
  RubAno  := THValComboBox(GetControl('RUBANO'));
  Dev     := THValComboBox(GetControl('DEV'));
  Etab    := THValComboBox(GetControl('ETAB'));
  QualPie := THValComboBox(GetControl('QUALPIE'));
  CbAxe   := THValComboBox(GetControl('CBAXE'));
  CbBudget:= THValComboBox(GetControl('CBBUDGET'));
  LibGene := THValComboBox(GetControl('LIBGENE'));
  SupCombo:= THValComboBox(GetControl('SUPERCOMBO'));

  BFerDetail := TToolbarButton97(GetControl('BFERDETAIL'));
  BDetail    := TToolbarButton97(GetControl('BDETAIL'));
  BStop      := TToolbarButton97(GetControl('BSTOP'));

  BGen     := TPopupMenu(GetControl('POPZOOM')).Items[0];
  BRub     := TPopupMenu(GetControl('POPZOOM')).Items[1];
  BCumP    := TPopupMenu(GetControl('POPZOOM')).Items[2];
  BSec     := TPopupMenu(GetControl('POPZOOM')).Items[3];
  SM       := THNumEdit(GetControl('SM'));

  FindDialog := TFindDialog.Create(Ecran);
  FindDialog.OnFind := FindDialogFind;

  OldBrush := TBrush.Create;
  OldPen   := TPen.Create;

  Dev     .OnChange := DevChange;
  CbAxe   .OnChange := CbAxeChange;
  Exo     .OnChange := ExoChange;
  TypRub  .OnChange := TypRubChange;
  Famrub  .OnChange := FamrubChange;
  CbBudget.OnChange := CbBudgetChange;
  (* 02/06/06 En attente de modification de l'AGL après la V7
  {JP 22/05/06 : FQ 17092}
  TFMul(Ecran).FFiltres.OnChange := FiltreChange;
  TFMul(Ecran).OnAfterSelectFiltre := AfterSelectFiltre;
  {JP 22/05/06 : FQ 17092}
  *)
  TToolbarButton97(GetControl('BCHERCHE' )).OnClick := BChercheClick;
  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := BImprimerClick;
  TToolbarButton97(GetControl('BEXPORT'  )).OnClick := bExportClick;
  TToolbarButton97(GetControl('BRECHERCHER')).OnClick := BRechercherClick;

  BFerDetail.OnClick := BFerDetailClick;
  BDetail   .OnClick := BDetailClick;
  BStop     .OnClick := BStopClick;
  BRub      .OnClick := BRubClick;
  BGen      .OnClick := BGenClick;
  BSec      .OnClick := BSecClick;
  BCumP     .OnClick := BCumPClick;

  FListe.OnDblClick       := FListeDblClick;
  FListe.OnTopLeftChanged := FListeTopLeftChanged;
  FListe.OnRowEnter       := FListeRowEnter;
  FListe.OnEnter          := FListeEnter;
  FListe.OnDrawCell       := FListeDrawCell;
  FListe2.OnEnter         := FListe2Enter;
  FListe2.OnTopLeftChanged:= FListe2TopLeftChanged;
  FListe2.OnDblClick      := FListe2DblClick;
  FListe2.OnDrawCell      := FListe2DrawCell;

  // GCO - 22/09/2006 - FQ 18545
  FormKeyDown := Ecran.OnKeyDown;
  Ecran.OnKeyDown := OnKeyDownEcran;

  Pd1.OnKeyPress := Pd1KeyPress;
  Pd1_.OnKeyPress := Pd1_KeyPress;
  //Pd1.OnExit     := Pd1Exit;
  //PD1_.OnExit     := PD1_Exit;

  {$IFDEF EAGLCLIENT}
  {JP 25/07/03 : Comme on ne travaille pas sur la grille du mul, on rend les boutons
                 invisibles, pour éviter les messages d'erreur}
  SetControlVisible('BPREV', False);
  SetControlVisible('BNEXT', False);
  {$ENDIF}

  {Initialisation des contôles et des variables}
  BFerDetail.Visible := False;
  BDetail   .Visible := True;
  SM        .Visible := False;

  TobRub := TOB.Create('', nil, -1);

  MemoLig := 0;
  Stop := False;

  InitPourBudget;
  CbAxe.ItemIndex := 0;

  TyRub := 'GEN'; Frub := ''; LExo  := ''; LRub1 := '';
  LQuali := ''  ; LDev := ''; LEtab := ''; LRub2 := '';
  MemoFam := '' ; MemoTyp := '';
  D1 := 0; D2 := 0;

  FListe2.Visible := False;
  BGen.Enabled := False;
  BSec.Enabled := False;
  RempliComboTypRub;
  if not Budget then FamRub.DataType := 'TTRUBFAMILLE'
                else RempliComboFamRub;
  InitChamp;
  FListe.ALign := alClient;

  SetControlVisible('TCBAXE', False);
  SetControlVisible('CBAXE' , False);

  if EstSerie(S3) or EstSerie(S5) then begin
    SetControlVisible('AVECREVISION' , False);
    SetControlChecked('AVECREVISION', False);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  FListe .VidePile(True);
  FListe2.VidePile(True);
  OldBrush.Free;
  OldPen.Free;
  FermeMsg;
  FreeAndNil(TobRub);
  FreeAndNil(FindDialog);
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.OnDisplay () ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.OnCancel () ;
{---------------------------------------------------------------------------------------}
begin
  if FListe2.Visible then begin
    FListe2.VidePile(True);
    FListe2.Visible := False;
  end ;

  FListe.VidePile(True);
  Inherited ;
end ;

{Constitution du filtre sur la tablette "Rubrique" pour les THEdit Rub1, Rub2
{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.PositionneZoomTable ;
{---------------------------------------------------------------------------------------}
var
  C1, C2,
  SynPlus   : string;
  ZoomTable : TZoomTable;
  Q         : TQuery; {FQ 21561 : 03/12/07}
begin
  ZoomTable := tzRubCPTA;
  SynPlus   := Famrub.Value ;

  if Budget then begin
         if FamRub.Value = 'CBG' then ZoomTable := tzRubBUDG
    else if FamRub.Value = 'CBS' then ZoomTable := tzRubBUDS
    else if FamRub.Value = 'G/S' then ZoomTable := tzRubBUDGS
    else if FamRub.Value = 'S/G' then ZoomTable := tzRubBUDSG;
    SynPlus := CbBudget.Value;
  end;

  Rub1.DataType := 'TZRUBRIQUE';
  Rub2.DataType := 'TZRUBRIQUE';
  {Constitution du filtre à proprement parler}
  Rub1.Plus := RetourneWhere(ZoomTable, SynPlus);
  Rub2.Plus := RetourneWhere(ZoomTable, SynPlus);

  {JP 03/12/07 : 21561 : Mise à jour de l'axe en fonction du budget}
  if Budget then begin
    Q := OpenSQL('SELECT BJ_AXE FROM BUDJAL WHERE BJ_BUDJAL = "' + CbBudget.Value + '"', True);
    if not Q.EOF then
      CbAxe.Value := Q.FindField('BJ_AXE').AsString;
    Ferme(Q);
  end;

  if not FiltreEncours then begin
    {Recherche des première et dernière valeurs de la tablette ainsi filtré}
    PremierDernierRub(ZoomTable, SynPlus, C1, C2);
    Rub1.Text := C1;
    Rub2.Text := C2;
  end;
end;

{Remplissage de la combo LibGene contenant les codes et les libellés des comptes
{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.RempliLibGene;
{---------------------------------------------------------------------------------------}
var
  QLoc : TQuery ;
  Sql  : string ;
begin
  LibGene.Values.Clear;
  LibGene.Items.Clear ;

  {Constitution de la requête en fonction du type de rubriques}
  case Lefb of
    fbAxe1..fbAxe5 : Sql := 'SELECT S_SECTION,S_LIBELLE FROM SECTION WHERE S_AXE="' + FbToAxe(Lefb) + '" ORDER BY S_SECTION';
    fbGene         : Sql := 'SELECT G_GENERAL,G_LIBELLE FROM GENERAUX ORDER BY G_GENERAL';
    fbAux          : Sql := 'SELECT T_AUXILIAIRE,T_LIBELLE FROM TIERS ORDER BY T_AUXILIAIRE';
    fbBudgen       : Sql := 'SELECT BG_BUDGENE,BG_LIBELLE FROM BUDGENE ORDER BY BG_BUDGENE' ;
  end ;

  QLoc:=OpenSql(Sql,True) ;

  while not QLoc.Eof do begin
    LibGene.Values.Add(QLoc.Fields[0].AsString) ;
    LibGene.Items.Add(QLoc.Fields[1].AsString) ;
    QLoc.Next ;
  end ;

  Ferme(QLoc) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.RempliComboTypRub;
{---------------------------------------------------------------------------------------}
begin
  TypRub.Values.Clear;
  TypRub.Items.Clear ;
  if not Budget then FactoriseComboTypeRub(TypRub)
                else TypRub.DataType := 'TTRUBTYPEBUD';
end;

{Remplissage de la combo des familles de rubriques
{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.RempliComboFamRub;
{---------------------------------------------------------------------------------------}
var
  Q    : TQuery ;
  Okok : Boolean ;
begin
  Q := OpenSql('Select CO_CODE,CO_LIBELLE from COMMUN Where CO_TYPE="RBB"',True);

  FamRub.Values.Clear;
  FamRub.Items.Clear;

  while not Q.Eof do begin
    Okok := True;
    if Q.Fields[0].AsString = 'GEN' then
      FamRub.Values.Add('CBG')
    else if Q.Fields[0].AsString = 'ANA' then
      FamRub.Values.Add('CBS')
    else if Q.Fields[0].AsString = 'G/A' then begin
      if EstSerie(S3) then Okok := False
                      else FamRub.Values.Add('G/S');
    end
    else if Q.Fields[0].AsString = 'A/G' then begin
      if EstSerie(S3) then Okok:=False
                      else FamRub.Values.Add('S/G') ;
    end ;

    if Okok then FamRub.Items.Add(Q.FindField('CO_LIBELLE').AsString) ;
    Q.Next ;
  end ;
  Ferme(Q);
end ;

{Initialisation de la fiche pour les rubriques budgétaire
{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.InitPourBudget;
{---------------------------------------------------------------------------------------}
begin
  SetControlVisible('TCBBUDGET', Budget);
  {JP 30/11/05 : FQ 17081 : CBBUDGET était toujours visible => Axe était toujours sur A1}
  SetControlVisible('CBBUDGET', Budget);
  
  SetControlVisible('CBAXE' , not Budget);
  SetControlVisible('TCBAXE' , not Budget); // VL 02/02/2005
  if Budget then SetControlCaption('CBBUDGET' , 'Budget')
            else SetControlCaption('CBBUDGET' , 'Axe');

  if not Budget then Exit ;

  SetControlVisible('TYPRUB' , False);
  SetControlVisible('TTYPRUB', False);

  FListe.ColCount := 7;
  FListe.ColWidths[FListe.ColCount - 2] := 128;
  FListe.ColWidths[FListe.ColCount - 1] := 585;
  FListe.Cells[FListe.ColCount - 2, 0]  := MsgBox.Mess[28];
  FListe.Cells[FListe.ColCount - 1, 0]  := MsgBox.Mess[29];

  if CbBudget.Values.Count > 0 then
    CbBudget.Value := CbBudget.Values[0];

  Exo.DataType := 'TTEXERCICEBUDGET';
  Exo.ReLoad ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.InitChamp ;
{---------------------------------------------------------------------------------------}
begin
  Famrub.ItemIndex := 0;
  TypRub.Value     := 'GEN';
  Exo.Value        := VH^.Entree.Code ;
  RubAno.Value     := 'TOU' ;
  Etab.ItemIndex   := 0;
  PositionneEtabUser(Etab);
  Dev.ItemIndex    := Dev.Values.IndexOf('EUR');
  QualPie.Value := 'NOR';
  DevChange(nil);
  if V_PGI.Controleur then SetControlVisible('AVECREVISION', True);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.PositionneLesVariables;
{---------------------------------------------------------------------------------------}
var
  Typiece : SetttTypePiece;
  St      : string;
begin
  Frub   := Famrub.Value;
  TyRub  := TypRub.Value;
  Lexo   := Exo.Value ;
  LRub1  := Rub1.Text;
  LRub2  := Rub2.Text;
  LRubAno:= RubAno.Value;
  LDev   := Dev.Value;
  LEtab  := Etab.Value;
  LQuali := QualPie.Value;

  Typiece:= WhatTypeEcr(LQuali, True, GetCheckBoxState('AVECREVISION'));

  SQuelTyp(Typiece,St) ;
       if LRubAno = 'SAN' then St := St + '-'
  else if LRubAno = 'ANO' then St := St + '#';
  LRubAno := St;

  D1 := StrToDate(Pd1.Text);
  D2 := StrToDate(PD1_.Text);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.LitFicheBudget(SQL, What : string ; var Cpt : string);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery ;
begin
  Cpt := '';
  Q := OpenSql(SQL + '"' + What + '"', True);
  if not Q.Eof then Cpt := Q.Fields[0].AsString ;
  Ferme(Q);
END ;


{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.QuelCptBud(TypRub, CodeRub, JalBud : string ; var Cpt1, Cpt2 : string);
{---------------------------------------------------------------------------------------}
var
  SQL1,
  SQL2,
  C1,C2 : string ;
  PosI : Integer ;
begin

  SQL1 := 'SELECT BG_BUDGENE FROM BUDGENE WHERE BG_RUB=';
  SQL2 := 'SELECT BS_BUDSECT FROM BUDSECT WHERE BS_RUB=';

  if TypRub='CBG' then begin
    C1 := Copy(CodeRub, 4 + Length(JalBud), Length(CodeRub) - 3 - Length(JalBud));
    LitFicheBudget(SQL1, C1, Cpt1);
  end

  else if TypRub='CBS' then begin
    C1 := Copy(CodeRub, 4 + Length(JalBud), Length(CodeRub) - 3 - Length(JalBud));
    LitFicheBudget(SQL2, C1, Cpt1);
  end

  else if TypRub='G/S' then begin
    PosI := Pos(';', CodeRub) ;
    C1 := Copy(CodeRub, 4 + Length(JalBud), PosI - 4 - Length(JalBud));
    C2 := Copy(CodeRub, PosI + 1, Length(CodeRub) - PosI);
    LitFicheBudget(SQL1, C1, Cpt1);
    LitFicheBudget(SQL2, C2, Cpt2);
  end

  else if TypRub='S/G' then begin
    PosI := Pos(';', CodeRub);
    C1 := Copy(CodeRub, 4 + Length(JalBud), PosI - 4 - Length(JalBud));
    C2 := Copy(CodeRub, PosI + 1, Length(CodeRub) - PosI);
    LitFicheBudget(SQL2, C1, Cpt1);
    LitFicheBudget(SQL1, C2, Cpt2);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.RempliGrilleCompte;
{---------------------------------------------------------------------------------------}
var
  i   : Integer;
  St,
  St1,
  St2 : string;
  X   : TCol;
  TempTStrings : TStringList;
begin
  ChangeMask(SM, Dec, '');

  TempTStrings := TStringList(FListe.Objects[0, FListe.Row]);
  if TempTStrings = nil then Exit;

  FListe2.VidePile(True);
  for i := 0 to TempTStrings.Count - 1 do begin
    St := TempTStrings[i];
    X := TCol.Create;
    X.LCol := StrToInt(Copy(St, Length(St), 1));

    System.Delete(St, Length(St) - 1, 2);
    FListe2.Cells[0, FListe2.RowCount - 1] := Copy(St, 1,Pos(':',St) - 1);

    System.Delete(St, 1, Pos(':', St));
    FListe2.Objects[0, FListe2.RowCount - 1] := X;
    St1 := FListe2.Cells[0, FListe2.RowCount - 1];

    if Pos('/', TypRub.Value) <= 0 then FListe2.Cells[1, FListe2.RowCount - 1] := LibGene.Items[LibGene.Values.Indexof(St1)]
                                   else FListe2.Cells[1, FListe2.RowCount - 1] := LibGene.Items[LibGene.Values.Indexof(ReadTokenSt(St1))];

    St1 := FaitSoldeMixte(StrToFloat(StrfMontant(StrToFloat(Copy(St, 1, Pos(':', St) - 1)), 0, Dec, Symb, False)));
    if St1 = '' then FListe2.Cells[7, FListe2.RowCount - 1] := Copy(St, 1, Pos(':', St) - 1)
                else FListe2.Cells[7, FListe2.RowCount - 1] := St1;

    System.Delete(St, 1, Pos(':',St));
    FListe2.Cells[6, FListe2.RowCount - 1] := StrfMontant(StrToFloat(Copy(St, 1, Pos(':', St) - 1)), 0, Dec, Symb, True);

    System.Delete(St, 1, Pos(':', St));
    FListe2.Cells[5, FListe2.RowCount - 1] := StrfMontant(StrToFloat(Copy(St, 1, Pos(':', St) - 1)), 0, Dec, Symb, True);

    System.Delete(St, 1, Pos(':', St));
    FListe2.Cells[3, FListe2.RowCount - 1] := StrfMontant(StrToFloat(Copy(St, 1, Pos(':', St) - 1)), 0, Dec, Symb, True);

    System.Delete(St, 1, Pos(':', St));
    FListe2.Cells[2, FListe2.RowCount - 1] := StrfMontant(StrToFloat(St), 0, Dec, Symb, True);

    case X.LCol of
      2: St1 := MsgBox.Mess[1];
      3: St1 := MsgBox.Mess[2];
      5: St1 := MsgBox.Mess[3];
      6: St1 := MsgBox.Mess[4];
      7: St1 := MsgBox.Mess[5];
      else St1 := '';
    end ;

    if X.LCol <> 7 then
       FListe2.Cells[4, FListe2.RowCount - 1] := FListe2.Cells[X.LCol, FListe2.RowCount - 1] + ' ' + St1
    else begin
      St2 := Copy(FListe2.Cells[X.LCol, FListe2.RowCount - 1], 1, Length(FListe2.Cells[X.LCol, FListe2.RowCount - 1]) - 1);
      if Trim(St2)='' then FListe2.Cells[4, FListe2.RowCount - 1] := '0 ' + St1
                      else FListe2.Cells[4, FListe2.RowCount - 1] := St2 + St1;
    end;
    FListe2.RowCount:=FListe2.RowCount+1 ;
  end;

  if FListe2.RowCount > 2 then FListe2.RowCount := FListe2.RowCount - 1;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.RendVisibleListeCpte(Avec : Boolean);
{---------------------------------------------------------------------------------------}
begin
  if Avec then begin
    FListe.ALign  := alTop;
    FListe.Height := 80;
    FListe2.Align := alClient;
 end
 else
   FListe.ALign:=alClient ;

  FListe2.Visible := Avec;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.BloqueControle(InRun : Boolean);
{---------------------------------------------------------------------------------------}
begin
  SetControlEnabled('PAGES'      , not InRun);
  SetControlEnabled('PFILTRES'   , not InRun);
  SetControlEnabled('BAGRANDIR'  , not InRun);
  SetControlEnabled('BREDUIRE'   , not InRun);
  SetControlEnabled('BRECHERCHER', not InRun);
  SetControlEnabled('BDETAIL'    , not InRun);
  SetControlEnabled('BMENUZOOM'  , not InRun);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.SwapFamille(St : String ; TypeRubri : Boolean);
{---------------------------------------------------------------------------------------}
var
  St1 : String ;
begin
  if not Budget then Exit;

  if TypeRubri then begin
         if St = 'CBG' then St1 := 'GEN'
    else if St = 'CBS' then St1 := 'ANA'
    else if St = 'G/S' then St1 := 'G/A'
    else if St = 'S/G' then St1 := 'A/G';

     TypRub.ItemIndex := TypRub.Values.IndexOf(St1) ;
  end

  else begin
         if St = 'GEN' then St1 := 'CBG'
    else if St = 'ANA' then St1 := 'CBS'
    else if St = 'G/A' then St1 := 'G/S'
    else if St = 'A/G' then St1 := 'S/G';

     Famrub.ItemIndex := Famrub.Values.IndexOf(St1) ;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.SwapHintZoom(St : string);
{---------------------------------------------------------------------------------------}
begin
  if St = '' then Exit;
  case St[1] of
    'A' : begin
            Case CbAxe.Value[2] of
              '1' : LeFb := fbAxe1 ;
              '2' : LeFb := fbAxe2 ;
              '3' : LeFb := fbAxe3 ;
              '4' : LeFb := fbAxe4 ;
              '5' : LeFb := fbAxe5 ;
            end ;
            BGen.Caption := MsgBox.Mess[9] ;
          end;
    'B' : begin LeFb := fbBudgen; BGen.Caption := MsgBox.Mess[8]; end;
    'G' : begin LeFb := fbGene  ; BGen.Caption := MsgBox.Mess[6]; end;
    'T' : begin LeFb := fbAux   ; BGen.Caption := MsgBox.Mess[7]; end;
  end;

  if Pos('/', St) > 0 then
    case St[3] of
      'A' : BSec.Caption := MsgBox.Mess[9] ;
      'G' : BSec.Caption := MsgBox.Mess[6] ;
      'T' : BSec.Caption := MsgBox.Mess[7] ;
    end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPTOTRUB.FaitRequeteBudget(Jal, CpteS, CpteG : string) : string;
{---------------------------------------------------------------------------------------}
var
  Sql : String ;
begin
  Sql := 'SELECT (SUM(BE_DEBIT) - SUM(BE_CREDIT)) FROM BUDECR '+
         'WHERE BE_EXERCICE = "' + Exo.Value +
         '" AND BE_DATECOMPTABLE >= "' + USDateTime(D1) + '" AND BE_DATECOMPTABLE <= "' + USDateTime(D2) + {JP 30/05/06 : FQ 18005}
         '" AND BE_QUALIFPIECE = "N" ';

  if Letab <> '' then Sql := Sql + 'AND BE_ETABLISSEMENT = "' + Letab + '" ';

  Sql := Sql + ' AND BE_BUDJAL= "' + Jal + '" ';

       if FRub='CBG' then Sql := Sql + 'AND BE_BUDGENE = "' + CpteG + '" '
  else if FRub='CBS' then Sql := Sql + 'AND BE_BUDSECT = "' + CpteS + '" '
  else if (FRub = 'G/S') or (FRub = 'S/G') then
    Sql := Sql + 'AND BE_BUDGENE = "' + CpteG + '" AND BE_BUDSECT = "' + CpteS + '" ';

  Result :=  Sql; 
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.FaitRequeteRubrique;
{---------------------------------------------------------------------------------------}
var
  St,
  St1 : string ;
  Q : TQuery;
begin
  St := 'SELECT * FROM RUBRIQUE ';

  if LRub1 <> '' then begin
    if LRub2 <> '' then
      St := St + 'WHERE RB_RUBRIQUE >= "' + LRub1 + '" AND RB_RUBRIQUE <= "' + LRub2 + '" '
    else
      St := St + 'WHERE RB_RUBRIQUE = "' + LRub1 + '" ';
  end ;

  if (LRub1 = '') and (LRub2 <> '') then
     St := St + 'WHERE RB_RUBRIQUE = "' + LRub2 + '" ';

  if TyRub <> '' then begin
    if Pos('WHERE', St) > 0 then St := St + ' AND RB_TYPERUB = "' + TyRub + '" '
                            else St := St + ' WHERE RB_TYPERUB = "' + TyRub + '" ';

    if Pos('A', TyRub) > 0 then begin
      if Pos('WHERE', St) > 0 then St := St + ' AND RB_AXE = "' + CbAxe.Value + '"'
                              else St := St + ' WHERE RB_AXE = "' + CbAxe.Value + '"';
    end;
  end ;

  if not Budget then begin
     if Pos('WHERE', St) > 0 then St1 := ' AND (RB_NATRUB="CPT")'
                             else St1 := ' WHERE (RB_NATRUB="CPT")';
  end
  else begin
     if Pos('WHERE', St) > 0 then St1 := ' AND (RB_NATRUB="BUD") AND RB_BUDJAL = "' + CbBudget.Value + '" '
                             else St1 := ' WHERE (RB_NATRUB="BUD") AND RB_BUDJAL = "' + CbBudget.Value + '" ';
  end ;

  St := St + St1;

  Q := OpenSql(St, True);        {FP FQ16085 Remplace False par True}
  TobRub.ClearDetail;
  TobRub.LoadDetailDB('RUBRIQUE', '', '', Q, False);
  Ferme(Q);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.RunCalculRubrique;
{---------------------------------------------------------------------------------------}
var
 { Pour TResult 1 pour SM ; 2 pour SC ; 3 pour SD ; 4 pour TC ; 5 pour TD ;}
  TResult : TabloExt;
  LesCpte : TStringList ;
  Compte1,
  Rubrique,
  Appel,
  JalBud,
  St, St1  : string;
  TBudgete : Double;
  n, p     : Integer;
begin
  PositionneLesVariables;
  FaitRequeteRubrique;
  FListe.VidePile(True);

  P := TobRub.Detail.Count;

  {$IFNDEF EAGLCLIENT}
  InitMove(p, MsgBox.Mess[0]);
  {$ENDIF}
  
  if Budget then Appel := 'BUDREA'
            else Appel := 'RUB';

  for n := 0 to p - 1 do begin
    if Pos(FRub, TobRub.Detail[n].GetValue('RB_FAMILLES')) > 0 then begin
      FillChar(TResult, SizeOf(TResult), 0);
      LesCpte := TStringList.Create;
      Rubrique := TobRub.Detail[n].GetValue('RB_RUBRIQUE');

      if GetCumul(Appel, Rubrique, '', LRubAno, LEtab, LDev, LExo, D1, D2, False, False, LesCpte, TResult, False) <> CodError then begin
        if LesCpte.Count > 0 then begin
          FListe.Cells  [0, FListe.RowCount - 1] := Rubrique;
          FListe.Objects[0, FListe.RowCount - 1] := LesCpte;
          FListe.Cells  [1, FListe.RowCount - 1] := TobRub.Detail[n].GetValue('RB_LIBELLE');
          FListe.Cells  [3, FListe.RowCount - 1] := StrfMontant(TResult[4], 0, Dec, Symb, True);
          FListe.Cells  [2, FListe.RowCount - 1] := StrfMontant(TResult[5], 0, Dec, Symb, True);
          FListe.Cells  [4, FListe.RowCount - 1] := StrfMontant(TResult[6], 0, Dec, Symb, True);

          St     := TobRub.Detail[n].GetValue('RB_EXCLUSION1');
          St1    := '';
          JalBud := '';

          if Budget then begin
            JalBud   := TobRub.Detail[n].GetValue('RB_BUDJAL');
            TBudgete := ChercheTotalBudgete(Rubrique, JalBud);
            FListe.Cells[5, FListe.RowCount - 1] := StrfMontant(TBudgete, 0, Dec, Symb, True);
          end;

          while St <> '' do begin
            St1 := ReadTokenSt(St);
            if St1 <> ''then Break ;
          end;

          Compte1 := TobRub.Detail[n].GetValue('RB_COMPTE1');

          if St1 <>'' then
            St := Compte1 + '<<>>' + TobRub.Detail[n].GetValue('RB_EXCLUSION1')
          else
            St := Compte1;

          FListe.Cells[5+Ord(Budget),FListe.RowCount-1]:=St ;
          FListe.RowCount:=FListe.RowCount+1 ;
        end
        else
          FreeAndNil(LesCpte);
      end;
    end;

    if TestStop then Break;
    {$IFNDEF EAGLCLIENT}
    MoveCur(False);
    {$ENDIF}
  end;

  if FListe.RowCount > 2 then FListe.RowCount := FListe.RowCount-1;
  {$IFNDEF EAGLCLIENT}
  FiniMove;
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.ExoChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  Lexo := Exo.Value;
  ExoToDates(LExo,Pd1,PD1_);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.FamrubChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Fam : String ;
begin
  Fam := Famrub.Value;
  if FiltreEncours then Exit;
  if (MemoFam = Fam) and Budget then Exit;
  MemoFam := Fam;
  SwapFamille(Fam, True);
  TyRub := TypRub.Value;
  SwapHintZoom(TyRub);
  SetControlEnabled('BCUMP', (TyRub = 'GEN'));
  PositionneZoomTable;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.TypRubChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  TyRub := TypRub.Value ;
  SetControlEnabled('TCBAXE', Pos('A',TyRub) > 0);
  SetControlEnabled('CBAXE',  Pos('A',TyRub) > 0);
  SetControlEnabled('BCUMP', (TyRub = 'GEN'));
  if (MemoTyp = TyRub) and Budget then Exit;
  MemoTyp := TyRub;
  SwapFamille(TyRub, False);
  SwapHintZoom(TyRub) ;
  if not FiltreEnCours then FamrubChange(FamRub) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.CbAxeChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if Budget then Exit;

  if Pos('A', TyRub) = 1 then begin
    if CbAxe.Value = '' then begin
      LeFb := fbAxe1;
      Exit;
    end;

    case CbAxe.Value[2] of
      '1' : LeFb := fbAxe1;
      '2' : LeFb := fbAxe2;
      '3' : LeFb := fbAxe3;
      '4' : LeFb := fbAxe4;
      '5' : LeFb := fbAxe5;
    end;

    if not FiltreEnCours then FamrubChange(FamRub) ;
  end ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.DevChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Ladev : RDevise ;
begin
  if (Dev.Value = V_PGI.DevisePivot) or (Dev.Value = '') then
    Dec := V_PGI.OkDecV
  else begin
    Ladev.Code := Dev.Value;
    GetInfosDevise(Ladev);
    Dec  := Ladev.Decimale;
    Symb := Ladev.Symbole;
  end ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.BDetailClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if FListe.Cells[0,1] = '' then Exit;

  BDetail   .Visible := False;
  BFerDetail.Visible :=True;
  FListeDblClick(nil);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.BFerdetailClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  FListe2.VidePile(True);
  BFerDetail.Visible := False;
  BDetail   .Visible := True;
  BGen.Enabled := False;
  RendVisibleListeCpte(False) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.FListeDblClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if Fliste.Cells[0, 1] = '' then Exit;

  RendVisibleListeCpte(True);
  FListe.TopRow := FListe.Row;
  BGen.Enabled:=True;
  BSec.Enabled:=(Pos('/',TypRub.Value)>0) ;
  MeMoLig := FListe.Row;
  BDetail   .Visible := False;
  BFerDetail.Visible := True;
  RempliGrilleCompte;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.FListeTopLeftChanged(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if not FListe2.Visible then Exit;
  FListe.Row := FListe.TopRow;
  if MemoLig = FListe.Row then Exit;
  RempliGrilleCompte;
  MemoLig := FListe.Row;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.FListeRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
{---------------------------------------------------------------------------------------}
begin
  if not FListe2.Visible then Exit ;
  MemoLig := Ou;
  RempliGrilleCompte;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.BStopClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  Stop := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.Pd1KeyPress(Sender: TObject; var Key: Char);
{---------------------------------------------------------------------------------------}
begin
  ParamDate(Ecran, Sender, Key);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.FListeEnter(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  CurGrid := 1;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.FListe2Enter(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  CurGrid := 2;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.BChercheClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {On n'exécute pas le traitement en entrée, car il est long et ne répond
   pas forcément aux citères voulus par l'utilisateur}
  if EnEntree then begin
    EnEntree := False;
    Exit;
  end;

  FListe.BeginUpdate ;
  try
    BGen.Enabled:=False ;
    BSec.Enabled:=False ;

    if FListe2.Visible then
      RendVisibleListeCpte(False);

    {Remplissage de la combo LibGene en fonction des critères de recherche}
    RempliLibGene;
    BStop.Cancel := True;
    Stop := False;

    {On desactive tous les contrôles pendant le calcul peut être long}
    BloqueControle(True);
    {Exécution du calcul des montants par rubriques}
    RunCalculRubrique;
    {On réactive les contrôles}
    BloqueControle(False);
    CurGrid:=1 ;
  finally
    FListe.EndUpdate ;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.BImprimerClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  LCpte   : TStringList;
  i, j    : Integer;
  TMere   : TOB;
  TFille  : TOB;
  Crit    : string;
  Revision: string;
  Res     : TRes;
begin
  TMere := TOB.Create('§TOTO', nil, -1);
  try
    {Récupération des données pour l'état}
    for j := 1 to Fliste.RowCount - 1 do begin
      LCpte := TStringList(FListe.Objects[0, j]);
      if LCpte=nil then continue; // VL 02/02/2005
      for i := 0 to LCpte.Count - 1 do begin
        PrepareValeurPourEdi(LCpte[i], Res);
        TFille := Tob.Create('$Fille',TMere,-1);
        TFille.AddChampSup('TR_RUBRIQUE', True);
        TFille.PutValue('TR_RUBRIQUE', Fliste.Cells[0, j]);
        TFille.AddChampSup('TR_LIBELLE', True);
        TFille.PutValue('TR_LIBELLE', Fliste.Cells[1, j]);
        TFille.AddChampSup('TR_COMPTE', True);
        TFille.PutValue('TR_COMPTE', Res.Cpt);
        TFille.AddChampSup('TR_LIBCPTE', True);
        TFille.PutValue('TR_LIBCPTE', LibGene.Items[LibGene.Values.IndexOf(Res.Cpt)]);
        TFille.AddChampSup('TR_DEBIT', True);
        TFille.PutValue('TR_DEBIT', Res.Deb);
        TFille.AddChampSup('TR_CREDIT', True);
        TFille.PutValue('TR_CREDIT', Res.Cdt);
        TFille.AddChampSup('TR_TOTCPTE', True);
        TFille.PutValue('TR_TOTCPTE', Res.Tot);
        TFille.AddChampSup('TR_SOLDEDEB', True);
        TFille.PutValue('TR_SOLDEDEB', Res.Sdb);
        TFille.AddChampSup('TR_SOLDECRED', True);
        TFille.PutValue('TR_SOLDECRED', Res.Scd);
        TFille.AddChampSup('TR_SOLDEMIXTE', True);
        TFille.PutValue('TR_SOLDEMIXTE', Res.Mix);
      end; {for i := 0}

    end;{for j := }

    {Constitution des critères d'édition}
    if GetControlVisible('AVECREVISION') then
      case GetCheckBoxState('AVECREVISION') of
        cbChecked   : Revision := 'Avec les écritures de révision';
        cbUnchecked : Revision := 'Sans les écritures de révision';
        cbGrayed    : Revision := 'Écritures de révision : non spécifié';
      end;

    Crit := 'E_ETABLISSEMENT=' + Etab.Text + '`FAMILLE=' + FamRub.Text + '`E_EXERCICE=' + Exo.Text +
            '`RUBDE=' + Rub1.Text + '`RUBA=' + Rub2.Text + '`ANOUVEAU=' + THValComboBox(GetControl('RUBANO')).Text +
            '`E_DEVISE=' + Dev.Value + '`AVECAFFICHAGE=' + Dev.Text + '`E_DATEDEB=' + PD1.Text + '`REV=' + Revision +
            '`E_DATEFIN=' + PD1_.Text + '`AVECQUALIFPIECE=' + QualPie.Text + '`TYPERUB=' + TypRub.Text;

    {lancement de l'état}
    LanceEtatTOB('E', 'RUB', 'TRU', TMere, True, False, False, nil, '', Ecran.Caption, False, 0, Crit);
  finally
    TMere.Free;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.BCumPClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  St1, StSup, TypRub,
  LAxe, LAxe1, LAxe2,
  Cpt1, Cpt1Ex       : string ;
  QRub               : TQuery ;
  Lefb1, Lefb2       : TFichierBase ;
  Multiple           : TTypeCalc ;
  Typiece  : SetttTypePiece;
  StQ      : string;
  ACritEdt : ClassCritEdt;
begin
  if FListe.Cells[0, 1] = '' then Exit ;

  {Constitution de la requête de l'état}
  QRub := OpenSql('SELECT * FROM RUBRIQUE WHERE RB_RUBRIQUE = "' + Fliste.Cells[0, Fliste.Row] + '"', True);

  if QRub.Eof then begin
    Ferme(QRub);
    Exit;
  end;

  TypRub := QRub.FindField('RB_TYPERUB').AsString;
  Laxe   := QRub.FindField('RB_AXE').AsString;
  Cpt1   := QRub.FindField('RB_COMPTE1').AsString;
  Cpt1Ex := QRub.FindField('RB_EXCLUSION1').AsString;
  Ferme(QRub);

  if (TypRub = 'A/B') or (TypRub = 'A/G') or (TypRub = 'B/A') or
     (TypRub = 'G/A') or (TypRub = 'G/T') or (TypRub = 'T/G') then Multiple := Deux
                                                              else Multiple := Un;

                                                              
  if Multiple = Deux then Exit;

  QuelFBRub(Lefb1, Lefb2, TypRub, LAxe, LAxe1, LAxe2, Multiple, True, False);
  St1 := AnalyseCompte(Cpt1, Lefb1, False, False);
  if St1 <> '' then StSup := '( ' + St1 + ' )';

  St1 := AnalyseCompte(Cpt1Ex, Lefb1, True, False);
  if St1 <> '' then StSup := StSup + ' AND ( ' + St1 + ' )';

  Typiece := WhatTypeEcr(LQuali, True, GetCheckBoxState('AVECREVISION'));
  if tpReel in Typiece then StQ := StQ + 'N;';
  if tpSim  in Typiece then StQ := StQ + 'S;';
  if tpPrev in Typiece then StQ := StQ + 'P;';
  if tpSitu in Typiece then StQ := StQ + 'U;';
  if tpRevi in Typiece then StQ := StQ + 'R;';
  if tpIfrs in Typiece then StQ := StQ + 'I;';

  ACritEdt := ClassCritEdt.Create;
  try
    Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
    ACritEdt.CritEdt.FiltreSup := StSup;
    ACritEdt.CritEdt.Exo.Code := Exo.Value;
    ACritEdt.CritEdt.Date1 := StrToDate(Pd1.Text);
    ACritEdt.CritEdt.Date2 := StrToDate(Pd1_.Text);
    ACritEdt.CritEdt.Qualifpiece := StQ;
    ACritEdt.CritEdt.Etab := Etab.Text;
    ACritEdt.CritEdt.DeviseSelect := Dev.Value;
    ACritEdt.CritEdt.Cpt1 := '';
    ACritEdt.CritEdt.Cpt2 := '';
    TheData := ACritEdt;
    CPLanceFiche_BalanceGeneral;
  finally
    FreeAndNil(ACritEdt);
    TheData := nil;
  end;

  (*
  if StSup <> '' then begin
    PREEJB(Lefb1, PreE, PreJ, PreB);
    StSup := StSup + ' AND ' + WhatExiste(PreE, PreJ, PreB, LRubAno, False, cbUnchecked, D1, D2, LExo, 1, '');
  end ;

  {Mise à jour de la table temporaire de l'état}
  ExecuteSQL('DELETE FROM CEDTBALANCE WHERE CED_USER = "' + V_PGI.User + '"');
  st1 := 'INSERT INTO CEDTBALANCE (CED_USER,CED_RUPTURE,CED_COMPTE, CED_LIBELLE, CED_DEBIT1, CED_CREDIT1, CED_DEBIT2, CED_CREDIT2) SELECT "' + V_PGI.User + '" , "N" ,G_GENERAL ,G_LIBELLE , G_TOTDEBE , G_TOTCREE , iif(((G_TOTCREE-G_TOTDEBE)<0), (G_TOTCREE-G_TOTDEBE), 0) AS CED_DEBIT2, iif(((G_TOTCREE-G_TOTDEBE)<0), 0, (G_TOTCREE-G_TOTDEBE)) FROM GENERAUX Q WHERE ';
  st1 := St1 + StSup;
  ExecuteSQL(St1);

  {Constitution des critères de l'état}
  StCri := 'COMPTE=' + Fliste2.Cells[0, 1] +
           '`COMPTE_=' + Fliste2.Cells[0, Fliste2.RowCount - 1];
           //'`DATECOMPTABLE=' + DateToStr(D1) +
           //'`DATECOMPTABLE_=' + DateToStr(D2) ;
           //'`BILAN=`AVECRUPTURE=SANS' +
           //'`AVECRUPTYPE=SANS`AVECCOLTYPE=TOTAL' + '`AVECAFFICHAGE='+ Dev.Text +
           //'`EXERCICE='+ Exo.Text +
           //'`ETABLISSEMENT=' + Etab.Text;
           //'`DEVISE=' + Dev.Value +
           //'`QUALIFPIECE='+ QualPie.Text+
           //'`NIVORUPTURE=0`MODESELECT=Comptes mouvementés sur la période' +
           //'`TITRECOL1=Total du ' + DateToStr(D1) + ' au ' + DateToStr(D2) +
           //'`TITRECOL2=Solde du ' + DateToStr(D1) + ' au ' + DateToStr(D2);

  {Lancement de l'état}
  St1 := 'SELECT CED_COMPTE CED_GENERAL, CED_LIBELLE, CED_COMPTE2 CED_CYCLE,'+
           'CED_LIBELLE2 CED_CYCLELIB, CED_RUPTURE,"" RUPTURELIB, CED_DEBIT1,'+
           'CED_CREDIT1, CED_DEBIT2, CED_CREDIT2, CED_COLLECTIF, CED_NATURE'+
           ' FROM CEDTBALANCE WHERE CED_USER = "' + V_PGI.User + '"'+
           ' ORDER BY CED_CYCLE, CED_RUPTURE, CED_COMPTE';

  LanceEtat('E', 'BAL', 'BAL', True, False, False, nil, St1, '', False,0,StCri);
  *)
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.BSecClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  St : String ;
begin
  St := FListe2.Cells[0, FListe2.Row];
  ReadTokenSt(St);
  case TypRub.Value[3] of
    'A' : FicheSection(nil, CbAxe.Value, St, taConsult, 0);
    'G' : FicheGene   (nil, '',          St, taConsult, 0);
    'T' : FicheTiers  (nil, '',          St, taConsult, 0);
  end ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.BRubClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var lCtxSurRubrique : TCtxSurRubrique;
begin
  if Budget then
    lCtxSurRubrique := CtxBudget
  else
    lCtxSurRubrique := Ctxrubrique;

  if FListe.Cells[0, Fliste.Row] <> '' then
    ParametrageRubrique(FListe.Cells[0, Fliste.Row], taConsult, lCtxSurRubrique);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.BGenClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  St : String ;
begin
  if not FListe2.Visible then Exit;

  St := FListe2.Cells[0, FListe2.Row];

  if Pos('/', TypRub.Value) > 0  then St := ReadTokenSt(St);
  case Lefb of
    fbAxe1..fbAxe5 : FicheSection(nil, CbAxe.Value, St, taConsult, 0);
    fbGene         : FicheGene   (nil, '', St, taConsult, 0);
    fbAux          : FicheTiers  (nil, '', St, taConsult, 0);
    fbBudgen       : FicheBudgene(nil, '', St, taConsult, 0);
  end ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.bExportClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  SD : TSaveDialog;
begin
  if not ExJaiLeDroitConcept(ccExportListe,True) then Exit;
  SD := TSaveDialog.Create(Ecran);
  SD.DefaultExt := 'XLS';
  SD.Filter := 'Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier ' +
               'Ascii (*.asc)|*.asc|Fichier Lotus (*.wks)|*.wks|Fichier HTML (*.' +
               'html)|*.html|Fichier XML (*.xml)|*.xml';
  SD.Options := [ofOverwritePrompt, ofHideReadOnly, ofPathMustExist, ofNoReadOnlyReturn, ofNoLongNames];
  SD.Title   := 'Export';

  try
    if SD.Execute then begin
      if CurGrid = 1 then ExportGrid(FListe , nil{FListe} , SD.FileName, SD.FilterIndex, True)
                     else ExportGrid(FListe2, nil{FListe2}, SD.FileName, SD.FilterIndex, True);
    end;
  finally
    SD.Free;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.BRechercherClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  FirstFind := True;
  FindDialog.Execute;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.FindDialogFind(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if CurGrid = 1 then Rechercher(FListe , FindDialog, FirstFind)
                 else Rechercher(FListe2, FindDialog, FirstFind);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.FListeDrawCell(Sender: TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  Text : array[0..255] of Char;
  F    : TAlignment ;
begin
  OldBrush.Assign(FListe.Canvas.Brush);
  OldPen.assign(FListe.Canvas.Pen);
  StrPCopy(Text, FListe.Cells[Col, Row]);

  FListe.Canvas.Font.Name  := pfName ;
  FListe.Canvas.Font.Size  := pfSize ;
  FListe.Canvas.Font.Style := pfStyle;
  FListe.Canvas.Font.Color := pfColor;

  if (gdFixed in State) then F := taCenter
  else if (Col = 0) or (Col = 1) or (Col = 5 + Ord(Budget)) then F := taLeftJustify
  else F := taRightJustify;

  if (gdFixed in State) then begin
    FListe.Canvas.Brush.Color := FListe.FixedColor;
    FListe.Canvas.Font .Color := FListe.Font.Color;
  end
  else begin
    if (gdSelected in State) then begin
      FListe.Canvas.Brush.Color := clHighlight;
      FListe.Canvas.Font .Color := clHighlightText;
    end
    else begin
      FListe.Canvas.Brush.Color := clWindow;
      FListe.Canvas.Font.Color  := clWindowText;
    end ;
  end ;

  case F of
    taRightJustify : ExtTextOut(FListe.Canvas.Handle,Rect.Right-FListe.Canvas.TextWidth(FListe.Cells[Col,Row])-3,
                     Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;

    taCenter       : ExtTextOut(FListe.Canvas.Handle,Rect.Left+((Rect.Right-Rect.Left-FListe.canvas.TextWidth(FListe.Cells[Col,Row])) div 2),
                     Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;

    else             ExtTextOut(FListe.Canvas.Handle,Rect.Left+2,
                     Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;
  end ;

  if((gdfixed in State) and FListe.Ctl3D) then begin
    DrawEdge(FListe.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_BOTTOMRIGHT);
    DrawEdge(FListe.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_TOPLEFT);
  end;

  FListe.Canvas.Brush.Assign(OldBrush);
  FListe.Canvas.Pen.Assign(OldPen);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.FListe2DrawCell(Sender: TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  Text : array[0..255] of Char;
  F    : TAlignment ;
begin
  OldBrush.Assign(FListe2.Canvas.Brush);
  OldPen  .Assign(FListe2.Canvas.Pen);
  StrPCopy(Text,FListe2.Cells[Col,Row]);

  FListe2.Canvas.Font.Name  := pfName ;
  FListe2.Canvas.Font.Size  := pfSize ;
  FListe2.Canvas.Font.Style := pfStyle ;
  FListe2.Canvas.Font.Color := pfColor ;

  if (gdFixed in State) then F := taCenter
  else if (Col=0) or (Col=1) then F := taLeftJustify
  else F := taRightJustify ;

  if (gdFixed in State) then begin
    FListe2.Canvas.Brush.Color := FListe2.FixedColor;
    FListe2.Canvas.Font .Color := FListe2.Font.Color;
  end
  else begin
    if (gdSelected in State) then begin
      FListe2.Canvas.Brush.Color := clHighlight;
      FListe2.Canvas.Font .Color := clHighlightText;
    end
    else begin
      FListe2.Canvas.Brush.Color := clWindow;
      if (TCol(FListe2.Objects[0,Row]).LCol = Col) or (Col = 4) then FListe2.Canvas.Font.Color := pfColor
                                                                else FListe2.Canvas.Font.Color := clWindowText;
      end;
    end;

  case F of
    taRightJustify : ExtTextOut(FListe2.Canvas.Handle,Rect.Right-FListe2.Canvas.TextWidth(FListe2.Cells[Col,Row])-3,
                     Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;
    taCenter       : ExtTextOut(FListe2.Canvas.Handle,Rect.Left+((Rect.Right-Rect.Left-FListe2.canvas.TextWidth(FListe2.Cells[Col,Row])) div 2),
                     Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;
    else             ExtTextOut(FListe2.Canvas.Handle,Rect.Left+2,
                     Rect.Top+2,ETO_OPAQUE or ETO_CLIPPED,@Rect,Text,StrLen(Text),nil) ;
  end;

  if((gdfixed in State) and FListe2.Ctl3D) then begin
     DrawEdge(FListe2.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_BOTTOMRIGHT);
     DrawEdge(FListe2.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_TOPLEFT);
  end;

  FListe2.Canvas.Brush.Assign(OldBrush) ; FListe2.Canvas.Pen.Assign(OldPen) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.FListe2TopLeftChanged(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  FListe2.Row := Fliste2.TopRow;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.FListe2DblClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  BGenClick(nil);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.CbBudgetChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  PositionneZoomTable;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.PrepareValeurPourEdi(Ch : string; var Res : TRes);
{---------------------------------------------------------------------------------------}
var
  StCpte, St1,
  St2, St3,
  RubComposite : string;
  LaDecimale   : Integer;
  X            : TCol;
  RIJT         : RDevise ;
begin
  if(Dev.Value = V_PGI.DevisePivot) or (Dev.Value = '') then
    LaDecimale := V_PGI.OkDecV
  else begin
    RIJT.Code := Dev.Value;
    GetInfosDevise(RIJT);
    LaDecimale := RIJT.Decimale;
  end;

  FillChar(Res,SizeOf(Res),#0);

  X := TCol.Create;
  X.LCol := StrToInt(Copy(Ch, Length(Ch), 1));
  System.Delete(Ch, Length(Ch) - 1, 2);

  StCpte := Copy(Ch, 1, Pos(':', Ch) - 1);
  System.Delete(Ch, 1, Pos(':', Ch));

  if (Pos(';', StCpte) > 0) then begin
    RubComposite := StCpte;
    StCpte := ReadTokenSt(RubComposite);
    Res.Cpt := StCpte + ';' + RubComposite;
  end
  else
    Res.Cpt := StCpte;

  St1 := FaitSoldeMixte(StrToFloat(StrfMontant(StrToFloat(Copy(Ch, 1, Pos(':', Ch) - 1)), 0, LaDecimale, '', False)));
  if St1 = '' then Res.Mix := Copy(Ch, 1, Pos(':', Ch) - 1)
              else Res.Mix := St1;
  System.Delete(Ch, 1, Pos(':', Ch));

  Res.Scd := StrfMontant(StrToFloat(Copy(Ch, 1, Pos(':', Ch) - 1)), 0, LaDecimale, '', True);
  System.Delete(Ch, 1, Pos(':', Ch));

  Res.Sdb := StrfMontant(StrToFloat(Copy(Ch, 1, Pos(':', Ch) - 1)), 0, LaDecimale, '', True);
  System.Delete(Ch, 1, Pos(':', Ch));

  Res.Cdt := StrfMontant(StrToFloat(Copy(Ch, 1, Pos(':', Ch) - 1)), 0, LaDecimale, '', True);
  System.Delete(Ch, 1, Pos(':', Ch));

  Res.Deb := StrfMontant(StrToFloat(Ch), 0, LaDecimale, '', True);

  case X.LCol of
    2 : begin
          St1 := MsgBox.Mess[1];
          St3 := Res.Deb;
        end;
    3 : begin
          St1 := MsgBox.Mess[2];
          St3 := Res.Cdt;
        end;
    5 : begin
          St1 := MsgBox.Mess[3];
          St3 := Res.Sdb;
        end;
    6 : begin
          St1 := MsgBox.Mess[4];
          St3 := Res.Scd;
        end;
    7 : begin
          St1 := MsgBox.Mess[5];
          St3 := Res.Mix;
        end;
    else begin
      St1 := '';
      St3 := '';
    end;
  end ;

  if X.LCol <> 7 then
   Res.Tot := St3 + ' ' + St1
  else begin
   St2 := Copy(St3, 1, Length(St3) - 1);
   if Trim(St2) = '' then Res.Tot := '0 '+ St1
                     else Res.Tot := St2 + St1;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPTOTRUB.ChercheTotalBudgete(St, JalBud : string) : Double;
{---------------------------------------------------------------------------------------}
var
  Sql,
  Cpt1,
  Cpt2 : string;
  Q    : TQuery;
begin
  QuelCptBud(FRub,St,JalBud,Cpt1,Cpt2) ;

       if FRub = 'CBG' then Sql := FaitRequeteBudget(JalBud, '', Cpt1)
  else if FRub = 'CBS' then Sql := FaitRequeteBudget(JalBud, Cpt1, '')
  else if FRub = 'G/S' then Sql := FaitRequeteBudget(JalBud, Cpt2, Cpt1)
  else if FRub = 'S/G' then Sql := FaitRequeteBudget(JalBud, Cpt1, Cpt2);

  {L'intérêt de la chose m'échappe, mais par superstition ...
     St1:='' ; SQuelTyp(WhatTypeEcr(LQuali,True,AvecRevision.State),St1);}

  Q := OpenSQL(Sql, True);        {FP FQ16085 Remplace False par True}
  Result := Q.Fields[0].AsFloat;
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
function TOF_CPTOTRUB.FaitSoldeMixte(Lesm : Double) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := '';
  if Lesm = 0 then Exit;

  if Lesm > 0 then AfficheLeSolde(SM, Abs(Lesm), 0)
              else AfficheLeSolde(SM, 0, Abs(Lesm));
  Result := SM.Text;
end ;

{---------------------------------------------------------------------------------------}
function TOF_CPTOTRUB.TestStop : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Application.ProcessMessages;
  Result := False;

  if not Stop then Exit ;

  Stop := False ;
  if AfficheMsg(30,'','') = mrYes then Result := True;
end ;

{---------------------------------------------------------------------------------------}
function TOF_CPTOTRUB.RetourneWhere(TT : TZoomTable; Syn : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := '';
  {JP 01/12/05 : FQ 17079 : Suppression du WHERE dans la clause Plus}
       if TT = tzRubCPTA  then Result := ' (RB_NATRUB="CPT")'
  else if TT = tzRubBUDG  then Result := ' RB_NATRUB="BUD" AND RB_FAMILLES LIKE "CBG%" '
  else if TT = tzRubBUDS  then Result := ' RB_NATRUB="BUD" AND RB_FAMILLES LIKE "CBS%" '
  else if TT = tzRubBUDSG then Result := ' RB_NATRUB="BUD" AND RB_FAMILLES LIKE "S/G%" '
  else if TT = tzRubBUDGS then Result := ' RB_NATRUB="BUD" AND RB_FAMILLES LIKE "G/S%" ';

  if (Syn <> '') and (Result <> '') then begin
    if TT = tzRubCPTA then Result := Result + 'AND RB_FAMILLES like "%' + Syn + '%"'
                      else Result := Result + 'AND RB_BUDJAL like "%'   + Syn + '%"';
  end;
end;

{Gestion des messages
{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.InitMsg;
{---------------------------------------------------------------------------------------}
begin
  MsgBox := THMsgBox.create(FMenuG);

  MsgBox.Mess.Add('Calcul en cours');
  MsgBox.Mess.Add('(TD)');
  MsgBox.Mess.Add('(TC)');
  MsgBox.Mess.Add('(SD)');
  MsgBox.Mess.Add('(SC)');
  MsgBox.Mess.Add('(SM)');
  MsgBox.Mess.Add('Voir Compte général');
  MsgBox.Mess.Add('Voir Compte auxiliaire');
  MsgBox.Mess.Add('Voir Compte budgétaire');
  MsgBox.Mess.Add('Voir Section analytique');
  MsgBox.Mess.Add('&Variable de');
  MsgBox.Mess.Add('11;');
  MsgBox.Mess.Add('Variable');
  MsgBox.Mess.Add('Libellé');
  MsgBox.Mess.Add('Total débit');
  MsgBox.Mess.Add('Total crédit');
  MsgBox.Mess.Add('Total de la variable');
  MsgBox.Mess.Add('Composition de la variable');
  MsgBox.Mess.Add('18;');
  MsgBox.Mess.Add('Consultation des rubriques de détail en montant');
  MsgBox.Mess.Add('Rubrique');
  MsgBox.Mess.Add('Libellé');
  MsgBox.Mess.Add('Total débit');
  MsgBox.Mess.Add('Total crédit');
  MsgBox.Mess.Add('Total Compte');
  MsgBox.Mess.Add('Solde débiteur');
  MsgBox.Mess.Add('Solde créditeur');
  MsgBox.Mess.Add('Solde mixte');
  MsgBox.Mess.Add('Total budgété');
  MsgBox.Mess.Add('Composition de la rubrique');
  MsgBox.Mess.Add('30;Contrôle des montants;Confirmez-vous l''arrêt du traitement en cours ?;Q;YN;N;N;');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.FermeMsg;
{---------------------------------------------------------------------------------------}
begin
  FreeAndNil(MsgBox);
end;

{---------------------------------------------------------------------------------------}
function TOF_CPTOTRUB.AfficheMsg(Num : Integer; Av, Ap : string ) : Word ;
{---------------------------------------------------------------------------------------}
begin
  Result := mrNone ;
  if Num > 30 then Exit ;
  Result := MsgBox.Execute(num,Av,Ap) ;
end;

//SG6 11.03.05
{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.Pd1_KeyPress(Sender : TObject; var Key: Char);
{---------------------------------------------------------------------------------------}
begin
  ParamDate(Ecran, Sender, Key);
end;

(* 02/06/06 En attente de modification de l'AGL après la V7
{JP 22/05/06 : FQ 17092 : Branchement de FiltreEnCours
{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.FiltreChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  FiltreEnCours := True;
  TFMul(Ecran).ListeFiltre.Select(GetControlText('FFILTRES'));
  FiltreEnCours := False;
end;

{JP 22/05/06 : FQ 17092 : Décryptage du contenu du filtre
{---------------------------------------------------------------------------------------}
function TOF_CPTOTRUB.DecrypteFiltre(Champ : string) : string;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := VarToStr(GetValeurDuFiltre(TFMul(Ecran).ListeFiltre.Current, Champ));
  n := RPosAbs('@', Result);
  {Y a-t-il deux @ dans la chaîne ?}
  if Pos('@', Result) <> n then begin
    {Dans ce cas, on coupe la fin de la chaine qui est gérée par l'agl}
    Result := Copy(Result, 1, n - 1);
  end;
end;

{JP 22/05/06 : FQ 17092 : Chargement des zones pouvant contenir un "@"
{---------------------------------------------------------------------------------------}
procedure TOF_CPTOTRUB.AfterSelectFiltre;
{---------------------------------------------------------------------------------------}
begin
  if FiltreEnCours then Exit;
  {22/05/06 : FQ 17092 : Les filtres ne peuvent gérer les @ car c'est un caractère réservé à la gestion
              interne des filtres ce qui entraine des interférence dans le décryptage du flux xml.
              Je le fais donc à la main.}
  SetControlText('FAMRUB', DecrypteFiltre('FAMRUB'));
  SetControlText('RUB1'  , DecrypteFiltre('RUB1'));
  SetControlText('RUB2'  , DecrypteFiltre('RUB2'));
end;
*)

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/09/2006
Modifié le ... : 30/11/2006
Description .. : GCO - 22/09/2006 - FQ 18545
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTRUB.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = VK_F12) then
  begin
    if not FListe.Focused then
      FListe.SetFocus
    else
      FamRub.SetFocus;
  end
  else
  if (Key = VK_F5) and (FListe.Focused) then
    FListe.OnDblClick(nil)
  else
  if (Key = 70) and (Shift = [ssCtrl]) then // GCO - 30/11/2006 - FQ 19248
    BRechercherClick (Sender)
  else  
    FormKeyDown(Sender, Key, Shift); // Pour touches standard AGL
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_CPTOTRUB ] ) ;


end.


