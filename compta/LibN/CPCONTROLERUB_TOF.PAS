{ Unité : Source TOF de la FICHE : CPMULRUB ()
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 4.3.1.101.001   27/06/03   JP   Création de l'unité
 4.3.1.101.001   25/07/03   JP   Rend invisibles les boutons Prev et Next dans le OnArgument
 6.02.001.002    02/02/05   VL   Prise en compte du budget
 8.00.001.011    17/04/07   JP   Petites modifications pour les bugets en 2/3
                                 Correction de l'algo qui ne fonctionnait pas !!
 8.00.001.017    28/05/07   JP   FQ 20382 : Le contrôle des sections et sections / généraux était mal géré
--------------------------------------------------------------------------------------}

Unit CPCONTROLERUB_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Windows,
    {$IFDEF EAGLCLIENT}
     MaineAGL, UtileAGL,
    {$ELSE}
     FE_Main, db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     PrintDBG,
    {$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, Ent1, HMsgBox, UTOF, UTOB, ExtCtrls,
     HTB97, Menus,
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     RUBRIQUE_TOM, // TCtxSurRubrique
     Dialogs;      // TFindDialog
//procedure ControleRubrique (Budget : Boolean = False; CtxRevision : Boolean = False);
procedure ControleRubrique ( vStTypeControle : string );

Type
  TInfoRub = class
    Libelle : string ;
    Famille : string ;
    Compte1 : string ;
    Exclu1  : string ;
    Compte2 : string ;
    Exclu2  : string ;
    TabLib  : string ;
    Axe     : string ;
  end ;

  TOF_CPCONTROLERUB = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure OnClickBRechercher   (Sender : Tobject); virtual;

  private
    Rub1     : THEdit;
    FListe   : THGrid;
    RgChoix  : TRadioGroup;
    CbAxe,
    TypRub, CbBud,
    FamRub   : THValComboBox;
    MsgBox   : THMsgBox;
    TEcrG,
    TEcrA,
    TEcrB,
    TRub,
    TGen     : TOB;
    BGen,
    BRub     : TMenuItem;

    FiltreEnCours : Boolean;

    FCtxSurRubrique : TCtxSurRubrique;

    LFam,
    LRub1,
    LRub2,
    Lcpt1,
    Lcpt2,
    MemoTyp,
    MemoFam : string ;
    Lefb       : TFichierBase ;
    ListeRub,
    ListeGen   : TStringList ;
    BRechercher   : TToolBarButton97;
    AFindDialog   : TFindDialog;
    FFindFirst    : Boolean;

    FormKeyDown  : TKeyEvent;
    procedure OnKeyDownEcran (Sender : TObject; var Key : Word; Shift : TShiftState); virtual;
    
    procedure TyprubChange  (Sender : TObject);
    procedure FamRubChange  (Sender : TObject);
    procedure BImprimerClick(Sender : TObject);
    procedure BChercheClick (Sender : TObject);
    procedure BRubClick     (Sender : TObject);
    procedure BGenClick     (Sender : TObject);
    procedure FListeDblClick(Sender : TObject);
    procedure RgChoixClick  (Sender : TObject);
    procedure CbAxeChange   (Sender : TObject);
    procedure C1Exit        (Sender : TObject);
    procedure C2Exit        (Sender : TObject);
    procedure ZoomMouseEnter(Sender : TObject);

    procedure RempliListeRubGen(AvecGen : Boolean) ;
    procedure TraiteRequeteGen(Ind : Integer ; Composite : Boolean ; Unfb : TFichierBase) ;

    procedure InitFListe ;
    procedure InitMsg ;
    procedure FermeMsg ;
    procedure QuelEstLeFb ;
    procedure PremierToDernier ;
    procedure ChargeRub ;
    procedure RunControle ;
    procedure ControleLesComptes ;
    procedure ControleLesRubriques ;
    procedure PositionneRequete ;
    procedure VideLaListe ;
    procedure ChargeTypRub;
    procedure RempliComboFamRub; {JP 19/04/07}

    function  QuelfbComposite : TFichierBase ;
    function  ChercheCompteDansRub(CptRub,ComptGene : String) : Boolean ;
    function  FaitRequeteEcr(Unfb : TFichierBase) : string;
    function  ChecheLeCompte  : string;
    function  FaitRequeteGene(Compte1,Cexlu1 : String ; Unfb : TFichierBase ; SurTabLib : Boolean) : String ;

    // Budgets
    procedure CbBudChange(Sender: TObject);
    function TradBud(Value : string; VerGen : Boolean = True) : string;

    // YMO 09/08/2006 FQ18639 Recherche par les jumelles
    procedure OnFindAFindDialog    (Sender : TObject);

  end;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  HStatus,
  ParamSoc, // GetParamSocSecur
  CPSECTION_TOM,
  CPGENERAUX_TOM,
  CPTIERS_TOM,
  BUDGENE_TOM,
  CPTEUTIL;

var
  EnEntree : Boolean;

{---------------------------------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/03/2005
Modifié le ... :   /  /
Description .. : Procédure de lancement du Controle des Rubriques
Suite ........ : RUBRIQUE; BUDGET; CYCLE; LIASSE
Mots clefs ... :
*****************************************************************}
procedure ControleRubrique ( vStTypeControle : string );
//procedure ControleRubrique ( Budget : Boolean = False; CtxRevision : Boolean = False);
//var szArg : String;
begin
  EnEntree := True;

  //if Budget then szArg := 'X'
  //          else szArg := '-';

  AGLLanceFiche('CP', 'CPCONTROLERUB', '', '', vStTypeControle);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.OnNew ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.OnDelete ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.OnUpdate ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.OnLoad ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
var Rub2 : THEdit;
    lStArg : string;
begin
  Inherited;
  lStArg := ReadTokenSt(S);

  if lStArg = 'CYCLE' then Ecran.HelpContext:=600300
  else
  if lStArg = 'LIASSE' then Ecran.HelpContext:=600400;

  if lStArg = 'BUDGET' then
    FCtxSurRubrique := CtxBudget
  else
    if lStArg = 'CYCLE' then
      FCtxSurRubrique := CtxCycle
    else
      if lStArg = 'LIASSE' then
        FCtxSurRubrique := CtxLiasse
      else
        FCtxSurRubrique := CtxRubrique;

  InitMsg;
  TGen  := TOB.Create('', nil, -1);
  TRub  := TOB.Create('', nil, -1);
  TEcrG := TOB.Create('', nil, -1);
  TEcrA := TOB.Create('', nil, -1);
  TEcrB := TOB.Create('', nil, -1);

  ListeRub := TStringList.Create;
  ListeGen := TStringList.Create;

  Rub1 := THEdit(GetControl('RUB1'));

  FListe   := THGrid       (GetControl('FLISTE1' ));
  RgChoix  := TRadioGroup  (GetControl('RGCHOIX'));
  CbAxe    := THValComboBox(GetControl('CBAXE'  ));
  TypRub   := THValComboBox(GetControl('TYPRUB' ));
  FamRub   := THValComboBox(GetControl('FAMRUB' ));
  BGen     := TPopupMenu(GetControl('POPZOOM')).Items[0];
  BRub     := TPopupMenu(GetControl('POPZOOM')).Items[1];
  CbBud := THValComboBox(GetControl('CBBUD'));

  TypRub.OnChange   := TyprubChange;
  FamRub.OnChange   := FamRubChange;
  Fliste.OnDblClick := FListeDblClick;
  RgChoix.OnClick   := RgChoixClick;
  cbAxe.OnChange    := CbAxeChange;                
  BRub.OnClick := BRubClick;
  BGen.OnClick := BGenClick;

  // GCO - 22/09/2006 - FQ 18545
  FormKeyDown := Ecran.OnKeyDown;
  Ecran.OnKeyDown := OnKeyDownEcran;

  TToolbarButton97(GetControl('BImprimer')).OnClick := BImprimerClick;
  TToolbarButton97(GetControl('BCherche' )).OnClick := BChercheClick;
  TToolBarButton97(GetControl('BMenuZoom')).OnMouseEnter := ZoomMouseEnter;
  THEdit(GetControl('C1')).OnExit := C1Exit;
  THEdit(GetControl('C2')).OnExit := C2Exit;

  {$IFDEF EAGLCLIENT}
  {JP 25/07/03 : Comme on ne travaille pas sur la grille du mul, on rend les boutons
                 invisibles, pour éviter les messages d'erreur}
  SetControlVisible('BPREV', False);
  SetControlVisible('BNEXT', False);
  {$ENDIF}

  ChargeTypRub;
  FamRub.ItemIndex := 0;
  CbAxe.ItemIndex  := 0;
  Typrub.Value:='GEN' ;

  FamRubChange(nil) ;
  RgChoixClick(nil) ;

  InitFliste ;

  if FCtxSurRubrique = CtxBudget then
  begin
    CbBud.OnChange := CbBudChange;

    SetControlCaption('LRubFam', '&Type');
    SetControlCaption('TTypRub', '&Bud');
    Typrub.Visible:=False;
    Rub2 := THEdit(GetControl('RUB2', True));
    CbBud.Left := Rub2.Left;
    CbBud.Width := Rub2.Width;
    CbBud.Visible := True;
    CbAxe.Visible := False;
    CbBud.ItemIndex:=0;
  end;

  if FCtxSurRubrique in [CtxCycle, CtxLiasse] then
  begin
    SetControlVisible('TTYPRUB', False);
    SetControlVisible('TYPRUB',  False);
    SetControlVisible('TCBAXE',  False);
    SetControlVisible('CBAXE',   False);

    if FCtxSurRubrique = CtxCycle then
    begin
      Ecran.Caption := TraduireMemoire('Contrôle du paramètrage des cycles de révision');
      UpDateCaption(Ecran);
      RgChoix.Items[0] := TraduireMemoire('Cycles de révision');
      SetControlProperty('LRUBFAM','CAPTION',  'Plan de révision');
      // GCO - 29/05/2007 - FQ 20259
      SetControlProperty('FAMRUB', 'DATATYPE', 'CREVPLAN');
      SetControlProperty('TRUB1',  'CAPTION',  'Cycle de révision');
      SetControlProperty('RUB1',   'DATATYPE', 'CREVCYCLEACTIF');
      SetControlProperty('RUB2',   'DATATYPE', 'CREVCYCLEACTIF');
      BRub.Caption := 'Voir le cycle de révision';
      FamRub.Value := VH^.Revision.Plan;
    end
    else
    begin // -> CtxLiasse
      SetControlProperty('LRUBFAM','CAPTION',  'Liasse');
      SetControlProperty('FAMRUB', 'DATATYPE', 'CPCONTROLELIASSE');
      BRub.Caption := 'Voir la liasse';
      RgChoix.Items[0] := TraduireMemoire('&Liasses');
    end;

    SetControlProperty('TCPT1',  'LEFT',     TLabel(GetControl('TRUB2')).Left);
    SetControlProperty('C2',     'LEFT',     THEdit(GetControl('RUB2')).Left);
    SetControlProperty('C1',     'WIDTH',    THEdit(GetControl('RUB1')).Width);
    SetControlProperty('C2',     'WIDTH',    THEdit(GetControl('RUB2')).Width);
  end;

  // YMO 09/08/2006 FQ18639 Recherche par les jumelles
  BRechercher   := TToolBarButton97(GetControl('BRechercher', True));
  BRechercher.OnClick   := OnClickBRechercher;

  AFindDialog := TFindDialog.Create(Ecran);
  AFindDialog.OnFind := OnFindAFindDialog;

end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  FermeMsg;
  FreeAndNil(TGen);
  FreeAndNil(TRub);
  FreeAndNil(TEcrG);
  FreeAndNil(TEcrA);
  FreeAndNil(TEcrB);
  FListe.VidePile(False) ;
  VideLaListe ;
  FreeAndNil(ListeRub);
  FreeAndNil(ListeGen);
  FreeAndNil(AFindDialog);
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.OnDisplay () ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.OnCancel () ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.C1Exit(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  SetControlText('C1', BourrelaDonc(GetControlText('C1'), LeFb));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.C2Exit(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  SetControlText('C2', BourrelaDonc(GetControlText('C2'), LeFb));
end;

{On bloque le rafraichissement de la grille pendant son remplissage
{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.BChercheClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {On n'exécute pas le traitement en entrée, car il est long et ne répond
   pas forcément aux citères voulus par l'utilisateur}
  if EnEntree then begin
    EnEntree := False;
    Exit;
  end;

  Fliste.BeginUpdate ;
  try
    if FListe.Cells[0,1] <> '' then FListe.VidePile(False) ;
    VideLaListe ;
    RunControle ;
  finally
    FListe.EndUpdate ;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.ZoomMouseEnter(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  case Lefb of
    fbAxe1..fbAxe5 : Bgen.Caption := 'Voir Section';
    fbGene         : Bgen.Caption := 'Voir Compte général';
    fbAux          : Bgen.Caption := 'Voir compte auxiliaire';
    fbBudgen       : Bgen.Caption := 'Voir compte bugétaire';
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.BRubClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  C : string ;
begin
  if FListe.Cells[0,1] = '' then Exit ;
  case RgChoix.ItemIndex of
    0 : begin
          if FListe.Cells[0, FListe.Row] = '' then C := ChecheLeCompte
                                              else C := FListe.Cells[0, FListe.Row] ;
          ParametrageRubrique(C, taConsult, FCtxSurRubrique);
        end ;
    1 : begin
          ParametrageRubrique(FListe.Cells[2, Fliste.Row], taConsult, FCtxSurRubrique) ;
        end ;
  end ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.BGenClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  C : String ;
begin
  if FListe.Cells[0,1]='' then Exit ;
  Case RgChoix.ItemIndex of
    0 : begin
          Case Lefb of
            fbAxe1..fbAxe5 : FicheSection(nil, CbAxe.Value, FListe.Cells[2,FListe.Row], taConsult, 0);
            fbGene         : FicheGene   (nil, '',          FListe.Cells[2,FListe.Row], taConsult, 0);
            fbAux          : FicheTiers  (Nil, '',          FListe.Cells[2,FListe.Row], taConsult, 0);
            fbBudgen       : FicheBudgene(nil, '',          FListe.Cells[2,FListe.Row], taConsult, 0);
          end ;
        end ;

    1 : begin
          if FListe.Cells[0,FListe.Row]='' then C := ChecheLeCompte
                                           else C := FListe.Cells[0, FListe.Row] ;

          Case Lefb of
            fbAxe1..fbAxe5 : FicheSection(Nil, CbAxe.Value, C, taConsult, 0);
            fbGene         : FicheGene   (Nil, '',          C, taConsult, 0);
            fbAux          : FicheTiers  (Nil, '',          C, taConsult, 0);
            fbBudgen     : FicheBudgene(Nil, '',          C, taConsult, 0);
        end ;
     end ;
  end ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.FListeDblClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  case RgChoix.ItemIndex of
    0 : BRubClick(nil) ;
    1 : BGenClick(nil) ;
  end ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.RgChoixClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if Pos('/',Typrub.Value) > 0 then RgChoix.ItemIndex := 0;

  Case RgChoix.ItemIndex of
    0 : begin
          SetControlEnabled('RUB1'  , True);
          SetControlEnabled('RUB2'  , True);
          SetControlEnabled('C1'    , False);
          SetControlEnabled('C2'    , False);

          if not (FCtxSurRubrique in [CtxCycle, CtxLiasse]) then
            SetControlEnabled('CBTOUS', False);
        end ;

    1 : begin
          SetControlEnabled('C1'    , True);
          SetControlEnabled('C2'    , True);
          SetControlEnabled('CBTOUS', True);
          PremierToDernier ;
          SetControlEnabled('RUB1'  , False);
          SetControlEnabled('RUB2'  , False);
        end ;
   end ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.CbAxeChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if not CbAxe.Enabled then Exit ;

  if Pos('A',TypRub.Value)=1 then
    case CbAxe.Value[2] of
      '1' : begin LeFb := fbAxe1; THEdit(GetControl('C1')).DataType := 'TZSECTION' ; THEdit(GetControl('C2')).DataType := 'TZSECTION' ; end;
      '2' : begin LeFb := fbAxe2; THEdit(GetControl('C1')).DataType := 'TZSECTION2'; THEdit(GetControl('C2')).DataType := 'TZSECTION2'; end;
      '3' : begin LeFb := fbAxe3; THEdit(GetControl('C1')).DataType := 'TZSECTION3'; THEdit(GetControl('C2')).DataType := 'TZSECTION3'; end;
      '4' : begin LeFb := fbAxe4; THEdit(GetControl('C1')).DataType := 'TZSECTION4'; THEdit(GetControl('C2')).DataType := 'TZSECTION4'; end;
      '5' : begin LeFb := fbAxe5; THEdit(GetControl('C1')).DataType := 'TZSECTION5'; THEdit(GetControl('C2')).DataType := 'TZSECTION5'; end;
     end ;

  PremierToDernier ;
  FamRubChange(nil) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.BImprimerClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
{$IFDEF EAGLCLIENT}
var
  MaTob : TOB;
  T1    : TOB;
  n, p  : Integer;
  CritEdition : string;
{$ENDIF}
begin
  {$IFDEF EAGLCLIENT}
  MaTob := TOB.Create('§LISTEPARAM', nil, -1);
  try
    for n := 1 to Fliste.RowCount - 1 do begin
      T1 := TOB.Create('', MaTob, -1);
      for p := 1 to 6 do begin
        T1.AddChampSup('CT_COL' + IntToStr(p) +'A', True);
        T1.PutValue('CT_COL' + IntToStr(p) + 'A', Fliste.Cells[p-1, n]);
      end;
    end;
    CritEdition := 'TITRE=Contrôle du paramétrage';
    for p := 1 to 6 do
      CritEdition := CritEdition + '`COL' + IntToStr(p) + '=' + Fliste.Cells[p-1, 0];
    LanceEtatTOB('E', 'RUB', 'CON', MaTob, True, False, False, nil, '', 'Contrôle du paramétrage', False, 0, CritEdition);
  finally
    FreeAndNil(MaTob);
  end;
  {$ELSE}
  PrintDBGrid (FListe, nil, Ecran.Caption, '');
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.TyprubChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  SetControlEnabled('TCbAxe', (Pos('A', TypRub.Value) > 0));
  SetControlEnabled('CbAxe' , (Pos('A', TypRub.Value) > 0));

  if TypRub.Value = '' then Exit ;

  if (Pos('/', TypRub.Value) > 0) and (RgChoix.ItemIndex = 1) then RgChoix.ItemIndex := 0 ;

  MemoTyp := TypRub.Value ;

  if FCtxSurRubrique = CtxBudget then
    Famrub.ItemIndex := Famrub.Values.IndexOf(TradBud(MemoTyp, False)) ;

  case TypRub.Value[1] of
    'A' : begin
            if CbAxe.Values.Count > 0 then CbAxe.Value := CbAxe.Values[0];
          end;
    'B' : begin
            Lefb := fbBudgen;
            THEdit(GetControl('C1')).DataType := 'TZBUDGENATT';
            THEdit(GetControl('C1')).Plus     := '';
            THEdit(GetControl('C2')).DataType := 'TZBUDGENATT';
            THEdit(GetControl('C2')).Plus     := '';
          end;
    'G' : begin
            Lefb := fbGene;
            THEdit(GetControl('C1')).DataType := 'TZGENERAL';
            THEdit(GetControl('C2')).DataType := 'TZGENERAL';
          end;
    'T' : begin
            Lefb := fbAux;
            THEdit(GetControl('C1')).DataType := 'TZTTOUS';
            THEdit(GetControl('C2')).DataType := 'TZTTOUS';
//            THEdit(GetControl('C2')).Plus := 'AND T_NATUREAUXI<>"NCP" AND T_NATUREAUXI<>"CON" AND T_NATUREAUXI<>"PRO" AND T_NATUREAUXI<>"SUS"'
          end;
  end ;

  if TypRub.Value[1] <> 'A' then begin
    PremierToDernier;
    FamRubChange(nil);
  end ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.FamRubChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Fam : string ;
begin
  Fam := TradBud(FamRub.Value);
  if (Pos('/',Fam) > 0) and (RgChoix.ItemIndex = 1) then RgChoix.ItemIndex := 0 ;
  MemoFam := Fam ;
  if FCtxSurRubrique = CtxBudget then
    TypRub.ItemIndex := TypRub.Values.IndexOf(Fam) ;
  ChargeRub;
end;

{---------------------------------------------------------------------------------------}
Procedure TOF_CPCONTROLERUB.RunControle ;
{---------------------------------------------------------------------------------------}
begin
  LFam  := FamRub.Value;
  LRub1 := GetControlText('RUB1');
  LRub2 := GetControlText('RUB2');
  InitFliste ;

  Lcpt1 := GetControlText('C1');
  Lcpt2 := GetControlText('C2');
  PositionneRequete ;

  FListe.VidePile(False);

  Case RgChoix.ItemIndex of
    0 : begin
          RempliListeRubGen(False) ;
          InitMove(ListeRub.Count,MsgBox.Mess[3]) ;
          ControleLesRubriques ;
          FiniMove ;
        end ;
    1 : begin
          RempliListeRubGen(True) ;
          InitMove(ListeGen.Count,MsgBox.Mess[3]) ;
          ControleLesComptes ;
          FiniMove ;
        end ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
Procedure TOF_CPCONTROLERUB.ControleLesComptes ;
{---------------------------------------------------------------------------------------}
Var
  Compte,
  CptRub,
  CptExRub    : string ;
  PremierTour : Boolean ;
  i, j        : Integer ;
  Temp        : TOB;
  Q           : TQuery;
  lstChampCpte: string;
begin
  TEcrG.ClearDetail;
  Q := OpenSql(FaitRequeteEcr(fbGene), True);
  TEcrG.LoadDetailDB('', '', '', Q, False);
  Ferme(Q);

  if GetParamSocSecur('SO_CPLIASSESURCORRESP',False) then
  begin
    if GetParamSocSecur('SO_CPLIASSEPLANCORRESP','')='1' then lstChampCpte := 'G_CORRESP1'
    else if GetParamSocSecur('SO_CPLIASSEPLANCORRESP','')='2' then lstChampCpte := 'G_CORRESP2'
    else lstChampCpte := 'G_GENERAL';
  end else lstChampCpte := 'G_GENERAL';

  for j:=0 to ListeGen.Count-1 do begin
    MoveCur(False) ;
    PremierTour := True ;
    for i := 0 to ListeRub.Count - 1 do begin
      if PremierTour then begin
        Compte := ListeGen.Strings[j];
        FListe.Cells[0,FListe.RowCount - 1] := Compte ;
        FListe.Cells[1,FListe.RowCount - 1] := TInfoRub(ListeGen.Objects[j]).Libelle ;

        if Pos('A',Typrub.Value)=1 then begin
           if VH^.Cpta[AxeToFb(CbAxe.Value)].Chantier then Fliste.Cells[5,FListe.RowCount-1] := MsgBox.Mess[6]
                                                      else Fliste.Cells[5,FListe.RowCount-1] := MsgBox.Mess[5] ;
        end

        else begin
          Temp := TEcrG.FindFirst(['G_GENERAL'], [Compte], True);
          if Temp <> nil then Fliste.Cells[5,FListe.RowCount-1] := MsgBox.Mess[6]
                         else Fliste.Cells[5,FListe.RowCount-1] := MsgBox.Mess[5] ;
        end ;
        PremierTour:=False ;
      end ;

      if Pos(LFam, TInfoRub(ListeRub.Objects[i]).Famille) > 0 then begin
        CptRub:=TInfoRub(ListeRub.Objects[i]).Compte1 ;
        CptExRub:=TInfoRub(ListeRub.Objects[i]).Exclu1 ;

      if TGen.Detail[j].GetString(lStChampCpte)<>'' then
        Compte := TGen.Detail[j].GetString(lStChampCpte);
        if ChercheCompteDansRub(CptRub,Compte)then begin
           FListe.Cells[2,FListe.RowCount-1] := ListeRub.Strings[i] ;
           FListe.Cells[3,FListe.RowCount-1] := TInfoRub(ListeRub.Objects[i]).Libelle ;
           FListe.Cells[4,FListe.RowCount-1] := MsgBox.Mess[5] ;
           FListe.RowCount := FListe.RowCount+1 ;
        end ;

        if CptExRub<>'' then begin
          if ChercheCompteDansRub(CptExRub,Compte) then begin
            FListe.Cells[2,FListe.RowCount-1]:=ListeRub.Strings[i] ;
            FListe.Cells[3,FListe.RowCount-1]:=TInfoRub(ListeRub.Objects[i]).Libelle ;
            FListe.Cells[4,FListe.RowCount-1]:=MsgBox.Mess[6] ;
            FListe.RowCount:=FListe.RowCount+1 ;
          end ;
        end ;
      end ; { if Pos(LFam }
    end ; { for i := 0 }

     if GetCheckBoxState('CBTOUS') = cbChecked  then
       if FListe.Cells[0,FListe.RowCount-1] <> '' then
         FListe.RowCount := FListe.RowCount + 1;
  end ; { for j := 0 }

  if Fliste.RowCount > 2 then
    Fliste.RowCount := Fliste.RowCount-1 ;

  if (Fliste.RowCount = 2) and (Fliste.Cells[2,1] = '') then
     for i := 0 to Fliste.ColCount-1 do Fliste.Cells[i,1] := '' ;
end ;

{---------------------------------------------------------------------------------------}
function TOF_CPCONTROLERUB.FaitRequeteGene(Compte1, Cexlu1: String; Unfb: TFichierBase; SurTabLib: Boolean): String;
{---------------------------------------------------------------------------------------}
Var
  Sql ,St,
  Where,
  Cod ,Lib,
  Table  : string ;
  lStCorresp : string;
begin
  Result:='' ;
  case Unfb of
      fbAxe1..fbAxe5 : begin
                         Cod   := 'S_SECTION';
                         Lib   := 'S_LIBELLE';
                         Table := 'SECTION'  ;
                         St    := 'AND S_AXE="' + CbAxe.Value + '" ';
                       end;
      fbGene         : begin
                         Cod   := 'G_GENERAL';
                         Lib   := 'G_LIBELLE';
                         Table := 'GENERAUX' ;
                       end;
      fbAux          : begin
                         Cod   := 'T_AUXILIAIRE';
                         Lib   := 'T_LIBELLE' ;
                         Table := 'TIERS' ;
                       end ;
      fbBudgen       : begin
                         Cod   := 'BG_BUDGENE' ;
                         Lib   := 'BG_LIBELLE' ;
                         Table := 'BUDGENE' ;
                       end ;
  end ;

  Where := AnalyseCompte(Compte1, Unfb, False, SurTabLib) ;


  if Where <> '' then begin
    { Analyse sur les comptes de correspondances }
    if (UnFb=fbGene) and GetParamSocSecur('SO_CPLIASSESURCORRESP',False) then
    begin
      if GetParamSocSecur('SO_CPLIASSEPLANCORRESP','')='1' then lStCorresp := 'G_CORRESP1'
      else if GetParamSocSecur('SO_CPLIASSEPLANCORRESP','')='2' then lStCorresp := 'G_CORRESP2'
      else lStCorresp := '';
      if lStCorresp <> '' then
      begin
        Where := '(('+Where + ' AND '+lStCorresp+'="") OR ('+lStCorresp+'<>"" AND ('+
          FindEtReplace(Where,'G_GENERAL',lStCorresp,True)+')))';
      end;
    end;

    Sql := 'Select ' + Cod + ',' + Lib + ' From ' + Table + '';
    Sql := Sql + ' Where ' + Where ;
  end
  else
    Exit;

  if Cexlu1 <>'' then
  begin
    Where := AnalyseCompte(Cexlu1, Unfb, True, False) ;
    if Where <> '' then
    begin
      { Analyse sur les comptes de correspondances }
      if (UnFb=fbGene) and GetParamSocSecur('SO_CPLIASSESURCORRESP',False) then
      begin
        if GetParamSocSecur('SO_CPLIASSEPLANCORRESP','')='1' then lStCorresp := 'G_CORRESP1'
        else if GetParamSocSecur('SO_CPLIASSEPLANCORRESP','')='2' then lStCorresp := 'G_CORRESP2'
        else lStCorresp := '';
        if lStCorresp <> '' then
        begin
          Where := '(('+FindEtReplace(Where,'G_GENERAL',lStCorresp,True)+' AND '+
            lStCorresp+'<>"") OR '+lStCorresp+'="")';
        end;
      end;
      Sql := Sql + ' And ' + Where ;
    end;
  end ;
  Sql := Sql + St + ' Order by ' + Cod ;
  Result := Sql;
END ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.PositionneRequete ;
{---------------------------------------------------------------------------------------}
var
  St,
  S  : string;
  Q  : TQuery;
begin
  if (LRub1 = '') or (LRub2 = '') then Exit ;

  TGen.ClearDetail;
  TRub.ClearDetail;

  case FCtxSurRubrique of
    CtxRubrique : St := ' AND (RB_NATRUB="CPT")' ;

    CtxBudget   : St := ' And (RB_NATRUB="BUD" And RB_BUDJAL="'+CbBud.Value+'")';

    CtxCycle    : St := ' AND RB_NATRUB = "CPT" AND RB_CLASSERUB = "CDR"';

    CtxLiasse   : ;
  end;

  case RgChoix.ItemIndex of
    0 : begin
          S := 'SELECT * FROM RUBRIQUE WHERE RB_RUBRIQUE >= "' + LRub1 + '" AND RB_RUBRIQUE <= "' +
               LRub2 + '" AND  RB_TYPERUB="' + Typrub.Value + '"' ;
          S := S + St;

          if (TypRub.Value = 'ANA') or (TypRub.Value = 'G/A') Or (TypRub.Value='A/G') then
            S := S + ' AND RB_AXE="' + CbAxe.Value + '" ' ;

          {A FAIRE !!!!
          // CA - 27/11/2001 - Ne pas prendre les enregistrements des autres dossiers
          QRub.SQL.Clear;
          QRub.SQL.Add (InsertWhereDPStd(S));
          // Fin ajout CA
          }
          Q := OpenSql(S, False);
          TRub.LoadDetailDB('RUBRIQUE', '', '', Q, False);
          Ferme(Q);
        end ;

    1 : begin
          S := 'SELECT * FROM RUBRIQUE WHERE RB_TYPERUB = "' + Typrub.Value + '"';
          S := S + St;
          if (TypRub.Value = 'ANA') or (TypRub.Value = 'G/A') Or (TypRub.Value='A/G') then
            S := S + ' AND RB_AXE="' + CbAxe.Value + '" ';

          case Lefb of
            fbGene   : St := 'SELECT G_GENERAL, G_LIBELLE, G_CORRESP1, G_CORRESP2 FROM GENERAUX WHERE G_GENERAL >= "' + Lcpt1 +
                             '" AND G_GENERAL <= "' + Lcpt2 + '" ORDER BY G_GENERAL';

            fbAux    : St := 'SELECT T_AUXILIAIRE, T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE >= "' + Lcpt1 +
                             '" AND T_AUXILIAIRE <= "' + Lcpt2 + '" ORDER BY T_AUXILIAIRE';

            fbBudgen : St := 'SELECT BG_BUDGENE, BG_LIBELLE FROM BUDGENE WHERE BG_BUDGENE >= "' + Lcpt1 +
                             '" AND B_BUDGENE <= "' + Lcpt2 + '" ORDER BY BG_BUDGENE';

            fbAxe1..fbAxe5 : St := 'SELECT S_SECTION,S_LIBELLE FROM SECTION WHERE S_SECTION >= "' + Lcpt1 +
                                   '" AND S_SECTION <= "' + Lcpt2 + '" AND S_AXE = "' + CbAxe.Value +
                                   '" ORDER BY S_SECTION';
          end;
          {A FAIRE !!!!
          // CA - 27/11/2001 - Ne pas prendre les enregistrements des autres dossiers
          stSql := QRub.SQL.Text;
          QRub.SQL.Clear;
          QRub.SQL.Add (stSQL);
          // Fin ajout CA
          }
          Q := OpenSql(S, False);
          TRub.LoadDetailDB('RUBRIQUE', '', '', Q, False);
          Ferme(Q);

          Q := OpenSql(St, False);
          TGen.LoadDetailDB('', '', '', Q, False);
          Ferme(Q);
        end ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.RempliListeRubGen(AvecGen : Boolean ) ;
{---------------------------------------------------------------------------------------}
var
  X   : TInfoRub;
  Cod,
  Lib : String ;
  n   : Integer;
begin

  // GCO - 20/09/2005 - FQ 16390
  if (FCtxSurRubrique = CtxCycle) and
     (RgChoix.ItemIndex = 0) and
     (GetCheckBoxState('CBTOUS') = cbChecked) then
  begin
    X := TInfoRub.Create ;
    X.Libelle := 'Comptes sans cycle de révision';
    X.Famille := VH^.Revision.Plan;
    X.Compte1 := '';
    X.Exclu1  := '';
    X.Compte2 := '';
    X.Exclu2  := '';
    X.TabLib  := '';
    X.Axe     := '';
    ListeRub.AddObject('SANS CYCLE',X) ;
  end;

  for n := 0 to TRub.Detail.Count - 1 do begin
    X := TInfoRub.Create ;
    X.Libelle := TRub.Detail[n].GetValue('RB_LIBELLE'   );
    X.Famille := TRub.Detail[n].GetValue('RB_FAMILLES'  );
    X.Compte1 := TRub.Detail[n].GetValue('RB_COMPTE1'   );
    X.Exclu1  := TRub.Detail[n].GetValue('RB_EXCLUSION1');
    X.Compte2 := TRub.Detail[n].GetValue('RB_COMPTE2'   );
    X.Exclu2  := TRub.Detail[n].GetValue('RB_EXCLUSION2');
    X.TabLib  := TRub.Detail[n].GetValue('RB_TABLELIBRE');
    X.Axe     := TRub.Detail[n].GetValue('RB_AXE'       );
    ListeRub.AddObject(TRub.Detail[n].GetValue('RB_RUBRIQUE'),X) ;
  end ;

  if AvecGen then begin
    case Lefb of
      fbGene         : begin Cod :='G_GENERAL'   ; Lib := 'G_LIBELLE'; end;
      fbAux          : begin Cod :='T_AUXILIAIRE'; Lib := 'T_LIBELLE'; end;
      fbAxe1..fbAxe5 : begin Cod :='S_SECTION'   ; Lib := 'S_LIBELLE'; end;
    end ;

    for n := 0 to TGen.Detail.Count - 1 do begin
      X := TInfoRub.Create ;
      X.Libelle := TGen.Detail[n].GetValue(Lib);
      ListeGen.AddObject(TGen.Detail[n].GetValue(Cod), X) ;
    end ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
Procedure TOF_CPCONTROLERUB.ControleLesRubriques;
{---------------------------------------------------------------------------------------}
var
  Compte1,
  Cexlu1,
  St,
  Compte2,
  Cexlu2 : String ;
  i      : Integer ;
  fb1,
  fb2    : TFichierBase ;
  Q      : TQuery;
begin
  fb1 := Lefb ;
  fb2 := QuelfbComposite ;

  TEcrG.ClearDetail;
  TEcrA.ClearDetail;
  TEcrB.ClearDetail;

  St := FaitRequeteEcr(fbGene);
  Q := OpenSql(St, True);
  TEcrG.LoadDetailDB('', '', '', Q, False);
  Ferme(Q);

  St := FaitRequeteEcr(fbAux);
  Q := OpenSql(St, True);
  TEcrA.LoadDetailDB('', '', '', Q, False);
  Ferme(Q);

  St := FaitRequeteEcr(fbBudGen);
  Q := OpenSql(St, True);
  TEcrB.LoadDetailDB('', '', '', Q, False);
  Ferme(Q);

  for i := 0 to ListeRub.Count-1 do
  begin
    if Pos(LFam, TInfoRub(ListeRub.Objects[i]).Famille) > 0 then
    begin
      FListe.Cells[0,FListe.RowCount-1] := ListeRub.Strings[i] ;
      FListe.Cells[1,FListe.RowCount-1] := TInfoRub(ListeRub.Objects[i]).Libelle ;

      Compte1 := TInfoRub(ListeRub.Objects[i]).Compte1 ;
      Cexlu1  := TInfoRub(ListeRub.Objects[i]).Exclu1 ;

      // GCO - 20/09/2005 - FQ 16390
      if (FCtxSurRubrique = ctxCycle) and
         (RgChoix.ItemIndex = 0) and
         (GetCheckBoxState('CBTOUS') = cbChecked) and
         (Compte1 = '') and (CExlu1= '') then
      begin
        TGen.ClearDetail;
        Q := OpenSql('SELECT * FROM GENERAUX WHERE G_CYCLEREVISION = "" ORDER BY G_GENERAL', True);
        TGen.LoadDetailDB('', '', '', Q, False);
        Ferme(Q);
        TraiteRequeteGen(0, False, fb1) ;
      end
      else
      begin
        St := FaitRequeteGene(Compte1,Cexlu1,fb1,TInfoRub(ListeRub.Objects[i]).TabLib='X') ;
        if St = '' then Continue ;
        TGen.ClearDetail;
        Q := OpenSql(St, False);
        TGen.LoadDetailDB('', '', '', Q, False);
        Ferme(Q);
        TraiteRequeteGen(0, False, fb1) ;

        St := FaitRequeteGene(Cexlu1, '', fb1, False) ;
        if St <> '' then
        begin
          TGen.ClearDetail;
          Q := OpenSql(St, False);
          TGen.LoadDetailDB('', '', '', Q, False);
          Ferme(Q);
          TraiteRequeteGen(1, False, fb1) ;
        end ;
        
        if Pos('/',Typrub.Value) > 0 then
        begin
          Compte2 := TInfoRub(ListeRub.Objects[i]).Compte2 ;
          Cexlu2  := TInfoRub(ListeRub.Objects[i]).Exclu2 ;

          St := FaitRequeteGene(Compte2,Cexlu2,fb2,False) ;

          if St <> '' then begin
            FListe.Cells[0,FListe.RowCount-1] := ListeRub.Strings[i] ;
            FListe.Cells[1,FListe.RowCount-1] := TInfoRub(ListeRub.Objects[i]).Libelle ;
            TGen.ClearDetail;
            Q := OpenSql(St, False);
            TGen.LoadDetailDB('', '', '', Q, False);
            Ferme(Q);

            (*
            if Pos('A',Typrub.Value) <> 3 then begin
              FaitRequeteEcr(fb2) ;
              QEcr.Close ;
              ChangeSQL(QEcr) ;
              PrepareSQLODBC(QEcr) ;
            end ; *)
            TraiteRequeteGen(0,True,fb2) ;
          end ;

          St:=FaitRequeteGene(Cexlu2,'',fb2,False) ;
          if St <> '' then begin
            TGen.ClearDetail;
            Q := OpenSql(St, False);
            TGen.LoadDetailDB('', '', '', Q, False);
            Ferme(Q);
            TraiteRequeteGen(1,True,fb2) ;
          end ;
        end ;
      end;
    end ;
    MoveCur(False) ;
  end ;
  if Fliste.RowCount > 2 then Fliste.RowCount := Fliste.RowCount-1 ;
end ;

{---------------------------------------------------------------------------------------}
Procedure TOF_CPCONTROLERUB.TraiteRequeteGen(Ind : Integer ; Composite : Boolean ; Unfb : TFichierBase) ;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  Temp,
  TEcr : TOB;
  chpC,
  chpL : string;
begin
  TEcr := nil;
       if Unfb = fbAux    then begin TEcr := TEcrA; chpC := 'T_AUXILIAIRE'; chpL := 'T_LIBELLE';  end
  else if Unfb = fbBudGen then begin TEcr := TEcrB; chpC := 'BG_BUDGENE';   chpL := 'BG_LIBELLE'; end
  else if Unfb = fbGene   then begin TEcr := TEcrG; chpC := 'G_GENERAL';    chpL := 'G_LIBELLE';  end
  {JP 28/05/07 : FQ 20382 : il serait pas mal de gérer les sections !!}
  else if UnFb in [fbAxe1..fbAxe5] then begin TEcr := TEcrG; chpC := 'S_SECTION'; chpL := 'S_LIBELLE'; end;

  if TEcr = nil then Exit;

  for n := 0 to TGen.Detail.Count - 1 do begin
    Fliste.Cells[2, FListe.RowCount - 1] := TGen.Detail[n].GetValue(chpC);
    Fliste.Cells[3, FListe.RowCount - 1] := TGen.Detail[n].GetValue(chpL);
    Fliste.Cells[4, FListe.RowCount - 1] := MsgBox.Mess[5 + Ind] ;
    if not Composite then begin
      if Pos('A',Typrub.Value)=1 then begin
        if VH^.Cpta[AxeToFb(CbAxe.Value)].Chantier then Fliste.Cells[5, FListe.RowCount-1] := MsgBox.Mess[6]
                                                   else Fliste.Cells[5, FListe.RowCount-1] := MsgBox.Mess[5] ;
      end
      else begin
        Temp := TEcr.FindFirst([chpC], [TGen.Detail[n].GetValue(chpC)], True);
        if Temp <> nil then Fliste.Cells[5,FListe.RowCount-1] := MsgBox.Mess[6]
                       else Fliste.Cells[5,FListe.RowCount-1] := MsgBox.Mess[5] ;
      end ;
    end
    else begin
      if UnFb in [fbAxe1..fbAxe5] then begin
        if VH^.Cpta[AxeToFb(CbAxe.Value)].Chantier then Fliste.Cells[5,FListe.RowCount-1] := MsgBox.Mess[6]
                                                   else Fliste.Cells[5,FListe.RowCount-1] := MsgBox.Mess[5] ;
      end
      else begin
        Temp := TEcr.FindFirst([chpC], [TGen.Detail[n].GetValue(chpC)], True);
        if Temp <> nil then Fliste.Cells[5,FListe.RowCount-1] := MsgBox.Mess[6]
                       else Fliste.Cells[5,FListe.RowCount-1] := MsgBox.Mess[5] ;
      end ;
    end ;
    FListe.RowCount := FListe.RowCount+1 ;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPCONTROLERUB.FaitRequeteEcr(Unfb : TFichierBase) : string;
{---------------------------------------------------------------------------------------}
begin

  case Unfb of
    fbGene : Result := 'SELECT G_GENERAL FROM GENERAUX WHERE ((EXISTS(SELECT E_GENERAL FROM ECRITURE WHERE ' +
                       'E_GENERAL = G_GENERAL)) OR (EXISTS(SELECT Y_GENERAL FROM ANALYTIQ WHERE Y_GENERAL = G_GENERAL)))';

    fbAux  : Result := 'SELECT T_AUXILIAIRE FROM TIERS WHERE EXISTS(SELECT E_AUXILIAIRE FROM ECRITURE WHERE ' +
                       'E_AUXILIAIRE = T_AUXILIAIRE)';

    fbBudgen : Result := 'SELECT BG_BUDGENE FROM BUDGENE WHERE (EXISTS(SELECT BE_BUDGENE FROM BUDECR WHERE ' +
                         'BE_BUDGENE = BG_BUDGENE))' ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.ChargeRub ;
{---------------------------------------------------------------------------------------}
var
  C1, C2, Value, Datatype : string ;
  TZ : TZoomTable;
begin
  TZ := TZRUBCPTA;
  if FCtxSurRubrique = CtxBudget then
  begin
    if TradBud(FamRub.Value) = 'GEN' then begin Datatype := 'TZRUBBUDG'; TZ := TZRUBBUDG; end else
    if TradBud(FamRub.Value) = 'ANA' then begin Datatype := 'TZRUBBUDS'; TZ := TZRUBBUDS; end else
    if TradBud(FamRub.Value) = 'G/A' then begin Datatype := 'TZRUBBUDGS'; TZ := TZRUBBUDGS; end else
    if TradBud(FamRub.Value) = 'A/G' then begin Datatype := 'TZRUBBUDSG'; TZ := TZRUBBUDSG end;
    if DataType='' then exit;
    THEdit(GetControl('RUB1')).DataType := Datatype;
    THEdit(GetControl('RUB2')).DataType := Datatype;
    THEdit(GetControl('RUB1')).Plus := ' AND RB_BUDJAL LIKE "%' + CbBud.Value + '%"';
    THEdit(GetControl('RUB2')).Plus := ' AND RB_BUDJAL LIKE "%' + CbBud.Value + '%"';
    Value := CbBud.Value;
  end
  else
  begin
    if FCtxSurRubrique = CtxCycle then
    begin
      //THEdit(GetControl('RUB1')).Plus := ' AND RB_NATRUB = "CPT" AND RB_CLASSERUB = "CDR"' +
      //                                  ' AND RB_FAMILLES LIKE "%' + TradBud(FamRub.Value) + '%"';
      //THEdit(GetControl('RUB2')).Plus := THEdit(GetControl('RUB1')).Plus;
      Value := TradBud(FamRub.Value);
    end;

    if FCtxSurRubrique in [CtxRubrique, CtxLiasse] then
    begin
      THEdit(GetControl('RUB1')).DataType := 'TZRUBRIQUE';
      THEdit(GetControl('RUB2')).DataType := 'TZRUBRIQUE';
      THEdit(GetControl('RUB1')).Plus := 'AND RB_NATRUB="CPT" AND RB_FAMILLES LIKE "%' + TradBud(FamRub.Value) + '%"';
      THEdit(GetControl('RUB2')).Plus := 'AND RB_NATRUB="CPT" AND RB_FAMILLES LIKE "%' + TradBud(FamRub.Value) + '%"';
      Value := TradBud(FamRub.Value);
    end
  end;

  if not FiltreEncours then
  begin
    PremierDernierRub(TZ, Value, C1, C2) ;
    SetControlText('RUB1', C1);
    SetControlText('RUB2', C2);
  end ;
  PremierToDernier ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.InitFliste ;
{---------------------------------------------------------------------------------------}
var
  Typr : string ;
begin
  Typr := Typrub.Value ;
  FListe.ColAligns[4] := taCenter ;
  FListe.ColAligns[5] := taCenter ;
  if Pos('/', Typr) <= 0 then begin
    FListe.Cells[1,0] := MsgBox.Mess[1] ;
    FListe.Cells[3,0] := MsgBox.Mess[1] ;

    if Pos('A',Typr) = 1 then FListe.Cells[5,0] := MsgBox.Mess[16]
                         else FListe.Cells[5,0] := MsgBox.Mess[7] ;

    if RgChoix.ItemIndex = 0 then begin
      FListe.Cells[0,0] := MsgBox.Mess[2] ;
      case Lefb of
        fbGene         : FListe.Cells[2,0] := MsgBox.Mess[0] ;
        fbAux          : FListe.Cells[2,0] := MsgBox.Mess[9] ;
        fbBudgen       : FListe.Cells[2,0] := MsgBox.Mess[11] ;
        fbAxe1..fbAxe5 : FListe.Cells[2,0] := MsgBox.Mess[10] ;
      end ;

      FListe.Cells[4,0] := MsgBox.Mess[4] ;
    end

    else begin
      case Lefb of
        fbGene         : FListe.Cells[0,0] := MsgBox.Mess[0] ;
        fbAux          : FListe.Cells[0,0] := MsgBox.Mess[9] ;
        fbBudgen       : FListe.Cells[0,0] := MsgBox.Mess[11] ;
        fbAxe1..fbAxe5 : FListe.Cells[0,0] := MsgBox.Mess[10] ;
      end ;
      FListe.Cells[2,0] := MsgBox.Mess[2] ;
    end ;
  end

  else begin
    FListe.Cells[0,0] := MsgBox.Mess[2] ;
    if Typr = 'G/A' then begin
      FListe.Cells[2,0] := MsgBox.Mess[27] ;
      FListe.Cells[5,0] := MsgBox.Mess[31] ;
    end
    else if Typr = 'G/T' then begin
      FListe.Cells[2,0] := MsgBox.Mess[28] ;
      FListe.Cells[5,0] := MsgBox.Mess[7] ;
    end
    else if Typr = 'A/G' then begin
       FListe.Cells[2,0] := MsgBox.Mess[29] ;
       FListe.Cells[5,0] := MsgBox.Mess[32] ;
    end
    else if Typr = 'T/G' then begin
      FListe.Cells[2,0] := MsgBox.Mess[30] ;
      FListe.Cells[5,0] := MsgBox.Mess[7] ;
    end ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.PremierToDernier ;
{---------------------------------------------------------------------------------------}
var
  Cpt1,
  Cpt2 : string ;
begin
  if FiltreEnCours then Exit ;

  QuelEstLeFb ;
  PremierDernier(Lefb, Cpt1, Cpt2) ;
  SetControlText('C1', Cpt1);
  SetControlText('C2', Cpt2);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.QuelEstLeFb ;
{---------------------------------------------------------------------------------------}
begin
  if Pos('A', TypRub.Value) = 1 then begin
    case CbAxe.Value[2] of
      '1' : LeFb := fbAxe1 ;
      '2' : LeFb := fbAxe2 ;
      '3' : LeFb := fbAxe3 ;
      '4' : LeFb := fbAxe4 ;
      '5' : LeFb := fbAxe5 ;
    end ;
  end

  else begin
    case TypRub.Value[1] of
      'B' : Lefb := fbBudgen ;
      'G' : Lefb := fbGene ;
      'T' : Lefb := fbAux ;
    end ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
Procedure TOF_CPCONTROLERUB.VideLaListe ;
{---------------------------------------------------------------------------------------}
var
  i : Integer ;
begin
  for i := 0 to ListeRub.Count-1 do TObject(ListeRub.Objects[i]).Free ;
  ListeRub.Clear ;

  for i := 0 to ListeGen.Count-1 do TObject(ListeGen.Objects[i]).Free ;
  ListeGen.Clear ;
end ;

{---------------------------------------------------------------------------------------}
function TOF_CPCONTROLERUB.ChecheLeCompte : String ;
{---------------------------------------------------------------------------------------}
var
  i : Integer ;
begin
  Result := '';
  for i := Fliste.Row Downto 1 do
    if Fliste.Cells[0, i] <> '' then begin
      Result := Fliste.Cells[0, i];
      Exit;
    end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.ChargeTypRub ;
{---------------------------------------------------------------------------------------}
begin
  {JP 19/07/04 : il y avait de gros problèmes en budget}
  if FCtxSurRubrique = CtxBudget then begin
    TypRub.Values.Clear ; TypRub.Items.Clear ;
    TypRub.DataType:='TTRUBTYPEBUD' ; RempliComboFamRub ;
  end
  else
    FactoriseComboTypeRub(TypRub);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.RempliComboFamRub;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSql('SELECT CO_CODE,CO_LIBELLE FROM COMMUN WHERE CO_TYPE="RBB"', True);
  FamRub.Values.Clear;
  FamRub.Items.Clear;
  while not Q.Eof do begin
    FamRub.Values.Add(TradBud(Q.Fields[0].AsString, False));
    FamRub.Items.Add(Q.FindField('CO_LIBELLE').AsString);
    Q.Next;
  end;
  Ferme(Q);
end;



{---------------------------------------------------------------------------------------}
function TOF_CPCONTROLERUB.QuelfbComposite : TFichierBase ;
{---------------------------------------------------------------------------------------}
begin
  Result := fbGene ;
  if Pos('/', Typrub.Value) <= 0 then Exit ;

  case Typrub.Value[3] of
    'A' : begin
            case CbAxe.Value[2] of
              '1' : Result := fbAxe1 ;
              '2' : Result := fbAxe2 ;
              '3' : Result := fbAxe3 ;
              '4' : Result := fbAxe4 ;
              '5' : Result := fbAxe5 ;
             end ;
           end ;
    'G' : Result := fbGene ;
    'T' : Result := fbAux ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
function TOF_CPCONTROLERUB.ChercheCompteDansRub(CptRub,ComptGene : String) : Boolean ;
{---------------------------------------------------------------------------------------}
var
  St,
  StTemp : String ;
begin
  Result:=False ;
  while CptRub <> '' do begin
    St := ReadTokenSt(CptRub) ;
    StTemp := '' ;

    if St = '' then Continue ;
    if Pos('(', St) > 0 then St := Copy(St, 1, Pos('(', St)-1) ;

    if Pos(':',St)>0 then begin
      StTemp := Copy(St, Pos(':',St)+1, 200) ;
      System.Delete(St, Pos(':',St), 200) ;
    end ;

    if StTemp<>'' then begin
      St := BourrelaDonc(St,Lefb) ;
      StTemp := BourreLaDonc(StTemp,Lefb) ;
      if(ComptGene >= St) and (ComptGene <= StTemp) then begin
        Result := True ;
        Exit ;
      end ;
    end

    else begin
      if Length(St) <= VH^.Cpta[Lefb].Lg then
        if(Pos(St,ComptGene) = 1) or (St = ComptGene) then begin
          Result := True ;
          Exit;
        end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.InitMsg;
{---------------------------------------------------------------------------------------}
begin
  MsgBox := THMsgBox.Create(Application) ;

  MsgBox.Mess.Add('Compte Général');
  MsgBox.Mess.Add('Intitulé');

  if (FCtxSurRubrique in [CtxCycle, CtxLiasse]) then
  begin
    if FCtxSurRubrique = CtxCycle then
      MsgBox.Mess.Add('Cycle')
    else
      MsgBox.Mess.Add('Liasse');
  end
  else
    MsgBox.Mess.Add('Rubrique');

  MsgBox.Mess.Add('Contrôle en cours');
  MsgBox.Mess.Add('Est exclu');
  MsgBox.Mess.Add('Non');
  MsgBox.Mess.Add('Oui');
  MsgBox.Mess.Add('Mouvementé');
  MsgBox.Mess.Add('8;Contrôle du paramétrage;Confirmez-vous l''arrêt du traitement en cours ?;Q;YN;N;N;');
  MsgBox.Mess.Add('Compte auxiliaire');
  MsgBox.Mess.Add('Section analytique');
  MsgBox.Mess.Add('Compte budgétaire');
  MsgBox.Mess.Add('Voir compte général');
  MsgBox.Mess.Add('Voir compte auxiliaire');
  MsgBox.Mess.Add('Voir compte budgétaire');
  MsgBox.Mess.Add('Voir section analytique');
  MsgBox.Mess.Add('Chantier');
  MsgBox.Mess.Add('17;');
  MsgBox.Mess.Add('Variables');
  MsgBox.Mess.Add('Familles de variables');
  MsgBox.Mess.Add('Rubriques de détail');
  MsgBox.Mess.Add('&Variables');
  MsgBox.Mess.Add('&Rub. de détail');
  MsgBox.Mess.Add('Consultation des rubriques de détail en paramétrage');
  MsgBox.Mess.Add('&Rubrique de');
  MsgBox.Mess.Add('25;');
  MsgBox.Mess.Add('V&ariables de');
  MsgBox.Mess.Add('Général/Section');
  MsgBox.Mess.Add('Général/Tiers');
  MsgBox.Mess.Add('Section/Général');
  MsgBox.Mess.Add('Tiers/Général');
  MsgBox.Mess.Add('Mvté/chantier');
  MsgBox.Mess.Add('Chantier/Mvté');
  MsgBox.Mess.Add('&Type');
  MsgBox.Mess.Add('&Bud');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.FermeMsg ;
{---------------------------------------------------------------------------------------}
begin
  FreeAndNil(MsgBox);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLERUB.CbBudChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  ChargeRub;
  Q := OpenSQL('SELECT BJ_AXE FROM BUDJAL WHERE BJ_BUDJAL="'+cbBud.Value+'" ', True);
  If Not Q.Eof Then cbAxe.Value := Q.Fields[0].AsString;
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
function TOF_CPCONTROLERUB.TradBud(Value : string; VerGen : Boolean = True) : string;
{---------------------------------------------------------------------------------------}
begin
  if FCtxSurRubrique <> CtxBudget then
    Result := Value
  else if VerGen then begin
         if Value = 'CBG' then Result := 'GEN'
    else if Value = 'CBS' then Result := 'ANA'
    else if Value = 'G/S' then Result := 'G/A'
    else if Value = 'S/G' then Result := 'A/G';
  end else begin
         if Value = 'GEN' then Result := 'CBG'
    else if Value = 'ANA' then Result := 'CBS'
    else if Value = 'G/A' then Result := 'G/S'
    else if Value = 'A/G' then Result := 'S/G';
  end;
end;

procedure TOF_CPCONTROLERUB.OnClickBRechercher(Sender: Tobject);
begin
  FFindFirst := True;
  AFindDialog.Execute;
end;

// YMO 09/08/2006 FQ18639 Recherche par les jumelles
procedure TOF_CPCONTROLERUB.OnFindAFindDialog(Sender: TObject);
begin
  Rechercher(FListe, AFindDialog, FFindFirst);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/09/2006
Modifié le ... : 30/11/2006
Description .. : GCO - 22/09/2006 - FQ 18545
Suite ........ : GCO - 30/11/2006 - FQ 19248
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONTROLERUB.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
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
    BRechercher.Click
  else
    FormKeyDown(Sender, Key, Shift); // Pour touches standard AGL
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_CPCONTROLERUB ] ) ;

end.

