Unit CPCONTROLELIASSE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
    {$IFDEF EAGLCLIENT}
     MaineAGL,
     UtileAGL,
     eMul,
    {$ELSE}
     FE_Main,
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Mul,
     PrintDBG,
    {$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     UTOF,
     UTOB,
     ExtCtrls,
     HTB97,
     HMsgBox,      // PgiAsk
     Menus,
     HStatus,
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     Dialogs,      // TFindDialog
     ParamSoc,     // GetParamSocSecur
     uTobDebug,    // TobDebug(TGenVide)
     Ent1;         // VH^.

Type
  TInfoRub = class
    Libelle : string ;
    Famille : string ;
    Compte1 : string ;
    Exclu1  : string ;
    Compte2 : string ;
    Exclu2  : string ;
    TabLib  : string ;
    Sens    : string ;
  end ;

  TOF_CPCONTROLELIASSE = Class (TOF)

    Pages     : TPageControl;
    RGChoix   : TRadioGroup;

    Ctrl0     : THValComboBox;
    Ctrl1     : THEdit;
    Ctrl1_    : THEdit;
    Compte    : THEdit;
    Compte_   : THEdit;

    FListe   : THGrid;

    MVCBCompte      : THMultiValComboBox;
    BCherche        : TToolbarButton97;
    BReChercher     : TToolBarButton97;
    BStop           : TToolBarButton97;
    BExport         : TToolBarButton97;
    BMontantAffecte : TToolBarButton97;

    FFindDialog : TFindDialog;

    PopZoom     : TPopUpMenu;

    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnClickBRechercher (Sender : Tobject); virtual;
    procedure OnClose                  ; override ;

    procedure OnChangeCtrl0     (Sender : TObject);
    procedure OnClickRgChoix    (Sender : TObject);
    procedure OnClickBImprimer  (Sender : TObject);
    procedure OnClickBCherche   (Sender : TObject);
    procedure OnFindFFindDialog (Sender : TObject);
    procedure BRubClick         (Sender : TObject);
    procedure BGenClick         (Sender : TObject);
    procedure FListeDblClick    (Sender : TObject);
    procedure OnExitCompte      (Sender : TObject);
    procedure OnExitCompte_     (Sender : TObject);
    procedure OnClickBStop      (Sender : TObject);
    procedure OnClickBExport    (Sender : TObject);
    procedure OnClickBMontantAffecte (Sender : TObject);
    procedure OnPopUpPopZoom    (Sender : TObject);

  private

    TEcrG    : Tob;
    TRub     : Tob;
    TGen     : TOB;
    TGenVide : Tob;
    BGen     : TMenuItem;
    BRub     : TMenuItem;

    FBoLiasse     : Boolean;
    FBoLoading    : Boolean;
    FFindFirst    : Boolean;
    FBoAvecAff    : Boolean;
    FBoAvecNaf    : Boolean;
    FBoQueMvt     : Boolean;
    FBoStop       : Boolean;

    ListeRub, ListeGen : TStringList ;

    FormKeyDown  : TKeyEvent;

    FStWhereRub  : HString;
    FStMinRub    : HString;
    FStMaxRub    : HString;

    FStLiasseDossier : string;
    FStArgument : string;

    procedure OnKeyDownEcran (Sender : TObject; var Key : Word; Shift : TShiftState); virtual;
    procedure RempliListeRubGen(AvecGen : Boolean) ;
    procedure TraiteRequeteGen(Ind : Integer ; Composite : Boolean ; Unfb : TFichierBase; vTobRub : Tob) ;

    procedure InitFListe ;
    procedure ControleLesRubriques ;
    procedure ControleLesComptes ;
    procedure PositionneRequete ;
    procedure VideLaListe;


    function  FaitRequeteEcr(Unfb : TFichierBase) : string;
    function  FaitRequeteGene(Compte1,Cexlu1 : String ; Unfb : TFichierBase ; SurTabLib : Boolean) : String ;

    procedure RemplitCriteres;

    procedure OnAfterFormShow;
    procedure OnAfterSelectFiltre;

  end;

procedure CPLanceFiche_CPControleCycle ( vStArgument : string );  
procedure CPLanceFiche_CPControleLiasse( vStArgument : string );

Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  uLibWindows,        // TestJoker
  Windows,            // VK_
  HSysMenu,           // THSystemMenu
  HXlsPas,            // ExportGrid
  CREvParamCycle_Tom, // CPLanceFiche_CREVPARAMCYCLE
  RUBRIQUE_TOM,        // ParametrageRubrique
  CPGENERAUX_TOM,      // FicheGene
  CPStatutDossier_TOF, // SelectAgregatDictionnaire
  CPTOTLIASSE_TOF,
  uLibAnalytique,      // ChercheCompteDansRub
  ZCumul,
  ed_Tools;



const cColAgregat   = 1;
      cColIntitule1 = 2;
      cColCompte    = 3;
      cColIntitule2 = 4;
      cColExclu     = 5;
      cColMvt       = 6;
      cColSens      = 7;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/07/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure CPLanceFiche_CPControleCycle( vStArgument : string );
begin
  AGLLanceFiche('CP', 'CPCONTROLELIASSE', '', '', vStArgument + ';-');
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/03/2005
Modifié le ... :   /  /
Description .. : Procédure de lancement du Controle des Rubriques
Suite ........ : RUBRIQUE; BUDGET; CYCLE; LIASSE
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_CPControleLiasse( vStArgument : string );
begin
  AGLLanceFiche('CP', 'CPCONTROLELIASSE', '', '', vStArgument + ';X');
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnArgument (S : String ) ;
begin
  inherited;

  FBoLiasse := False;

  FStArgument := ReadTokenSt(S);

  // Récupérations des composants
  Pages    := TPageControl(GetControl('PAGES', True));
  FListe   := THGrid(GetControl('FLISTE1', True));
  RgChoix  := TRadioGroup(GetControl('RGCHOIX', True));
  Ctrl0    := THValComboBox(GetControl('CTRL0', True));
  Ctrl1    := THEdit(GetControl('CTRL1', True));
  Ctrl1_   := THEdit(GetControl('CTRL1_', True));
  Compte   := THEdit(GetControl('COMPTE', True));
  Compte_  := THEdit(GetControl('COMPTE_', True));

  MVCBCompte := THMultiValComboBox(GetControl('MVCBCOMPTE', True));

  BCherche        := TToolbarButton97(GetControl('BCHERCHE', True));
  BGen            := TPopupMenu(GetControl('POPZOOM')).Items[0];
  BRub            := TPopupMenu(GetControl('POPZOOM')).Items[1];
  BRechercher     := TToolBarButton97(GetControl('BRECHERCHER', True));
  BStop           := TToolBarButton97(GetControl('BSTOP', True));
  BExPort         := TToolBarButton97(GetControl('BEXPORT', True));
  BMontantAffecte := TToolBarButton97(GetControl('BMONTANTAFFECTE', True));

  FFindDialog := TFindDialog.Create(Ecran);
  PopZoom     := TPopUpMenu(GetControl('POPZOOM', True));

  // Branchement des événements
  TFMul(Ecran).OnAfterFormShow     := OnAfterFormShow;
  TFMul(Ecran).OnAfterSelectFiltre := OnAfterSelectFiltre;

  // GCO - 22/09/2006 - FQ 18545
  FormKeyDown     := Ecran.OnKeyDown;
  Ecran.OnKeyDown := OnKeyDownEcran;

  Fliste.OnDblClick   := FListeDblClick;
  BRechercher.OnClick := OnClickBRechercher;
  FFindDialog.OnFind  := OnFindFFindDialog;

  TToolbarButton97(GetControl('BIMPRIMER')).OnClick  := OnClickBImprimer;
  BCherche.OnClick        := OnClickBCherche;
  BStop.OnClick           := OnClickBStop;
  BExport.OnClick         := OnClickBExport;
  BMontantAffecte.OnClick := OnClickBMontantAffecte;

  Ctrl0.OnChange  := OnChangeCtrl0;
  RgChoix.OnClick := OnClickRgChoix;
  Compte.OnExit   := OnExitCompte;
  Compte_.OnExit  := OnExitCompte_;
  BRub.OnClick    := BRubClick;
  BGen.OnClick    := BGenClick;
  PopZoom.OnPopup := OnPopUpPopZoom;

  // Création des variables globales
  //InitMsg;
  TGen     := TOB.Create('', nil, -1);
  TRub     := TOB.Create('', nil, -1);
  TEcrG    := TOB.Create('', nil, -1);
  TGenVide := Tob.Create('', nil, -1);

  ListeRub := TStringList.Create;
  ListeGen := TStringList.Create;

{$IFDEF EAGLCLIENT}
  SetControlVisible('BPREV', False);
  SetControlVisible('BNEXT', False);
{$ENDIF}

  FStLiasseDossier := GetParamSocSecur('SO_CPCONTROLELIASSE', '');
  if FStLiasseDossier <> '' then
    Ctrl0.Value := FStLiasseDossier
  else
    Ctrl0.ItemIndex := 0;

  BStop.Visible := False;

  OnClickRgChoix(nil);

  MVCBCompte.SelectAll;
  if MVCBCompte.Tous then MVCBCompte.Text := TraduireMemoire('<<Tous>>');

  InitFliste ;

  FBoLiasse := ReadTokenSt(S) = 'X';

  if FBoLiasse then
  begin
    Ecran.HelpContext := 600400;
    Ecran.Caption := TraduireMemoire('Contrôle du paramétrage des rubriques');
    UpDateCaption(Ecran);
    SetControlProperty('LRUBFAM','CAPTION',  'Liasse');
    SetControlProperty('FAMRUB', 'DATATYPE', 'CPCONTROLELIASSE');
    BRub.Caption := 'Voir la liasse';
    FStWhereRub := 'RB_NATRUB = "CPT" AND ' +
                 'RB_CLASSERUB = "CPT" AND ' +
                 'RB_TYPERUB = "GEN" AND ' +
                 '(RB_PREDEFINI <> "DOS" OR ' +
                 '(RB_PREDEFINI = "DOS" AND RB_NODOSSIER = "' + V_PGI.NoDossier + '"))';
  end
  else
  begin
    Ecran.HelpContext := 600300;
    Ecran.Caption := TraduireMemoire('Contrôle du paramétrage des cycles de révision');
    UpDateCaption(Ecran);
    RGChoix.Buttons[0].Caption := TraduireMemoire('&Cycles');
    SetControlProperty('LRUBFAM','CAPTION',  'Plan');
    SetControlProperty('CTRL0', 'DATATYPE',  'CREVPLAN');
    SetControlProperty('TCTRL1', 'CAPTION',  'Cycle de');
    SetControlProperty('CTRL1',  'DATATYPE', 'CREVPARAMCYCLE');
    SetControlProperty('CTRL1_', 'DATATYPE', 'CREVPARAMCYCLE');
    BRub.Caption := TraduireMemoire('Voir le cycle de révision');
    Ctrl0.Value := VH^.Revision.Plan;
    BMontantAffecte.Visible := False;
  end;

  FBoLoading := True;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnAfterFormShow;
begin
  FBoLoading := False;

  if FStArgument <> '' then
  begin
    RgChoix.ItemIndex := 1;
    Compte.Text  := ReadTokenSt(FStArgument);
    Compte_.Text := Compte.Text;
    MVCBCompte.SelectAll;
    MVCBCompte.Text := TraduireMemoire('<<Tous>>');
    SetControlChecked('CBMVT', False);
    TToolbarButton97(GetControl('BCHERCHE', True)).Click;
  end;

// GCO - 02/10/2007 - FQ 21566
{$IFDEF CCSTD}
  SetControlChecked('CBMVT', False);
{$ENDIF}

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnLoad ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnDisplay () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnAfterSelectFiltre;
begin
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPCONTROLELIASSE.OnUpdate ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.BRubClick(Sender: TObject);
var lStTemp      : HString;
    lTobTemp : Tob;
begin
  if RgChoix.ItemIndex = 0 then
    lStTemp := (FListe.Cells[cColAgregat, FListe.Row])
  else
    lStTemp := (FListe.Cells[cColCompte, Fliste.Row]);

  if lStTemp <> '' then
  begin
    if FBoLiasse then
    begin
      lTobTemp := TRub.FindFirst(['RB_RUBRIQUE'],[lStTemp], False);
      if lTobTemp <> nil then
        ParametrageRubrique(lTobTemp.GetString('RB_RUBRIQUE') + ';' +
                            lTobTemp.GetString('RB_NODOSSIER') + ';' +
                            lTobTemp.GetString('RB_PREDEFINI') + ';' +
                            DateToStr(lTobTemp.GetDateTime('RB_DATEVALIDITE')), taConsult, CtxRubrique);
    end
    else
    begin
      // GCO - 30/07/2007 - FQ 21184
      lTobTemp := TRub.FindFirst(['CPC_CODECYCLE'],[lStTemp], False);
      CPLanceFiche_CREVPARAMCYCLE( lTobTemp.GetString('CPC_CODECYCLE') + ';' +
                                   lTobTemp.GetString('CPC_PREDEFINI') + ';', TaConsult );
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.BGenClick(Sender: TObject);
var lStTemp : HString;
begin
  if RgChoix.ItemIndex = 0 then
    lStTemp := (FListe.Cells[cColCompte, FListe.Row])
  else
    lStTemp := (FListe.Cells[cColAgregat, Fliste.Row]);

  if lStTemp <> '' then
    FicheGene(nil, '', lStTemp, taConsult, 0);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.FListeDblClick(Sender: TObject);
begin
  (*
  if (RgChoix.ItemIndex = 0) then
  begin
    if (FListe.Col = cColCompte or FListe.Col = cColIntitule2) then

  end
  else
  begin

  end;

    lStTemp := (FListe.Cells[cColCompte, FListe.Row])
  else
    lStTemp := (FListe.Cells[cColAgregat, Fliste.Row]);



  if FListe.Col =

   then

  else *)


  BRubClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnExitCompte(Sender: TObject);
begin
  if not TestJoker(Compte.Text) and (Compte.Text <> '') then
    Compte.Text := BourreEtLess(Compte.Text, fbGene);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnExitCompte_(Sender: TObject);
begin
  if not TestJoker(Compte_.Text) and (Compte_.Text <> '') then
    Compte_.Text := BourreEtLess(Compte_.Text, fbGene);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2006
Modifié le ... : 02/10/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnChangeCtrl0(Sender: TObject);
var lQuery   : TQuery;
    lStSQL   : HString;
    lStWhere : HString;
begin
  Ctrl1.Text  := '';
  Ctrl1_.Text := '';

  if FBoLiasse then
    lStWhere := ' RB_FAMILLES LIKE "' + Ctrl0.Value + '%" AND ' + FStWhereRub
  else
    lStWhere := ' CPC_PLANASSOCIE LIKE "' + Ctrl0.Value + '%"';

  CTRL1.Plus  := lStWhere;
  CTRL1_.Plus := CTRL1.Plus;

  try
    if FBoLiasse then
      lStSql := 'SELECT MIN(RB_RUBRIQUE) MIN, MAX(RB_RUBRIQUE) MAX ' +
                'FROM RUBRIQUE WHERE ' + lStWhere
    else
      lStSql := 'SELECT MIN(CPC_CODECYCLE) MIN, MAX(CPC_CODECYCLE) MAX ' +
                'FROM CREVPARAMCYCLE WHERE ' + lStWhere;

    lQuery := OpenSQL( lStSql, True );

    Ctrl1.Text  := lQuery.FindField('MIN').AsString;
    Ctrl1_.Text := lQuery.FindField('MAX').AsString;

    FStMinRub := Ctrl1.Text;
    FStMAxRub := Ctrl1_.Text;

  finally
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.RemplitCriteres;
var lQuery : TQuery;
begin
  if not TestJoker(Ctrl1.Text) then
  begin
    if (CTrl1.Text <> '') and (Ctrl1_.Text = '') then
      Ctrl1_.Text := Ctrl1.Text;
  end;

  if not TestJoker(Compte.Text) then
  begin
    if (Compte.Text = '') and (Compte_.Text = '') then
    begin
      try
        lQuery := OpenSql('SELECT MIN(G_GENERAL) MIN, MAX(G_GENERAL) MAX FROM GENERAUX', True);
        Compte.Text  := lQuery.FindField('MIN').AsString;
        Compte_.Text := lQuery.FindField('MAX').AsString;
      finally
        Ferme(lQuery);
      end;  
    end
    else
    begin
      if (Compte.Text <> '') and (Compte_.Text = '') then
        Compte_.Text := Compte.Text;
    end;    
  end;

  if (Ctrl1.Text <> FStMinRub) or (Ctrl1_.Text <> FStMaxRub) then
  begin
    MVCBCompte.Value := 'AFF;';
  end; // MBAMF
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnClickRgChoix(Sender: TObject);
begin
  Ctrl1.Enabled   := (RgChoix.ItemIndex = 0);
  Ctrl1_.Enabled  := (RgChoix.ItemIndex = 0);

  Compte.Enabled  := (RgChoix.ItemIndex = 1);
  Compte_.Enabled := (RgChoix.ItemIndex = 1);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnClickBCherche(Sender: TObject);
begin
  inherited;
  if FBoLoading then Exit;

  Fliste.BeginUpdate;
  try
    VideLaListe;
    InitFListe;
    RemplitCriteres;
    PositionneRequete;

    if MVCBCompte.Tous then
    begin
      FBoAvecAff := True;
      FBoAvecNaf := True;
    end
    else
    begin
      FBoAvecAff := (Pos('AFF', MVCBCompte.Value) > 0);
      FBoAvecNaf := (Pos('NAF', MVCBCompte.Value) > 0);
    end;

    FBoQueMvt := GetCheckBoxState('CBMVT') = CbChecked;

    BStop.Visible := True;

    case RgChoix.ItemIndex of

      0 : begin
            RempliListeRubGen(False);
            InitMove(ListeRub.Count, '');
            ControleLesRubriques;
            FiniMove;
          end;

      1 : begin
            RempliListeRubGen(True);
            InitMove(ListeGen.Count, '');
            ControleLesComptes;
            FiniMove;
          end;
    end;

  finally
    if Fliste.RowCount > 2 then Fliste.RowCount := Fliste.RowCount-1;
    FListe.Row := 1;
    THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
    FListe.EndUpdate ;
    if FListe.CanFocus then FListe.SetFocus;
    BStop.Visible := False;
    TFMUL(Ecran).CritModified := False;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnClickBImprimer(Sender: TObject);
{$IFDEF EAGLCLIENT}
var
  MaTob : TOB;
  T1    : TOB;
  n, p  : Integer;
  CritEdition : HString;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  MaTob := TOB.Create('§LISTEPARAM', nil, -1);
  try
    for n := 1 to Fliste.RowCount - 1 do
    begin
      T1 := TOB.Create('', MaTob, -1);
      for p := 1 to 6 do
      begin
        T1.AddChampSup('CT_COL' + IntToStr(p) +'A', True);
        T1.PutValue('CT_COL' + IntToStr(p) + 'A', Fliste.Cells[p, n]);
      end;
    end;

    CritEdition := 'TITRE=Contrôle du paramétrage des liasses';
    for p := 1 to 6 do
      CritEdition := CritEdition + '`COL' + IntToStr(p) + '=' + Fliste.Cells[p, 0];
    LanceEtatTOB('E', 'RUB', 'CON', MaTob, True, False, False, nil, '', 'Contrôle du paramétrage des liasses', False, 0, CritEdition);
  finally
    FreeAndNil(MaTob);
  end;
{$ELSE}
  PrintDBGrid (FListe, nil, Ecran.Caption, '');
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/10/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnClickBMontantAffecte(Sender: TObject);
var lStTemp : HString;
begin
  if RgChoix.ItemIndex = 0 then
    lStTemp := (FListe.Cells[cColAgregat, FListe.Row])
  else
    lStTemp := (FListe.Cells[cColCompte, Fliste.Row]);

  if lStTemp <> '' then
    CPLanceFiche_CPTOTLiasse(lStTemp);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnClickBExport (Sender : TObject);
var SD : TSaveDialog;
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
    if SD.Execute then
      ExportGrid(FListe, nil, SD.FileName, SD.FilterIndex, True);
  finally
    SD.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.PositionneRequete;
var lStSql : HString;
    lQuery : TQuery;
begin
  TGen.ClearDetail;
  TRub.ClearDetail;

  case RgChoix.ItemIndex of

  0 : begin
        if FBoLiasse then
          SelectAgregatDictionnaire('*', ConvertitCaractereJokers(Ctrl1, Ctrl1_, 'RB_RUBRIQUE'),
                                    VH^.EnCours.Fin, Ctrl0.Value, TRub)
        else
          TRub.LoadDetailFromSQL('SELECT CPC_CODECYCLE, CPC_LIBELLECYCLE, ' +
               'CPC_LISTECOMPTE, CPC_LISTEEXCLUSION, CPC_PLANASSOCIE, ' +
               'CPC_PREDEFINI FROM CREVPARAMCYCLE WHERE ' +
               'CPC_PLANASSOCIE LIKE "' + Ctrl0.Value + '%" AND ' +
               'CPC_CODECYCLE >= "' + Ctrl1.Text + '" AND ' +
               'CPC_CODECYCLE <= "' + Ctrl1_.Text + '" ' +
               'ORDER BY CPC_CODECYCLE');
      end;

  1 : begin
        if FBoLiasse then
          SelectAgregatDictionnaire('*', '', VH^.EnCours.Fin, Ctrl0.Value, TRub)
        else
          TRub.LoadDetailFromSQL('SELECT CPC_CODECYCLE, CPC_LIBELLECYCLE, ' +
               'CPC_LISTECOMPTE, CPC_LISTEEXCLUSION, CPC_PLANASSOCIE, ' +
               'CPC_PREDEFINI FROM CREVPARAMCYCLE WHERE ' +
               'CPC_PLANASSOCIE LIKE "' + Ctrl0.Value + '%" ORDER BY CPC_CODECYCLE');
        try
          // GCO - 11/07/2007 - FQ 20991 ajout de G_CORRESP1, G_CORRESP2
          lStSql := 'SELECT G_GENERAL, G_LIBELLE, G_CORRESP1, G_CORRESP2 ' +
                    'FROM GENERAUX WHERE ' + ConvertitCaractereJokers(Compte, Compte_, 'G_GENERAL') +
                    ' ORDER BY G_GENERAL';
          lQuery := OpenSql(lStSql, False);
          TGen.LoadDetailDB('', '', '', lQuery, False);
        finally
          Ferme(lQuery);
        end;
      end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOF_CPCONTROLELIASSE.ControleLesRubriques;
var Compte1, Cexlu1, Sens, St : HString;
    i,j     : Integer ;
    Q       : TQuery;
    lStDico : HString;
    lTobRub : Tob;
    lBoCptMvt : Boolean;
    lStCompte : HString;
    lBoAjouteEntete : Boolean;
begin
  TEcrG.ClearDetail;

  St := FaitRequeteEcr(fbGene);
  Q := OpenSql(St, True);
  TEcrG.LoadDetailDB('', '', '', Q, False);
  Ferme(Q);

  lStDico := Ctrl0.Value;
  lBoAjouteEntete := False;
  FBoStop         := False;

  TGenVide.ClearDetail;
  if (ListeRub.Count <> 0) and ( FBoAvecNaf )   then
  begin
    Q := OpenSql('SELECT G_GENERAL, G_LIBELLE FROM GENERAUX ORDER BY G_GENERAL', True);
    TGenVide.LoadDetailDB('', '', '', Q, False);
    Ferme(Q);
  end;

  for i := 0 to ListeRub.Count-1 do
  begin
    MoveCur(False);
    if Pos(lStDico, TInfoRub(ListeRub.Objects[i]).Famille) > 0 then
    begin
      lTobRub := Tob.Create('', nil, -1);
      lTobRub.AddChampSupValeur('AGREGAT', ListeRub.Strings[i]);
      lTobRub.AddChampSupValeur('INTITULE1', TInfoRub(ListeRub.Objects[i]).Libelle);

      Compte1 := TInfoRub(ListeRub.Objects[i]).Compte1;
      Cexlu1  := TInfoRub(ListeRub.Objects[i]).Exclu1;

      Sens := '';
      if Pos('(', Compte1) > 0 then
        Sens := Copy(Compte1, Pos('(', Compte1)+1, 2);

      St := FaitRequeteGene(Compte1, CExlu1, FbGene, TInfoRub(ListeRub.Objects[i]).TabLib='X') ;
      if St = '' then Continue ;
      TGen.ClearDetail;
      Q := OpenSql(St, False);
      TGen.LoadDetailDB('', '', '', Q, False);
      Ferme(Q);
      TraiteRequeteGen(0, False, fbGene, lTobRub) ;

      St := FaitRequeteGene(Cexlu1, '', fbGene, False) ;
      if St <> '' then
      begin
        TGen.ClearDetail;
        Q := OpenSql(St, False);
        TGen.LoadDetailDB('', '', '', Q, False);
        Ferme(Q);
        TraiteRequeteGen(1, False, fbGene, lTobRub) ;
      end ;

      Application.ProcessMessages;
      if FBoStop then
      begin
        if PgiAsk('Confirmez-vous l''arrêt du traitement en cours ?', Ecran.CAPtion) = MrYes then
          Break // Arrêt demandé par l'utilisateur
        else
          FBoStop := False;
        Application.ProcessMessages;
      end;

      if lTobRub.Detail.Count <> 0 then
      begin
        FListe.InsertRow(FListe.RowCount-1);
        FListe.Cells[cColAgregat,FListe.RowCount-2]   := lTobRub.GetString('AGREGAT');
        FListe.Cells[cColIntitule1,FListe.RowCount-2] := lTobRub.GetString('INTITULE1');
        for j := 0 to lTobRub.Detail.Count - 1 do
        begin
          if j <> 0 then
          begin
            FListe.InsertRow(FListe.RowCount-1);
          end;
          FListe.Cells[cColCompte,FListe.RowCount-2]    := lTobRub.Detail[j].GetString('COMPTE');
          FListe.Cells[cColIntitule2,FListe.RowCount-2] := lTobRub.Detail[j].GetString('INTITULE2');
          FListe.Cells[cColExclu, FListe.RowCount-2]    := lTobRub.Detail[j].GetString('EXCLU');
          FListe.Cells[cColMvt, FListe.RowCount-2]      := lTobRub.Detail[j].GetString('MVT');
          FListe.Cells[cColSens, FListe.RowCount-2]     := Sens;
        end;
        lTobRub.ClearDetail;
      end;
      FreeAndNil(lTobRub);
    end;
  end;

  if TGenVide = nil then Exit;

  for i := TGenVide.Detail.Count - 1 downto 0 do
  begin
    lStCompte := TGenVide.Detail[i].GetString('G_GENERAL');

    lBoCptMvt := (TEcrG.FindFirst(['G_GENERAL'],[lStCompte], True) <> nil);
    if (FBoQueMvt) and (not lBoCptMvt) then Continue;

    lBoAjouteEntete := True;
    FListe.InsertRow(1);

    FListe.Cells[cColCompte,1]    := lStCompte;
    FListe.Cells[cColIntitule2,1] := TGenVide.Detail[i].GetString('G_LIBELLE');
    FListe.Cells[cColMvt, 1]      := IIF(lBoCptMvt, 'X', '-');
    FListe.Cells[cColSens, 1]     := '';
  end;

  if lBoAjouteEntete then
  begin
    if FBoLiasse then
    begin
      FListe.Cells[cColAgregat,1]   := TraduireMemoire('<< Sans agrégat >>');
      FListe.Cells[cColIntitule1,1] := TraduireMemoire('Comptes hors liasse fiscale');
    end
    else
    begin
      FListe.Cells[cColAgregat,1]   := TraduireMemoire('<< Sans cycle >>');
      FListe.Cells[cColIntitule1,1] := TraduireMemoire('Comptes hors cycle de révision');
    end;
  end;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOF_CPCONTROLELIASSE.ControleLesComptes ;
var Compte       : HString;
    CptRub,
    CptExRub     : HString ;
    lStChampCpte : HString;
    i, j         : Integer ;
    Q            : TQuery;
    lStDico      : HString;
    lBoCptMvt    : Boolean;
    lTobGen      : Tob;
    lTobFille    : Tob;
    lTobFind     : Tob;
    lTobTemp     : Tob;
begin
  TEcrG.ClearDetail;

  Q := OpenSql(FaitRequeteEcr(fbGene), True);
  TEcrG.LoadDetailDB('', '', '', Q, False);
  Ferme(Q);

  lStDico := Ctrl0.Value;
  FBoStop := False;

  if GetParamSocSecur('SO_CPLIASSESURCORRESP',False) and (FBoLiasse) then
  begin
    if GetParamSocSecur('SO_CPLIASSEPLANCORRESP','') = '1' then
      lStChampCpte := 'G_CORRESP1'
    else
      if GetParamSocSecur('SO_CPLIASSEPLANCORRESP','') = '2' then
        lStChampCpte := 'G_CORRESP2'
      else
        lStChampCpte := 'G_GENERAL';
  end
  else
    lStChampCpte := 'G_GENERAL';

  // GCO - 11/07/2007 - FQ 20991
  for i := 0 to TGen.Detail.Count -1 do
  //for i := 0 to ListeGen.Count - 1 do
  begin
    MoveCur(False) ;

    lTobTemp  := TGen.Detail[i];
    Compte    := lTobTemp.GetString('G_GENERAL');
    lTobFind  := TEcrG.FindFirst(['G_GENERAL'], [Compte], True);
    lBoCptMvt := (lTobFind <> nil);

    if (FBoQueMvt) and (not lBoCptMvt) then Continue;

    // Si on travaille sur les comptes de correspondance, on regarde si le compte
    // de correspondance est renseigné.
    if lstChampCpte <> 'G_GENERAL' then
    begin
      if Trim(lTobTemp.GetString(lstChampCpte)) <> '' then
        Compte := BourreOuTronque(lTobTemp.GetString(lstChampCpte), fbGene);
    end;

    lTobGen := Tob.Create('', nil, -1);
    lTobGen.AddChampSupValeur('AGREGAT', lTobTemp.GetString('G_GENERAL'));
    lTobGen.AddChampSupValeur('INTITULE1', lTobTemp.GetString('G_LIBELLE'));
    lTobGen.AddChampSupValeur('AFFECTE', '-');
    lTobGen.AddChampSupValeur('MVT', IIF(lBoCptMvt, 'X', '-'));

    for j := 0 to ListeRub.Count - 1 do
    begin
      if (Pos(lStDico, TInfoRub(ListeRub.Objects[j]).Famille) > 0) then
      begin
        CptRub   := TInfoRub(ListeRub.Objects[j]).Compte1;
        CptExRub := TInfoRub(ListeRub.Objects[j]).Exclu1;

        if ChercheCompteDansRub(CptRub, Compte) then
        begin
          lTobGen.SetString('AFFECTE', 'X');
          if (not FBoAvecAff) and (FBoAvecNaf) then Break;

          lTobFille := Tob.Create('', lTobGen, -1);
          lTobFille.AddChampSupValeur('COMPTE', ListeRub.Strings[j]);
          lTobFille.AddChampSupValeur('INTITULE2', TInfoRub(ListeRub.Objects[j]).Libelle);
          lTobFille.AddChampSupValeur('EXCLU', '-');
        end;

        if CptExRub <> '' then
        begin
          if ChercheCompteDansRub(CptExRub,Compte) then
          begin
            lTobGen.SetString('AFFECTE', 'X');
            if (not FBoAvecAff) and (FBoAvecNaf) then Break;
            lTobFille := Tob.Create('', lTobGen, -1);
            lTobFille.AddChampSupValeur('COMPTE', ListeRub.Strings[j]);
            lTobFille.AddChampSupValeur('INTITULE2', TInfoRub(ListeRub.Objects[j]).Libelle);
            lTobFille.AddChampSupValeur('EXCLU', 'X');
          end ;
        end ;
      end ; // if Pos(lStDico
    end ; // for j := 0

    // Affichage du compte général et de ses agrégats
    if lTobGen = nil then Continue;

    if ((FBoAvecAff) and (lTobGen.GetString('AFFECTE')='X')) or
       ((FBoAvecNaf) and (lTobGen.GetString('AFFECTE')='-')) then
    begin
      FListe.InsertRow(FListe.RowCount-1);
      FListe.Cells[cColAgregat,FListe.RowCount-2]   := lTobGen.GetString('AGREGAT');
      FListe.Cells[cColIntitule1,FListe.RowCount-2] := lTobGen.GetString('INTITULE1');
      FListe.Cells[cColMvt, FListe.RowCount-2]      := lTobGen.GetString('MVT');

      for j := 0 to lTobGen.Detail.Count - 1 do
      begin
        if j <> 0 then
        begin
          FListe.InsertRow(FListe.RowCount-1);
        end;
        FListe.Cells[cColCompte,FListe.RowCount-2]    := lTobGen.Detail[j].GetString('COMPTE');
        FListe.Cells[cColIntitule2,FListe.RowCount-2] := lTobGen.Detail[j].GetString('INTITULE2');
        FListe.Cells[cColExclu, FListe.RowCount-2]    := lTobGen.Detail[j].GetString('EXCLU');
        FListe.Cells[cColSens, FListe.RowCount-2]     := lTobGen.Detail[j].GetString('SENS');
      end;
    end;

    lTobGen.ClearDetail;
    FreeAndNil(lTobGen);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/10/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOF_CPCONTROLELIASSE.FaitRequeteGene(Compte1, Cexlu1: String; Unfb: TFichierBase; SurTabLib: Boolean): String;
Var
  Sql ,St,
  Where,
  Cod ,Lib,
  Table  : string ;
  lStCorresp : string;
begin
  Result:='' ;
  Cod   := 'G_GENERAL';
  Lib   := 'G_LIBELLE';
  Table := 'GENERAUX' ;

  Where := AnalyseCompte(Compte1, Unfb, False, SurTabLib) ;

  if Where <> '' then
  begin
    // Analyse sur les comptes de correspondances 
    if GetParamSocSecur('SO_CPLIASSESURCORRESP',False) and (FBoLiasse) then
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

    Sql := 'SELECT ' + Cod + ',' + Lib + ' FROM ' + Table + '';
    Sql := Sql + ' Where ' + Where ;
  end
  else
    Exit;

  if Cexlu1 <>'' then
  begin
    Where := AnalyseCompte(Cexlu1, Unfb, True, False) ;
    if Where <> '' then
    begin
      // Analyse sur les comptes de correspondances
      if GetParamSocSecur('SO_CPLIASSESURCORRESP',False) and (FBoLiasse) then
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

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOF_CPCONTROLELIASSE.TraiteRequeteGen(Ind : Integer ; Composite : Boolean ; Unfb : TFichierBase; vTobRub : Tob);
var i : Integer;
  TEcr : TOB;
  chpC, chpL : string;
  lStCompte  : string;
  lBoCptMvt  : Boolean;
  lTobTemp   : Tob;
  lTobGen    : Tob;
begin
  TEcr := TEcrG;
  chpC := 'G_GENERAL';
  chpL := 'G_LIBELLE';

  if TEcr = nil then Exit;

  for i := 0 to TGen.Detail.Count - 1 do
  begin
    lStCompte := TGen.Detail[i].GetValue(chpC);

    // On enleve le compte de la liste des comptes qui sert à trouver les non affectés
    if TGenVide <> nil then
    begin
      lTobTemp := TGenVide.FindFirst(['G_GENERAL'],[lStCompte], True);
      if lTobTemp <> nil then
        lTobTemp.Free;
    end;

    if (not FBoAvecAff) and (FBoAvecNaf) then Continue;

    lBoCptMvt := (TEcr.FindFirst([chpC], [lStCompte], True) <> nil);
    // Si que les mouvementés et qu'il ne l'est pas, on passe au suivant
    if (FBoQueMvt) and (not lBoCptMvt) then Continue;

    lTobGen := Tob.Create('', vTobRub, -1);
    lTobGen.AddChampSupValeur('COMPTE', lStCompte);
    lTobGen.AddChampSupValeur('INTITULE2', TGen.Detail[i].GetString(chpL));
    lTobGen.AddChampSupValeur('MVT', IIF(lBoCptMvt, 'X', '-'));
    lTobGen.AddChampSupValeur('EXCLU', IIF( ind = 0, '-', 'X'));
//    lTobGen.AddChampSupValeur('SENS',     );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.RempliListeRubGen(AvecGen : Boolean);
var X : TInfoRub;
    Cod, Lib : String ;
    n : Integer;
begin
  for n := 0 to TRub.Detail.Count - 1 do
  begin
    X := TInfoRub.Create ;

    if FBoLiasse then
    begin
      X.Libelle := TRub.Detail[n].GetString('RB_LIBELLE'   );
      X.Famille := TRub.Detail[n].GetString('RB_FAMILLES'  );
      X.Compte1 := TRub.Detail[n].GetString('RB_COMPTE1'   );
      X.Exclu1  := TRub.Detail[n].GetString('RB_EXCLUSION1');
      X.Compte2 := TRub.Detail[n].GetString('RB_COMPTE2'   );
      X.Exclu2  := TRub.Detail[n].GetString('RB_EXCLUSION2');
      X.TabLib  := TRub.Detail[n].GetString('RB_TABLELIBRE');
      X.Sens    := TRub.Detail[n].GetString('RB_SIGNERUB');
      ListeRub.AddObject(TRub.Detail[n].GetString('RB_RUBRIQUE'),X) ;
    end
    else
    begin
      X.Libelle := TRub.Detail[n].GetString('CPC_LIBELLECYCLE');
      X.Famille := TRub.Detail[n].GetString('CPC_PLANASSOCIE');
      X.Compte1 := TRub.Detail[n].GetString('CPC_LISTECOMPTE');
      X.Exclu1  := TRub.Detail[n].GetString('CPC_LISTEEXCLUSION');
      X.Compte2 := TRub.Detail[n].GetString('');
      X.Exclu2  := TRub.Detail[n].GetString('');
      X.TabLib  := TRub.Detail[n].GetString('');
      X.Sens    := TRub.Detail[n].GetString('');
      ListeRub.AddObject(TRub.Detail[n].GetValue('CPC_CODECYCLE'),X) ;
    end;

  end ;

  if AvecGen then
  begin
    Cod := 'G_GENERAL';
    Lib := 'G_LIBELLE';
  end ;

  for n := 0 to TGen.Detail.Count - 1 do
  begin
    X := TInfoRub.Create ;
    X.Libelle := TGen.Detail[n].GetValue(Lib);
    ListeGen.AddObject(TGen.Detail[n].GetValue(Cod), X) ;
  end ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCONTROLELIASSE.FaitRequeteEcr(Unfb : TFichierBase) : string;
begin
  Result := 'SELECT G_GENERAL FROM GENERAUX WHERE ((EXISTS(SELECT E_GENERAL FROM ECRITURE WHERE ' +
            'E_GENERAL = G_GENERAL)) OR (EXISTS(SELECT Y_GENERAL FROM ANALYTIQ WHERE Y_GENERAL = G_GENERAL)))';
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.InitFliste;
begin
  FListe.ColALigns[cColAgregat]   := TaCenter;
  FListe.ColALigns[cColIntitule1] := TaLeftJustify;
  FListe.ColALigns[cColCompte]    := TaCenter;
  FListe.ColALigns[cColIntitule2] := TaLeftJustify;
  FListe.ColTypes[cColExclu]      := 'B';
  FListe.ColAligns[cColExclu]     := TaCenter;
  FListe.ColFormats[cColExclu]    := IntToStr( Byte( csCheckBox ) );
  FListe.ColTypes[cColMvt]        := 'B';
  FListe.ColAligns[cColMvt]       := TaCenter;
  FListe.ColFormats[cColMvt]      := IntToStr( Byte( csCheckBox ) );

  FListe.ColALigns[cColSens]      := TaCenter;
  if not FBoLiasse then
  begin
    FListe.ColWidths[cColSens]    := -1;
    FListe.ColLengths[cColSens]   := -1;
  end;

(*
  FListe.ColAligns[cColPredefini] := TaCenter;
{$IFDEF GIL}
  FListe.ColWidths[cColPredefini]  := 100;
  FListe.ColLengths[cColPredefini] := 100;
{$ELSE}
  FListe.ColWidths[cColPredefini]  := -1;
  FListe.ColLengths[cColPredefini] := -1;
{$ENDIF}
*)

  if RgChoix.ItemIndex = 0 then
  begin
    if FBoLiasse then
      FListe.Cells[cColAgregat, 0] := TraduireMemoire('Agrégat')
    else
      FListe.Cells[cColAgregat, 0] := TraduireMemoire('Cycle');

    FListe.Cells[cColCompte, 0]  := TraduireMemoire('Compte général');
  end
  else
  begin
    FListe.Cells[cColAgregat, 0] := TraduireMemoire('Compte général');
    if FBoLiasse then
      FListe.Cells[cColCompte, 0]  := TraduireMemoire('Agrégat')
    else
      FListe.Cells[cColCompte, 0] := TraduireMemoire('Cycle');
  end;

  FListe.Cells[cColIntitule1, 0] := TraduireMemoire('Intitulé');
  FListe.Cells[cColIntitule2, 0] := TraduireMemoire('Intitulé');
  FListe.Cells[cColExclu, 0]     := TraduireMemoire('Est Exclu');
  FListe.Cells[cColMvt, 0]       := TraduireMemoire('Mouvementé');
  FListe.Cells[cColSens, 0]      := TraduireMemoire('Sens');

  FListe.VidePile(False);

  THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
Procedure TOF_CPCONTROLELIASSE.VideLaListe ;
var i : Integer ;
begin
  for i := 0 to ListeRub.Count-1 do TObject(ListeRub.Objects[i]).Free ;
  ListeRub.Clear ;

  for i := 0 to ListeGen.Count-1 do TObject(ListeGen.Objects[i]).Free ;
  ListeGen.Clear ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnClickBRechercher(Sender: Tobject);
begin
  FFindFirst := True;
  FFindDialog.Execute;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnFindFFindDialog(Sender: TObject);
begin
  Rechercher(FListe, FFindDialog, FFindFirst);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/09/2006
Modifié le ... :   /  /
Description .. : GCO - 22/09/2006 - FQ 18545
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = VK_F11) then
  begin
    PopZoom.Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);
  end
  else
  if (Key = VK_F12) then
  begin
    if not FListe.Focused then
      FListe.SetFocus
    else
      Ctrl0.SetFocus;
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
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnPopUpPopZoom(Sender: TObject);
//var lStGeneral : string;
//    lStRubrique : string;
begin
  (*
  BGen.Enabled := False;
  BRub.Enabled := False;

  if RgChoix.ItemIndex = 0 then
  begin
    lStRubrique := (FListe.Cells[cColAgregat, FListe.Row]);
    lStGeneral  := (FListe.Cells[cColCompte, FListe.Row]);
  end
  else
  begin
    lStRubrique := (FListe.Cells[cColCompte, FListe.Row]);
    lStGeneral  := (FListe.Cells[cColAgregat, FListe.Row]);
  end;

  if lStRubrique <> '' then
  begin
    if ExisteSQL('SELECT RB_RUBRIQUE FROM RUBRIQUE WHERE RB_RUBRIQUE = "' + lStRubrique + '" AND ' +
                'RB_FAMILLES LIKE "' + Ctrl0.Value + '%" AND ' + FStWhereRub) then
      BRub.Enabled := True;
  end;

  if lStGeneral <> '' then
  begin
    if ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + lStGeneral + '"') then
      BGen.Enabled := True;
  end;
  *)
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnClose;
begin
  VideLaListe ;
  FreeAndNil(TGenVide);
  FreeAndNil(TGen);
  FreeAndNil(TRub);
  FreeAndNil(TEcrG);
  FreeAndNil(ListeRub);
  FreeAndNil(ListeGen);
  FreeAndNil(FFindDialog);
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONTROLELIASSE.OnClickBStop(Sender: TObject);
begin
  FBoStop := True;
end;

////////////////////////////////////////////////////////////////////////////////


Initialization
  registerclasses ( [ TOF_CPCONTROLELIASSE ] ) ;

end.

