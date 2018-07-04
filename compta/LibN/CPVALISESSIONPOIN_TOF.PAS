{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 21/05/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : IACCEPTBUDGETENG ()
Mots clefs ... : TOF;IACCEPTBUDGETENG
*****************************************************************}
Unit CPVALISESSIONPOIN_TOF ;

Interface

Uses StdCtrls, windows,
     Controls, 
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     fe_Main,
{$else}
     eMul,
     MaineAgl,
{$ENDIF}
     uTob,
     forms, Vierge, menus, Hsysmenu,Graphics,Grids, 
     sysutils, types, lookup,
     ComCtrls,
     HCtrls,
     Ent1, 
     HEnt1,
     HMsgBox,
     UTOF,
     HTb97,
     EntImo,
     IMAPLIST,
     IMAP,
     IBUDGET_ANCETRE_TOF,
     IOBJBUDGET,
     IOBJBUDGETLIGNE,
     IOBJBUDGETENG,
     IENTBUDGET,
     IBUDGETENGINE ;

Type
  TOF_CPVALISESSIONPOIN = Class (TOF_IBUDGET_ANCETRE)

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    optionBudget : TOptionBudget ;

    bModifFg            : Boolean ;
    bForceUpdateFg      : Boolean ;
    bAffinageFg         : Boolean ;
    onLoadFg, onCloseFg : Boolean ;
    bValiderFg          : Boolean ;

    iBudgetList : TIMapList ;
    GRID        : THGrid ;
    BUDGET      : THEdit ;
    RowColor1   : Integer ;
    RowColor2   : Integer ;
    RowColor3   : Integer ;
    bUpdateFg   : Boolean ;
    svChoix     : String ;


    procedure bInfoBudgetOnclick(Sender: TObject) ;
    procedure bInfoLigneOnclick(Sender: TObject) ;
    procedure bInfoEngagementOnclick(Sender: TObject) ;

    procedure BAffinageOnClick(Sender: TObject) ;


    procedure bChoixBudgetOnclick(Sender : TObject) ;

    procedure bValiderOnclick(Sender : TObject) ;

    procedure bMettreEnAttenteOnclick(Sender : TObject) ;
    procedure bRefuserOnclick(Sender : TObject) ;
    procedure bAccepterOnclick(Sender : TObject) ;
    procedure bDefaireOnclick(Sender : TObject) ;

    procedure ChargeLesBitmaps();
    procedure ChangeCellValue(ACol,ARow : integer ; Value : string);
    procedure ChangeValueSelection(MyRect : TGridRect ; Value : string);

    Function ControlerAvantValidation() : Boolean ;

    procedure EdtCODEBUDOnElipsisClick(Sender : TObject) ;

    procedure FormaterColonnes() ;
    procedure FormaterGRID(theIBudget : TIBudget) ;
    procedure FormaterGRIDAllBudget() ;

    Function  GererEtatModifies() : Boolean ;
    procedure GoAcceptation() ;

    procedure GRID_OnPopUp(sender: TObject);
    procedure GRID_AccepterOnClick(sender: TObject);
    procedure GRID_RefuserOnClick(sender: TObject);
    procedure GRID_MettreEnAttenteOnClick(sender: TObject);
    procedure GRID_RevenirEtatInitialOnClick(sender: TObject);

    procedure GRIDFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GRIDClick(Sender : TObject) ;
    procedure GRIDDblClick(Sender: TObject);
    procedure GRIDPostDrawCell (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);

    Procedure GRIDGetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    procedure GRIDMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GRIDMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GRIDMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure GRIDDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure GRIDEndDrag(Sender, Target: TObject; X, Y: Integer);

    procedure InitialiserTitre();
    procedure MajHelp ;

    Function MonterData(theBudgetCd : String) : Boolean  ;

    procedure SelectLaCol(ACol : integer);
    procedure SelectLaRow(ARow : integer);
 end ;

  procedure LancerAcceptationBudgetEng(theOptionBudget : TOptionBudget) ;

  const

     COL_NIVEAU     = 1 ;
     COL_BUDGET     = 2 ;
     COL_LIGNE      = 3 ;
     COL_ENGAGEMENT = 4 ;
     COL_INTITULE   = 5 ;
     COL_MONTANT    = 6 ;
     COL_ETAT       = 7 ;

     IcoAccepter = '#ICO#47';
     IcoAccepterUpdate = '#ICO#89';

     IcoRefuser = '#ICO#46' ;
     IcoRefuserUpdate = '#ICO#88' ;

     IcoEnAttente = '#ICO#45' ;
     IcoEnAttenteUpdate = '#ICO#87' ;

     IcoBudget     =  '#ICO#57';
     IcoLigne      =  '#ICO#76';
     IcoEngagement =  '#ICO#14';

Implementation
 uses IBUDGET_TOF, ISAIBUDGETLIG_TOF, IAFFINEBUDGET_TOF, ISAIBUDGETENG_TOF ;

procedure LancerAcceptationBudgetEng(theOptionBudget : TOptionBudget) ;
var parmetre : String ;
BEGIN
 If Not ExJaiLeDroitConcept(tConcept(ciAcceptEng),FALSE) Then
 begin
   PgiBox(TraduireMemoire('Vous ne pouvez utiliser cette fonction car vous n''avez pas accès à l''acceptation des engagements. (Concept groupe utilisateur)'),'');
   Exit ;
 end;
 parmetre := parmetre + '<OPTION>' + IntToStr(Integer(theOptionBudget)) + '</OPTION>' ;
 AGLLanceFiche('IM','CPVALISESSIONPOIN','','',parmetre) ;
END ;


procedure TOF_CPVALISESSIONPOIN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPVALISESSIONPOIN.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPVALISESSIONPOIN.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPVALISESSIONPOIN.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPVALISESSIONPOIN.OnCancel () ;
begin
  Inherited ;
end ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.OnArgument (S : String ) ;
VAR
 BTN : TToolbarButton97 ;
begin
  Inherited ;
  OptionBudget  := ToptionBudget(StrToInt(RechercherParameterXML('OPTION', S))) ;


  iBudgetList := Nil ;
  onLoadFg := True ;
  onCloseFg := False ;
  bForceUpdateFg := False ;
  bAffinageFg := False ;
  bValiderFg := False ;



  ChargeLesBitmaps();

  RowColor1 := THEdit(GetControl('COLOR1')).Color;
  RowColor2 := THEdit(GetControl('COLOR2')).Color;
  RowColor3 := THEdit(GetControl('COLOR3')).Color;

  RegLoadToolbarPos(Ecran, 'CPVALISESSIONPOIN') ; //* Recup position fenêtre

  TOBAFFINAGE :=  TOB.Create('', Nil, -1) ;
  InitTobAffinage('ACCLIGNEBUD') ;


  bModifFg := False ;
  bUpdateFg := False ;

  TPopupMenu(GetControl('MNU_GRID')).OnPopup  := GRID_OnPopUp ;

  BUDGET := THEdit(GetControl('BUDGET')) ;
  if (BUDGET <> Nil) then
  begin
    BUDGET.OnElipsisClick := EdtCODEBUDOnElipsisClick ;
  end ;

  TFVierge(Ecran).OnKeyDown := GRIDFormKeyDown;

  BTN := TToolbarButton97(GetControl('BAPPLIQUER')) ;
  if (BTN <> Nil) then BTN.OnClick := bChoixBudgetOnclick ;

  BTN := TToolbarButton97(GetControl('BVALIDER')) ;
  if (BTN <> Nil) then BTN.OnClick := bValiderOnclick ;

  BTN := TToolbarButton97(GetControl('BENATTENTE')) ;
  if (BTN <> Nil) then BTN.OnClick := bMettreEnAttenteOnclick ;

  BTN := TToolbarButton97(GetControl('BREFUSER')) ;
  if (BTN <> Nil) then BTN.OnClick := bRefuserOnclick ;

  BTN := TToolbarButton97(GetControl('BAUTORISER')) ;
  if (BTN <> Nil) then BTN.OnClick := bAccepterOnclick ;

  BTN := TToolbarButton97(GetControl('BDEFAIRE')) ;
  if (BTN <> Nil) then BTN.OnClick := bDefaireOnclick ;

 BTN := TToolbarButton97(GetControl('BINFOBUDGET')) ;
 if (BTN <> Nil) then
 begin
   BTN.OnClick := bInfoBudgetOnclick ;
 end ;

 BTN := TToolbarButton97(GetControl('BINFOLIGNE')) ;
 if (BTN <> Nil) then
 begin
   BTN.Enabled := False ;
   BTN.OnClick := bInfoLigneOnclick ;
 end ;

 BTN := TToolbarButton97(GetControl('BINFOENGAGEMENT')) ;
 if (BTN <> Nil) then
 begin
   BTN.Enabled := False ;
   BTN.OnClick := bInfoEngagementOnclick ;
 end ;

 BTN := tToolbarButton97(GetCOntrol('BAFFINEBUDGET')) ;
 if (BTN <> Nil) then BTN.OnClick :=  BAffinageOnClick ;

  GRID := THGrid(GetControl('GRID')) ;
  if GRID <> Nil then
  BEGIN
     GRID.OnClick       := GRIDClick ;
     GRID.OnDblClick    := GRIDDblClick ;
     GRID.PostDrawCell  := GRIDPostDrawCell ;
     GRID.GetCellCanvas := GRIDGetCellCanvas ;
     GRID.OnMouseUp     := GRIDMouseUp ;
     GRID.OnMousedown   := GRIDMouseDown ;
     GRID.OnMouseMove   := GRIDMouseMove ;
     GRID.OnDragOver    := GRIDDragOver ;
     GRID.OnEndDrag     := GRIDEndDrag ;
  end ;
 SetControlEnabled('GRID', False);
 SetControlVisible('GRID', False);
 SetControlText('BUDGET', '') ;
 MajHelp();
 InitialiserTitre();
 SetControlVisible('LABEL_AFFINAGE', (RecupWhereTOBAffinagePourBudgetEtLigne() <> ''));
end ;
{
*******************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.InitialiserTitre();
var StTitre : String ;
BEGIN
 StTitre := TraduireMemoire('Acceptation des engagements') ;
 Ecran.Caption := StTitre ;
 UpdateCaption(Ecran) ;
END ;
{
*******************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.MajHelp();
BEGIN
   TFVierge(Ecran).HelpContext := 28444330 ;
 end ;{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.OnClose ;
var 
    F: TfVierge ;
begin
  onCloseFg := True ;
  F := TfVierge(Ecran) ;
  F.ModalResult := mrNone ;
  Inherited ;
  if Not bUpdateFg then
  begin
    if Not AbandonnerAccepationEngagement() then Exit ;
  end ;

  if (iBudgetList <> Nil) then FreeAndNil(iBudgetList) ;
  if (TOBAFFINAGE <> Nil) then FreeAndNil(TOBAFFINAGE);

  RegSaveToolbarPos(Ecran,'CPVALISESSIONPOIN') ; //* Sauvegarde position fenêtre
  F.ModalResult := mrOk ;
end ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.OnLoad ;
begin
  Inherited ;
 onLoadFg       := True ;
 onCloseFg      := False ;

 if Not bValiderFg
  and Not bAffinageFg
  and Not IsAffichageMultiBudgetAcceptation() then
                    SetControlText('BUDGET', RechercherPremierBudget()) ;

 if (GetControlText('BUDGET') <> '')
   or bValiderFg then bChoixBudgetOnclick(Nil) ;

 onLoadFg := False ;
 bValiderFg :=  False ;
end ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GRID_OnPopUp(sender: TObject);
Var T    : TMenuItem ;
    Pop  : TPopupMenu;
    ligne : Integer ;
    niveau : String ;
    BTN : TToolbarButton97 ;
BEGIN
 Pop:=TPopupMenu(Sender);
 PurgePopup(Pop);

 ligne   := GRID.Row ;
 niveau  :=  GRID.Cells[COL_NIVEAU, ligne] ;


 BTN := TToolbarButton97(GetControl('BINFOBUDGET')) ;
 if (BTN <> Nil) then BTN. Enabled := (niveau = 'B') ;

 BTN := TToolbarButton97(GetControl('BINFOLIGNE')) ;
 if (BTN <> Nil) then BTN. Enabled := (niveau = 'L') ;

 BTN := TToolbarButton97(GetControl('BINFOENGAGEMENT')) ;
 if (BTN <> Nil) then BTN. Enabled := (niveau = 'E') ;

 T := TMenuItem.Create(Pop);
 T.Name    := 'GRID_AccepterOnClick';
 T.Caption := TraduireMemoire('Accepter <Ctrl + A>') ;
 T.OnClick := GRID_AccepterOnClick ;
 V_PGI.GraphList.GetBitmap(46, T.Bitmap) ;
 T.Enabled := (niveau = 'E') ;
 Pop.Items.Add(T);

 T := TMenuItem.Create(Pop);
 T.Name    := 'GRID_RefuserOnClick';
 T.Caption := TraduireMemoire('Refuser <Ctrl + R>') ;
 T.OnClick := GRID_RefuserOnClick ;
 V_PGI.GraphList.GetBitmap(45, T.Bitmap) ;
 T.Enabled := (niveau = 'E')  ;
 Pop.Items.Add(T);

 T := TMenuItem.Create(Pop);
 T.Name    := 'GRID_MettreEnAttenteOnClick';
 T.Caption := TraduireMemoire('Mettre en attente <Ctrl + U>') ;
 V_PGI.GraphList.GetBitmap(44, T.Bitmap) ;
 T.OnClick := GRID_MettreEnAttenteOnClick ;
 T.Enabled := (niveau = 'E')  ;
 Pop.Items.Add(T);

 T := TMenuItem.Create(Pop);
 T.Name    := 'GRID_RevenirEtatInitialOnClick';
 T.Caption := TraduireMemoire('Revenir à l''état initial <Ctrl + Z>') ;
 T.OnClick := GRID_RevenirEtatInitialOnClick ;
 V_PGI.GraphList.GetBitmap(79, T.Bitmap) ;
 T.Enabled := (niveau = 'E')  ;
 Pop.Items.Add(T);

 T := TMenuItem.Create(Pop);
 T.Name    := 'LIGNE_Separator';
 T.Caption := '-' ;
 T.OnClick := Nil ;
 T.Enabled := True ;
 Pop.Items.Add(T);
 T := TMenuItem.Create(Pop);
 T.Name    := 'bInfoBudgetOnclick';
 T.Caption := TraduireMemoire('Voir les propriétés du budget') ;
 T.OnClick := bInfoBudgetOnclick ;
 V_PGI.GraphList.GetBitmap(57, T.Bitmap) ;
 T.Enabled := (niveau = 'B') ;
 Pop.Items.Add(T);

 T := TMenuItem.Create(Pop);
 T.Name    := 'bInfoLigneOnclick';
 T.Caption := TraduireMemoire('Voir les propriétés de la ligne budgétaire') ;
 T.OnClick := bInfoLigneOnclick ;
 V_PGI.GraphList.GetBitmap(76, T.Bitmap) ;
 T.Enabled := (niveau = 'L') ;
 Pop.Items.Add(T);

 T := TMenuItem.Create(Pop);
 T.Name    := 'bInfoEngagementOnclick';
 T.Caption := TraduireMemoire('Voir les propriétés de l''engagement') ;
 T.OnClick := bInfoEngagementOnclick ;
 V_PGI.GraphList.GetBitmap(14, T.Bitmap) ;
 T.Enabled := (niveau = 'E') ;
 Pop.Items.Add(T);

END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GRID_AccepterOnClick(sender: TObject);
BEGIN
  ChangeValueSelection(GRID.Selection,  'VAL')
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GRID_RefuserOnClick(sender: TObject);
BEGIN
  ChangeValueSelection(GRID.Selection,  'REF')
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GRID_MettreEnAttenteOnClick(sender: TObject);
BEGIN
  ChangeValueSelection(GRID.Selection,  'ATT')
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GRID_RevenirEtatInitialOnClick(sender: TObject);
BEGIN
  ChangeValueSelection(GRID.Selection,  '')
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.bInfoBudgetOnclick(Sender: TObject) ;
BEGIN
 GRIDDblClick(Nil) ;
END ;
procedure TOF_CPVALISESSIONPOIN.bInfoEngagementOnclick(sender: TObject);
BEGIN
 GRIDDblClick(Nil) ;
END ;
procedure TOF_CPVALISESSIONPOIN.bInfoLigneOnclick(Sender: TObject) ;
BEGIN
 GRIDDblClick(Nil) ;
END ;
{
********************************************************************************
}
Function TOF_CPVALISESSIONPOIN.MonterData (theBudgetCd : String) : Boolean ;
var
   bLigneFg,bEngagementFg, bRealiseFg : Boolean ;
   bEnfObligFg, bIsPresImmoAffectFg : Boolean ;
   msgError : String ;
   xx_whereBudget : String ;
   wIbudget : TIBudget ;
BEGIN
 result := False ;
 bLigneFg            := True ;
 bEngagementFg       := True ;
 bRealiseFg          := False ;
 bEnfObligFg         := True ;
 bIsPresImmoAffectFg := False ;

 xx_whereBudget := '' ;
 if (theBudgetCd <> '') then
 begin
   if Not IsPresenceBudget(theBudgetCd) then
   begin
     PGIError(RechercherMessage(101, [theBudgetCd]) , Ecran.Caption);
     Exit ;
   end  ;
   xx_whereBudget := ' IBU_CODEBUD = "' + theBudgetCd + '" ' ;
 end ;

 if (iBudgetList <> Nil) then  FreeAndNil(iBudgetList) ;
 iBudgetList := TIMapList.Create();
 RechercherArboPourTousLesBudgets(xx_whereBudget, '', '', '', bLigneFg, bEngagementFg, bRealiseFg, False,
                                  iBudgetList, TOBAFFINAGE,  bEnfObligFg, bIsPresImmoAffectFg, Ecran)  ;

 if (iBudgetList.GetCount() = 0) then
 begin
  if (theBudgetCd <> '') then
  begin
    wIbudget := TIBudget.Create();
    wIbudget.LoadDataFromDB(GetControlText('BUDGET'));
    msgError := RechercherMessage(530, [wIbudget.GetIBU_CODEBUD(), wIbudget.GetIBU_LIBELLE()]) ;
    FreeAndNil(wIbudget) ;
  end
  else msgError := RechercherMessage(508, ['']) ;
  PGIError(msgError, Ecran.Caption);
  Exit ;
 end ;
 Result := True ;
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.bChoixBudgetOnclick(Sender : TObject) ;
var budget : String ;
   wiBudget : TIBudget ;
   HMT : THSystemMenu ;
BEGIN
 wiBudget := Nil ;
 budget := GetControlText('BUDGET') ;

 //* En cas de changement du budget dans la gestion courante
 //* On va rechercher les données.

 if Not onLoadFg  and (budget <> svChoix)  then
 begin
   if Not GererEtatModifies() then
   begin
     budget := svChoix ;
     SetControlText('BUDGET', budget) ;
     Exit ;
   end ;
 end ;

 If  Not MonterData(budget) then
 begin
   SetControlEnabled('GRID', False);
   SetControlVisible('GRID', False);
   Exit ;
 end ;

 SetControlEnabled('GRID', True);
 SetControlVisible('GRID', True);

 //* Controle de l'existence du budget

 if (budget <> '') then
 begin
   wiBudget := TIBudget(iBudgetList.GetValue(budget)) ;
   if (wiBudget = Nil) then
   begin
     if bAffinageFg then
     begin
       budget := '' ;
       SetControlText('BUDGET', budget) ;
       bAffinageFg := False ;
     end
     else
     begin
       PGIError(RechercherMessage(504, [budget]), ecran.Caption);
       SetFocusControl('BUDGET');
       SetControlVisible('GRID', False);
       svChoix := GetControlText('BUDGET') ;
       Exit ;
     end ;
   end ;
 end ;

 if (wiBudget = Nil) then FormaterGRIDAllBudget()
                     else FormaterGRID(wiBudget) ;

 svChoix := GetControlText('BUDGET') ;
 HMT:=THSystemMenu(GetControl('HMTRAD')) ;
 If (HMT<>NIL) Then HMT.ResizeGridColumns(GRID);
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.bValiderOnclick(Sender : TObject) ;
var F : TFVierge ;
BEGIN
 F := TfVierge(Ecran) ;
 F.ModalResult := mrNone ;

 if Not ControlerAvantValidation()  then Exit ;

 if Not bForceUpdateFg then
   if  Not ConfirmerAccepationEngagement() then Exit ;

 if (Transactions(GoAcceptation,0) <> oeOK) Then
 begin
    PGIError(TraduireMemoire('Erreur pendant le traitement.'), '') ;
    exit ;
 end
 else
 begin
   PGIInfo(TraduireMemoire('Traitement terminé'), TraduireMemoire('Acceptation')) ;
 end ;
 bModifFg := False ;
 bForceUpdateFg := False ;
 OnLoad() ;
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GoAcceptation() ;
var indexB, indexL, indexE : Integer ;
    iBudget  : TIBudget ;
    iLigne   : TIBudgetLigne ;
    iEng     : TIBudgetEng ;
    wMapList : TIMapList ;
    wEngList : TIMapList ;
    bFlagLigneFg  : Boolean ;
    bFlagBudgetFg : Boolean ;
BEGIN
 //* Forcage du marquage sur tous les budgets qui ont une demande d'acceptation

  for indexB := 0 to iBudgetList.GetCount() - 1 do
  begin
    iBudget  := TIBudget(iBudgetList.Detail(indexB)) ;
    wMapList := iBudget.GetIBudgetLigneList() ;
    bFlagBudgetFg  := False ;
    if (wMapList <> Nil) and (wMapList.GetCount() <> 0) then
    begin
      for indexL := 0 to wMapList.GetCount() -1 do
      begin
        iLigne := TIBudgetLigne(wMapList.Detail(indexL)) ;
        bFlagLigneFg  := False ;
        wEngList := iLigne.GetENGAGEMENTList() ;
        if (wEngList <> Nil) and (wEngList.GetCount() <> 0) then
        begin
          for indexE := 0 to wEngList.GetCount() -1 do
          begin
            iEng := TIBudgetEng(wEngList.Detail(indexE)) ;
            if (iEng <> Nil) and iEng.GetDEMACCEPTFg() then
            begin
              iEng.SetMarquageFg(True) ;
              bFlagLigneFg := True ;
              bFlagBudgetFg := True ;
            end ;
          end ;
        end ;
        iLigne.SetMarquageFg(bFlagLigneFg) ;
      end ;
    end ;
    iBudget.SetMarquageFg(bFlagBudgetFg) ;
  end ;

 //* Mise à jour de la base

 bValiderFg := AccepterEngagementListByBudgetList(iBudgetList) ;
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.bMettreEnAttenteOnclick(Sender : TObject) ;
BEGIN
  ChangeValueSelection(GRID.Selection, 'ATT')
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.bRefuserOnclick(Sender : TObject) ;
BEGIN
  ChangeValueSelection(GRID.Selection,  'REF')
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.bAccepterOnclick(Sender : TObject) ;
BEGIN
  ChangeValueSelection(GRID.Selection,  'VAL')
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.bDefaireOnclick(Sender : TObject) ;
BEGIN
  ChangeValueSelection(GRID.Selection,  '')
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.ChargeLesBitmaps();
begin
  ChargeImageList() ;
end;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.FormaterColonnes() ;
BEGIN
 GRID.DefaultRowHeight := 19 ;
 GRID.ColCount         := 8 ;

 ClearTHGRID(GRID, GRID.ColCount , [0,COL_NIVEAU, COL_BUDGET, COL_LIGNE, COL_ENGAGEMENT,
                                    COL_INTITULE, COL_MONTANT, COL_ETAT]) ;

 GRID.Titres.Clear ;
 GRID.Cells[0,0] := '' ;
 GRID.Cells[COL_LIGNE,0]    := TraduireMemoire('Ligne')  ;
 GRID.Cells[COL_INTITULE,0] := TraduireMemoire('Intitulé')  ;
 GRID.Cells[COL_MONTANT,0]  := TraduireMemoire('Montant')  ;
 GRID.Cells[COL_ETAT, 0]    := TraduireMemoire('Etat') ;

 GRID.ColWidths[0]              := -1 ;
 GRID.ColWidths[COL_NIVEAU]     := -1 ;
 GRID.ColWidths[COL_BUDGET]     := -1 ;
 GRID.ColWidths[COL_LIGNE]      := -1 ;
 GRID.ColWidths[COL_ENGAGEMENT] := -1 ;
 GRID.ColWidths[COL_INTITULE]   := 400  ;
 GRID.ColWidths[COL_MONTANT]    := 110  ;
 GRID.ColWidths[COL_ETAT]       := 90 ;

 GRID.ColAligns[0]              := taCenter ;
 GRID.ColAligns[COL_NIVEAU]     := taCenter ;
 GRID.ColAligns[COL_BUDGET]     := taCenter ;
 GRID.ColAligns[COL_LIGNE]      := taCenter  ;
 GRID.ColAligns[COL_ENGAGEMENT] := taCenter  ;
 GRID.ColAligns[COL_INTITULE]   := taLeftJustify ;
 GRID.ColAligns[COL_MONTANT]    := taRightJustify ;
 GRID.ColAligns[COL_ETAT]       := taCenter ;

 GRID.ColEditables[0]              := False ;
 GRID.ColEditables[COL_NIVEAU]     := False ;
 GRID.ColEditables[COL_BUDGET]     := False ;
 GRID.ColEditables[COL_LIGNE]      := False  ;
 GRID.ColEditables[COL_ENGAGEMENT] := False  ;
 GRID.ColEditables[COL_INTITULE]   := False ;
 GRID.ColEditables[COL_MONTANT]    := False ;
 GRID.ColEditables[COL_ETAT]       := False ;

 GRID.ColTypes[COL_MONTANT]  := 'R' ;

END ;
{
********************************************************************************
* CHARGEMENT DU TABLEAU TOUS LES BUDGETS / TOUTES LES LIGNES
*
}
procedure TOF_CPVALISESSIONPOIN.FormaterGRIDAllBudget() ;
var
    iLigneBudgetList : TIMapList ;
    iEngagementList : TIMapList ;
    wIb  : TIBudget ;
    wIbl : TIBudgetLigne ;
    wIBe : TIBudgetEng ;
    indexb, index, ligne, indexE : Integer ;
    IBL_CODEBUD, IBL_LIGNE, IBL_LIBELLE : String ;
    IBU_CODEBUD, IBU_LIBELLE : String ;
    IBE_CODEBUD, IBE_LIGNE, IBE_ENGAGEMENT, IBE_LIBELLE : String ;
    IBE_ETAT, ACCEPTATION : String ;
    IBE_MONTANT : Double ;

BEGIN
 GRID := THGrid(GetControl('GRID')) ;
 if (GRID = Nil) then Exit ;

 FormaterColonnes() ;

 if (iBudgetList = Nil) then Exit ;

 if (iBudgetList.GetCount() <> 0) then
 begin
   GRID.RowCount := 1 ;

   ligne := 0 ;

   GRID.RowCount := GRID.RowCount + iBudgetList.GetCount() ;

   for indexb := 0 to iBudgetList.GetCount() -1 do
   begin

     wIb := TIBudget(iBudgetList.Detail(indexb)) ;
     if (wIb <> Nil) then
     begin

       IBU_CODEBUD := wIb.GetIBU_CODEBUD() ;
       IBU_LIBELLE := wIb.GetIBU_LIBELLE() ;

       Inc(ligne)  ;

       GRID.Cells[COL_NIVEAU, ligne]     := 'B' ;
       GRID.Cells[COL_BUDGET, ligne]     := IBU_CODEBUD ;
       GRID.Cells[COL_LIGNE, ligne]      := '' ;
       GRID.Cells[COL_ENGAGEMENT, ligne] := '' ;
       GRID.Cells[COL_INTITULE, ligne]   :=  '   ' + IBU_CODEBUD + ' - '+ IBU_LIBELLE ;
       GRID.Cells[COL_ETAT, ligne]       := '' ;

       //****** LIGNES BUDGETAIRES

       iLigneBudgetList := wIb.GetIBudgetLigneList() ;
       if (iLigneBudgetList.GetCount <> 0) then
       begin
         GRID.RowCount := GRID.RowCount + iLigneBudgetList.GetCount() ;
         for index := 0 to iLigneBudgetList.GetCount -1 do
         begin
           wIbl := TIBudgetLigne(iLigneBudgetList.Detail(index)) ;
           if (wIbl <> Nil) then
           begin
             IBL_CODEBUD := wIbl.GetIBL_CODEBUD() ;
             IBL_LIGNE   := wIbl.GetIBL_CODELIGBUD() ;
             IBL_LIBELLE := wIbl.GetIBL_LIBELLE() ;

             Inc(ligne)  ;
             GRID.Cells[COL_NIVEAU, ligne] := 'L' ;
             GRID.Cells[COL_BUDGET, ligne] := IBL_CODEBUD ;
             GRID.Cells[COL_LIGNE, ligne] := IBL_LIGNE ;
             GRID.Cells[COL_ENGAGEMENT, ligne] := '' ;
             GRID.Cells[COL_INTITULE, ligne] := '      ' + IBL_LIGNE + ' - '+ IBL_LIBELLE ;
             GRID.Cells[COL_ETAT, ligne] := '' ;

            //****** ENGAGEMENTS

             iEngagementList := wIbl.GetENGAGEMENTList() ;
             if (iEngagementList.GetCount <> 0) then
             begin
               GRID.RowCount := GRID.RowCount + iEngagementList.GetCount() ;
               for indexE := 0 to iEngagementList.GetCount -1 do
               begin
                 wIBe := TIBudgetEng(iEngagementList.Detail(indexE)) ;
                 if (wIBe <> Nil) then
                 begin
                   IBE_CODEBUD    := wIBe.GetIBE_CODEBUD() ;
                   IBE_LIGNE      := wIBe.GetIBE_CODELIGBUD() ;
                   IBE_ENGAGEMENT := wIBe.GetIBE_CODEENGBUD() ;
                   IBE_LIBELLE    := wIBe.GetIBE_LIBELLE() ;
                   IBE_ETAT       := wIBe.GetIBe_ETAT() ;
                   IBE_MONTANT    := wIbe.GetIBE_MONTANT() ;
                   ACCEPTATION    := wIbe.GetEtatDemCd() ;

                   Inc(ligne)  ;

                   GRID.Cells[COL_NIVEAU, ligne]   := 'E' ;
                   GRID.Cells[COL_BUDGET, ligne]   := IBL_CODEBUD ;
                   GRID.Cells[COL_LIGNE, ligne]    := IBL_LIGNE ;
                   GRID.Cells[COL_ENGAGEMENT, ligne]    := IBE_ENGAGEMENT ;
                   GRID.Cells[COL_INTITULE, ligne] := '          ' + IBE_ENGAGEMENT + ' - '+ IBE_LIBELLE ;
                   GRID.Cells[COL_MONTANT, ligne]  := FloatToStrF(IBE_MONTANT, ffNumber, 20, NbDecV) ;

                   if (ACCEPTATION = '') then
                   begin
                     if (IBE_ETAT = 'VAL')      then GRID.Cells[COL_ETAT, ligne] := IcoAccepter
                     else if (IBE_ETAT = 'REF') then GRID.Cells[COL_ETAT, ligne] := IcoRefuser
                                                else GRID.Cells[COL_ETAT, ligne] := IcoEnAttente ;
                   end
                   else
                   begin
                     if (ACCEPTATION = 'VAL')      then GRID.Cells[COL_ETAT, ligne] := IcoAccepterUpdate
                     else if (ACCEPTATION = 'REF') then GRID.Cells[COL_ETAT, ligne] := IcoRefuserUpdate
                                                   else GRID.Cells[COL_ETAT, ligne] := IcoEnAttenteUpdate ;
                   end ;
                 end ;
               end ;
             end ;

           end ;
         end ;
       end ;
     end ;
   end ;
   SetControlEnabled('GRID', True);
   SetControlVisible('GRID', True);
 end
 else
 begin
   SetControlEnabled('GRID', False);
   SetControlVisible('GRID', False);
 end ;
 if (GRID.RowCount <= 1) then GRID.RowCount := 2 ;
 GRID.FixedRows := 1  ;
END ;
{
********************************************************************************
* CHARGEMENT DU TABLEAU TOUTES LES LIGNES POUR UN BUDGET SELECTIONNE
*
}
procedure TOF_CPVALISESSIONPOIN.FormaterGRID(theIBudget : TIBudget) ;
var
    iLigneBudgetList : TIMapList ;
    iEngagementList  : TIMapList ;
    wIbl             : TIBudgetLigne ;
    wIBe             : TIBudgetEng ;
    index, indexE, ligne : Integer ;
    IBL_CODEBUD, IBL_LIGNE, IBL_LIBELLE, IBU_CODEBUD, IBU_LIBELLE : String ;
    IBE_CODEBUD, IBE_LIGNE, IBE_ENGAGEMENT, IBE_LIBELLE : String ;
    IBE_ETAT, ACCEPTATION : String ;
    IBE_MONTANT : Double ;
BEGIN
 GRID := THGrid(GetControl('GRID')) ;
 if (GRID = Nil) then Exit ;

 FormaterColonnes() ;

 iLigneBudgetList := theIBudget.GetIBudgetLigneList() ;
 if (iLigneBudgetList.GetCount <> 0) then
 begin
   GRID.RowCount := iLigneBudgetList.GetCount() + 2 ;
   if (GRID.RowCount <= 1) then GRID.RowCount := 2 ;
   GRID.FixedRows := 1  ;

   ligne := 0 ;

   IBU_CODEBUD := theIBudget.GetIBU_CODEBUD() ;
   IBU_LIBELLE := theIBudget.GetIBU_LIBELLE() ;

   Inc(ligne)  ;

   GRID.Cells[COL_NIVEAU, ligne] := 'B' ;
   GRID.Cells[COL_BUDGET, ligne] := IBU_CODEBUD ;
   GRID.Cells[COL_LIGNE, ligne] := '' ;
   GRID.Cells[COL_ENGAGEMENT, ligne] := '' ;
   GRID.Cells[COL_INTITULE, ligne] :=  '    ' + IBU_CODEBUD + ' - '+ IBU_LIBELLE ;
   GRID.Cells[COL_ETAT, ligne] := '' ;

   for index := 0 to iLigneBudgetList.GetCount -1 do
   begin
     wIbl := TIBudgetLigne(iLigneBudgetList.Detail(index)) ;
     if (wIbl <> Nil) then
     begin
       IBL_CODEBUD := wIbl.GetIBL_CODEBUD() ;
       IBL_LIGNE   := wIbl.GetIBL_CODELIGBUD() ;
       IBL_LIBELLE := wIbl.GetIBL_LIBELLE() ;

       Inc(ligne)  ;

       GRID.Cells[COL_NIVEAU, ligne]   := 'L' ;
       GRID.Cells[COL_BUDGET, ligne]   := IBL_CODEBUD ;
       GRID.Cells[COL_LIGNE, ligne]    := IBL_LIGNE ;
       GRID.Cells[COL_ENGAGEMENT, ligne]    :=  '' ;
       GRID.Cells[COL_INTITULE, ligne] := '             ' + IBL_LIGNE + ' - '+ IBL_LIBELLE ;
       GRID.Cells[COL_ETAT, ligne] := '' ;

       //****** ENGAGEMENTS

       iEngagementList := wIbl.GetENGAGEMENTList() ;
       if (iEngagementList.GetCount <> 0) then
       begin
         GRID.RowCount := GRID.RowCount + iEngagementList.GetCount() ;
         for indexE := 0 to iEngagementList.GetCount -1 do
          begin
            wIBe := TIBudgetEng(iEngagementList.Detail(indexE)) ;
            if (wIBe <> Nil) then
            begin
              IBE_CODEBUD    := wIBe.GetIBE_CODEBUD() ;
              IBE_LIGNE      := wIBe.GetIBE_CODELIGBUD() ;
              IBE_ENGAGEMENT := wIBe.GetIBE_CODEENGBUD() ;
              IBE_LIBELLE    := wIBe.GetIBE_LIBELLE() ;
              IBE_ETAT       := wIBe.GetIBe_ETAT() ;
              IBE_MONTANT    := wIbe.GetIBE_MONTANT() ;
              ACCEPTATION    := wIbe.GetEtatDemCd() ;

              Inc(ligne)  ;

              GRID.Cells[COL_NIVEAU, ligne]   := 'E' ;
              GRID.Cells[COL_BUDGET, ligne]   := IBL_CODEBUD ;
              GRID.Cells[COL_LIGNE, ligne]    := IBL_LIGNE ;
              GRID.Cells[COL_ENGAGEMENT, ligne]    := IBE_ENGAGEMENT ;
              GRID.Cells[COL_INTITULE, ligne] := '                   ' + IBE_ENGAGEMENT + ' - '+ IBE_LIBELLE ;
              GRID.Cells[COL_MONTANT, ligne]  := FloatToStrF(IBE_MONTANT, ffNumber, 20, NbDecV) ;

              if (ACCEPTATION = '') then
              begin
                if (IBE_ETAT = 'VAL')      then GRID.Cells[COL_ETAT, ligne] := IcoAccepter
                else if (IBE_ETAT = 'REF') then GRID.Cells[COL_ETAT, ligne] := IcoRefuser
                                           else GRID.Cells[COL_ETAT, ligne] := IcoEnAttente ;
              end
              else
              begin
                if (ACCEPTATION = 'VAL')      then GRID.Cells[COL_ETAT, ligne] := IcoAccepterUpdate
                else if (ACCEPTATION = 'REF') then GRID.Cells[COL_ETAT, ligne] := IcoRefuserUpdate
                                              else GRID.Cells[COL_ETAT, ligne] := IcoEnAttenteUpdate ;
              end ;
            end ;
          end ;
        end ;
     end ;
   end ;
   SetControlEnabled('GRID', True);
   SetControlVisible('GRID', True);
 end
 else
 begin
   SetControlEnabled('GRID', False);
   SetControlVisible('GRID', False);
 end ;
END ;


{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GRIDPostDrawCell (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
VAR
  niveau : String ;
  Arect  : Trect ;
BEGIN
  if Arow < GRID.Fixedrows then exit;

  niveau := GRID.cells[COL_NIVEAU, ARow] ;
  if (niveau = 'E') then Exit ;

  Arect := GRID.CellRect(Acol,Arow) ;

  if (niveau = 'B') then Canvas.Brush.Color := RowColor1 ;
  if (niveau = 'L') then Canvas.Brush.Color := RowColor2 ;

  Canvas.FillRect(ARect);
  Canvas.TextRect(ARect, ARect.Left + 2, ARect.Top + 2, GRID.Cells[ACol,ARow]);
END ;
{
********************************************************************************
}
Procedure TOF_CPVALISESSIONPOIN.GRIDGetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
VAR
  niveau : String ;
BEGIN

  niveau := GRID.cells[COL_NIVEAU, ARow] ;
  if (niveau = 'B') then
  begin
    Canvas.Font.Style := [fsBold];
    Canvas.Font.Color := clWhite;
  end
  else if (niveau = 'L') then
  begin
    Canvas.Font.Style := [];
    Canvas.Font.Color := clBlack;
  end ;
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GRIDMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
VAR
    ARect : TGridRect;
    ACol,ARow,iInd : integer;
    okStop : boolean;
    wBudget,  niveau, savNiveau  : String ;
BEGIN
 GRID.MouseToCell(X,Y,ACol,ARow);

 case Button of
   mbLeft :
     begin
       if (ACol = 0) and (ARow = 0) then exit;
       if (ACol < 0) or (ARow < 0) then exit;

       if (ARow = 0) then
       begin SelectLaCol(ACol) ; exit ; end ;

       if (ACol = COL_INTITULE)
        or (ACol = COL_MONTANT)  then
       begin SelectLaRow(ARow); exit; end;

       okStop := False;

       savNiveau  := GRID.Cells[COL_NIVEAU, ARow] ;
       if (savNiveau = 'E') then Exit ;

       for iInd := ARow + 1 to GRID.RowCount do
       begin
         niveau  := GRID.Cells[COL_NIVEAU, iInd] ;
         wBudget := GRID.Cells[COL_BUDGET, iInd] ;

         if (niveau = savNiveau) then  okStop := True;
         if (iInd = GRID.RowCount-1) then okStop := True;

         if okStop then
         begin
           if ACol = 0 then
           begin
             ARect.Left := 1;
             ARect.Right := GRID.ColCount;
             ARect.Top := ARow + 1;
           end
           else
           begin
             ARect.Left := ACol;
             ARect.Right := ACol;
             ARect.Top := ARow +1;
           end;

           if (iInd = GRID.RowCount-1) then ARect.Bottom := iInd
                                  else ARect.Bottom := iInd - 1;
           GRID.Selection := ARect;
           exit;
        end;
      end;
    end;
  mbRight :
    begin
      if (ACol = 0) and (ARow = 0) then exit;
      if (ACol < 0) or (ARow < 0) then exit;
      IF (ACol = COL_ETAT) then Exit ;
      SelectLaRow(ARow) ;
    end ;
 end ;
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.SelectLaCol(ACol : integer);
var ARect : TGridRect;
begin
  ARect.Left     := ACol;
  ARect.Top      := 1;
  ARect.Right    := ACol;
  ARect.Bottom   := GRID.RowCount;
  GRID.Selection := ARect;
end;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.SelectLaRow(ARow : integer);
var ARect : TGridRect;
begin
  ARect.Left := 1;
  ARect.Top := ARow;
  ARect.Right := GRID.ColCount;
  ARect.Bottom := ARow;
  GRID.Selection := ARect;
end;
{
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GRIDMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var MyRect : TGridRect;
    ACol, Arow : integer;
BEGIN
 //pour éviter la sélection d'un groupe (visuel)
 if Button <> mbLeft then exit;
 GRID.MouseToCell(X,Y,ACol,ARow);

 if (ACol <= 0) or (ARow <= 0) then exit;

 myRect := GRID.Selection;
 if (ACol = COL_MONTANT)
   or (ACol = COL_INTITULE) then  MyRect.Left := 1 ;
 GRID.Selection := myRect;
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GRIDMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
var ACol, ARow : integer;
BEGIN
  GRID.MouseToCell(X,Y,ACol,ARow);
  if ((ARow = 0) and (ACol > 0)) or
     ((ACol = 0) and (ARow > 0)) then
  begin
    GRID.Hint := Trim(GRID.Cells[ACol,ARow]);
    GRID.ShowHint := True;
  end
  else GRID.ShowHint := False;
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GRIDDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
BEGIN
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GRIDEndDrag(Sender, Target: TObject; X, Y: Integer);
BEGIN
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.ChangeCellValue(ACol, ARow : integer ; Value : string);
var
    wBudget, wLigne,wEng : String ;
    wBudgetLigneListe : TIMapList ;
    wEngagementListe : TIMapList ;
    wB  : TIBudget ;
    wBL : TIBudgetLigne ;
    wBe : TIBudgetEng ;
    wVal : String ;
    IBE_ETAT : String ;
    niveau : String ;
begin
  wVal := Value ;
  if (ACol = 0) or (ARow = 0) then exit;

  niveau := GRID.cells[COL_NIVEAU, ARow] ;
  if (niveau <> 'E') then Exit ;

  wBudget :=  GRID.cells[COL_BUDGET, ARow] ;
  if (wBudget = '') then Exit ;

  wLigne := GRID.cells[COL_LIGNE, ARow] ;
  if (wLigne = '') then Exit ;

  wEng := GRID.cells[COL_ENGAGEMENT, ARow] ;
  if (wEng = '') then Exit ;


  wB := TIBudget(iBudgetList.GetValue(wBudget)) ;
  if (wB = Nil) then Exit ;

  wBudgetLigneListe := wB.GetIBudgetLigneList() ;
  if (wBudgetLigneListe =  Nil) then Exit ;

  wBL := TIBudgetLigne(wBudgetLigneListe.GetValue(wLigne)) ;
  if (wBL =  Nil) then Exit ;


  wEngagementListe := wBL.GetENGAGEMENTList() ;
  if (wEngagementListe =  Nil) then Exit ;


  wBe := TIBudgetEng(wEngagementListe.GetValue(wEng)) ;
  if (wBe =  Nil) then Exit ;

  IBE_ETAT := wBe.GetIBE_ETAT() ;

  if (pos('#ICO#',GRID.Cells[ACol,ARow]) > 0) then
  begin
    if (wVal <> '') then
    begin
      if (wVal = 'VAL') then GRID.Cells[COL_ETAT, ARow] := IcoAccepterUpdate
      else if (wVal = 'REF') then GRID.Cells[COL_ETAT, ARow] := IcoRefuserUpdate
      else if (wVal = 'ATT') then GRID.Cells[COL_ETAT, ARow] := IcoEnAttenteUpdate ;
      wBe.SetEtatDemCd(wVal);
      wBe.SetDEMACCEPTFg(True);
      bModifFg := True;
    end
    else
    begin
      if (IBE_ETAT = 'VAL') then GRID.Cells[COL_ETAT, ARow] := IcoAccepter
      else if (IBE_ETAT = 'REF') then GRID.Cells[COL_ETAT, ARow] := IcoRefuser
      else if (IBE_ETAT = 'ATT') then GRID.Cells[COL_ETAT, ARow] := IcoEnAttente ;
      wBe.SetEtatDemCd('');
      wBe.SetDEMACCEPTFg(False);
    end ;
  end ;
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.ChangeValueSelection(MyRect : TGridRect ; Value : string);
var iCol,iRow : integer;
BEGIN
//  if (MyRect = Nil) then myRect := GRID.Selection ;
  for iRow := myRect.Top to myRect.Bottom  do
  begin
    for iCol := myRect.Left to myRect.Right do
    begin
      ChangeCellValue(iCol,iRow,Value);
    end;
  end;
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GRIDFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var FocusGrid : Boolean;
begin
 FocusGrid   := (Screen.ActiveControl.Name = 'GRID');
 Case Key of
    VK_F9 : begin   bChoixBudgetOnclick(Nil) end ;
   {Ctrl+A}  65 : if Shift=[ssCtrl] then if FocusGrid then begin Key:=0 ; ChangeValueSelection(GRID.Selection,'VAL'); end ;
   {Ctrl+R}  82 : if Shift=[ssCtrl] then if FocusGrid then begin Key:=0 ; ChangeValueSelection(GRID.Selection,'REF'); end ;
   {Ctrl+U}  85 : if Shift=[ssCtrl] then if FocusGrid then begin Key:=0 ; ChangeValueSelection(GRID.Selection,'ATT'); end ;
   {Ctrl+Z}  90 : if Shift=[ssCtrl] then if FocusGrid then begin Key:=0 ; ChangeValueSelection(GRID.Selection,''); end ;
 end;
end;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GRIDClick(Sender : TObject) ;
var
   ligne : Integer ;
   niveau : String ;
   BTN : TToolbarButton97 ;
BEGIN

 ligne   := GRID.Row ;
 niveau  :=  GRID.Cells[COL_NIVEAU, ligne] ;

 BTN := TToolbarButton97(GetControl('BINFOBUDGET')) ;
 if (BTN <> Nil) then BTN. Enabled := (niveau = 'B') ;

 BTN := TToolbarButton97(GetControl('BINFOLIGNE')) ;
 if (BTN <> Nil) then BTN. Enabled := (niveau = 'L') ;

 BTN := TToolbarButton97(GetControl('BINFOENGAGEMENT')) ;
 if (BTN <> Nil) then BTN. Enabled := (niveau = 'E') ;
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.GRIDDblClick(Sender: TObject);
var
 wBudget, wLigne, wEng, niveau : String ;
 MyRect : TGridRect ;
 ARow : Integer ;
BEGIN
  MyRect := GRID.Selection ;
  ARow := myRect.Top ;

  if (ARow <= 0) then Exit ;

  niveau := GRID.cells[COL_NIVEAU, ARow] ;

  wBudget := GRID.cells[COL_BUDGET, ARow] ;
  wLigne  := GRID.cells[COL_LIGNE, ARow] ;
  wEng    := GRID.cells[COL_ENGAGEMENT, ARow] ;

  if (niveau = 'B') then
  begin
    if (wBudget = '') then Exit ;
    LanceConsultationBudget(wBudget) ;
  end
  else if (niveau = 'L') then
  begin
    if (wBudget = '') or  (wLigne = '') then Exit ;
    LanceSaisieBudgetLignes(consultLigne, 'ACTION=CONSULTATION', wBudget, wLigne) ;
  end
  else if (niveau = 'E') then
  begin
    if (wBudget = '') or  (wLigne = '') or (wEng = '') then Exit ;
    LancerConsultationEngagement(optionBudget, wBudget, wLigne, wEng);
  end ;
END;
{
********************************************************************************
}
Function TOF_CPVALISESSIONPOIN.ControlerAvantValidation() : Boolean ;
var
  index, indexL, indexE : Integer ;
  wB      : TIBudget ;
  wBLList : TIMapList ;
  wBEList : TIMapList ;
  wBl     : TIBudgetLigne ;
  wBE     : TIBudgetEng ;
  ACCEPTATION : String ;
  IBE_ETAT : String  ;
  IBU_CODEBUD, IBU_LIBELLE : String ;
  IBL_CODELIGBUD, IBL_LIBELLE : String ;
  IBE_CODEENGBUD, IBE_LIBELLE : String ;
BEGIN
 result := False ;

 for index := 0 to iBudgetList.GetCount() - 1 do
 begin
    wB := TIBudget(iBudgetList.Detail(index)) ;
    if (wB <> Nil) then
    begin
      wBLList := wB.GetIBudgetLigneList() ;
      if (wBLList <> Nil) and (wBLList.GetCount() <> 0) then
      begin
        for indexL := 0 to wBLList.GetCount() - 1 do
        begin
          wBl := TIBudgetLigne(wBLList.Detail(indexL)) ;
          if (wBl <> Nil) then
          begin
            wBEList := wBl.GetENGAGEMENTList() ;
            if (wBEList <> Nil) and (wBEList.GetCount() <> 0) then
            begin
              for indexE := 0 to wBEList.GetCount() - 1 do
              begin
                wBE :=  TIBudgetEng(wBEList.Detail(indexE)) ;
                if (wBE <> Nil) then
                begin
                  IBE_ETAT    := wBE.GetIBE_ETAT() ;
                  ACCEPTATION := wBE.GetEtatDemCd() ;

                  if (ACCEPTATION <> '') and (ACCEPTATION <> 'VAL') then
                  begin
                    if (IBE_ETAT = 'VAL') then
                    begin
                      IBU_CODEBUD := wBE.GetIBE_CODEBUD() ;
                      IBL_CODELIGBUD := wBE.GetIBE_CODELIGBUD() ;
                      IBE_CODEENGBUD := wBE.GetIBE_CODEENGBUD() ;
                      if IsPresenceBudgetRealise(IBU_CODEBUD, IBL_CODELIGBUD,IBE_CODEENGBUD) then
                      begin
                        IBU_LIBELLE := wB.GetIBU_LIBELLE() ;
                        IBL_LIBELLE := wBl.GetIBL_LIBELLE() ;
                        IBE_LIBELLE := wBE.GetIBE_LIBELLE() ;
                        PGIError(RechercherMessage(8503, [IBL_LIBELLE, IBL_CODELIGBUD,
                                                            IBU_LIBELLE, IBU_CODEBUD,
                                                            IBE_LIBELLE, IBE_CODEENGBUD]), Ecran.caption);
                        Exit ;
                      end ;
                    end ;
                  end ;
                  if (ACCEPTATION = 'VAL') then
                  begin
                     if (wBE.GetIBE_MONTANT() = 0) then
                     begin
                        PGIError(RechercherMessage(8515, [IBL_LIBELLE, IBL_CODELIGBUD,
                                                            IBU_LIBELLE, IBU_CODEBUD,
                                                            IBE_LIBELLE, IBE_CODEENGBUD]), Ecran.caption);
                        Exit ;
                     end ;
                  end ;

                end ;
              end ;
            end ;
          end ;
        end ;
      end ;
    end ;
 end ;
 result := True ;
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.BAffinageOnClick(Sender: TObject) ;
var
 bRetourFg : Boolean ;
BEGIN
 if Not GererEtatModifies() then Exit ;

 bRetourFg := LancerFenetreAffinageCritere(optionBudget, TOBAFFINAGE, True, False, False) ;
 if Not bRetourFg then exit ;

 SetControlVisible('LABEL_AFFINAGE', (RecupWhereTOBAffinagePourBudgetEtLigne() <> ''));
 bAffinageFg := True ;
 bChoixBudgetOnclick(Nil);
END ;
{
********************************************************************************
}
Function TOF_CPVALISESSIONPOIN.GererEtatModifies() : Boolean ;
BEGIN
 result := True ;
 if bModifFg then
 begin
  case PGIAskCancel(RechercherMessage(8501, [''])) of
   mrYes : begin
             bForceUpdateFg := True ;
             bValiderOnclick(Nil);
           end ;
   mrNo : begin  bModifFg := False ; Exit ; end ;
   mrCancel : begin result := False ; Exit ; end ;
  end ;
 end ;
END ;
{
********************************************************************************
}
procedure TOF_CPVALISESSIONPOIN.EdtCODEBUDOnElipsisClick(Sender : TObject) ;
VAR
  titre : String ;
  tableName : String ;
  codeName : String ;
  libelleName : String ;
  where : String ;
  tri : String ;
BEGIN

 titre := TraduireMemoire('Budgets') ;
 tableName := 'IBUDGET'  ;
 codeName := 'IBU_CODEBUD' ;
 libelleName := 'IBU_LIBELLE' ;

 where := RecupWhereTOBAffinagePourBudgetEtLigne() ;

 tri := 'IBU_CODEBUD' ;
 LookupList(THEdit(Sender),titre, tableName, codeName, libelleName, where, tri, True,0);
END ;

Initialization
  registerclasses ( [ TOF_CPVALISESSIONPOIN ] ) ;
end.
