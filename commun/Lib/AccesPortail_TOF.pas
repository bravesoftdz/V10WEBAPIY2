{***********UNITE*************************************************
Auteur  ...... : JS
Créé le ...... : 08/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CONFIDENTIALITE ()
          ...Gestion des droits d'accès
          ...Consultation et modification des accès par groupe d'utilisateurs
Mots clefs ... : TOF;CONF_VUE;CONFIDENTIALITE;GROUPE;UTILISATEUR;MENU
*****************************************************************}
Unit AccesPortail_TOF ;

Interface

Uses StdCtrls,
     Controls,Dialogs,
     Classes,Windows,
{$IFDEF EAGLCLIENT}
     MaineAGL, UtileAgl,
{$ELSE}
     db, EdtREtat,
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
     FE_main,
{$ENDIF}
     forms, Graphics,
     sysutils,
     ComCtrls,Menus,
     HCtrls, Grids, Vierge,
     HEnt1,HSysMenu,HTB97,
     HMsgBox,
     UTOF,UTOB,ED_TOOLS,LicUtil
     {$IFNDEF PGIMAJVER}
     {$IFDEF GCGC}, UtilGC, EntGC {$ENDIF}
     {$ENDIF},
     ListUsersAutorises;
     
function YYLanceFiche_AccesPortail(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_ACCESPORTAIL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    private

      TobVignette,TobUserGrp,TobUser,TobUserName, TobAssoc, TobMyResult: TOB; //js1 160206
      GC,GUser,GRole,GListUsers : THGrid;              //js1 140206
      HMTrad: THSystemMenu;
      Action   : TActionFiche;
      bListUsers,bAcc,bRef,bUser,bCop,bUserGrp,bMnAss : TMenuItem; //js1 130206
      FFirstFind,bModif : boolean;
      FindLigne: TFindDialog;
      Modules,cbDomaine_Value,LesCols : string;
      RowColor1,RowColor2,RowColor3,ColGrp,CopierColler : integer;
      TTag : array of string;
      bForceMaj : boolean;

      //Form
      procedure AffecteLesControls;

      //Grid
      procedure DessineCell (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
      Procedure GCGetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
      procedure GCDblClick(Sender: TObject);
      procedure GCRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
      procedure GCMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
      procedure GCFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
      procedure GCMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
      procedure GCMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
      procedure GCDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
      procedure GCEndDrag(Sender, Target: TObject; X, Y: Integer);
      procedure GUserDblClick(Sender: TObject);
      procedure GUserRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);

      //Splitter
      procedure SplitMoved(Sender: TObject);

      //Actions sur le Grid
      procedure ChangeValueSelection(MyRect : TGridRect ; Value : string);
      procedure ChangeCellValue(ACol,ARow : integer ; Value : string = '');
      procedure FormateLesCols;
      procedure SelectLaCol(ACol : integer);
      procedure SelectLaRow(ARow : integer);
      procedure ResizeLesGrids;
      function RecupWhereDomaine(Module : string ) : string;
      procedure cbMenuChange(Sender: TObject);
      procedure ZoomUtilisateur;


      //Boutons
      procedure bChercherClick(Sender: TObject);
      procedure bUserGrpClick(Sender: TObject);
      procedure FindLigneFind(Sender: TObject);
      procedure BAccClick(Sender: TObject);
      procedure BRefClick(Sender: TObject);
      procedure BUserClick(Sender: TObject);
      procedure BCopClick(Sender: TObject);
      procedure BApercuClick(Sender: TObject);
      procedure BReapplicClick(Sender: TObject);
      //js1 130206
      procedure BListUsersClick(Sender: TObject);

      //Infos complémentaires
      procedure SetInfoCompVisible(Visu : boolean ; TabSheetName : string);

      procedure AfficheGrpAssocies;

      procedure ChargeLesBitmaps;

      //Edition
      procedure ValUserGrpEnter(Sender: TObject);

      //Gestion des Tobs
      procedure CreerLesTobs;
      procedure LibereLesTobs;
      procedure ChargeTobVignette(TobAcces : tob ; sModule : string);
      procedure AddLesSups(TOBM : TOB ; InitAcces : boolean);
      procedure SauveLesModifs;
      {$IFDEF GCGC}
      procedure MajAccesMenu (TobAcces, TobModule : tob ; TRecap : TStringList);
      procedure MajAccesGroupe(TobAcces : tob ;TRecap : TStringList);
      {$ELSE GCGC}
      procedure MajAccesMenu (TobAcces, TobModule : tob);
      procedure MajAccesGroupe  (TobAcces : tob);
      {$ENDIF GCGC}
      function  UpdateTableMenu(TOBM : TOB  ; ForceMaj : boolean) : boolean;
      function  AccesGroupe(TOBM : TOB ; NumGrp : integer) : string;
      function  ChangeAccesMenu(TOBM : TOB ; NumGrp : integer ; okVisu : boolean) : string;
      procedure FindUserGroupe(Sender: TObject);
      function  isModule(TOBM : TOB) : boolean;
      function  isMenu(TOBM : TOB) : boolean;
      procedure DupliqueLaCol(ColSource, ColTarget : integer);
      //
      function  TagDejaTraite(Tag : String) : boolean;

  end ;

  Const Espacement = 6 ;   //pour décallage des menus
        NbEditGrp = 6 ;    //Edition : nb maxi de groupe utilisat
        //Accès autorisé
        IcoYes = '#ICO#47';
        IcoNewY = '#ICO#44';  //MN_VERSIONDEV='X'
        //Accès refusé
        IcoNo = '#ICO#46';
        IcoNewN = '#ICO#43';  //MN_VERSIONDEV='X'
        //Accès modifié
        IcoYes2 = '#ICO#89';
        IcoNo2 = '#ICO#88';

        TexteMessage: array[1..5] of string 	= (
          {1}  'Enregistrer les modifications ?'
          {2} ,'Confirmez vous la mise à jour ?'
          {3} ,'La mise à jour s''est correctement effectuée'
          {4} ,'Problème rencontré. La mise à jour ne s''est pas correctement effectuée'
          {5} ,'Les droits d''accès vont être réappliqués, voulez-vous continuer?'
              );

Implementation

Uses
  UiUtil, HPanel {JP 19/08/05 :  FQ COMPTA 15453}
  {$IFDEF COMPTA}
  , Ent1, UtilPGI
  {$ENDIF !COMPTA}
  ;

function YYLanceFiche_AccesPortail(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
  Result:='';
  if Nat='' then exit;
  if Cod='' then exit;
  Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_ACCESPORTAIL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_ACCESPORTAIL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_ACCESPORTAIL.OnUpdate ;
begin
  Inherited ;
  if Action = taConsult then exit;
  if PGIAsk(TexteMessage[2],Ecran.Caption) <> mrYes then exit;
  bForceMaj := false;
  Transactions(SauveLesModifs,0);
  GC.VidePile(False);
  TobVignette.PutGridDetail(GC,False,False,LesCols,True);
end ;

procedure TOF_ACCESPORTAIL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_ACCESPORTAIL.OnArgument (S : String) ;
var st,stArgument,stPlus,ValeurCritere : string;
    iInd : integer;
    DOMAINE: THValComboBox; //PMJ 20/12/04
begin
  Inherited ;
  RowColor1 := THEdit(GetControl('COLOR1')).Color;
  RowColor2 := THEdit(GetControl('COLOR2')).Color;
  RowColor3 := THEdit(GetControl('COLOR3')).Color;
  Action := taModif ; bModif := False;
  stArgument := S ; Modules :=  '';
  Repeat
    st:=Trim(ReadTokenSt(StArgument)) ;
    if st <> '' then
        begin
        iInd:=pos('=',st);
        if iInd<>0 then
           begin
           ValeurCritere:=copy(st,iInd+1,length(st));
           if ValeurCritere='CREATION' then begin Action:=taCreat ; end ;
           if ValeurCritere='MODIFICATION' then begin Action:=taModif ; end ;
           if ValeurCritere='CONSULTATION' then begin Action:=taConsult ; end ;
           end else if isNumeric(st) then Modules := Modules + ';' + st;
        end;
  until st = '';

  //Initialisation
  AffecteLesControls;
  CreerLesTobs;
  TobUserGrp.LoadDetailFromSQL('SELECT UG_GROUPE,UG_NUMERO,UG_LIBELLE FROM USERGRP');
  TobUser.LoadDetailFromSQL('SELECT US_UTILISATEUR,US_LIBELLE,UG_LIBELLE,UG_GROUPE,US_GRPSDELEGUES FROM UTILISAT ' +
           'LEFT JOIN USERGRP ON US_GROUPE=UG_GROUPE ORDER BY US_LIBELLE,UG_LIBELLE');

  ColGrp := 0; //ColGrp contient le n° de colonne sur l'evt mousedown
  CopierColler := -1;

  setControlText('DOMAINE','');
  cbDomaine_Value := '';
  stPlus :=  RecupWhereDomaine(cbDomaine_Value);
  if stPlus <> '' then stPlus := ' AND ' + stPlus;

  setControlProperty('DOMAINE','Plus',stPlus);
  DOMAINE := THValComboBox(Getcontrol('DOMAINE')); //PMJ 20/12/04, traduction
  for iind := 0 to DOMAINE.ITems.Count - 1 do
    DOMAINE.ITems[iind] := TraduireMemoire(copy(DOMAINE.ITems[iind],1,pos('(',DOMAINE.ITems[iind])-2))+' '+copy(DOMAINE.ITems[iind],pos('(',DOMAINE.ITems[iind]),255);

  ChargeTobVignette(TobVignette, GetControlText('DOMAINE'));

  cbDomaine_Value := GetControlText('DOMAINE');

  if Assigned(GetControl('BREAPPLIC')) then SetControlVisible('BREAPPLIC',V_PGI.SAV);

  ChargeLesBitmaps;

  HMTrad.ResizeGridColumns(GC) ;
  FormateLesCols;

  TobVignette.PutGridDetail(GC,False,False,LesCols,True);
end ;

procedure TOF_ACCESPORTAIL.OnClose ;
begin
 // Inherited ;
  bForceMaj := false;
  if (Action <> taConsult) and (bModif) then
    if PGIAsk(TexteMessage[1],Ecran.Caption) = mrYes then Transactions(SauveLesModifs,0);

  LibereLesTobs;
  if assigned(FindLigne) then  FindLigne.Free;

  {JP 19/08/05 :  FQ COMPTA 15453 : Avec les voutons en bas de fiche, lorsque l'on clique sur la croix
                  rouge, puis sur Génération PGI, on a une violation d'accès}
  if IsInside(Ecran) then
  begin
  //YMO   23/11/2005  FQ17013 Pb d'affichage à la fermeture(fond d'écran garde les menus PGI)
 //   CloseInsidePanel(Ecran) ;
    CloseInsidePanel(Ecran) ;
    THPanel(Ecran.parent).InsideForm := nil;
    THPanel(Ecran.parent).VideToolBar;
  end;

  {FIN JP 19/08/05}

end ;

procedure TOF_ACCESPORTAIL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_ACCESPORTAIL.OnCancel () ;
begin
  Inherited ;
end ;

///// Form
procedure TOF_ACCESPORTAIL.AffecteLesControls;
var iInd : integer;
begin
  TFVierge(Ecran).OnKeyDown := GCFormKeyDown;
  //Grid
  GC := THGrid(GetControl('GC'));
  GUser := THGrid(GetControl('GUSER'));
  //js1 140206
  GListUsers := THGrid(GetControl('GListUsers'));
  if assigned(GetControl('GROLE')) then
    GRole := THGrid(GetControl('GROLE'));

  GC.PostDrawCell := DessineCell;
  GC.GetCellCanvas := GCGetCellCanvas;
  GC.OnRowEnter := GCRowEnter;
  GC.OnDblClick := GCDblClick;
  GUser.OnDblClick := GUserDblClick;
  GUser.OnRowEnter := GUserRowEnter;
  GC.OnMouseUp := GCMouseUp; GC.OnMousedown := GCMouseDown;
  GC.OnMouseMove := GCMouseMove;
  GC.OnDragOver := GCDragOver; GC.OnEndDrag := GCEndDrag;

  THSplitter(GetControl('SPLITTER')).OnMoved := SplitMoved;

  //Sélection des modules
  THValComboBox(GetControl('DOMAINE')).OnChange := cbMenuChange;

  //Boutons autorisation
  //js1 130206
  bListUsers := TMenuItem(GetControl('MNListUsers'));
  bAcc := TMenuItem(GetControl('MNAUTORISE'));
  bRef := TMenuItem(GetControl('MNREFUSE'));
  bUser := TMenuItem(GetControl('MNUSER'));
  bCop := TMenuItem(GetControl('MNCOPIE'));
  if Assigned(GetControl('MNUSERGRP')) then
  begin
     bUserGrp := TMenuItem(GetControl('MNUSERGRP'));
     bUserGrp.OnClick := bUserGrpClick;
   end;
  
  //js1 130206
  bListUsers.Onclick := bListUsersClick;
  bAcc.OnClick := bAccClick;
  bRef.OnClick := bRefClick;
  bUser.OnClick := bUserClick;
  bCop.OnClick := bCopClick;

  TToolBarButton97(GetControl('BAUTORISER')).OnClick := bAccClick;
  TToolBarButton97(GetControl('BREFUSER')).OnClick := bRefClick;

  //Recherche UserGroupe
  THEdit(GetControl('USERNAME')).OnChange := FindUserGroupe;
  THValComboBox(GetControl('VALUSERGRP')).OnChange := FindUserGroupe;

  //Recherche
  TToolBarButton97(GetControl('BCHERCHER')).OnClick := bChercherClick;
  FindLigne := TFindDialog.Create(TFVierge(Ecran));
  FindLigne.OnFind := FindLigneFind;

  //Edition
  if Assigned(GetControl('BAPERCU')) then TToolbarButton97(GetControl('BAPERCU')).OnClick := BApercuClick;

  //Réaplique les critères
  if Assigned(GetControl('BREAPPLIC')) then TToolbarButton97(GetControl('BREAPPLIC')).OnClick := BReapplicClick;

  for iInd := 1 to NbEditGrp do
    THValComboBox(GetControl('USERGRP' + intToStr(iInd))).OnEnter := ValUserGrpEnter;
end;

///// Grid
procedure TOF_ACCESPORTAIL.DessineCell (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
var Arect: Trect ;
    TOBM : TOB;
    NumMenu : integer;
begin

  if Arow < GC.Fixedrows then exit;

  TOBM := TOB(GC.Objects[0,ARow]);
  if TOBM = nil then exit;
  if not TOBM.FieldExists('NUMVIG') then exit;
  Arect:=GC.CellRect(Acol,Arow) ;
  NumMenu := TOBM.GetInteger('NUMVIG');
  case NumMenu of
    1 : Canvas.Brush.Color := RowColor1;
    2 : if ACol > 0 then exit else Canvas.Brush.Color := clWindow;
  end;
  Canvas.FillRect(ARect);
  Canvas.TextRect(ARect,ARect.Left+2,ARect.Top+2,GC.Cells[ACol,ARow]);
end;

Procedure TOF_ACCESPORTAIL.GCGetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
var TOBM : TOB;
begin

  if ACol = 0 then
  begin
    TOBM := TOB(GC.Objects[0,ARow]);
    if TOBM = nil then exit;
    if isModule(TOBM) then
    begin
      Canvas.Font.Style := [fsBold];
      Canvas.Font.Color := clWhite;
    end;
  end;
end;

procedure TOF_ACCESPORTAIL.GCRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
//js1 150206
  if GetControlVisible('PINFO') then BListUsersClick(Self); //js1 150206
end;

procedure TOF_ACCESPORTAIL.GCDblClick(Sender: TObject);
var ACol, ARow : integer;
    Arect : TGridRect;
begin
  ARect := GC.Selection ;
  if (ARect.Top <> Arect.Bottom) or (ARect.Left <> Arect.Right) then exit; //Multi sélection
  ACol := GC.Col; ARow := GC.Row;
  ChangeCellValue(ACol, ARow);
end;

procedure TOF_ACCESPORTAIL.GUserDblClick(Sender: TObject);
begin
  ZoomUtilisateur;
end;

procedure TOF_ACCESPORTAIL.GUserRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  AfficheGrpAssocies;
end;

procedure TOF_ACCESPORTAIL.GCMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var ARect : TGridRect;
    ACol,ARow,iInd : integer;
    okStop : boolean;
    TOBM1,TOBM2 : TOB;
begin

  GC.MouseToCell(X,Y,ACol,ARow);

  ///////////////////////////////
  // Gestion des boutons
  bAcc.Visible := (ARow <> 0);
  bRef.Visible := (ARow <> 0);
  bListUsers.Visible := (ARow <> 0); //js1 150206
  bUser.Visible := (ARow = 0);
  bCop.Visible := (ARow = 0);
  ColGrp := ACol;

    ///////////////////////////////
  // Gestion de la sélection
  if Button <> mbLeft then exit;
  if (ACol = 0) and (ARow = 0) then exit;
  if (ACol < 0) or (ARow < 0) then exit;

  if ARow = 0 then
  begin
    SelectLaCol(ACol);
    exit;
  end;

  TOBM1 := TOB(GC.Objects[0,ARow]);
  if TOBM1 = nil then exit;
  if (ACol = 0) and isMenu(TOBM1) then
  begin
    SelectLaRow(ARow);
    exit;
  end;

  if (ACol > 0) and isMenu(TOBM1) then exit;

  okStop := False;
  for iInd := ARow+1 to GC.RowCount do
  begin

    TOBM2 := TOB(GC.Objects[0,iInd]);
    if TOBM2 = nil then okStop := True;
    case TOBM2.GetInteger('NUMVIG') of
      1 : okStop := True;
    end;

    if iInd = GC.RowCount-1 then okStop := True;

    if okStop then
    begin
      if ACol = 0 then
      begin
        ARect.Left := 1;
        ARect.Right := GC.ColCount;
        ARect.Top := ARow+1;
      end else
      begin
        ARect.Left := ACol;
        ARect.Right := ACol;
        ARect.Top := ARow+1;
      end;
      if iInd = GC.RowCount-1 then ARect.Bottom := iInd
      else ARect.Bottom := iInd-1;
      GC.Selection := ARect;
      exit;
    end;
  end;
end;

procedure TOF_ACCESPORTAIL.GCMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var MyRect : TGridRect;
    ACol, Arow : integer;
begin
 //pour éviter la sélection d'un groupe (visuel)
 if Button <> mbLeft then exit;
  GC.MouseToCell(X,Y,ACol,ARow);
  if (ACol <= 0) or (ARow <= 0) then exit;
  if GC.Cells[ACol,ARow] <> '' then exit;
  myRect := GC.Selection;
  myRect.Top := myRect.Top+1 ;
  GC.Selection := myRect;
end;

procedure TOF_ACCESPORTAIL.GCMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
var ACol, ARow : integer;
begin
  GC.MouseToCell(X,Y,ACol,ARow);
  if ((ARow = 0) and (ACol > 0)) or
     ((ACol = 0) and (ARow > 0)) then
  begin
    GC.Hint := Trim(GC.Cells[ACol,ARow]);
    GC.ShowHint := True;
  end
  else GC.ShowHint := False;
end;

procedure TOF_ACCESPORTAIL.GCDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
var ACol, ARow : integer;
begin
  if Source is ThGrid then
  begin
    ThGrid(Source).MouseToCell(X,Y,ACol,ARow);
    Accept := ((ARow = 0) and (ACol > 0) and (ACol <> CopierColler));
  end else Accept := false;
end;

procedure TOF_ACCESPORTAIL.GCEndDrag(Sender, Target: TObject; X, Y: Integer);
var ACol,ARow : integer;
begin
  if Target is ThGrid then
  begin
    ThGrid(Target).MouseToCell(X,Y,ACol,ARow);
    if (ARow = 0) and (ACol > 0) and (ACol <> CopierColler) then
    begin
      SelectLaCol(ACol);
      DupliqueLaCol(CopierColler,ACol);
    end;
  end;
end;

procedure TOF_ACCESPORTAIL.GCFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var FocusGrid : Boolean;
begin
FocusGrid := (Screen.ActiveControl = GC);
Case Key of
    VK_RETURN : Key:=VK_TAB ;
    {Ctrl+A}  65 : if Shift=[ssCtrl] then if FocusGrid then begin Key:=0 ; ChangeValueSelection(GC.Selection,'X'); end ;
    {Ctrl+R}  82 : if Shift=[ssCtrl] then if FocusGrid then begin Key:=0 ; ChangeValueSelection(GC.Selection,'-'); end ;
    {Ctrl+U}  85 : if Shift=[ssCtrl] then if GC.Col > 0 then begin Key:=0 ; ColGrp := GC.Col ; bUserClick(nil); end ;
    {Ctrl+C}  67 : if Shift=[ssCtrl] then if FocusGrid and (GC.Col > 0) then begin Key:=0 ; bCopClick(nil); end ;
    end;
end;

procedure TOF_ACCESPORTAIL.SplitMoved(Sender: TObject);
begin
  ResizeLesGrids;
end;

///////Actions liées au Grid
// Change la valeur d'une plage cellule
procedure TOF_ACCESPORTAIL.ChangeValueSelection(MyRect : TGridRect ; Value : string);
var iCol,iRow : integer;
begin
if Action = taConsult then exit;
  myRect := GC.Selection;
  for iRow := myRect.Top to myRect.Bottom  do
  begin
    for iCol := myRect.Left to myRect.Right do
    begin
      ChangeCellValue(iCol,iRow,Value);
    end;
  end;
end;

// Change la valeur d'une cellule par value ou son contraire
procedure TOF_ACCESPORTAIL.ChangeCellValue(ACol,ARow : integer ; Value : string = '');
var TOBM : TOB;
    Ico : string;
begin
  TOBM := TOB(GC.Objects[0,ARow]);
  if TOBM = nil then exit;
  if (ACol = 0) or (ARow = 0) then exit;

  if pos('#ICO#',GC.Cells[ACol,ARow]) > 0 then
  begin
    if Value = '' then
      if TOBM.GetString('ACCES'+intToStr(ACol)) = 'X' then Value := '-'
      else Value := 'X';
      if Value = 'X' then Ico := IcoYes2 else Ico := IcoNo2;
    GC.Cells[ACol,ARow] := Ico;
    TOBM.SetString('ACCES'+intToStr(ACol),Value);
  end;
  bModif := True;
end;

procedure TOF_ACCESPORTAIL.FormateLesCols;
var iCol : integer;
begin
  GC.ColCount := TobUserGrp.Detail.Count + 1;
  GC.Cells[0,0] := TraduireMemoire('Domaine');
  GC.ColWidths[0] := 200;
  LesCols := 'NIVEAU';
  for iCol := 0 to TobUserGrp.Detail.Count -1 do
  begin
    GC.Cells[iCol+1,0] := TobUserGrp.Detail[iCol].GetString('UG_LIBELLE');
    LesCols := LesCols + ';ACCESICO' + IntToStr(iCol+1);
    GC.ColAligns[iCol+1] := taCenter;
  end;
end;

procedure TOF_ACCESPORTAIL.SelectLaCol(ACol : integer);
var ARect : TGridRect;
begin
  ARect.Left := ACol;
  ARect.Top := 1;
  ARect.Right := ACol;
  ARect.Bottom := GC.RowCount;
  GC.Selection := ARect;
end;

procedure TOF_ACCESPORTAIL.SelectLaRow(ARow : integer);
var ARect : TGridRect;
begin
  ARect.Left := 1;
  ARect.Top := ARow;
  ARect.Right := GC.ColCount;
  ARect.Bottom := ARow;
  GC.Selection := ARect;
end;

procedure TOF_ACCESPORTAIL.ResizeLesGrids;
begin
  HMTrad.ResizeGridColumns(GUser) ;
  HMTrad.ResizeGridColumns(GRole) ;
end;

function TOF_ACCESPORTAIL.RecupWhereDomaine(Module : string ) : string;
var st,stNumMenu : string;
begin
  Result := '';
  stNumMenu := Module;
  if stNumMenu <> '' then
  begin
    Result := 'WHERE SWV_VREGROUPE="' + stNumMenu + '"';
    exit
  end;
  stNumMenu := Modules;
  st := ReadTokenSt(stNumMenu);
  while (stNumMenu <> '') or (st <> '') do
  begin
    if st <> '' then
    begin
      if Result <> '' then Result := Result + ' OR ';
      Result := Result + 'SWV_VREGROUPE="' + st + '"';
    end;
    st := ReadTokenSt(stNumMenu);
  end;
  if Result <> '' then Result := ' (' + Result + ') ';
end;

procedure TOF_ACCESPORTAIL.cbMenuChange(Sender: TObject);
begin
  if GetControlText('DOMAINE') = cbDomaine_Value then exit;
  if (Action <> taConsult) and bModif then
    begin
      bForceMaj := false;
      Case PGIAskCancel(TexteMessage[1],Ecran.Caption) of
           mrYes : Transactions(SauveLesModifs,0);
           mrCancel : exit;
      end;
      bModif := False;
    end;
  TobVignette.ClearDetail;
  GC.VidePile(False);
  ChargeTobVignette(TobVignette, GetControlText('DOMAINE'));

  TobVignette.PutGridDetail(GC,False,False,LesCols,True);
  cbDomaine_Value := GetControlText('DOMAINE');
end;

procedure TOF_ACCESPORTAIL.ZoomUtilisateur;
var TobU : TOB;
begin
  TobU := TOB(GUser.Objects[GUser.Col,GUser.Row]);
  if TobU <> nil then
    AGLLanceFiche('YY','YYUTILISAT','',TobU.GetString('US_UTILISATEUR'),'ACTION=CONSULTATION');
end;

//Boutons
procedure TOF_ACCESPORTAIL.bChercherClick(Sender: TObject);
begin
  FFirstFind := true;
  FindLigne.Execute;
end;

procedure TOF_ACCESPORTAIL.bUserGrpClick(Sender: TObject);
begin
  //SetInfoCompVisible(not (GetControlVisible('PINFO')),'TABUSER');
  SetInfoCompVisible(true,'TABUSER');
  SetControlText('USERNAME','');
  SetControlText('VALUSERGRP','');
end;


procedure TOF_ACCESPORTAIL.AfficheGrpAssocies;
var TobU,TobRole,TobR : TOB;
    sTokenRoles,sRole,GrpDefaut : string;

    function GetGrpLibelle(Grp : string) : string;
    var TobLib :tob;
    begin
      TobLib := TobUserGrp.FindFirst(['UG_GROUPE'],[Grp],false);
      if assigned(TobLib) then result := TobLib.GetString('UG_LIBELLE')
      else result := '';
    end;
begin
  GRole.VidePile(false);
  TobU := TOB(GUser.Objects[0,GUser.Row]);
  if TobU <> nil then
  begin
    TobRole := TOB.create('',nil, -1);
    try
      GrpDefaut := TobU.GetString('UG_GROUPE');
      sTokenRoles := TobU.GetString('US_GRPSDELEGUES');
      if pos(GrpDefaut + ';', sTokenRoles) <= 0 then
      sTokenRoles := GrpDefaut + ';' + sTokenRoles;
      sRole := ReadTokenSt(sTokenRoles);
      while (sTokenRoles <> '') or (sRole <> '') do
      begin
        if sRole <> '' then
        begin
          TobR := TOB.create('', TobRole, -1);
          TobR.AddChampSupValeur('USROLE', GetGrpLibelle(sRole));
        end;
        sRole := ReadTokenSt(sTokenRoles);
      end;
      TobRole.PutGridDetail(GRole,False,False,'USROLE',True);
    finally
      TobRole.free;
    end;
    GRole.SortGrid(0,false);
  end;
end;

procedure TOF_ACCESPORTAIL.FindLigneFind(Sender: TObject);
begin
  Rechercher(GC, FindLigne, FFirstFind);
end;

procedure TOF_ACCESPORTAIL.BAccClick(Sender: TObject);
var myRect : TGridRect;
begin
  myRect := GC.Selection;
  ChangeValueSelection(myRect,'X');
end;

procedure TOF_ACCESPORTAIL.BRefClick(Sender: TObject);
var myRect : TGridRect;
begin
  myRect := GC.Selection;
  ChangeValueSelection(myRect,'-');
end;

procedure TOF_ACCESPORTAIL.BUserClick(Sender: TObject);
var ACol : integer;
begin
  ACol := ColGrp;
  if (ACol <= 0) or (ACol > GC.ColCount) then exit;
  setControlText('VALUSERGRP',TobUserGrp.Detail[ACol-1].GetString('UG_GROUPE'));
  if not GetControlVisible('PINFO') then
  begin
    SetControlText('USERNAME','');
    SetInfoCompVisible(True,'TABUSER');
  end;
  FindUserGroupe(nil);
end;

procedure TOF_ACCESPORTAIL.BCopClick(Sender: TObject);
begin
  SelectLaCol(ColGrp);
  GC.BeginDrag(true);
  CopierColler := GC.Col;
end;

procedure TOF_ACCESPORTAIL.SetInfoCompVisible(Visu : boolean ; TabSheetName : string);
begin
  if Visu and (not GetControlVisible('PINFO')) then
  begin
    SetControlVisible('PINFO', Visu);
    SetActiveTabSheet(TabSheetName);
    SetControlVisible('SPLITTER',Visu);
    SetControlProperty('PINFO','Width',(GC.Width/4)); // 1/4 du grid
    HMTrad.ResizeGridColumns(GUser) ;
    SetControlProperty('SPLITTER','Align',alRight);
    ResizeLesGrids;
  end else  SetActiveTabSheet(TabSheetName);
end;

procedure TOF_ACCESPORTAIL.BApercuClick(Sender: TObject);
var iInd, iInd2 : integer;
    TobEdit, TOBE, TOBU : TOB;
    ValGrp : string;
    ColEdit : array[1..NbEditGrp] of string;
    LanceEdt : boolean;
begin
  TobEdit := TOB.Create('',nil,-1);
  try
    LanceEdt := false;
    for iInd := 1 to NbEditGrp do
    begin
      ColEdit[iInd] := '';
      ValGrp := GetControlText('USERGRP' + intToStr(iInd));
      if ValGrp = ''  then continue;
      TOBU := TobUserGrp.FindFirst(['UG_GROUPE'],[ValGrp],false);
      if TOBU = nil then continue;
      if pos('ACCES' + intToStr(TOBU.GetIndex+1),ColEdit[iInd]) <= 0 then
      begin
        ColEdit[iInd] := 'ACCES' + intToStr(TOBU.GetIndex+1);
        LanceEdt := true;
      end;
    end;
    if LanceEdt then
      for iInd := 0 to TobVignette.Detail.Count -1 do
      begin
        TOBE := TOB.Create('',TobEdit,-1);
        TOBE.AddChampSupValeur('NIVEAU',TobVignette.Detail[iInd].GetString('NIVEAU'),false);
        TOBE.AddChampSupValeur('NUMVIG',TobVignette.Detail[iInd].GetString('NUMVIG'),false);
        TOBE.AddChampSupValeur('VIGNETTE',TobVignette.Detail[iInd].GetString('SWV_LIBELLE'),false);
        for iInd2 := 1 to NbEditGrp do
        begin
          TOBE.AddChampSupValeur('ACCESGRP'+intToStr(iInd2),'',false);
          if ColEdit[iInd2] <> '' then TOBE.SetString('ACCESGRP'+intToStr(iInd2),TobVignette.Detail[iInd].GetString(ColEdit[iInd2]));
        end;
      end;
    SetControlText('EDTMODULE',THValComboBox(GetControl('DOMAINE')).Text);
    LanceEtatTob('E','CFD','CFV',TobEdit,True,False,False,TPageControl(GetControl('Pages')),'',Ecran.Caption,False);
  finally
    TobEdit.Free;
  end;
  SetControlVisible('TEDIT',false);
end;

//Pour réappliquer les droits..
procedure TOF_ACCESPORTAIL.BReapplicClick(Sender: TObject);
begin
  if Action = taConsult then exit;
  if PGIAsk(TexteMessage[5],Ecran.Caption) <> mrYes then exit;
  bForceMaj := true;
  Transactions(SauveLesModifs,0);
  GC.VidePile(False);
  TobVignette.PutGridDetail(GC,False,False,LesCols,True);
end;

procedure TOF_ACCESPORTAIL.ChargeLesBitmaps;
begin
  ChargeImageList;
  V_PGI.GraphList.GetBitmap(32-1, bListUsers.Bitmap);      //js1 150206
  V_PGI.GraphList.GetBitmap(47-1, bacc.Bitmap);     
  V_PGI.GraphList.GetBitmap(46-1, bRef.Bitmap);
  V_PGI.GraphList.GetBitmap(96-1, bUser.Bitmap);
  V_PGI.GraphList.GetBitmap(51-1, bCop.Bitmap);
  if Assigned(bUserGrp) then V_PGI.GraphList.GetBitmap(96-1, bUserGrp.Bitmap);
  if Assigned(bMnAss) then V_PGI.GraphList.GetBitmap(65-1, bMnAss.Bitmap);
end;

//Pour édition..
procedure TOF_ACCESPORTAIL.ValUserGrpEnter(Sender: TObject);
var Champ,Valeur : string;
    Plus, St : string;
    Indice, iInd : integer;
begin
  Champ := THValComboBox(Sender).Name;
  Indice := strToInt(copy(Champ,length(Champ),1));
  Valeur := GetControlText(Champ);

  for iInd := 1 to NbEditGrp do
  begin
    if iInd = Indice then continue;
    St := GetControlText('USERGRP'+intToStr(iInd));
    if St <> '' then Plus := Plus + ' AND UG_GROUPE<>"' + St + '"' ;
  end;
  SetControlProperty('USERGRP' + intToStr(Indice),'Plus',Plus);
  SetControlText('USERGRP' + intToStr(Indice),Valeur);
end;

//
procedure TOF_ACCESPORTAIL.CreerLesTobs;
begin
  TobVignette := TOB.Create('',nil,-1);
  TobUserGrp := TOB.Create('',nil,-1);//Contient les groupes des users
  TobUser := TOB.Create('',nil,-1); //Contient tous les users
  TobUserName := TOB.Create('',nil,-1); //Contient les users du groupe sélectionné
  TobAssoc := TOB.Create('',nil,-1); //Contient l'ensemble des menus pour la visu des tags associés
//js1 160206
  TobMyResult := TOB.Create('',nil,-1);
end;

procedure TOF_ACCESPORTAIL.LibereLesTobs;
begin
  TobVignette.Free;
  TobUserGrp.Free;
  TobUser.Free;
  TobUserName.Free;
  TobAssoc.Free;
//js1 160206
  TobMyResult.Free;
end;

{ sModule peut contenir une chaine de modules séparées par des ; }
procedure TOF_ACCESPORTAIL.ChargeTobVignette(TobAcces : tob ; sModule : string);
var iInd,NumMenu : integer;
    TobModules,TOBM : TOB;
    QMn : TQuery;
    stWhere : string;

    function SpaceStr ( nb : integer) : string;
    var St_Chaine : string ;
        i_ind : integer ;
    begin
      St_Chaine := '' ;
      for i_ind := 1 to nb do St_Chaine:=St_chaine+' ';
      Result:=St_Chaine;
    end;
begin
  { Chargement du domaines }
  TobModules := TOB.Create('',nil,-1);
  try
    stWhere :=  RecupWhereDomaine(sModule);
    QMn :=  OpenSQL('SELECT DISTINCT(SWV_VREGROUPE) FROM SWVIGNETTES ' + stWhere,True);
    try
      if not QMn.Eof then TobModules.LoadDetailDB('SWVIGNETTE','','',QMn,true);
    finally
      Ferme(QMn);
    end;
    for iInd := 0 to TobModules.Detail.count - 1 do  //PMJ 20/12/04, traduction
    begin
      TobModules.Detail[iInd].PutValue('SWV_VREGROUPE',Traduirememoire(TobModules.Detail[iInd].GetValue('SWV_VREGROUPE')));
      TobModules.Detail[iInd].AddChampSupValeur('SWV_CODEVIGNETTE','');
      TobModules.Detail[iInd].AddChampSupValeur('SWV_LIBELLE','');
      TobModules.Detail[iInd].AddChampSupValeur('SWV_ACCESGRP','');
      TobModules.Detail[iInd].AddChampSupValeur('SWV_TYPEVIGNETTE','');
    end;

  { Chargement des vignettes }
    stWhere :=  RecupWhereDomaine(sModule);
    QMn := OpenSQL('SELECT * FROM SWVIGNETTES ' + stWhere + ' ORDER BY SWV_VREGROUPE,SWV_LIBELLE',True);
    try
      if not QMn.Eof then TobAcces.LoadDetailDB('SWVIGNETTE','','',QMn,true);
    finally
      Ferme(QMn);
    end;
    for iInd := 0 to TobAcces.Detail.count - 1 do  //PMJ 20/12/04, traduction
      TobAcces.Detail[iInd].PutValue('SWV_LIBELLE',Traduirememoire(TobAcces.Detail[iInd].GetValue('SWV_LIBELLE')));

    for iInd := TobModules.Detail.Count -1 downto 0 do
    begin
      TOBM := TobAcces.FindFirst(['SWV_VREGROUPE'],[TobModules.Detail[iInd].GetString('SWV_VREGROUPE')],False);
      if TOBM = nil then continue;
      TobModules.Detail[iInd].ChangeParent(TobAcces,TOBM.GetIndex);
    end;
  finally
    TobModules.Free;
  end;

  { Création des niveaux dans NUMMENU }
  { 1 = module } { 2 = groupe } { 3 = ss groupe } { 4 = menu de base }
  for iInd := 0 to TobAcces.Detail.Count -1 do
  begin
    TOBM := TobAcces.Detail[iInd];
    TOBM.AddChampSup('NUMVIG',False);
    if TOBM.GetString('SWV_LIBELLE')= '' then TOBM.SetInteger('NUMVIG',1)
      else TOBM.SetInteger('NUMVIG',2);

    NumMenu  := TOBM.GetInteger('NUMVIG');
    case NumMenu of
      1 : begin AddLesSups(TOBM,False); TOBM.SetString('NIVEAU',VireSouligne(RechDom('SWREGROUPEMENT',TOBM.GetString('SWV_VREGROUPE'),false))); end;
      2 : begin AddLesSups(TOBM,True); TOBM.SetString('NIVEAU',SpaceStr((TOBM.GetInteger('NUMVIG')-1)*Espacement)+VireSouligne(TOBM.GetString('SWV_LIBELLE'))); end;
    end;
  end;
end;

procedure TOF_ACCESPORTAIL.SauveLesModifs;
var TobAccesParent, TobModule : tob;
    sModule, sInfo : string;
    iIndex, iMod : integer;
   {$IFDEF GCGC}
    TLBRecap : TStringList;
   {$ENDIF}
begin
  {$IFDEF GCGC}
  TLBRecap := TStringList.Create;
  try
    TLBRecap.Add('********************************');
    TLBRecap.Add(TraduireMemoire('Module : ') + THValComboBox(GetControl('DOMAINE')).Text);
    TLBRecap.Add('********************************');
    TLBRecap.Add('');
    TLBRecap.Add(TraduireMemoire('Tag') + '      ' + TraduireMemoire('DOMAINE'));
    TLBRecap.Add('----------------');
  {$ENDIF}
    SetLength(TTag,0);

    TobVignette.GetGridDetail(GC,GC.RowCount,'',LesCols);

    { On modifie les accès des menus de bas niveau }
    TobModule := tob.Create('',nil,-1);
    try
      {$IFDEF GCGC}
      MajAccesMenu (TobVignette, TobModule, TLBRecap);
      {$ELSE GCGC}
      MajAccesMenu (TobVignette, TobModule);
      {$ENDIF GCGC}

      { On reconstruit les accès des menus parents en fonctions des modules modifiés contenu dans tobmodule }
      TobAccesParent := tob.create('',nil,-1);
      try
        InitMoveProgressForm(nil,Ecran.caption, 'Traitement en cours, veuillez patienter...', TobModule.Detail.Count, False, true) ;
        try
          Application.ProcessMessages;
          for iMod := 0 to TobModule.Detail.Count -1 do
          begin
            TobAccesParent.ClearDetail;
            sModule := TobModule.Detail[iMod].GetString('MODULE');
            ChargeTobVignette(TobAccesParent, sModule);
            {$IFDEF GCGC}
            MajAccesGroupe (TobAccesParent, TLBRecap);
            {$ELSE GCGC}
            MajAccesGroupe (TobAccesParent);
            {$ENDIF GCGC}
            iIndex := ThValComboBox(GetControl('DOMAINE')).Values.IndexOf(sModule);
            if iIndex > 0 then sInfo := TraduireMemoire('Mise à jour du module ') + ThValComboBox(GetControl('DOMAINE')).Items[iIndex]
            else sInfo := TraduireMemoire('Mise à jour du module ') + sModule;
            MoveCurProgressForm(sInfo);
          end;
        finally
          FiniMoveProgressForm ;
        end;
      finally
        TobAccesParent.Free;
      end;
      {$IFDEF GCGC}
      {$IFNDEF PGIMAJVER}
      MAJJnalEvent('ADM','OK', Ecran.Caption, TLBRecap.Text);
      {$ENDIF}
      {$ENDIF}
    finally
      TobModule.Free;
    end;
  {$IFDEF GCGC}
  finally
    TLBRecap.Free;
  end;
  {$ENDIF}
end;

{$IFDEF GCGC}
procedure TOF_ACCESPORTAIL.MajAccesGroupe (TobAcces : tob ; TRecap : TStringList);
{$ELSE GCGC}
procedure TOF_ACCESPORTAIL.MajAccesGroupe (TobAcces : tob);
{$ENDIF GCGC}
var iInd:integer;
    Tag : String;
    TOBM : TOB;
begin
  for iInd := 0 to TobAcces.Detail.Count -1 do
  begin
    TOBM := TobAcces.Detail[iInd];
    Tag := TOBM.GetString('SWV_VREGROUPE') + ';' + TOBM.GetString('SWV_CODVIGNETTE');
    if not TagDejaTraite(Tag) or isModule(TOBM) then
      if UpdateTableMenu(TOBM, bForceMaj) then
      begin
        {$IFDEF GCGC}
        TRecap.Add(TOBM.GetString('SWV_CODEVIGNETTE') + '  ' + TOBM.GetString('SWV_ACCESGRP'));
        {$ENDIF}
      end;
  end;
  bModif := False;
end;


{ MAJ des menus de bas niveaux }
{$IFDEF GCGC}
procedure TOF_ACCESPORTAIL.MajAccesMenu (TobAcces, TobModule : tob ; TRecap : TStringList);
{$ELSE GCGC}
procedure TOF_ACCESPORTAIL.MajAccesMenu (TobAcces, TobModule : tob);
{$ENDIF GCGC}
var iInd,iGrp,NumeroGrp,Ico : integer;
    TOBM : TOB;
    Tag : String;
begin
  InitMoveProgressForm(nil,Ecran.caption,'Traitement en cours, veuillez patienter...', TobAcces.Detail.Count, False, true) ;
  try
    Application.ProcessMessages;
    for iInd := 0 to TobAcces.Detail.Count -1 do
    begin
      TOBM := TobAcces.Detail[iInd];

      MoveCurProgressForm();
      Tag := TOBM.GetString('SWV_VREGROUPE') + ';' + TOBM.GetString('SWV_CODEVIGNETTE');
      { on vérifie pour le changement pour chaque groupe }
      for iGrp := 0 to TobUserGrp.Detail.Count -1 do
      begin
        Ico := TOBM.GetNumChamp('ACCESICO'+IntToStr(iGrp+1));
        if TOBM.GetValeur(Ico) = IcoYes2 then TOBM.PutValeur(Ico,IcoYes) ///réinitialisation pour putgriddetail
        else if TOBM.GetValeur(Ico) = IcoNo2 then TOBM.PutValeur(Ico,IcoNo);
        NumeroGrp := TobUserGrp.Detail[iGrp].GetInteger('UG_NUMERO');
        if isMenu(TOBM) then
        begin
          ChangeAccesMenu(TOBM,NumeroGrp,(TOBM.GetString('ACCES'+IntToStr(iGrp+1)) = 'X'));
        end;
      end;
      if not TagDejaTraite(Tag) then
        if UpdateTableMenu(TOBM, bForceMaj) then
        begin
          {$IFDEF GCGC}
          TRecap.Add(Tag + '  ' + TOBM.GetString('SWV_ACCESGRP'));
          {$ENDIF}
        end;
    end;
  finally
    FiniMoveProgressForm ;
  end;
  bModif := False;
end;

function TOF_ACCESPORTAIL.ChangeAccesMenu(TOBM : TOB ; NumGrp : integer ; okVisu : boolean) : string;
var Acces : string;
begin
  if TOBM = nil then exit;
  Acces := TOBM.GetString('SWV_ACCESGRP');
  if okVisu then
  begin
    if Acces[NumGrp] <> '0' then begin Acces[NumGrp] := '0'; TOBM.SetString('SWV_ACCESGRP',Acces); end;
  end else
  begin
    if Acces[NumGrp] <> '-' then begin Acces[NumGrp] := '-'; TOBM.SetString('SWV_ACCESGRP',Acces); end;
  end;
end;

function TOF_ACCESPORTAIL.AccesGroupe(TOBM : TOB ; NumGrp : integer) : string;
begin
  if TOBM = nil then exit;
  Result := copy(TOBM.GetString('SWV_ACCESGRP'),NumGrp,1);
end;

procedure TOF_ACCESPORTAIL.AddLesSups(TOBM : TOB ; InitAcces : boolean);
var iGrp : integer;
    TOBU : TOB;
begin
  TOBM.AddChampSup('NIVEAU',False);
  for iGrp := 0 to TobUserGrp.Detail.Count -1 do
  begin
    TOBU := TobUserGrp.Detail[iGrp];
    TOBM.AddChampSupValeur('UG_GROUPE'+IntToStr(iGrp+1),TOBU.GetString('UG_GROUPE'));
    TOBM.AddChampSupValeur('UG_LIBELLE'+IntToStr(iGrp+1),TOBU.GetString('UG_LIBELLE'));
    TOBM.AddChampSupValeur('UG_NUMERO'+IntToStr(iGrp+1),TOBU.GetInteger('UG_NUMERO'));
    TOBM.AddChampSupValeur('ACCESICO'+IntToStr(iGrp+1),'');
    TOBM.AddChampSupValeur('ACCES'+IntToStr(iGrp+1),'');
    if not InitAcces then continue;
    if AccesGroupe(TOBM,TOBU.GetInteger('UG_NUMERO')) = '0' then
    begin
     TOBM.SetString('ACCESICO'+IntToStr(iGrp+1),IcoYes);
     TOBM.SetString('ACCES'+IntToStr(iGrp+1),'X');
    end else
    begin
     TOBM.SetString('ACCESICO'+IntToStr(iGrp+1),IcoNo);
     TOBM.SetString('ACCES'+IntToStr(iGrp+1),'-');
    end;
  end;
end;

procedure TOF_ACCESPORTAIL.FindUserGroupe(Sender: TObject);
var stU,stG : string;
    iInd, Lg : integer;
    TobU : TOB;
begin
  GUser.VidePile(False);
  stU := UpperCase(GetControlText('USERNAME'));
  stG := GetControlText('VALUSERGRP');
  if (Trim(stU) = '') and (stG = '') then
  begin
    TobUser.PutGridDetail(GUser,False,False,'US_LIBELLE;UG_LIBELLE',True);
  end else
  begin
    TobUserName.ClearDetail;
    Lg := Length(stU);
    for iInd := 0 to TobUser.Detail.Count -1 do
    begin
      if stU <> '' then
        if stU <> Uppercase(copy(TobUser.Detail[iInd].GetString('US_LIBELLE'),1,Lg)) then continue;
      if stG <> '' then
        if (stG <> TobUser.Detail[iInd].GetString('UG_GROUPE')) and
           (pos(stG+';',TobUser.Detail[iInd].GetString('US_GRPSDELEGUES')) <= 0) then continue;
      TobU := TOB.Create('',TobUserName,-1);
      TobU.Dupliquer(TobUser.Detail[iInd],True,True);
    end;
    TobUserName.PutGridDetail(GUser,False,False,'US_LIBELLE;UG_LIBELLE',True);
  end;
  AfficheGrpAssocies;
end;

function TOF_ACCESPORTAIL.isModule(TOBM : TOB) : boolean;
begin
  result := (TOBM.GetInteger('NUMVIG') = 1);
end;

function TOF_ACCESPORTAIL.isMenu(TOBM : TOB) : boolean;
begin
  result := (TOBM.GetInteger('NUMVIG') = 2)
         and (TOBM.GetString('SWV_LIBELLE') <> '');
end;

procedure TOF_ACCESPORTAIL.DupliqueLaCol(ColSource, ColTarget : integer);
var iRow : integer;
begin
  for iRow := 0 to TobVignette.Detail.Count -1 do
  begin
    ChangeCellValue(ColTarget,iRow+1,TobVignette.Detail[iRow].GetString('ACCES'+intToStr(ColSource)));
  end;
end;

///
function TOF_ACCESPORTAIL.TagDejaTraite(Tag : String) : boolean;
var iInd, NbTag : integer;
begin
  Result := False;
  if (Tag = '') then
  begin
    Result := True;
    exit;
  end;
  NbTag := High(TTag);
  for iInd := 0 to NbTag do
    if TTag[iInd] = Tag then
    begin
      Result := True;
      exit;
    end;
  inc(NbTag);
  if NbTag <= 0 then NbTag := 1;
  SetLength(TTag,NbTag);
  TTag[NbTag-1] := Tag;
end;

function TOF_ACCESPORTAIL.UpdateTableMenu(TOBM : TOB ; ForceMaj : boolean) : boolean;
var SetChamps : string;
begin
  SetChamps := '';
  Result := False;
  //if TOBM.IsFieldModified('SWV_ACCESGRP') or ForceMaj then
  SetChamps := 'SWV_ACCESGRP="' + TOBM.GetString('SWV_ACCESGRP') + '"';
  if SetChamps <> '' then
  begin
    if isMenu(TOBM) then
    begin
      ExecuteSQL('UPDATE SWVIGNETTES SET ' + SetChamps + ' WHERE SWV_VREGROUPE="' + TOBM.GetString('SWV_VREGROUPE') + '" AND'+
      ' SWV_CODEVIGNETTE ="' + TOBM.GetString('SWV_CODEVIGNETTE') + '"' );
    end;
    TOBM.Modifie := False;
    Result := True;
  end;
end;


//js1 : 130206 : nouvel item de menu sur le click droit permettant de consulter
//l'ensembles des users ayant acces a un tag de menu : lance la fonction ListUsers
// de ListUsersAutorises.pas et affiche la tob retour dans un grid
procedure TOF_ACCESPORTAIL.BListUsersClick(Sender: TObject);
var myrect:TGridRect;
    TobMyRow : tob;

begin

  myRect := GC.Selection;

  GListUsers.VidePile(true);

// si rien de sélectionné exit
  if ((myRect.Top = 1) and (myRect.Bottom = 1)) then exit;

// si plus d'une ligne sélectionnée, on vide le grid
  if ((myRect.Bottom - myRect.Top) > 0) then exit;

  TobMyRow := TOB(GC.Objects[0,GC.Row]);

  ListUsersVignettes(TobMyRow.GetString('SWV_TYPEVIGNETTE') + ';' + TobMyRow.GetString('SWV_CODEVIGNETTE'),TobMyResult);

  if not GetControlVisible('PINFO') then
    SetInfoCompVisible(True,'LISTUSERS');

  TobMyResult.PutGridDetail(GListUsers,False,False,'US_LIBELLE;UG_LIBELLE',True);

end;

Initialization
  registerclasses ( [ TOF_ACCESPORTAIL ] ) ;
end.

