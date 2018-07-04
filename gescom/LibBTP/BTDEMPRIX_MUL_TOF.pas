{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 01/03/2012
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTDEMPRIX_MUL ()
Mots clefs ... : TOF;BTDEMPRIX_MUL
*****************************************************************}
Unit BTDEMPRIX_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     fe_main,
     mul,
{$else}
     MaineAGL,
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     Hpanel,
     uTOFComm,
     uEntCommun,
     uTob,
     Menus,
     Udefexport,
     grids,
     vierge,
     AglInit,
     DialogEx,
     UTOF ;

Type
  TtypeModeTrait = (TtmtSelect,TtmtModif);
  //
  TOF_BTDEMPRIX_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  Private
    fTtypeTrait : TtypeModeTrait;
    FfromSaisie : Boolean;
    //
    TOBparam    : TOB;
    TobPiece    : TOB;
    TobDemPrix  : TOB;
    TOBArtDemPx : TOB;
    TobFrsDemPx : TOB;
    TobDetDemPx : TOB;
    //
    BtCreate    : TToolbarButton97;
    BtDelete    : TToolBarButton97;
    BtExport    : TToolbarButton97;
    BtDetailDdePrix : TToolbarButton97;
    BtSaisieDde : TToolbarButton97;
    //
    PanelDetail : THpanel;
    //
    GrilleDdePx : THGrid;
    GrilleDetail: THGrid;
    //
    //
    Unique      : Integer;
    ColUnique   : Integer;
    ColLibelle  : Integer;
    ColTraite   : Integer;
    ColEnvoye   : Integer;
    ColRefArt   : Integer;
    ColPrixAchat: Integer;

    Cledoc      : R_CLEDOC;

    LibDdePrix    : String;
    LesColDemPrix : String;
    LesColDetail  : String;
    FF            : String;

    //
    procedure AfficheEcran;
    procedure AfficheDetailDemandedePrix;
    procedure DefinitZone;
    procedure DessineLaGrille(Grille : THGrid; LibColonne : String);
    procedure DetailDdePrix(Sender: Tobject);
    procedure ExportOnClick(Sender: Tobject);
    procedure GrilleDetailOnDblClick(Sender: Tobject);
    procedure GrilleOnClick(Sender: Tobject);
    procedure GrilleOnDblClick(Sender: Tobject);
    procedure SaisieDdePrix(Sender: Tobject);
    //
  end ;

const
  // libellés des messages de la TOM Affaire
  TexteMessage : array [1..6] of string = (
    {1} 'Confirmez-vous la suppression de cette demande de prix ?',
    {2} 'Erreur de suppression de la demande de Prix',
    {3} 'Type d''export non défini',
    {4} 'Veuillez renseigner l''emplacement de stockage des exports',
    {5} 'Modèle d''export non défini',
    {6} 'Excel n''est pas installé sur ce poste'
    ) ;

Implementation
uses  UdemandePrix,
      UtilTOBPiece, TntGrids;

Procedure TOF_BTDEMPRIX_MUL.AfficheEcran;
begin

  if TobDemPrix.detail.count = 0 then
  begin
    BtDetailDdePrix.visible := False;
    BtSaisieDde.Visible     := False;
    BtExport.Visible        := False;
    BtDelete.Visible        := False;
  end
  else if TOBArtDemPx.detail.count = 0 then
  begin
    BtDetailDdePrix.visible := False;
    BtSaisieDde.Visible     := False;
    BtExport.Visible        := False;
    BtDelete.Visible        := True;
  end;


  TobDemPrix.detail.Sort('BPP_UNIQUE');

  ChargementGrilleWithTob(TobDemPrix, GrilleDdePx, LesColDemPrix);

  TFVierge (Ecran).HMTrad.ResizeGridColumns(GrilleDdePx);

  GrilleDdePx.Refresh;

  Unique := 0;

end;

procedure TOF_BTDEMPRIX_MUL.OnNew ;
begin
  Inherited ;

  CreateDemandePrix(TobParam, 0);
  if (TOBparam.GetValue('RETOUR')='X') then TobDemPrix.PutValue('MODIFIED','X');
  AfficheEcran;

end ;

procedure TOF_BTDEMPRIX_MUL.OnDelete ;
Var Arow  : Integer;
begin
  Inherited ;

  ARow    := GrilleDdePx.Row-1;

  Unique := StrToInt(TobDemPrix.Detail[Arow].GetString('BPP_UNIQUE'));

  //lecture entete demande de prix
  if (PGIAsk(TexteMessage[1], ecran.Caption) = mrYes) then
  begin
  	TobDemPrix.PutValue('MODIFIED','X');
    SupprimeDdePrix(TobDemPrix,  Unique,FfromSaisie);
    SupprimeDdePrix(TobArtDemPx, Unique,FfromSaisie);
    SupprimeDdePrix(TobFrsDemPx, Unique,FfromSaisie);
    SupprimeDdePrix(TobDetDemPx, Unique,FfromSaisie);
    AfficheEcran;
  end;

end ;

procedure TOF_BTDEMPRIX_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMPRIX_MUL.OnLoad ;
begin
  Inherited ;

  Cledoc.NaturePiece  := TobPiece.GetString('GP_NATUREPIECEG');
  Cledoc.Souche       := TobPiece.GetString('GP_SOUCHE');
  Cledoc.NumeroPiece  := TobPiece.GetInteger('GP_NUMERO');
  Cledoc.Indice       := TobPiece.GetInteger('GP_INDICEG');

  AfficheEcran;
  
end ;

procedure TOF_BTDEMPRIX_MUL.OnArgument (S : String ) ;
var i : integer;
begin
  Inherited ;
  FfromSaisie := False;
  LesColDemPrix := 'FIXED;BPP_UNIQUE;BPP_LIBELLE;BPP_TRAITE;BPP_ENVOYE';
  LesColDetail  := 'FIXED;BDP_ARTICLE;BDP_LIBELLE;BDP_DPA';

  ecran.Caption := 'Gestion des demandes de prix';

  FF:='#';
  if V_PGI.OkDecV>0 then
   begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecV-1 do
       begin
       FF:=FF+'0';
       end;
    FF:=FF+'0';
   end;

  fTTypeTrait := TtmtModif;
  if pos('SELECTION',S)> 0 then fTtypeTrait := TtmtSelect;

  if LaTOB <> nil then
  begin
    TOBparam    := LaTOB;
    TobPiece    := TOB(TOBParam.data);
    TobDemPrix  := TOB(TOBPiece.data);
    TOBArtDemPx := TOB(TOBDemPrix.data);
    TobFrsDemPx := TOB(TOBArtDemPx.data);
    TobDetDemPx := TOB(TobFrsDemPx.data);
    if TOBParam.GetString('FROMSAISIE')='X' then FfromSaisie := True;
  end;

  DefinitZone;

  if TOBArtDemPx.detail.count = 0 then
  begin
    BtDetailDdePrix.visible := False;
    BtSaisieDde.Visible     := False;
    BtExport.Visible        := False;
    BtDelete.Visible        := False;
  end;

  if assigned(TobParam) then
  begin
    TobParam.AddChampSupValeur('NATUREPIECEG', TobPiece.GetString('GP_NATUREPIECEG'));
    TobParam.AddChampSupValeur('SOUCHE', TobPiece.GetString('GP_SOUCHE'));
    TobParam.AddChampSupValeur('NUMERO', TobPiece.GetInteger('GP_NUMERO'));
    TobParam.AddChampSupValeur('INDICEG', TobPiece.GetInteger('GP_INDICEG'));
  end;

  //dessin de la grille Article demande de prix
  DessineLaGrille(GrilleDdePx, LesColDemPrix);

end ;

procedure TOF_BTDEMPRIX_MUL.DefinitZone;
begin

  if assigned(Getcontrol('BInsert')) then BtCreate := TToolbarButton97(ecran.FindComponent('BInsert'));

  if Assigned(GetControl('BDelete')) then BtDelete := TToolbarButton97(ecran.findComponent('BDelete'));

  if Assigned(GetControl('GRILLE_DEMPRIX')) then
  begin
    GrilleDdePx := THGrid(ecran.FindComponent('GRILLE_DEMPRIX'));
    GrilleDdePx.OnDblClick := GrilleOnDblClick;
    GrilleDdePx.OnClick    := GrilleOnClick;
  end;

  if Assigned(GetControl('GRILLE_DETAILDDEPRIX')) then
  begin
    GrilleDetail  := THGrid(ecran.FindComponent('GRILLE_DETAILDDEPRIX'));
    GrilleDetail.OnDblclick := GrilleDetailOnDblClick; 
  end;

  if Assigned(GetControl('PANEL_DETAIL')) then
  begin
    PanelDetail := THPanel(ecran.FindComponent('PANEL_DETAIL'));
  end;

  if assigned(GetControl('BEXPORTS')) then
  begin
    BtExport := TToolbarButton97(Ecran.Findcomponent('BEXPORTS'));
    BtExport.OnClick := ExportOnClick;
  end;

  if assigned(GetControl('BTDETAILDDEPRIX')) then
  begin
    BtDetailDdePrix := TToolbarButton97(Ecran.Findcomponent('BTDETAILDDEPRIX'));
    BtDetailDdePrix.OnClick := DetailDdePrix;
  end;

  if assigned(GetControl('BTSAISIEDDE')) then
  begin
    BtSaisieDde := ttoolbarButton97(Ecran.Findcomponent('BTSAISIEDDE'));
    BtSaisieDde.OnClick := SaisieDdePrix;
  end;

end;
procedure TOF_BTDEMPRIX_MUL.ExportOnClick(Sender: Tobject);
var Arow          : integer;
    TOBDetailArt  : TOB;
    TOBA          : TOB;
    iInd          : Integer;
begin
  if TobDemPrix.detail.Count = 0 then Exit;
  ARow    := GrilleDdePx.Row-1;

  TOBDetailArt  := Tob.create('DESTINATION', Nil, -1);

  Unique := TobDemPrix.Detail[Arow].GetInteger('BPP_UNIQUE');
  LibDdePrix := TobDemPrix.Detail[Arow].GetString('BPP_LIBELLE');

  For iInd := 0 To TobArtDemPx.detail.count - 1 do
  begin
    if TobArtDemPx.Detail[iInd].GetInteger('BDP_UNIQUE') = Unique then
    begin
      TOBA := Tob.create('ARTDEMPRIX', TOBDetailArt, -1);
      TOBA.Dupliquer(TobArtDemPx.detail[iInd], False, True);
    end;
  end;

  //génération de la tob nécessaire à l'envoi dans excel
  ExporteDocumentToExcel(TobPiece, TOBDetailArt, TobFrsDemPx, Unique, LibDdePrix);
  TobDemPrix.putValue('MODIFIED','X');
  TobDemPrix.Detail[Arow].SetBoolean('MODIFIED',True);
  GrilleDdePx.Refresh();

end;

Procedure TOF_BTDEMPRIX_MUL.DetailDdePrix(Sender: Tobject);
begin

  if PanelDetail.Visible then
  begin
    PanelDetail.visible := False;
    exit;
  end;

  PanelDetail.Visible := True;

  DessineLaGrille(GrilleDetail, LesColDetail);

  AfficheDetailDemandeDePrix;

end;

Procedure TOF_BTDEMPRIX_MUL.SaisieDdePrix(Sender: Tobject);
begin

  TobFrsDemPx.Data  := TobDetDemPx;
  TobArtDemPx.data  := TobFrsDemPx;
  TobDemPrix.Data   := TobArtDemPx;
  TobPiece.data     := TobDemPrix;
  TOBParam.data     := TOBPiece;
  TheTOB            := TOBParam;

  AGLLanceFiche('BTP','BTVALIDEDDEPRIX','','','UNIQUE='+ IntToStr(TobDemPrix.Detail[GrilleDdePx.Row - 1].GetInteger('BPP_UNIQUE')));

end;

Procedure TOF_BTDEMPRIX_MUL.AfficheDetailDemandedePrix;
Var Arow  : Integer;
    TobAffDetail : Tob;
    TobA  : TOB;
    TobB  : TOB;
    iInd  : Integer;
begin

  Arow := GrilleDdePx.Row - 1;

  Unique := TobDemPrix.Detail[Arow].GetInteger('BPP_UNIQUE');
  LibDdePrix := TobDemPrix.Detail[Arow].GetString('BPP_LIBELLE');

  TobAffDetail := Tob.Create('LES ARTICLEDEMPRIX', nil, -1);

  for iInd := 0 to TOBArtDemPx.detail.count -1 do
  begin
		TOBA := TOBArtDemPx.detail[iInd];
     if TOBA.GetInteger('BDP_UNIQUE') = Unique then
    begin
      TOBB := TOB.Create ('ARTICLEDEMPRIX',TobAffDetail,-1);
      TOBB.Dupliquer(TOBA,False,true);
      TOBB.data := TOBA;
    end;
  end;

  if TobAffDetail <> Nil then
  begin
    TobAffDetail.Detail.Sort('GL_TYPEARTICLE,GL_FAMILLENIV1,GL_FAMILLENIV2,GL_FAMILLENIV3,BDP_ARTICLE');
    ChargementGrilleWithTob(TobAffDetail, GrilleDetail, LesColDetail);
    TFVierge (Ecran).HMTrad.ResizeGridColumns(GrilleDetail);
    GrilleDetail.Refresh;
  end;

  FreeandNil(TobAffDetail);

end;

Procedure TOF_BTDEMPRIX_MUL.GrilleOnClick(Sender: Tobject);
begin
  if PanelDetail.Visible then AfficheDetailDemandedePrix;
end;

procedure TOF_BTDEMPRIX_MUL.GrilleOnDblClick(Sender: Tobject);
var Arow      : Integer;
begin

  ARow    := GrilleDdePx.Row-1;

  unique := StrToInt(TobDemPrix.Detail[Arow].GetString('BPP_UNIQUE'));

  if (fTtypeTrait = TtmtModif ) then
  begin
    Cledoc.NaturePiece := TobDemPrix.Detail[Arow].GetString('BPP_NATUREPIECEG');

    Cledoc.Souche := TobDemPrix.Detail[Arow].GetString('BPP_SOUCHE');

    Cledoc.NumeroPiece := TobDemPrix.Detail[Arow].GetInteger('BPP_NUMERO');

    Cledoc.Indice := TobDemPrix.Detail[Arow].GetInteger('BPP_INDICEG');

    //CreateDemandePrix(TobDemPrix,Unique);
    CreateDemandePrix(TobParam,Unique);

    AfficheEcran;
  end else
  begin
    // selection et sortie
    Tfmul(ecran).Retour := IntToStr(TobDemPrix.Detail[Arow].GetInteger('BPP_UNIQUE'));
    ecran.Close;
  end;

end;

Procedure TOF_BTDEMPRIX_MUL.GrilleDetailOnDblClick(Sender: Tobject);
var ArtAff  : String;
begin

  if TOBArtDemPx.detail.count = 0 then exit;

  ArtAff := GrilleDetail.Cells[ColRefArt,GrilleDetail.Row];

  AGLLanceFiche('BTP', 'BTARTICLE', '', ArtAff, '');

end;

procedure TOF_BTDEMPRIX_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMPRIX_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMPRIX_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMPRIX_MUL.DessineLaGrille(Grille : THGrid; LibColonne : String);
var st      : string;
    Nam     : String;
    i       : Integer;
    depart  : Integer;
begin
  //
  Grille.FixedCols := 1;
  Grille.FixedRows := 1;
  //
  Grille.DefaultRowHeight := 18;
  //
  Grille.ColCount     := findOccurenceString(LibColonne, ';') + 1;
  Grille.ColWidths[0] := 10;
  //
  St := LibColonne;
  //
  ColUnique     := -1;
  ColLibelle    := -1;
  ColTraite     := -1;
  ColEnvoye     := -1;
  ColRefArt     := -1;
  ColPrixAchat  := -1;
  //
  for i := 0 to Grille.ColCount - 1 do
  begin
    depart := 1;
    if i > depart then Grille.ColWidths[i] := 100;
    Nam := ReadTokenSt(St);
    if Nam = 'BPP_UNIQUE' then
    begin
      Grille.ColWidths[i] := 50;
      Grille.ColAligns[i] := taCenter;
      Grille.ColEditables[i] := False;
      ColUnique := i;
    end
    else if  (Pos(Nam ,'BPP_LIBELLE;BDP_LIBELLE')>0) then
    begin
      Grille.ColWidths[i] := 300;
      Grille.ColEditables[i] := False;
      ColLibelle := i;
    end
    else if Nam = 'BPP_TRAITE' then
    begin
      Grille.ColWidths[i] := 60;
      Grille.ColAligns[i] := taCenter;
      Grille.ColEditables[i] := False;
      ColTraite := i;
    end
    else if Nam = 'BPP_ENVOYE' then
    begin
      Grille.ColWidths[i] := 60;
      Grille.ColAligns[i] := taCenter;
      Grille.ColEditables[i] := False;
      ColEnvoye := i;
    end
    Else if Nam = 'BDP_ARTICLE' then
    begin
      Grille.ColWidths[i] := 100;
      Grille.ColEditables[i] := False;
      ColRefArt := i;
    end
    else if Nam = 'BDP_DPA' then
    begin
      Grille.ColWidths[i] := 60;
      Grille.ColFormats[i] := FF;
      Grille.ColAligns[i] := taRightJustify;
      Grille.ColEditables[i] := False;
      ColPrixAchat := i;
    end;
  end;
  //
  if ColUnique     >=0 then Grille.Cells[ColUnique,0]    := 'N° Dde';
  if ColLibelle    >=0 then Grille.Cells[ColLibelle,0]   := 'Libellé';
  if ColTraite     >=0 then Grille.Cells[ColTraite,0]    := 'Traité';
  if ColEnvoye     >=0 then Grille.Cells[ColEnvoye,0]    := 'Envoyé';
  if ColRefArt     >=0 then Grille.Cells[ColRefArt,0]    := 'Ref. Article';
  if ColPrixAchat  >=0 then Grille.Cells[ColPrixAchat,0] := 'Px Achat';
  //

end ;


Initialization
  registerclasses ( [ TOF_BTDEMPRIX_MUL ] ) ; 
end.

