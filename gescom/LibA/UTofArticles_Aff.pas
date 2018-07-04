unit UTofArticles_Aff;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,AGLInit,SaisUtil,paramsoc,//windows,messages,
{$IFDEF EAGLCLIENT}
      MaineAGL,emul,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, FE_Main,mul,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOF, FactUtil,
      AffaireUtil,UTofAfBaseCodeAffaire,
      utob,grids, DicoBTP, HSysMenu,HTB97,ConfidentAffaire;
Type
     TOF_ARTICLES_AFF = Class (TOF_AFBASECODEAFFAIRE)
      procedure OnArgument (stArgument : String ); override;

      procedure BoutonClick (Sender : TObject);
      procedure GridDblClick (Sender : TObject);
      procedure FormClose (Sender : TObject; var Action : TCloseAction);
      procedure GereValidation;
      procedure FormateZoneSaisie (ACol, ARow : Longint );
      procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
     public
         TobArticles, TobArtSelect : TOB;
         CodeArticle : string;
         TypeArticle : string;
         MonGrid : THGrid;
         bouton : TToolBarButton97;

     END ;

Procedure AFLanceFiche_ArticleParAff (Argument : string);

implementation

procedure TOF_ARTICLES_AFF.OnArgument (stArgument : String);
Var
   Tmp,St,affaire, LibAffaire, sSansPrix : String;
   wi, i, Row : Integer;
   HTrad:THSystemMenu;
//   Aff1, Aff2, Aff3,AfAve : THEdit;
   LibAff : THLabel;
begin
//  Inherited;

  tmp := ''; St := stArgument; {LibAff := nil;}
//  Aff1 := nil; Aff2 := nil; Aff3 := nil; AfAve := nil;
  wi := 0;

  While St <> '' do
    BEGIN
      Tmp := ReadTokenSt (St);
      If (wi = 0) Then
        Begin
          affaire := Tmp;
          SetControlText('AFF_AFFAIRE', affaire);

(*          Aff1 := THEdit(GetControl ('AFF_AFFAIRE1'));
          Aff2 := THEdit(GetControl ('AFF_AFFAIRE2'));
          Aff3 := THEdit(GetControl ('AFF_AFFAIRE3'));
          Afave := THEdit(GetControl ('AFF_AVENANT'));
          ChargeCleAffaire (Nil, THCritMaskEdit(Aff1), THCritMaskEdit(Aff2), THCritMaskEdit(Aff3),
                          THCritMaskEdit(AfAve), Nil, taConsult, Tmp, False);

          LibAff := THLabel(GetControl ('LAFF_AFFAIRE'));
          LibAffaire := RechDom ('AFLAFFAIRE', Tmp, False);
          if (LibAffaire <> 'Error') then
            LibAff.Caption := RechDom ('AFLAFFAIRE', Tmp, False);
          //      SetControlText('LAFF_AFFAIRE',RechDom('AFLAFFAIRE',Tmp,False));
          *)
        end;

      If (wi = 1) Then
        Begin
          CodeArticle := Tmp;
        end;

      If (wi = 2) Then
        Begin
          TypeArticle := Tmp;
        end;

      If (wi = 3) Then
        Begin
          sSansPrix := Tmp;
        end;

      wi := wi+1;
    End;


// Init du code affaire dans la tof ancêtre
inherited;

  // libellé de l'affaire recadré
  LibAff := THLabel(GetControl ('LAFF_AFFAIRE'));
  LibAffaire := RechDom ('AFLAFFAIRE', affaire, False);
  if (LibAffaire <> 'Error') then
    LibAff.Caption := LibAffaire;
  // Protection des champ affaire
  EditAff.ReadOnly := true;
  EditAff1.ReadOnly := true;
  EditAff2.ReadOnly := true;
  EditAff3.ReadOnly := true;
  EditAff4.ReadOnly := true;



  if (EditAff1 <> nil) then
    begin
      if (EditAff2.visible = false) then
        begin
          LibAff.Left := EditAff.Left;
          if (EditAff4.visible = True) then libaff.left := EditAff4.left + EditAff4.width + 5;
        end
      else
      if (EditAff3.visible = false) then
        begin
          LibAff.Left := EditAff3.Left;
          EditAff4.left := EditAff3.Left;
          if (EditAff4.visible = True) then libaff.left := EditAff4.left + EditAff4.width + 5;
        end
      else
      if (EditAff4.visible = False) then
        LibAff.left := EditAff3.left + EditAff3.width + 5
      else
        libaff.left := EditAff4.left + EditAff4.width + 5;
    end;

  MonGrid := THGRID(GetControl ('GRID_LISTEARTICLESAFF'));
  Bouton := TToolBarButton97(GetControl ('BValider'));

  // PL le 09/10/02 : on doit cacher le prix si on ne valorise pas par affaire
  if (GetParamSoc ('SO_AFVALOACTPR') <> 'AFF') and (GetParamSoc ('SO_AFVALOACTPV') <> 'AFF') then
    sSansPrix := 'S';
  //////////////////////////

  if (LaTOB = nil) or (LaTOB.Detail.Count = 0) or (Not LaTOB.TOBNameExist ('ARTICLE')) then Ecran.Close
  else
    begin
      TobArticles := LaTOB.Detail[0];
      TobArticles.Detail[0].AddChampSup ('Curseur', TRUE);
      if (sSansPrix = 'S') then
        // Si on ne doit pas accéder à la valorisation on n'affiche pas les prix
        begin
          TobArticles.PutGridDetail (MonGrid, false, false, 'Curseur;GA_CODEARTICLE;GA_LIBELLE;GA_QUALIFUNITEVTE');
          MonGrid.ColCount := 4;
          MonGrid.ColWidths[0] := 10;
          MonGrid.ColWidths[2] := 250;
          MonGrid.ColWidths[3] := 20;
        end
      else
        begin
          TobArticles.PutGridDetail (MonGrid, false, false, 'Curseur;GA_CODEARTICLE;GA_LIBELLE;GA_QUALIFUNITEVTE;GA_PMRP;GA_PVHT');
          MonGrid.ColCount := 6;
          MonGrid.ColWidths[0] := 10;
          MonGrid.ColWidths[2] := 180;
          MonGrid.ColWidths[3] := 20;
          MonGrid.ColAligns[4] := taRightJustify;
          MonGrid.ColAligns[5] := taRightJustify;
        end;
    end;


  if (TobArticles <> nil) then
    for Row := 1 to TobArticles.Detail.Count do
      for i := 1 to MonGrid.ColCount - 1 do FormateZoneSaisie (i, Row);

  HTrad := THSystemMenu.Create (MonGrid.Parent);
  HTrad.ResizeGridColumns (MonGrid);

  Ecran.OnClose := FormClose;
  Bouton.OnClick := BoutonClick;
  MonGrid.OnDblClick := GridDblClick;
End;

procedure TOF_ARTICLES_AFF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff0:=nil;
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
end;

procedure TOF_ARTICLES_AFF.BoutonClick(Sender: TObject);
begin
GereValidation;
end;

procedure TOF_ARTICLES_AFF.GridDblClick(Sender: TObject);
begin
GereValidation;
Ecran.Close;
end;

procedure TOF_ARTICLES_AFF.GereValidation;
begin
if (TobArticles=nil) then exit;
if (TobArticles.Detail.Count=0) then exit;
if (MonGrid.Row > TobArticles.Detail.Count) then exit;

if (TheTOB<>nil) then begin TheTOB.Free; TheTOB:=nil; end;
TheTOB:=TOB.Create('la TOB',nil,-1);
TobArtSelect:= TOB.Create('ARTICLE', TheTOB, -1);

TobArtSelect.Dupliquer(TobArticles.Detail[MonGrid.Row-1], true, true, false);
TobArtSelect.DelChampSup('Curseur',false) ;
end;

procedure TOF_ARTICLES_AFF.FormClose(Sender: TObject; var Action: TCloseAction);
begin
TobArticles.Free;

Action:=caFree ;
end;

procedure TOF_ARTICLES_AFF.FormateZoneSaisie (ACol,ARow : Longint ) ;
Var St,StC : String ;
BEGIN
St:=MonGrid.Cells[ACol,ARow];

StC:=St ;

if (ACol=3) then
    StC:=AnsiUppercase(Trim(St));

if (ACol=4) or (ACol=5) then
    StC:=StrF00(Valeur(St),V_PGI.OkDecP);

MonGrid.Cells[ACol,ARow]:=StC;
END ;
Procedure AFLanceFiche_ArticleParAff(Argument:string);
begin
AGLLanceFiche('AFF','AFARTICLESAFFAIRE','','',argument);
end;


Initialization
registerclasses([TOF_ARTICLES_AFF]);
end.
