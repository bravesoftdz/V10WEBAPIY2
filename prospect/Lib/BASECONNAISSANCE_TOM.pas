{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 12/12/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : BASECONNAISSANCE (BASECONNAISSANCE)
Mots clefs ... : TOM;BASECONNAISSANCE
requêtes utiles : select * from lienbconnaissance order by lbc_numero,lbc_numratt
*****************************************************************}
Unit BASECONNAISSANCE_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db, DbCtrls,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,
     Fe_Main,
{$else}
     eFiche,UtileAGL,
     eFichList,
     MaineAgl,
{$ENDIF}
     forms, HTB97,
     sysutils,
     ComCtrls,
     HCtrls,Hdb,
     HEnt1,
     HMsgBox,
     UTOM,wTom,
     UTob,shellapi // pour Shellexecute
     ,windows,lookup,graphics,Grids,menus
      ;

Type
  TOM_BASECONNAISSANCE = Class (tWTOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnLoadAlerte; override;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  Private
    oldFichier : String;
    ValeurOri : Array [1..3] of String;
    MajGrid : Array [1..3] of Boolean;
    GS : THGRID;
    AfterLoad : boolean;
    bDuplication : boolean;
    procedure RabTheGoodField( NumChamp : Integer);
    procedure Val1_NewClick ( Sender : tObject);
    procedure Val2_NewClick ( Sender : tObject);
    procedure Val3_NewClick ( Sender : tObject);
    procedure Rattachement1_OnChangeClick ( Sender : tObject);
    procedure Rattachement2_OnChangeClick ( Sender : tObject);
    procedure Rattachement3_OnChangeClick ( Sender : tObject);
    procedure BCO_EMETTEUR_OnElipsisClick( Sender : tObject);
    Procedure BCO_VALIDEE_OnClick( Sender : tObject);
    Procedure BCO_PERIMEE_OnClick( Sender : tObject);
    Procedure DuplicationClick( Sender : tObject);
    Procedure LienDocOnElipsisClick( Sender : tObject);
    Procedure LienDocOnExit( Sender : tObject);
    Procedure OuvreLienDoc( Sender : tObject);
    procedure AfficheChampsLibres ;
    procedure my_YFileStd_Delete (CodeProduit, Nom, Langue, Predefini, Crit1 : String);
    Procedure my_YFileStd_update (CodeProduit, Nom, Langue, Predefini : String);
    Procedure DoImportMyFic(FileName : string);
    Procedure MajLesRattachements(iNum : integer; var numRatt : integer) ;
    Procedure RecupLesRattachements ;
    procedure GSLigneDClick (Sender: TObject);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
    procedure GSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
    procedure AddElemClick(iNUm : integer);
    procedure GSExit(Sender: TObject);
    procedure DessineCellGS ( ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
    procedure AfficheLeGrid( i : integer) ;
    procedure DessineCell (GG : Thgrid; ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DoVideLesGrids ;
    Procedure MajLesLibelles(NumGrid,numDeb,numFin : integer) ;
    procedure ListAlerte_OnClick_BCO(Sender: TObject);
    procedure Alerte_OnClick_BCO(Sender: TObject);
    end ;

Implementation
uses
 UtilBaseConnaissance,UtilConfid,MessagesErreur,uYFileStd,EntRT
{$IFDEF SAV}
  ,EntGC
{$ENDIF SAV}
  ,UtilAlertes,YAlertesConst,EntPgi
 ;

procedure TOM_BASECONNAISSANCE.OnNewRecord ;
begin
  Inherited ;
  SetField('BCO_VERSIONBC',1);
  SetControlText('LIENDOC','');
  if VH_RT.RTResponsable <> '' then
    setfield ('BCO_EMETTEUR',VH_RT.RTResponsable);
  if not bDuplication then
    DoVideLesGrids;  
end ;

procedure TOM_BASECONNAISSANCE.OnDeleteRecord ;
begin
  Inherited ;
 { GC/GRC : MNG / gestion des alertes }
  if (not V_Pgi.SilentMode) and (not AfterInserting)
    and (AlerteActive (TableToPrefixe(TableName))) then
    if (not ExecuteAlerteDelete(TForm(ecran),true)) then
      begin
      LastError := 99;
      exit;
      end;
  if (GetControlText ('LIENDOC') <> '' )  then
    my_YFileStd_Delete (CODEPRODUIT, GetControlText ('LIENDOC'), 'FRA', 'NEW', GetString ('BCO_NUMERO'));
  ExecuteSql ('DELETE FROM '+TableLien+' WHERE LBC_NUMERO = '+GetString ('BCO_NUMERO'));
end ;

procedure TOM_BASECONNAISSANCE.OnUpdateRecord ;
var nbco : integer;
    NomChamp,Select : String;
    Q : TQuery ;
begin
  Inherited ;
  NomChamp:='';
  NomChamp:=VerifierChampsObligatoires(Ecran,'');
  if NomChamp<>'' then
    begin
    NomChamp:=ReadTokenSt(NomChamp);
    SetFocusControl(NomChamp) ;
    LastError:=1; LastErrorMsg:=MSG_BaseConnaissance[LastError]+champToLibelle(NomChamp);
    exit;
    end;
  { GC/GRC : MNG / gestion des alertes }
  if (ds<>nil) and (not V_Pgi.SilentMode)
    and (AlerteActive (TableToPrefixe(TableName))) then
      if (not ExecuteAlerteUpdate(TForm(Ecran),true)) then
        LastError := 99;
  { GC/GRC : MNG fin / gestion des alertes }
  if getfield ('BCO_NUMERO') = 0 then
  begin
    Select := 'SELECT MAX(BCO_NUMERO) FROM BASECONNAISSANCE';
    Q:=OpenSQL(Select, True);
    try
      if not Q.Eof then
         begin
         nbco := Q.Fields[0].AsInteger;
         setfield ('BCO_NUMERO',nbco+1);
         end
         else
         setfield ('BCO_NUMERO',1);
    finally
      Ferme(Q) ;
    end;
  end;
end ;

procedure TOM_BASECONNAISSANCE.OnAfterUpdateRecord ;
var i : integer;
    numRatt : integer;
begin
  Inherited ;
  { maj table lien }
  numRatt:=0;
  for i :=1 to nb_rattachement Do
    MajLesRattachements(i,numRatt);

  if oldFichier<>GetControlText ('LIENDOC') then
    DoImportMyFic(GetControlText ('LIENDOC'))
  else
    if (GetControlText ('LIENDOC') <> '' ) and (action = 'CREATION' ) then
      my_YFileStd_update (CODEPRODUIT, GetControlText ('LIENDOC'), 'FRA', 'NEW');
end ;

procedure TOM_BASECONNAISSANCE.OnAfterDeleteRecord ;
begin
  Inherited ;
  
end ;

procedure TOM_BASECONNAISSANCE.OnLoadRecord ;
var sSQL : string;
    Q : TQuery;
    i : integer;
begin
  Inherited ;
  AppliquerConfidentialite(Ecran,'');
  { vidage des grids + accès table lien }
  RecupLesRattachements;
  sSQL := 'SELECT YFS_NOM FROM YFILESTD WHERE '
    + 'YFS_CODEPRODUIT =  "' + CODEPRODUIT + '" AND '
    + 'YFS_CRIT1 =  "' + GetString ('BCO_NUMERO') + '"';
  Q := OpenSQL (sSQL, true) ; // ouverture de la requete
  if not Q.Eof then // si un enregistrement a ete trouve
    SetControlText('LIENDOC',Q.findfield('YFS_NOM').AsString);
  Ferme (Q) ;
  oldFichier:=GetControlText ('LIENDOC');
  if oldFichier = '' then SetControlEnabled  ('OUVRELIENDOC',false)
  else SetControlEnabled  ('OUVRELIENDOC',true);
  SetControlEnabled('BDUPLICATION',action <>'CREATION');
  for i:= 1 to nb_rattachement Do
    AfficheLeGrid(i);
  AfterLoad:=true;
end ;

procedure TOM_BASECONNAISSANCE.DoVideLesGrids ;
var i : integer;
begin
  for i:=1 to nb_rattachement Do
    begin
    GS:=THGRID(GetControl('GRATTACHEMENT'+IntToStr(i)));
    if Assigned(GS) then
      GS.videpile (False);
    SetControlText('RATTACHEMENT'+IntToStr(i),'');
    end;
end;

procedure TOM_BASECONNAISSANCE.AfficheLeGrid( i : integer) ;
begin
  //LesColonnesGrid :=Colonnes_BaseConnaissance[i];
  GS:=THGRID(GetControl('GRATTACHEMENT'+IntToStr(i)));
  if not Assigned(GS) then exit;
  GS.OnRowEnter:=GSRowEnter ;
  GS.OnRowExit:=GSRowExit ;
  GS.OnCellExit:=GSCellExit ;
  GS.OnCellEnter:=GSCellEnter ;
  GS.OnExit:=GSExit ;
  GS.OnDblClick:=GSLigneDClick ;
  GS.PostDrawCell:= DessineCellGS;
  GS.ColWidths[0]:=10;
  GS.ColWidths[1]:=80; GS.ColAligns[1]:=taLeftJustify;
  GS.ColWidths[2]:=160; GS.ColAligns[2]:=taLeftJustify;
  TFFiche(Ecran).Hmtrad.ResizeGridColumns(GS) ;
  TFFiche(Ecran).OnKeyDown:=FormKeyDown ;
  MajGrid[i]:=false;
end;

procedure TOM_BASECONNAISSANCE.GSLigneDClick (Sender: TObject);
var chps,prefixe,nomFiche,domaine,stArg,stLequel,stRange : string;
    i : integer;
    Q : TQuery;
begin
  if (copy(Screen.ActiveControl.Name,1,13) <> 'GRATTACHEMENT' ) then exit;

  GS:=THGRID(GetControl(Screen.ActiveControl.Name));
  if not Assigned(GS) then exit;
  if not IsNumeric(copy(Screen.ActiveControl.Name,14,1)) then exit;
  prefixe:=GetControlText('RATTACHEMENT'+copy(Screen.ActiveControl.Name,14,1));

  if prefixe = '' then exit;
  for i:=1 to max_rattachement Do
    if copy(Colonnes_BC[i],1,length(prefixe)) = prefixe then
      chps:=Colonnes_BC[i];
  ReadToKenSt(chps);
  ReadToKenSt(chps);
  domaine:=ReadToKenSt(chps);
  nomFiche:=ReadToKenSt(chps);
  if ctxaffaire in v_pgi.PgiCOntexte then
  begin //mcd 30/08/07 GIGA14457 pour bonne fiche article
    if Nomfiche='WARTICLE_FIC' then
    begin
      domaine:='AFF';
      NomFiche:='AFARTICLE';
    end;
  end;
  stRange:='';
  stLequel:='';
  { cas du tiers où il faut la nature }
  if prefixe = 'T' then
    begin
    Q:=OpenSql('SELECT T_NATUREAUXI,T_AUXILIAIRE FROM TIERS WHERE T_TIERS = "'+GS.Cells[1,GS.Row]+'"',true);
    if not Q.EOF then
      begin
      stArg:='T_NATUREAUXI='+Q.FindField('T_NATUREAUXI').AsString;
      stLequel:=Q.FindField('T_AUXILIAIRE').AsString;
      end;
    Ferme(Q);
    end
  else
    { type d'action il faut GRC et --- }
    if prefixe = 'RPA' then
      begin
      stArg:='GRC;ACTION=CONSULTATION';
      stRange:='---;GRC;'+GS.Cells[1,GS.Row];
      end
    else
      begin
      stArg:=''; stLequel:=GS.Cells[1,GS.Row];
      end;
  AGLLanceFiche(domaine,nomFiche,stRange,stLequel,stArg);
end;

procedure TOM_BASECONNAISSANCE.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GS.InvalidateRow(ou) ;
end;

procedure TOM_BASECONNAISSANCE.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  GS.InvalidateRow(ou) ;
end;

procedure TOM_BASECONNAISSANCE.GSCellEnter(Sender: TObject; var ACol,
                                        ARow: Integer; var Cancel: Boolean);
begin
end;

procedure TOM_BASECONNAISSANCE.GSCellExit(Sender: TObject; var ACol,
                                        ARow: Integer; var Cancel: Boolean);
begin
end;

procedure TOM_BASECONNAISSANCE.DessineCellGS ( ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
begin
DessineCell( GS,ACol,ARow,Canvas,AState);
end;

procedure TOM_BASECONNAISSANCE.DessineCell (GG : Thgrid; ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
  Arect: Trect ;
begin
If Arow < GG.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=GG.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := GG.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = GG.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=ARect.Right-2 ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
end;

procedure TOM_BASECONNAISSANCE.GSExit(Sender: TObject);
var cancel: boolean;
  ACOl, ARow: integer;
begin
  ACol := GS.Col;
  ARow := GS.Row;
  Cancel := False;
  GSCellExit(Sender, ACol, ARow, Cancel);
end;


procedure TOM_BASECONNAISSANCE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_BASECONNAISSANCE.OnArgument ( S: String ) ;
var i : integer;
begin
  Inherited ;
  AfterLoad:=false;
  bDuplication  := false;
  if Assigned(GetControl('RATTACHEMENT1')) then
    THVALCOMBOBOX(GetControl('RATTACHEMENT1')).OnChange := Rattachement1_OnChangeClick;
  if Assigned(GetControl('RATTACHEMENT2')) then
    THVALCOMBOBOX(GetControl('RATTACHEMENT2')).OnChange := Rattachement2_OnChangeClick;
  if Assigned(GetControl('RATTACHEMENT3')) then
    THVALCOMBOBOX(GetControl('RATTACHEMENT3')).OnChange := Rattachement3_OnChangeClick;

  if Assigned(GetControl('BNOUVEAU1')) then
    TToolBarButton97(GetControl('BNOUVEAU1')).OnClick := Val1_NewClick;
  if Assigned(GetControl('BNOUVEAU2')) then
    TToolBarButton97(GetControl('BNOUVEAU2')).OnClick := Val2_NewClick;
  if Assigned(GetControl('BNOUVEAU3')) then
    TToolBarButton97(GetControl('BNOUVEAU3')).OnClick := Val3_NewClick;

  if Assigned(GetControl('BCO_EMETTEUR')) then
    THDBEdit(GetControl('BCO_EMETTEUR')).OnElipsisClick := BCO_EMETTEUR_OnElipsisClick;

  if Assigned(GetControl('BCO_VALIDEE')) then
    TDBCheckBox(GetControl('BCO_VALIDEE')).OnClick := BCO_VALIDEE_OnClick;
  if Assigned(GetControl('BCO_PERIMEE')) then
    TDBCheckBox(GetControl('BCO_PERIMEE')).OnClick := BCO_PERIMEE_OnClick;
  { duplication }
  if Assigned(GetControl('BDUPLICATION')) then
    TToolBarButton97(GetControl('BDUPLICATION')).OnClick := DuplicationClick;
  { document lié : LienDocOnElipsisClick }
  if Assigned(GetControl('LIENDOC')) then
    THEdit(GetControl('LIENDOC')).OnElipsisClick := LienDocOnElipsisClick;
  if Assigned(GetControl('LIENDOC')) then
    THEdit(GetControl('LIENDOC')).OnExit := LienDocOnExit;

  { document lié : LienDocClick }
  if Assigned(GetControl('OUVRELIENDOC')) then
    TToolBarButton97(GetControl('OUVRELIENDOC')).OnClick := OuvreLienDoc;

  AfficheChampsLibres;
{$IFDEF SAV}
  { si SAV pas sérialisé, on enlève la table du combo }
  if not VH_GC.SAVSeria then
{$ENDIF SAV}
    begin
    for i := 1 to nb_rattachement Do
      if Assigned(GetControl('RATTACHEMENT'+IntToStr(i))) then
        THVALCOMBOBOX(GetControl('RATTACHEMENT'+IntToStr(i))).plus := ' AND CO_CODE <> "'+PrefixeArtParc+'"';
    end;
{$IFNDEF EAGLSERVER}
  if Assigned(GetControl('MnAlerte')) then
    if AlerteActive('BCO') then
      TMenuItem(GetControl('MnAlerte')).OnClick := Alerte_OnClick_BCO
    else
      TMenuItem(GetControl('MnAlerte')).visible:=false;

  if Assigned(GetControl('MnListAlerte')) then
    if AlerteActive('BCO') then
      TMenuItem(GetControl('MnListAlerte')).OnClick := ListAlerte_OnClick_BCO
    else
      TMenuItem(GetControl('MnListAlerte')).visible:=false;

  if Assigned(GetControl('MnAlerte')) and Assigned(GetControl('MnListAlerte')) and
      Assigned(GetControl('BZOOM')) then
         TToolBarButton97(GetControl('BZOOM')).visible := (TMenuItem(GetControl('MnAlerte')).visible)
          and (TMenuItem(GetControl('MnListAlerte')).visible);
{$ENDIF EAGLSERVER}
end ;

procedure TOM_BASECONNAISSANCE.AfficheChampsLibres ;
var i : integer;
begin
  for i := 1 to 3 Do
    if copy(RechDom('RBBOOLLIBREBC','BB'+IntToStr(i),FALSE),1,1) = '.' then
      SetControlVisible ('BCO_BOOLLIBRE'+IntToStr(i),false)
    else
      SetControlCaption ('BCO_BOOLLIBRE'+IntToStr(i),RechDom('RBBOOLLIBREBC','BB'+IntToStr(i),FALSE));
  SetControlCaption ('TBCO_CHARLIBRE1',RechDom('RBMULTILIBREBC','BM1',FALSE));
end ;

procedure TOM_BASECONNAISSANCE.OnClose ;
begin
  Inherited ;
  if (GetControlText ('BCO_NUMERO') = '0' ) and (GetControlText ('LIENDOC') <> '' )  then
    my_YFileStd_Delete (CODEPRODUIT, GetControlText ('LIENDOC'), 'FRA', 'NEW', GetString ('BCO_NUMERO'));
end ;

procedure TOM_BASECONNAISSANCE.OnCancelRecord ;
begin
  Inherited ;
  DoVideLesGrids;
  RecupLesRattachements;
end ;

procedure TOM_BASECONNAISSANCE.AddElemClick(iNum : integer);
var retour,stValeur : string;
    bBreak : boolean;
    i,numDeb : integer;
begin
  retour:= LanceTheGoodMul(THVALCOMBOBOX(GetControl('RATTACHEMENT'+IntToStr(iNum))),nil,nil,nil,1);
  if retour = '' then exit;
  GS:=THGRID(GetControl('GRATTACHEMENT'+IntToStr(iNum)));
  if not assigned(GS) then exit;
  numDeb:=GS.RowCount-1;
  repeat
    stValeur:=ReadToKenSt(retour);
    if stValeur <> '' then
      begin
      bBreak:=false;
      { si existe déjà dans le grid on zappe }
      for i:=1 to Pred(GS.RowCount) Do
        if GS.Cells[1,i] = stValeur then bBreak:=true;
      if bBreak then continue;
      GS.RowCount:=GS.RowCount+1;
      GS.Cells[1,GS.RowCount-1]:=stValeur;
      end;
  until stValeur ='';
  { cas du Grid vide au départ, la ligne 1 est vide on la detruit }
  if (GS.Cells[1,1]='') then
    GS.DeleteRow (1);

  ForceUpdate;
  MajGrid[iNum]:=true;
  MajLesLibelles(iNum,numDeb,GS.RowCount-1)
end;

procedure TOM_BASECONNAISSANCE.Val1_NewClick( Sender : tObject);
begin
  AddElemClick(1);
end ;

procedure TOM_BASECONNAISSANCE.Val2_NewClick( Sender : tObject);
begin
  AddElemClick(2);
end ;

procedure TOM_BASECONNAISSANCE.Val3_NewClick( Sender : tObject);
begin
  AddElemClick(3);
end ;

procedure TOM_BASECONNAISSANCE.Rattachement1_OnChangeClick ( Sender : tObject);
begin
  RabTheGoodField(1);
end;

procedure TOM_BASECONNAISSANCE.Rattachement2_OnChangeClick ( Sender : tObject);
begin
  RabTheGoodField(2);
end;

procedure TOM_BASECONNAISSANCE.Rattachement3_OnChangeClick ( Sender : tObject);
begin
  RabTheGoodField(3);
end;


procedure TOM_BASECONNAISSANCE.RabTheGoodField( NumChamp : Integer);
var i : integer;
begin
  if not AfterLoad then exit;
  { contrôle que le rattachement n'existe pas déjà }
  if GetControlText('RATTACHEMENT'+IntToStr(NumChamp)) <> '' then
    for i:=1 to nb_rattachement Do
      if ( i <> NumChamp ) and ( GetControlText('RATTACHEMENT'+IntToStr(NumChamp)) = GetControlText('RATTACHEMENT'+IntToStr(i)) ) then
        PgiError ('Ce rattachement existe déjà',Ecran.caption);

  GS:=THGRID(GetControl('GRATTACHEMENT'+IntToStr(NumChamp)));
  if Assigned(GS) then
    GS.videpile (False);
  { une remise à blanc de 1 ou 2 entraine une maj de tous à cause du n° rattachement }
  if NumChamp = 3 then
    MajGrid[3]:=true
  else
    for i:= 1 to nb_rattachement Do
      MajGrid[i]:=true;
  ForceUpdate;
end ;


procedure TOM_BASECONNAISSANCE.BCO_EMETTEUR_OnElipsisClick( Sender : tObject);
var Emetteur : String;
begin
Emetteur:=AGLLanceFiche('AFF', 'RESSOURCERECH_MUL','ARS_RESSOURCE='+GetControlText('BCO_EMETTEUR'), '', '');
if Emetteur <> '' then
  begin
    SetControlText('BCO_EMETTEUR',Emetteur);
    SetField('BCO_EMETTEUR',Emetteur)
  end;
end;

Procedure TOM_BASECONNAISSANCE.BCO_VALIDEE_OnClick( Sender : tObject);
begin
  if GetCheckBoxState('BCO_VALIDEE') = cbChecked then
    begin
    SetString('BCO_UTILVALID',v_pgi.user);
    SetDateTime('BCO_DATEVALIDATION',Date);
    end
  else
    begin
    SetString('BCO_UTILVALID','');
    SetDateTime('BCO_DATEVALIDATION',iDate1900);
    end;
end;

Procedure TOM_BASECONNAISSANCE.BCO_PERIMEE_OnClick( Sender : tObject);
begin
  if GetCheckBoxState('BCO_PERIMEE') = cbChecked then
    begin
    SetString('BCO_UTILPEREMPTION',v_pgi.user);
    SetDateTime('BCO_DATEPEREMPTION',Date);
    end
  else
    begin
    SetString('BCO_UTILPEREMPTION','');
    SetDateTime('BCO_DATEPEREMPTION',iDate1900);
    end;
end;

Procedure TOM_BASECONNAISSANCE.DuplicationClick( Sender : tObject);
var tobBase : tob;
    i : integer;
begin
  if (Ecran = Nil) or ( not(Ecran is TFFiche) ) then exit;
  tobBase:=TOB.Create ('BASECONNAISSANCE', Nil, -1);
  tobBase.GetEcran (TFfiche(Ecran),Nil);
  bDuplication  := true;
  if not TFFiche (Ecran).Bouge (TNavigateBtn(NbInsert)) then
    begin
    FreeAndNil(tobBase);
    exit;
    end;
  tobBase.SetInteger('BCO_NUMERO',0);
  tobBase.SetInteger('BCO_VERSIONBC',1);
  tobBase.PutEcran(TFfiche(Ecran));
  SetControlEnabled('BDUPLICATION',FALSE);
  ForceUpdate;
  FreeAndNil(tobBase);
  SetFocusControl('BCO_SUJET');
  For i:=1 to nb_rattachement Do
    if GetControlText ('RATTACHEMENT'+IntToStr(i)) <> '' then
      MajGrid[i]:=true;
end;

Procedure TOM_BASECONNAISSANCE.LienDocOnElipsisClick( Sender : tObject);
var FileName : String;
begin
  //if GetInteger ('BCO_NUMERO') = 0 then exit;
  if not GetFileName(tfOpen, '*.*', FileName) then exit;
  if FileName = '' then exit;
  DoImportMyFic(FileName);
end;

Procedure TOM_BASECONNAISSANCE.LienDocOnExit( Sender : tObject);
begin
  //if GetInteger ('BCO_NUMERO') = 0 then exit;
  if oldFichier=GetControlText ('LIENDOC') then exit;
  DoImportMyFic(GetControlText ('LIENDOC'));
end;

Procedure TOM_BASECONNAISSANCE.DoImportMyFic(FileName : string);
var Num : integer;
begin
  SetControlText ('LIENDOC',ExtractFileName(FileName));
  if Filename <>'' then
    begin
    Num:=AGL_YFILESTD_IMPORT(FileName, CODEPRODUIT,  ExtractFileName(FileName),
        '', GetString ('BCO_NUMERO'),
        '', '', '', '', '-', '-', '-', '-', '-', 'FRA', 'NEW', GetField('BCO_LIBELLE'));
    if num <> -1 then
      begin
      PgiInfo(AGL_YFILESTD_GET_ERR(num),'Import document lié') ;
      SetControlText ('LIENDOC','');
      end;
    end;

  if (ExtractFileName(FileName) <> oldFichier) and (oldFichier <> '') then
    my_YFileStd_Delete (CODEPRODUIT, oldFichier, 'FRA', 'NEW', GetString ('BCO_NUMERO'));
  oldFichier:=GetControlText ('LIENDOC');
  if GetControlText ('LIENDOC') = '' then SetControlEnabled  ('OUVRELIENDOC',false)
  else SetControlEnabled  ('OUVRELIENDOC',true);
end;

Procedure TOM_BASECONNAISSANCE.OuvreLienDoc( Sender : tObject);
var Num : integer;
    Path : string;
begin
   if GetControlText ('LIENDOC') = '' then exit;
   Num:=AGL_YFILESTD_EXTRACT(path, CODEPRODUIT, GetControlText ('LIENDOC'),
        GetString ('BCO_NUMERO'), '', '', '', '', FALSE, 'FRA', 'NEW');
   if num <> -1 then
     PgiInfo(AGL_YFILESTD_GET_ERR(num),'Extraction document lié')
   else
      if (FileExists (Path) = TRUE) then
        ShellExecute (0, pchar ('open'), pchar (Path), nil, nil, SW_RESTORE);
end;

Procedure TOM_BASECONNAISSANCE.my_YFileStd_Delete (CodeProduit, Nom, Langue, Predefini, Crit1 : String);
var Q : TQuery;
    Guid : string;
begin
  Guid:='';
  Q := OpenSql ('SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="' + Trim (CodeProduit) + '" AND YFS_NOM="' + Trim (Nom) + '" AND YFS_LANGUE="' + Trim (Langue) + '" AND YFS_PREDEFINI="' + Trim (Predefini) + '" AND YFS_CRIT1="' + Trim (Crit1) + '"', True) ;
  if not Q.Eof then
     Guid:=Q.FindField ('YFS_FILEGUID').AsString;
  Ferme (Q);
  ExecuteSQL('DELETE FROM YFILESTD WHERE YFS_CODEPRODUIT="' + Trim (CodeProduit) + '" AND YFS_NOM="' + Trim (Nom) + '" AND YFS_LANGUE="' + Trim (Langue) + '" AND YFS_PREDEFINI="' + Trim (Predefini) + '" AND YFS_CRIT1="' + Trim (Crit1) + '"');
  ExecuteSQL('DELETE FROM NFILES WHERE NFI_FILEGUID = "'+Guid+'"');
  ExecuteSQL('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID = "'+Guid+'"');
end;

Procedure TOM_BASECONNAISSANCE.my_YFileStd_update (CodeProduit, Nom, Langue, Predefini : String);
begin
  ExecuteSQL('UPDATE YFILESTD set YFS_CRIT1="'+GetString('BCO_NUMERO')+'" WHERE YFS_CODEPRODUIT="' + Trim (CodeProduit) + '" AND YFS_NOM="' +
     Trim (Nom) + '" AND YFS_LANGUE="' + Trim (Langue) + '" AND YFS_PREDEFINI="' + Trim (Predefini) + '" AND YFS_CRIT1="0"');
end;

Procedure TOM_BASECONNAISSANCE.RecupLesRattachements ;
var TobValeurs : tob;
    i,grid,lig : integer;
    sql,stValeurs,prefixe : string;
    GS : THGRID;
begin
  for i := 1 to nb_rattachement Do
    ValeurOri[i] := '';
  stValeurs:='';
  sql := 'SELECT LBC_RATTACHEMENT,LBC_VALEURR FROM '+TableLien+' WHERE LBC_NUMERO='+GetString('BCO_NUMERO')+
     ' ORDER BY LBC_NUMRATT';
  TobValeurs:=TOB.Create (TableLien, nil, -1);
  try
    TobValeurs.LoadDetailDBFromSQL(TableLien, sql);
  finally
    lig:=1;
    grid:=1;
    prefixe:='';
    GS:=THGRID(GetControl('GRATTACHEMENT'+IntToStr(grid)));
    if TobValeurs.Detail.count <> 0 then
      begin
      prefixe:=TobValeurs.Detail[0].GetString('LBC_RATTACHEMENT');
      ValeurOri[1] := prefixe;
      SetControlText('RATTACHEMENT'+IntToStr(grid),prefixe);
      for i:=0 to pred(TobValeurs.Detail.count) Do
        begin
        if prefixe <> TobValeurs.Detail[i].GetString('LBC_RATTACHEMENT') then
          begin
          if Assigned(GS) then
            begin
            GS.RowCount:=lig;
            MajLesLibelles(grid,1,GS.RowCount-1);
            end;
          inc(grid);
          prefixe:=TobValeurs.Detail[i].GetString('LBC_RATTACHEMENT');
          SetControlText('RATTACHEMENT'+IntToStr(grid),prefixe);
          ValeurOri[grid] := prefixe;
          lig:=1;
          end;
        GS:=THGRID(GetControl('GRATTACHEMENT'+IntToStr(grid)));
        // Cells [colonne,ligne)
        if Assigned(GS) then
          GS.Cells[1,lig]:=TobValeurs.Detail[i].GetString('LBC_VALEURR');
        inc(lig);
        end;
      end;
      if Assigned(GS) and (TobValeurs.Detail.count <> 0) then
        begin
        GS.RowCount:=lig;
        MajLesLibelles(grid,1,GS.RowCount-1);
        end;
    FreeAndNil(TobValeurs);
  end;
end;

procedure TOM_BASECONNAISSANCE.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    Case key of
      VK_DELETE :
      begin
      if (Shift=[ssCtrl]) and (copy(Screen.ActiveControl.Name,1,13) = 'GRATTACHEMENT' ) then
        begin
        GS:=THGRID(GetControl(Screen.ActiveControl.Name));
        if not Assigned(GS) then exit;
        if ( GS.Cells[1,GS.Row] <> '' ) then
          if GS.RowCount = 2 then
            GS.videpile (False)
          else
            GS.DeleteRow (GS.Row);
        if IsNumeric(copy(Screen.ActiveControl.Name,14,1)) then
          MajGrid[ValeurI(copy(Screen.ActiveControl.Name,14,1))]:=true;
        ForceUpdate;
        end;
      end;
    81 : {Ctrl + Q - Création d'1 alerte} if (ssCtrl in Shift) then
          begin
          Key:=0 ;
          Alerte_OnClick_BCO(Sender);
          end;
    85 : {Ctrl + U - liste des alertes du tiers} if (ssCtrl in Shift) then
          begin
          Key:=0 ;
          ListAlerte_OnClick_BCO(Sender);
          end;
    end;
end;

Procedure TOM_BASECONNAISSANCE.MajLesRattachements(iNum : integer; var numRatt : integer) ;
var TobMere, TobValeurs : tob;
    i : integer;
begin
  if GetControlText('RATTACHEMENT'+IntToStr(iNum)) <> '' then
    inc(numRatt);
  if not MajGrid[iNum] then exit;

  if ValeurOri[iNum] <> '' then
    ExecuteSql ('DELETE FROM '+TableLien+' WHERE LBC_NUMERO='+GetString('BCO_NUMERO')+
      ' AND LBC_RATTACHEMENT = "'+ ValeurOri[iNum] + '"');

  if GetControlText('RATTACHEMENT'+IntToStr(iNum)) = '' then exit;

  GS:=THGRID(GetControl('GRATTACHEMENT'+IntToStr(iNum)));

  if not Assigned(GS) then exit
    else
      if GS.RowCount = 1 then exit;

  TobMere:=TOB.Create ('Mere des liens', nil, -1);
  for i:= 1 to pred(GS.RowCount) Do
    begin
    if GS.Cells[1,i] <> '' then
      begin
      TobValeurs:=TOB.Create (TableLien, TobMere, -1);
      TobValeurs.SetInteger ('LBC_NUMERO',GetInteger('BCO_NUMERO'));
      TobValeurs.SetInteger ('LBC_NUMRATT',numRatt);
      TobValeurs.SetString ('LBC_RATTACHEMENT',GetControlText('RATTACHEMENT'+IntToStr(iNum)));
      TobValeurs.SetString ('LBC_VALEURR',GS.Cells[1,i]);
      end;
    end;

  if TobMere.Detail.Count <> 0 then
    try
      TobMere.InsertDB(nil,false);
    finally
    end;

  FreeAndNil(TobMere);
end;

Procedure TOM_BASECONNAISSANCE.MajLesLibelles(NumGrid,numDeb,numFin : integer) ;
var i : integer;
    chps,colLib,colCode : string;
begin
  if GetControlText('RATTACHEMENT'+IntToStr(NumGrid)) = '' then exit;
  for i:=1 to max_rattachement Do
    if GetControlText('RATTACHEMENT'+IntToStr(NumGrid)) =
      copy(Colonnes_BC[i],1,length(GetControlText('RATTACHEMENT'+IntToStr(NumGrid)))) then
         chps:=Colonnes_BC[i];

  GS:=THGRID(GetControl('GRATTACHEMENT'+IntToStr(NumGrid)));
  if not Assigned(GS) then exit;
  colLib:=ReadToKenSt(chps);
  colCode:=ReadToKenSt(chps);
  for i:=numDeb to numFin Do
    //function GetColonneSQL(Table, Colonne: string; Where: hstring): hstring;
    // attention faudra adapter si la clé devient autre qu'un seul champ de type varchar
    GS.Cells[2,i]:=GetColonneSQL(PrefixeToTable(GetControlText('RATTACHEMENT'+IntToStr(NumGrid))),
       colLib,colCode+' = "'+GS.Cells[1,i]+'"');
end;

{ GC/GRC : MNG / gestion des alertes }
procedure TOM_BASECONNAISSANCE.ListAlerte_OnClick_BCO(Sender: TObject);
begin
  if (GetString ('BCO_NUMERO') <> '') and(AlerteActive(TableToPrefixe(TableName))) then
    AGLLanceFiche('Y','YALERTES_MUL','YAL_PREFIXE=BCO','','ACTION=CREATION;MONOFICHE;CHAMP=BCO_NUMERO;VALEUR='
      +GetString ('BCO_NUMERO')+';LIBELLE='+GetString('BCO_LIBELLE')) ;
end ;

procedure TOM_BASECONNAISSANCE.Alerte_OnClick_BCO(Sender: TObject);
begin
  if (GetString ('BCO_NUMERO') <> '') and(AlerteActive(TableToPrefixe(TableName))) then
    AGLLanceFiche('Y','YALERTES','','','ACTION=CREATION;MONOFICHE;CHAMP=BCO_NUMERO;VALEUR='
      +GetString ('BCO_NUMERO')+';LIBELLE='+GetString('BCO_LIBELLE')) ;
VH_EntPgi.TobAlertes.ClearDetail;
end;

procedure TOM_BASECONNAISSANCE.OnLoadAlerte;
begin
  {$IFNDEF EAGLSERVER}
  { GC/GRC : MNG / gestion des alertes }
  if (not V_Pgi.SilentMode) and (not AfterInserting)
    and (AlerteActive (TableToPrefixe(TableName))) then
      ExecuteAlerteLoad(TForm(Ecran),true);
  { GC/GRC : MNG / gestion des alertes }
  {$ENDIF !EAGLSERVER}
end;
Initialization
  registerclasses ( [ TOM_BASECONNAISSANCE ] ) ;
end.


