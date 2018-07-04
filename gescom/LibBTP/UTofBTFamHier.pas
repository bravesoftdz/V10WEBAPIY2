unit UTofBTFamHier;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,Hpanel, Math
      ,HCtrls,HEnt1,HMsgBox,UTOF,vierge,UTOB,AglInit,LookUp,EntGC,SaisUtil,graphics
      ,grids,windows,FactUtil,Ent1,
{$IFDEF EAGLCLIENT}
     HPdfPrev,UtileAGL,Maineagl,eTablette,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,EdtEtat,FE_Main,Tablette,
{$ENDIF}
{$IFNDEF SANSCOMPTA}
      Ventil,
{$ENDIF}
      M3FP,HTB97,Dialogs, AGLInitGC;

Type
     TOF_BTFamHier = Class (TOF)
        procedure OnNew                    ; override ;
        procedure OnDelete                 ; override ;
        procedure OnUpdate                 ; override ;
        procedure OnLoad                   ; override ;
        procedure OnArgument (S : String ) ; override ;
        procedure OnDisplay                ; override ;
        procedure OnClose                  ; override ;
        procedure OnCancel                 ; override ;

     private
        //
        Action        : TActionFiche;
        TypeAction    : String; //C=Création - M=Modification - S=Suppression
        //
        Famille       : THValComboBox;
        //
        GSFAM2        : THGRID ;
        GSFAM3        : THGRID ;
        //
        TOBFAM1       : TOB;
        TOBFAM2       : TOB;
        TOBFAM3       : TOB;
        //
        Ok_CtrlSaisie : Boolean;
        //
        LibFam        : THLabel;
        LibFam2       : THLabel;
        LibFam3       : THLabel;
        //
        PFAMILLE2     : THPanel;
        PFAMILLE3     : THPanel;
        //
        Niveau1       : String;
        Niveau2       : String;
        Niveau3       : String;
        //
        Famille1      : String;
        Famille2      : String;
        Fam2Defaut    : String;
        Famille3      : String;
        //
        TypeAct       : String;
        //
        BValider      : TToolbarButton97;
        BImprimer     : TToolbarButton97;
        BDelete1      : TToolbarButton97;
        BDelete2      : TToolbarButton97;
        BInser1       : TToolbarButton97;
        BInser2       : TToolbarButton97;
        BNewLine      : TToolbarButton97;
        BDelLine      : TToolbarButton97;
        //
        fColNamesGS   : string;
        Falignement   : string;
        Ftitre        : string;
        fLargeur      : string;
        //
        QQ            : TQuery ;
        StSQL         : String ;
        //
        ColCode       : Integer;
        ColLibelle    : Integer;
        ColAbrege     : Integer;
        //
        //
        //
        procedure AfficheLagrille(GS: THGRID; TOBGrille: TOB);
        procedure AfficheLagrille2;
        procedure AfficheLagrille3;
        //
        procedure AfficheLesGrilles;
        procedure ChargeFamilleNiv1;
        procedure ChargeFamilleNiv2;
        procedure ChargeFamilleNiv3;
        procedure ChargementFamille;
        procedure ControleChamp(Champ, Valeur: String);
        procedure ControlCode2(var St: string);
        procedure ControlCode3(var St: string);
        procedure ControleSaisie;
        procedure CreateTOB;
        procedure DefinieGrid;
        procedure DeleteFamille3;
        procedure DeleteLigneGFAM2(Sender: Tobject);
        procedure DeleteLigneGFAM3(Sender: Tobject);
        procedure DessineGrille(GS: THGRID);
        procedure DestroyTOB;
        procedure EnableDisable(OK_Zone: Boolean; GS: THGrid);
        //
        procedure Fam2CellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure Fam2CellExit(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean);
        procedure Fam2RowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure Fam2RowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
        procedure Fam2GSOnDblClick(Sender: TObject);
        //
        procedure Fam3CellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure Fam3CellExit(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean);
        procedure Fam3RowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure Fam3RowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
        procedure Fam3GSOnDblClick(Sender: TObject);
        //
        procedure FamilleOnChange(Sender: Tobject);
        function  GetTOBGrille(TOBGS: TOB; ARow: integer): TOB;
        procedure GetObjects;
        procedure GridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        //
        procedure MAJTOBFamille2(TOBL: TOB; Arow: Integer);
        procedure MAJTOBFamille3(Arow: Integer);
        //
        procedure NewLigne(GS: THGRID; TOBFAM: TOB; Niveau, Libre: String);
        procedure NewLigneGFAM2(Sender: Tobject);
        procedure NewLigneGFAM3(Sender: Tobject);
        //
        procedure SetGrilleEvents(Etat: boolean);
        procedure SetScreenEvents;
        //
        procedure ValideSaisie(Sender: Tobject);
        //
        function  ZoneAccessible  (GS : THGrid; var ACol, ARow: integer): boolean;
        procedure ZoneSuivanteOuOk(GS : THGrid; var ACol, ARow: integer; var Cancel: boolean);

        //
     public

     END ;

Implementation
uses  Messages,
      ParamSoc,
      BTPUtil,
      DateUtils,
      StrUtils,
      UPrintScreen, TntGrids;

procedure TOF_BTFamHier.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTFamHier.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTFamHier.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTFamHier.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTFamHier.OnArgument (S : String );
var i       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
    Indice  : Integer;
begin
  Inherited ;
  //
  Ok_CtrlSaisie := False;
  //
  //Chargement des zones ecran dans des zones programme
  GetObjects;
  //
  CreateTOB;
  //
  Critere := S;

  While (Critere <> '') do
  BEGIN
    i:=pos(':',Critere);
    if i = 0 then i:=pos('=',Critere);
    if i <> 0 then
       begin
       Champ:=copy(Critere,1,i-1);
       Valeur:=Copy (Critere,i+1,length(Critere)-i);
       end
    else
       Champ := Critere;
    Controlechamp(Champ, Valeur);
    Critere:=(Trim(ReadTokenSt(S)));
  END;

  if Famille2 <> '' then Fam2Defaut:=famille2;

  //Gestion des évènement des zones écran
  SetScreenEvents;

  //Gestion des évènement de la grille
  SetGrilleEvents(True);

  // définition de la grille
  DefinieGrid;

  //dessin de la grille
  DessineGrille(GSFAM2);
  //
  DessineGrille(GSFAM3);
  //
  if Famille1='' then
  begin
    if Famille.Items.Count<>0 then Famille.ItemIndex:=0;
    Famille1:=Famille.Value;
  end
  else Famille.Value:=Famille1;

  // gestion de famille ouvrage NIV1 BATIPRIX
  If Niveau1 = 'BO1' then
  Begin
  	Famille.DataType:='BTFAMILLEOUV1';
    Ecran.Caption   := 'Familles Ouvrages';
    LibFam.Caption  :=RechDom('GCLIBOUVRAGE','BP'+Copy(Niveau1,3,1),false);
    LibFam2.Caption :=RechDom('GCLIBOUVRAGE','BP'+Copy(Niveau2,3,1),false);
    LibFam3.Caption :=RechDom('GCLIBOUVRAGE','BP'+Copy(Niveau3,3,1),false);
  end
  else If Niveau1 = 'BP1' then
  Begin
  	Famille.DataType:='BTFAMILLEPARC1';
    Ecran.Caption   := 'Familles Parc';
    LibFam.Caption  :=RechDom('GCLIBFAMILLE','BP'+Copy(Niveau1,3,1),false);
    LibFam2.Caption :=RechDom('GCLIBFAMILLE','BP'+Copy(Niveau2,3,1),false);
    LibFam3.Caption :=RechDom('GCLIBFAMILLE','BP'+Copy(Niveau3,3,1),false);
  end
  else
  begin
    Famille.DataType:='GCFAMILLENIV'+Copy(Niveau1,3,1);
    Ecran.Caption   := 'Familles Articles';
    LibFam.Caption  :=RechDom('GCLIBFAMILLE','LF'+Copy(Niveau1,3,1),false);
    LibFam2.Caption :=RechDom('GCLIBFAMILLE','LF'+Copy(Niveau2,3,1),false);
    LibFam3.Caption :=RechDom('GCLIBFAMILLE','LF'+Copy(Niveau3,3,1),false);
  end;

  ChargementFamille;

  GSFAM2.Options := GSFAM2.Options - [goEditing] + [goRowSelect];
  GSFAM3.Options := GSFAM2.Options - [goEditing] + [goRowSelect];

  if (Fam2Defaut <> famille2) and (Fam2Defaut <> '') then
  begin
    Famille2 := Fam2defaut;
    for Indice := 0 to GSFAM2.RowCount - 1 do
    begin
      if GSFAM2.Cells[ColCode, Indice]=Famille2 then
      begin
        GSFAM2.Row := Indice;
        Break;
      end;
    end;
    //
    AfficheLagrille3;
  end;

end;

Procedure TOF_BTFamHier.ChargementFamille;
begin

  //chargement de la Famille de Niveau 1
  ChargeFamilleNiv1;

  //Affichage de la grille sous-famille niveau 2
  AfficheLagrille2;

  //Affichage de la grille sous-famille niveau 2
  AfficheLagrille3;

  GSFAM2.SetFocus;

end;

Procedure TOF_BTFamHier.AfficheLagrille2;
begin

  //chargement de la Famille de Niveau 2 En fonction de famille
  ChargeFamilleNiv2;

  AffichelaGrille(GSFAM2, TOBFAM2);

  //on se positionne sur la première ligne après rechargement
  if TypeAction = '' then
  begin
    if TOBFAM2.detail.count < 0 then
      exit
    else
    begin
      GSFAM2.col := ColCode;
      GSFAM2.row := 1;
    end;
  end;

  if TOBFAM2.Detail.Count > 0 then Famille2 := TOBFAM2.Detail[GSFAM2.Row-1].GetString('CC_CODE');

end;

Procedure TOF_BTFamHier.AfficheLagrille3;
begin

  //chargement de la famille de Niveau 3 en fonction de Niveau 1 et Niveau 2
  ChargeFamilleNiv3;
  //
  AffichelaGrille(GSFAM3, TOBFAM3);

  //on se positionne sur la première ligne après rechargement
  if TypeAction = '' then
  begin
    if TOBFAM3.detail.count < 0 then
      exit
    else
    begin
      GSFAM3.col := ColCode;
      GSFAM3.row := 1;
    end;
  end;

  if TOBFAM3.Detail.Count > 0 then Famille3 := TOBFAM3.Detail[GSFAM3.Row-1].GetString('CC_CODE');

end;

Procedure TOF_BTFamHier.AffichelesGrilles;
begin

  AfficheLagrille(GSFAM2, TOBFAM2);

  AfficheLagrille(GSFAM3, TOBFAM3);

end;

procedure TOF_BTFamHier.OnClose ;
begin
  Inherited ;

  //controle si au moins une saisie à été faite...
  If Ok_CtrlSaisie then
  begin
    if PGIAsk('Des modifications ont été effectuées, Voulez-vous les enregistrer ?', 'Sauvegarde')=MrYes then
    begin
      ValideSaisie(Self);
    end;
  end;

  DestroyTOB;

end ;

procedure TOF_BTFamHier.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTFamHier.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTFamHier.CreateTOB;
begin

  TOBFAM1  := TOB.Create('FAMILLE NIVEAU 1', nil, -1);
  TOBFAM2  := TOB.Create('FAMILLE NIVEAU 2', nil, -1);
  TOBFAM3  := TOB.Create('FAMILLE NIVEAU 3', nil, -1);

end;

procedure TOF_BTFamHier.DestroyTOB;
begin

  FreeAndNil(TOBFAM1);
  FreeAndNil(TOBFAM2);
  FreeAndNil(TOBFAM3);

end;

procedure TOF_BTFamHier.GetObjects;
begin
  //
  GSFAM2      := THgrid(GetControl('GFAM2'));
  GSFAM3      := THgrid(GetControl('GFAM3'));
  //
  PFAMILLE2   := THPanel(GetControl('PARTICLE'));
  PFAMILLE3   := THPanel(GetControl('PARTICLE1'));
  //
  BValider    := TToolbarButton97 (GetControl('BVALIDER'));
  BImprimer   := TToolbarButton97 (GetControl('BIMPRIMER'));
  BDelete1    := TToolbarButton97 (GetControl('BDelete1'));
  BDelete2    := TToolbarButton97 (GetControl('BDelete2'));
  BInser1     := TToolbarButton97 (GetControl('BInsert1'));
  BInser2     := TToolbarButton97 (GetControl('BInsert2'));
  BNewLine    := TToolbarButton97 (GetControl('BNEWLINE'));
  BDelLine    := TToolbarButton97 (GetControl('BDELLINE'));
  //
  Famille := THValComboBox(GetControl('FAMNVSUP'));
  //
  LibFam  := THLabel(GetControl('TFAMNVSUP'));
  LibFam2 := THLabel(GetControl('TFAMNVSUP2'));
  LibFam3 := THLabel(GetControl('TFAMNVSUP3'));

end;

procedure TOF_BTFamHier.SetGrilleEvents (Etat : boolean);
begin

  if Etat then
  begin
    GSFAM2.OnCellEnter  := Fam2CellEnter;
    GSFAM2.OnCellExit   := Fam2CellExit;
    GSFAM2.OnRowEnter   := Fam2RowEnter;
    GSFAM2.OnRowExit    := Fam2RowExit;
    GSFAM2.OnKeyDown    := GridKeyDown;
    GSFAM2.OnDblClick   := Fam2GSOnDblClick;
    //
    GSFAM3.OnCellEnter  := Fam3CellEnter;
    GSFAM3.OnCellExit   := Fam3CellExit;
    GSFAM3.OnRowEnter   := Fam3RowEnter;
    GSFAM3.OnRowExit    := Fam3RowExit;
    GSFAM3.OnKeyDown    := GridKeyDown;
    GSFAM3.OnDblClick   := Fam3GSOnDblClick;
  end
  else
  begin
    GSFAM2.OnCellEnter  := nil;
    GSFAM2.OnCellExit   := nil;
    GSFAM2.OnRowEnter   := Nil;
    GSFAM2.OnKeyDown    := nil;
    GSFAM2.OnDblClick   := Nil;
    //
    GSFAM3.OnCellEnter  := nil;
    GSFAM3.OnCellExit   := nil;
    GSFAM3.OnRowEnter   := nil;
    GSFAM3.OnKeyDown    := nil;
    GSFAM3.OnDblClick   := Nil;
  end;

end;

procedure TOF_BTFamHier.SetScreenEvents;
begin

  Famille.OnChange := FamilleOnChange;
  //
  BValider.OnClick := ValideSaisie;

  BInser1.OnClick  := NewLigneGFAM2;
  BDelete1.OnClick := DeleteLigneGFAM2;

  BInser2.OnClick  := NewLigneGFAM3;
  BDelete2.OnClick := DeleteLigneGFAM3;

end;

Procedure TOF_BTFAMHIER.FamilleOnChange(Sender : Tobject);
begin


  ControleSaisie;


  If Ok_CtrlSaisie then
  begin
    if PGIAsk('Des modifications ont été effectuées, Voulez-vous les enregistrer ?', 'Sauvegarde')=MrYes then
    begin
      ValideSaisie(SELF);
      Exit;
    end;
  end;

  DestroyTOB;

  GSFAM2.Options := GSFAM2.Options + [goEditing] + [goRowSelect];
  GSFAM3.Options := GSFAM3.Options + [goEditing] + [goRowSelect];

  Famille1:=Famille.Value;

  TypeAction := '';

  DestroyTOB;
  CreateTOB;

  ChargementFamille;

end;

Procedure TOF_BTFAMHIER.ControleSaisie;
Var indice  : Integer;
    TOBL    : TOB;
begin

  Ok_CtrlSaisie := False;

  //contrôle de la famille 2
  for Indice :=0 to TOBFAM2.Detail.count - 1 do
  begin
    TOBL := TOBFAM2.Detail[Indice];
    if (TOBL.GetString('CC_CODE')    <> GSFAM2.Cells[ColCode, Indice+1])    or
       (TOBL.GetString('CC_LIBELLE') <> GSFAM2.Cells[ColLibelle, Indice+1]) or
       (TOBL.GetString('CC_ABREGE')  <> GSFAM2.Cells[ColAbrege, Indice+1]) then
    begin
      Ok_CtrlSaisie := True;
      break;
    end;
  end;

  if Ok_CtrlSaisie then
  begin
    For indice :=0 To TOBFAM3.Detail.Count - 1 do
    begin
      TOBL := TOBFAM3.Detail[indice];
      if (TOBL.GetString('CC_CODE')    <> GSFAM3.Cells[ColCode, Indice+1]) or
         (TOBL.GetString('CC_LIBELLE') <> GSFAM3.Cells[ColLibelle, Indice+1]) or
         (TOBL.GetString('CC_ABREGE')  <> GSFAM3.Cells[ColAbrege, Indice+1]) then
      begin
        Ok_CtrlSaisie := True;
        break;
      end;
    end;
  end;

end;

procedure TOF_BTFAMHIER.ValideSaisie(Sender : Tobject);
Var liggrille : integer;
    LigTob    : Integer;
begin

  //FV1 : 25/06/2017 - FS#1564 - SOPREL : Pb pour créer famille article niveau 2

  LigTob    := GSFam2.Row-1;
  LigGrille := GSFAM2.row;

  MAJTOBFamille2(TOBFAM2.detail[LigTob], LigGrille);

  MAJTOBFamille3(GSFAM3.row);

end;

procedure TOF_BTFAMHIER.DeleteLigneGFAM2(Sender : Tobject);
Var TOBL  : TOB;
    StSQL : String;
    Libre : string;
begin

  if Action = taConsult then Exit;

  TOBL := TOBFAM2.FindFirst(['CC_TYPE','CC_CODE'],[Niveau2,Famille2], False);
  If TOBL = nil then exit;

  //On Contrôle si cette sous- famille n'est pas sur un article
  StSQL := 'SELECT GA_ARTICLE FROM ARTICLE WHERE GA_FAMILLENIV1="' + Famille1 + '" AND GA_FAMILLENIV2="' + Famille2 + '"';
  If ExisteSQL(StSQL) then
  begin
    PGIError('Suppression de ' + LibFam2.Caption + ' :' + Famille2 + ' impossible. Elle se trouve sur un article !', 'Suppression');
    Exit;
  end;

  //on Contrôle si cette sous- famille n'est pas sur une ligne document
  StSQL := 'SELECT GL_NUMERO FROM LIGNE WHERE GL_FAMILLENIV1="' + Famille1 + '" AND GL_FAMILLENIV2="' + Famille2 + '"';
  If ExisteSQL(StSQL) then
  begin
    PGIError('Suppression de ' + LibFam2.Caption + ' :' + Famille2 + ' impossible. Elle se trouve sur un document !', 'Suppression');
    Exit;
  end;

  //Suppression de la ligne
  if PGIAsk('Confirmez-vous la suppression de ' + LibFam2.Caption + ' :' + Famille2 + '?','Suppression')=MrNO  then Exit;

  if TOBFAM3.Detail.count > 0 then
  begin
    if PGIAsk('Attention : toutes les ' + LibFam3.Caption + ' associées seront supprimées. confirmez-vous ?', 'suppression')=mrNo then Exit;
  end;

  EnableDisable(False, GSFAM3);

  GSFAM2.CacheEdit;
  GSFAM2.SynEnabled := False;

  Famille2 := GSFAM2.Cells[ColCode, GSFAM2.Row];

  Libre := TOBL.GetString('CC_LIBRE');

  //Si la sous Famille est sur plusieurs famille
  if Pos(';',Libre) > 0 then
  begin
    if Pos(Famille1,Libre) < Pos(';',Libre) then
      Libre := AnsiReplaceStr(libre, Famille1 + ';', '')
    else
      Libre := AnsiReplaceStr(libre, ';' + Famille1, '');
    TOBL.PutValue('CC_Libre', Libre);
    TOBL.UpdateDB(false, False, '');
    //suppression de la  ligne dans la grille
    GSFAM2.DeleteRow(GSFAM2.Row);
    //
    FreeAndNil(TOBL);
    //
    //suppression des sous-famille niuveau 3
    if TOBFAM3.Detail.count > 0 then DeleteFamille3;
  end
  else
  begin
    //Suppression de la famille 2 dans la table
    StSQL := 'DELETE FROM CHOIXCOD WHERE CC_TYPE ="' + Niveau2 + '" AND CC_CODE="' + Famille2 + '"';
    if ExecuteSQL(StSQL) > 0 then
    begin
      //suppression de la  ligne dans la grille
      GSFAM2.DeleteRow(GSFAM2.Row);
      //
      FreeAndNil(TOBL);
      //
      //suppression des sous-famille niuveau 3
      if TOBFAM3.Detail.count > 0 then DeleteFamille3;
    end
    else
      PGIError('Suppression ' + LibFam2.Caption + ' Impossible');
  end;


  AfficheLesGrilles;

  GSFAM2.SynEnabled := True;

  EnableDisable(True, GSFAM3);

end;

Procedure TOF_BTFAMHIER.DeleteFamille3;
begin

  StSQL := 'DELETE FROM CHOIXCOD WHERE CC_TYPE = "'+ Niveau3 +'" AND CC_LIBRE like "%'+ Famille1+Famille2 +'%"';

  if ExecuteSQL(StSQL)>0 then
    //On Recharge les code famille de niveau 3 même si on sait quil n'y aura rien !!!!
    ChargeFamilleNiv3
  else
    PGIError('Suppression ' + LibFam3.Caption + ' Impossible');

end;

procedure TOF_BTFAMHIER.NewLigneGFAM2(Sender : Tobject);
Var ARow      : Integer;
    Cancel    : Boolean;
begin

  Cancel := False;

  EnableDisable(False, GSFAM3);

  If TOBFAM2.detail.Count = 0 then
    ARow := TOBFAM2.detail.Count+2
  else
    ARow := TOBFAM2.detail.Count+1;

  GSFAM3.Enabled := False;

  NewLigne(GSFAM2, TOBFAM2, Niveau2, Famille1);

  Fam2CellEnter(Self, ColCode, Arow, Cancel);

end;

Procedure TOF_BTFamHier.Fam2GSOnDblClick(Sender: TObject);
begin

  if TypeAct = 'SAIART' then Exit;

  TypeAction := 'M';

  GSFAM2.Options := GSFAM2.Options + [goEditing] - [goRowSelect];

  EnableDisable(False, GSFAM3);

end;

Procedure TOF_BTFamHier.Fam3GSOnDblClick(Sender: TObject);
begin

  if TypeAct = 'SAIART' then Exit;

  TypeAction := 'M';

  GSFAM3.Options := GSFAM3.Options + [goEditing] - [goRowSelect];

  EnableDisable(False, GSFAM2);

  GSFAM3.Setfocus;

end;

procedure TOF_BTFamHier.Fam2RowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin

  if Action = TaConsult then Exit;

  Famille2 := GSFAM2.Cells[ColCode,GSFAM2.row];

end;
//
procedure TOF_BTFamHier.Fam2CellEnter(Sender: TObject; var ACol,  ARow: Integer; var Cancel: Boolean);
begin

  if Action = TaConsult then Exit;

  ZoneSuivanteOuOk(GSFAM2,ACol, ARow, Cancel);

end;

Procedure TOf_BTFAMHIER.Fam2RowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
VAR Arow    : Integer;
    TOBL    :  TOB;
begin

  if Action = taConsult then Exit;

  ARow:= Ou;

  GSFAM2.Options := GSFAM2.Options - [goEditing] + [goRowSelect];

  if (GSFAM2.Cells[ColCode, ARow] = '')     OR
     (GSFAM2.Cells[COLLibelle, ARow] = '')  OR
     (GSFAM2.Cells[ColAbrege, ARow] = '') then
  begin
     //Si on est en création...
     if TypeAction = 'C' then
     begin
       TOBL := TOBFAM2.detail[Arow-1];
       FreeAndNil(TOBL);
       //
       GSFam2.DeleteRow(Arow);
       if ARow > 1 then GSFAM2.Row := ARow - 1 Else GSFAM2.Row := 1;
     end
     else if TypeAction = 'M' then
     begin
       //Si on est en Modification...
       GSFAM2.Cells[COLLibelle, ARow] := TOBFAM2.Detail[Arow-1].GetValue('CC_LIBELLE');
       GSFAM2.Cells[ColAbrege, ARow]  := TOBFAM2.Detail[Arow-1].GetValue('CC_ABREGE');
     end;
  end
  else
  begin
    if TypeAction <> '' then
    begin
      MAJTOBFamille2(TOBFAM2.detail[Arow-1], Arow);
    end;
  end;

  TypeAction := '';

  EnableDisable(True, GSFAM2);

end;

Procedure TOF_BTFamhier.MAJTOBFamille2(TOBL : TOB; Arow : Integer);
Var Code    : string;
    Libre   : string;
    Indice  : Integer;
begin

  Code := GSFAM2.Cells[ColCode, Arow];

  if Code = '' then exit;

  If (TOBL.GetValue('CC_CODE')='') then
    TOBL.PutValue('CC_CODE',Code)
  else
  begin
    //Si le code Famille a été modifé modifier le CC_Libre de l'ensemble des Sous-famille de niveau 3
    if (Code <> TOBL.GetValue('CC_CODE')) then
    begin
      TOBL.PutValue('CC_CODE', Code);
      For Indice := 0 TO TOBFAM3.Detail.count -1 do
      begin
        Libre := TOBFAM3.Detail[Indice].GetValue('CC_LIBRE');
        Libre := StringReplace(Libre,TOBL.GetValue('CC_CODE'),Code,[rfReplaceAll]);
        TOBFAM3.Detail[Indice].PutValue('CC_LIBRE', Libre);
      end;
    end;
  end;

  TOBL.PutValue('CC_LIBELLE',GSFAM2.Cells[ColLibelle, Arow]);
  TOBL.PutValue('CC_ABREGE', GSFAM2.Cells[ColAbrege, Arow]);
  //
  If TOBL.Getvalue('CC_LIBRE') <> '' then
  begin
    If Pos(Famille1,TOBL.Getvalue('CC_LIBRE')) = 0 then TOBL.PutValue('CC_LIBRE', TOBL.GetValue('CC_LIBRE') + ';' + Famille1)
  end;

  TOBL.InsertOrUpdateDB(false);

  GSFAM2.Options := GSFAM2.Options - [goEditing] + [goRowSelect];

  EnableDisable(True, GSFAM3);

end;

procedure TOF_BTFamHier.Fam2CellExit(Sender: TObject; var ACol,  ARow: Integer; var Cancel: Boolean);
Var St : String;
begin

  if Action = taConsult then Exit;

  St := GSFAM2.Cells[ACol,Arow];

  if St=  '' then
  begin
    Fam2CellEnter(Self, ColCode, ARow,Cancel);
    exit;
  end;

  If ACol = ColCode  then
  begin
    St := GSFAM2.Cells [ColCode, ARow];
    ControlCode2(St);
    GSFAM2.Cells [ColCode, ARow] := St;
  end
  else if Acol = ColLibelle then
  begin
    if GSFAM2.Cells[ACol,Arow] = '' then
    begin
      PGIInfo('Le libellé de la famille est obligatoire',Ecran.Caption);
      //Cancel:=True;
    end
    else
      GSFAM2.Cells[ColAbrege,ARow]:= copy(GSFAM2.Cells[ColLibelle,ARow],1,17);
  end
  else if Acol = ColAbrege then
  begin
    Fam2RowExit(Self, ARow, Cancel, False);
  end;


end;

Procedure TOF_BTFAMHIER.ControlCode2(var St : string);
Var TOBL : TOB;
    Arow : Integer;
    code : string;
begin

  if TypeAction = '' then exit;

  Arow := GSFAM2.Row;

  TOBL := TOBFAM2.detail[Arow-1];

  //On charge automatiquement la zone en majuscule
  St := UpperCase(St);

  //on vérifie si le code saisit est le même que celui de la tob
  //si c'est le cas on ne fait aucun contrôle...
  if TOBL.GetValue('CC_CODE') = St then Exit;

  //On vérifie d'abord l'existence au niveau de la tob affichée
  TOBL := TOBFAM2.FindFirst(['CC_TYPE','CC_CODE'],[Niveau2,St], False);
  if TOBL <> nil then
  begin
    PGIInfo('Ce code famille ' + St + ' existe déjà sur cette famille',Ecran.Caption);
    St := '';
    Exit;
  end;

  //Si il n'existe pas dans la TOB on vérifie s'il n'existe pas dans la table mais sous une autre famille
  StSQL := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="' + Niveau2 + '" AND CC_CODE="' + st + '"';
  QQ := OpenSQL(StSQl, False);
  If Not QQ.eof then
  begin
    if PGIAsk('Cette Sous-Famille ' + St + ' existe sur une ou plusieurs autres Familles.' + CHR(10) + 'Voulez-vous ajouter cette Famille en plus ?') = MrYes then
    begin
      St := QQ.findField('CC_CODE').AsString;
      //chargement de la zone CC_LIBRE de la TOB avec la zone CC_LIBRE de la TABLE
      //Et chargement du libellé de l'autre sous-famille
      GSFAM2.Cells  [ColLibelle, GSFAM2.row]  := QQ.findfield('CC_LIBELLE').asString;
      GSFAM2.Cells  [ColAbrege,GSFAM2.row]    := copy(GSFAM2.Cells[ColLibelle,GSFAM2.row],1,17);
      TOBFAM2.detail[GSFAM2.Row-1].PutValue('CC_LIBRE',   QQ.findfield('CC_LIBRE').asString);
    end
    else
    begin
      St := '';
    End;
  end;

end;


procedure TOF_BTFAMHIER.ControlCode3(var St : string);
Var TOBL : TOB;
    Arow : Integer;
    code : string;
begin

  if TypeAction =  '' then exit;
  
  Arow := GSFAM3.Row;

  TOBL := TOBFAM3.detail[Arow-1];

  //On charge automatiquement la zone en majuscule
  St := UpperCase(St);

  //on vérifie si le code saisit est le même que celui de la tob
  //si c'est le cas on ne fait aucun contrôle...
  if TOBL.GetValue('CC_CODE') = St then Exit;

  //On vérifie d'abord l'existence au niveau de la tob affichée
  TOBL := TOBFAM3.FindFirst(['CC_TYPE','CC_CODE'],[Niveau3,St], False);
  if TOBL <> nil then
  begin
    PGIInfo('Ce code ' + Libfam3.caption + 'existe déjà',Ecran.Caption);
    St := '';
    Exit;
  end;

  //Si il n'existe pas dans la TOB on vérifie s'il n'existe pas dans la table mais sous une autre famille
  StSQL := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="' + Niveau3 + '" AND CC_CODE="' + st + '"';
  QQ := OpenSQL(StSQl, False);
  If Not QQ.eof then
  begin
    if PGIAsk('Cette ' + LibFam3.caption + ' :' + St + ' existe sur une ou plusieurs autres ' + LibFam2.caption + '.' + CHR(10) + 'Voulez-vous l''ajouter ?') = MrYes then
    begin
      St := QQ.findfield('CC_CODE').asString;
      //chargement de la zone CC_LIBRE de la TOB avec la zone CC_LIBRE de la TABLE
      //Et chargement du libellé de l'autre sous-famille
      GSFAM3.Cells[ColLibelle, GSFAM3.row]  := QQ.findfield('CC_LIBELLE').asString;
      GSFAM3.Cells[ColAbrege,GSFAM3.row]    := copy(GSFAM3.Cells[ColLibelle,GSFAM3.row],1,17);
      TOBFAM3.detail[GSFAM3.Row-1].PutValue('CC_LIBRE', QQ.findfield('CC_LIBRE').asString);
    end
    else
    begin
      St := '';
    End;
  end;

end;    
//
procedure TOF_BTFamHier.Fam3CellEnter(Sender: TObject; var ACol,  ARow: Integer; var Cancel: Boolean);
begin

  if Action = TaConsult then Exit;

  ZoneSuivanteOuOk(GSFAM3,ACol, ARow, Cancel);

end;

procedure TOF_BTFamHier.Fam3CellExit(Sender: TObject; var ACol,  ARow: Integer; var Cancel: Boolean);
Var St : string;
begin

  if Action = taConsult then Exit;

  St := GSFAM3.Cells[ACol,Arow];

  if St=  '' then
  begin
    Fam3CellEnter(Self, ColCode, ARow,Cancel);
    exit;
  end;

  If ACol = ColCode  then
  begin
    St := GSFAM3.Cells [ColCode, ARow];
    ControlCode3(St);
    GSFAM3.Cells [ColCode, ARow] := St;
  end
  else if Acol = ColLibelle then
  begin
    if GSFAM3.Cells[ACol,Arow] = '' then
    begin
      PGIInfo('Le libellé de la Sous-famille est obligatoire',Ecran.Caption);
      //Cancel:=True;
    end
    else
      GSFAM3.Cells[ColAbrege,ARow]:= copy(GSFAM3.Cells[ColLibelle,ARow],1,17);
  end
  else if Acol = ColAbrege then
  begin
    Fam3RowExit(Self, ARow, Cancel, False);
  end;

end;

procedure TOF_BTFamHier.Fam3RowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin

  if Action = TaConsult then Exit;

  Famille3 := GSFAM3.Cells[ColCode,GSFAM3.row];

end;

procedure TOF_BTFamHier.Fam3RowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
VAR Arow    : Integer;
    TOBL    :  TOB;
begin

  if Action = taConsult then Exit;

  ARow:= Ou;

  GSFAM3.Options := GSFAM3.Options - [goEditing] + [goRowSelect];

  if (GSFAM3.Cells[ColCode, ARow] = '')     OR
     (GSFAM3.Cells[COLLibelle, ARow] = '')  OR
     (GSFAM3.Cells[ColAbrege, ARow] = '') then
  begin
     //Si on est en création...
     if TypeAction = 'C' then
     begin
       TOBL := TOBFAM3.detail[Arow-1];
       FreeAndNil(TOBL);
       //
       GSFam3.DeleteRow(Arow);
       if ARow > 1 then GSFAM3.Row := ARow - 1 Else GSFAM3.Row := 1;
     end
     else if TypeAction = 'M' then
     begin
       //Si on est en Modification...
       GSFAM3.Cells[ColLibelle, ARow] := TOBFAM2.Detail[Arow-1].GetValue('CC_LIBELLE');
       GSFAM3.Cells[ColAbrege, ARow]  := TOBFAM2.Detail[Arow-1].GetValue('CC_ABREGE');
     end;
  end
  else
  begin
    if TypeAction <> '' then
    begin
      MAJTOBFamille3(Arow);
    end;
  end;

  TypeAction := '';

  EnableDisable(True, GSFAM2);

end;

Procedure TOF_BTFamhier.MAJTOBFamille3(Arow : Integer);
Var Code    : string;
    TOBL    : TOB;
begin

  Code := GSFAM3.Cells[ColCode, Arow];

  if code = '' then exit;

  TOBL := TOBFAM3.detail[Arow-1];

  TOBL.PutValue('CC_CODE',Code);
  //
  TOBL.PutValue('CC_LIBELLE',GSFAM3.Cells[ColLibelle, Arow]);
  TOBL.PutValue('CC_ABREGE', GSFAM3.Cells[ColAbrege, Arow]);
  //
  If TOBL.Getvalue('CC_LIBRE') <> '' then
  begin
    If Pos(Famille1+Famille2,TOBL.Getvalue('CC_LIBRE')) = 0 then TOBL.PutValue('CC_LIBRE', TOBL.GetValue('CC_LIBRE') + ';' + Famille1+Famille2);
  end
  else
    TOBL.PutValue('CC_LIBRE', Famille1+Famille2);

  TOBL.InsertOrUpdateDB(false);

  GSFAM3.Options := GSFAM3.Options - [goEditing] + [goRowSelect];

  EnableDisable(True, GSFAM3);


end;


procedure TOF_BTFamHier.ZoneSuivanteOuOk(GS : THGrid; var ACol, ARow: integer; var Cancel: boolean);
var Sens    : Integer;
    Lim     : Integer;
    OldEna  : Boolean;
    ChgLig  : Boolean;
    ChgSens : Boolean;
    ii      : Integer;
begin

  OldEna := GS.SynEnabled;
  GS.SynEnabled := False;

  ii := 0;

  Sens := -1;

  ChgLig  := (GS.Row <> ARow);
  ChgSens := False;

  if GS.Row > ARow then Sens := 1 else if ((GS.Row = ARow) and (ACol <= GS.Col)) then Sens := 1;

  ACol := GS.Col;
  ARow := GS.Row;

  while not ZoneAccessible(GS,ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      Lim := GS.RowCount ;
      // ---
      if ((ACol = GS.ColCount - 1) and (ARow >= Lim)) then
      begin
        if ChgSens then
          Break
        else
          // Ajout d'une ligne
          break;
      end;
      if ChgLig then
      begin
        ACol := GS.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < GS.ColCount - 1 then
        Inc(ACol)
      Else
      begin
        Inc(ARow);
        ACol := GS.FixedCols;
      end;
    end
    else
    begin
      if ((ACol = GS.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then
          Break
        else
        begin
          Sens := 1;
          Continue;
        end;
      end;
      if ChgLig then
      begin
        ACol := GS.ColCount;
        ChgLig := False;
      end;
      if ACol > GS.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := 1;
      end;
    end;
  end;

  GS.SynEnabled := OldEna;

end;

function TOF_BTFamHier.ZoneAccessible(GS : THGrid; var ACol, ARow: integer): boolean;
begin

  result := true;

  if GS.ColWidths[acol] = 0 then
  Begin
    result := false;
    exit;
  End;

  //A partir de là on vérifie si la zone est gérable dans la grille ou non.
  if (GS.Name='GFAM2') then
  begin
    if (ACol = ColCode) Or (ACol = ColLibelle) Or (ACol = ColAbrege) then
    begin
      Famille2 := GSFAM2.Cells[ColCode,GSFAM2.row];
      //
      AfficheLagrille3;
      //
      result := true;
    end;
  end
  Else if (GS.Name='GFAM3') then
  begin
    if (ACol = ColCode) Or (ACol = ColLibelle) Or (ACol = ColAbrege) then Result := True;
  end;

end;

function TOF_BTFamHier.GetTOBGrille(TOBGS : TOB;ARow: integer): TOB;
begin

  result := nil;

  if Arow > TOBGS.detail.count Then exit;

  Result := TOBGS.detail[Arow -1];
  
end;

procedure TOF_BTFamHier.DefinieGrid;
begin

  // Définition de la liste de saisie pour la grille Détail
  fColNamesGS := 'SEL;CC_CODE;CC_LIBELLE;CC_ABREGE';
  Falignement := 'C.0  ---;C.0  ---;G.0  ---;G.0  ---';
  Ftitre      := ' ;Code;Libellé;Abregé';
  fLargeur    := '2;4;40;20';

end;

procedure TOF_BTFamHier.DessineGrille(GS : THGRID);
var st              : String;
    lestitres       : String;
    lesalignements  : String;
    FF              : String;
    alignement      : String;
    Nam             : String;
    leslargeurs     : String;
    lalargeur       : String;
    letitre         : String;
    lelement        : string;
    //
    Obli            : Boolean;
    OkLib           : Boolean;
    OkVisu          : Boolean;
    OkNulle         : Boolean;
    OkCumul         : Boolean;
    Sep             : Boolean;
    Okimg           : boolean;
    //
    dec             : Integer;
    NbCols          : integer;
    indice          : Integer;
    //
    TailleEcran     : Variant;
begin
  //
  //Calcul du nombre de Colonnes du Tableau en fonction des noms
  st := fColNamesGS;

  NbCols := 0;

  repeat
    lelement := READTOKENST (st);
    if lelement <> '' then
    begin
      inc(NbCols);
    end;
  until lelement = '';
  //
  GS.ColCount     := Nbcols;
  //
  st              := fColNamesGS ;
  lesalignements  := Falignement;
  lestitres       := Ftitre;
  leslargeurs     := Flargeur;

  //Mise en forme des colonnes
  for indice := 0 to Nbcols -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    //
    alignement  := ReadTokenSt(lesalignements);
    lalargeur   := readtokenst(leslargeurs);
    letitre     := readtokenst(lestitres);
    //
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
    //
    GS.cells[Indice,0]     := leTitre;
    GS.ColNames [Indice]   := Nam;
    //
    if      nam = 'CC_CODE'           then
    begin
      GS.ColFormats[Indice]:='UPPER';
      GS.ColLengths[Indice]:=3;
      ColCode   := Indice;
    end
    else if nam = 'CC_LIBELLE'        then
    begin
      GS.ColLengths[Indice]:=35;
      ColLibelle:= Indice;
    end
    else if nam = 'CC_ABREGE'         then
    begin
      GS.ColLengths[Indice]:=17;
      ColAbrege := Indice;
    end;
    //
    //Alignement des cellules
    if copy(Alignement,1,1)='G'       then //Cadré à Gauche
      GS.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D'  then  //Cadré à Droite
      GS.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C'  then
      GS.ColAligns[indice] := taCenter; //Cadré au centre

    //Colonne visible ou non
    if OkVisu then
  		GS.ColWidths[indice] := strtoint(lalargeur)*GS.Canvas.TextWidth('W')
    else
    	GS.ColWidths[indice] := -1;


    //Affichage d'une image ou du texte
    okImg := (copy(Alignement,8,1)='X');
    if (OkLib) or (okImg) then
    begin
    	GS.ColFormats[indice] := 'CB=' + Get_Join(Nam);
      if OkImg then
      begin
      	GS.ColDrawingModes[Indice]:= 'IMAGE';
      end;
    end;

    if (Dec<>0) or (Sep) then
    begin
    	if OkNulle then
        GS.ColFormats[indice] := FF+';; ;' //'#'
      else
      	GS.ColFormats[indice] := FF; //'#';
    end;

  end ;

  TailleEcran := Ecran.Width;

  PFAMILLE2.Width := TailleEcran/2;
  PFAMILLE3.Width := PFAMILLE2.Width;

end;

procedure TOF_BTFamHier.GridKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_DELETE: if (Shift = [ssCtrl]) then
    begin
      //Suppression de la ligne
      Key := 0;
    end;
    VK_INSERT : if (Shift = []) then
    begin
      //Insertion d'une ligne
      Key := 0;
    end;
    VK_RETURN : if (Shift = []) then
    begin
    	Key := 0;
      SendMessage(THedit(Sender).Handle, WM_KeyDown, VK_TAB, 0);
    end;
    VK_ESCAPE:
    BEGIN
    	Key := 0;
      //sortie du programme
    END;
  end;
end;


procedure TOF_BTFamHier.AfficheLagrille (GS : THGRID; TOBGrille : TOB);
var indice : integer;
begin

  if TOBGrille.detail.count = 0 then
  begin
		GS.RowCount := TOBGrille.detail.count + 2;
    GS.row := 1;
    GS.Cells [1,GS.row] := '';
    GS.Cells [2,GS.Row] := '';
    GS.Cells [3,GS.row] := '';
  end
  else
		GS.RowCount := TOBGrille.detail.count + 1;

  GS.DoubleBuffered := true;
  GS.BeginUpdate;

  TRY
    GS.SynEnabled := false;
    for Indice := 0 to TOBGrille.detail.count -1 do
    begin
      GS.row := Indice+1;
      GS.Cells [1,GS.row] := TOBGrille.Detail[Indice].GetString('CC_CODE');
      GS.Cells [2,GS.Row] := TOBGrille.Detail[Indice].GetString('CC_LIBELLE');
      GS.Cells [3,GS.row] := TOBGrille.Detail[Indice].GetString('CC_ABREGE');
    end;
  FINALLY
    GS.SynEnabled := true;
    GS.EndUpdate;
  END;

  TFVierge(ecran).HMTrad.ResizeGridColumns(GS);

end;

procedure TOF_BTFamHier.ControleChamp(Champ, Valeur: String);
begin

  if Champ='ACTION' then
  begin
    if Valeur='CREATION'          then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end;

  if champ ='NIV1'  then Niveau1  := Valeur;

  if champ ='NIV2'  then Niveau2  := Valeur;

  if champ ='NIV3'  then Niveau3  := Valeur;

  if champ='TYPE'   then TypeAct  := Valeur;

  if champ='FAMNV1' then Famille1 := Valeur;

  if champ='FAMNV2' then Famille2 := Valeur;

end;

procedure TOF_BTFamHier.ChargeFamilleNiv1;
BEGIN

  TobFam1.ClearDetail;

  if Niveau1  = '' then Exit;
  if Famille1 = '' then Exit;

  StSQL := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "'+Niveau1+'" AND CC_CODE="' + Famille1 + '"';
  QQ:=OpenSql(StSQL, True) ;

  if not QQ.Eof then TOBFAM1.LoadDetailDB('FAMNIV1','','',QQ,false,true) ;

  Ferme(QQ) ;

end;

procedure TOF_BTFamHier.ChargeFamilleNiv2;
BEGIN

  TobFam2.ClearDetail;

  if Niveau2  = '' then Exit;
  if Famille1 = '' then Exit;

  StSQL := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "'+Niveau2+'" AND CC_LIBRE like "%'+ Famille1 +'%"';
  QQ:=OpenSql(StSQL, True) ;

  if not QQ.Eof then TOBFAM2.LoadDetailDB('CHOIXCOD','','',QQ,false,true) ;

  Ferme(QQ) ;

end;

procedure TOF_BTFamHier.ChargeFamilleNiv3;
BEGIN

  TobFam3.ClearDetail;

  if Niveau3  = '' then Exit;
  if Famille2 = '' then Exit;

  StSQL := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "'+Niveau3+'" AND CC_LIBRE like "%'+ Famille1 + Famille2 +'%"';
  QQ:=OpenSql(StSQL, True) ;

  if not QQ.Eof then TOBFAM3.LoadDetailDB('CHOIXCOD','','',QQ,false,true) ;

  Ferme(QQ) ;

end;

procedure TOF_BTFAMHIER.DeleteLigneGFAM3(Sender : Tobject);
Var TOBL  : TOB;
    StSQL : String;
    Libre : string;
begin

  if Action = taConsult then Exit;

  TOBL := TOBFAM3.FindFirst(['CC_TYPE','CC_CODE'],[Niveau3,Famille3], False);
  If TOBL = nil then exit;

  //On Contrôle si cette sous- famille n'est pas sur un article
  StSQL := 'SELECT GA_ARTICLE FROM ARTICLE WHERE GA_FAMILLENIV1="' + Famille1 + '" AND GA_FAMILLENIV2="' + Famille2 + '" AND GA_FAMILLENIV3="' + Famille3 + '"';
  If ExisteSQL(StSQL) then
  begin
    PGIError('Suppression de ' + LibFam3.Caption + ' :' + Famille3 + ' impossible. Elle se trouve sur un article !', 'Suppression');
    Exit;
  end;

  //on Contrôle si cette sous- famille n'est pas sur une ligne document
  StSQL := 'SELECT GL_NUMERO FROM LIGNE WHERE GL_FAMILLENIV1="' + Famille1 + '" AND GL_FAMILLENIV2="' + Famille2 + '" AND GL_FAMILLENIV3="' + Famille3 + '"';
  If ExisteSQL(StSQL) then
  begin
    PGIError('Suppression de ' + LibFam3.Caption + ' :' + Famille3 + ' impossible. Elle se trouve sur un document !', 'Suppression');
    Exit;
  end;

  //Suppression de la ligne
  if PGIAsk('Confirmez-vous la suppression de ' + LibFam3.Caption + ' :' + Famille3 + '?','Suppression')=MrNO  then Exit;

  GSFAM3.CacheEdit;
  GSFAM3.SynEnabled := False;

  Famille3 := GSFAM3.Cells[ColCode, GSFAM3.Row];

  Libre := TOBL.GetString('CC_LIBRE');

  //si la zone Libre comporte un ; alors nous avons une famille 3 sur plusieurs famille 2
  if Pos(';',Libre)<>0 then
  begin
    if PGIAsk('Attention : ' + LibFam3.Caption + ' est associées à plusieurs ' + LibFam2.caption+ '. Voulez-vous ne supprimer que ' + Famille2 + ' ?', 'Suppression')=mryes then
    begin
      //on ne supprime que famille1+Famill2 dans la zone CC_LIBRE
      //on vérifie si c'est la première Famille/sous-famille
      if Pos(Famille1+Famille2,Libre) < Pos(';',Libre) then
        Libre := AnsiReplaceStr(libre, Famille1 + Famille2 + ';', '')
      else
        Libre := AnsiReplaceStr(libre, ';' + Famille1 + Famille2, '');
      TOBL.PutValue('CC_Libre', Libre);
      TOBL.UpdateDB(false, False, '');
      //suppression des sous-famille niuveau 3
      GSFAM3.DeleteRow(GSFAM3.Row);
      //
      FreeAndNil(TOBL);
      //
      AfficheLagrille(GSFAM3, TOBFAM3);
      GSFAM3.SynEnabled := True;
      Exit;
    end;
  end;

  StSQL := 'DELETE FROM CHOIXCOD WHERE CC_TYPE ="' + Niveau3 + '" AND CC_CODE="' + Famille3 + '"';
  if ExecuteSQL(StSQL)>0 then
  begin
    GSFAM3.DeleteRow(GSFAM3.Row);
    FreeAndNil(TOBL);
  end
  else
    PGIError('Suppression Sous-Famille ' + Famille3 + ' Impossible');

  AfficheLagrille(GSFAM3, TOBFAM3);

  GSFAM3.SynEnabled := True;

end;

procedure TOF_BTFAMHIER.NewLigneGFAM3(Sender : Tobject);
Var Cancel    : Boolean;
    Arow      : Integer;
begin

  Cancel := False;

  EnableDisable(False, GSFAM2);

  If TOBFAM3.detail.Count = 0 then
    ARow := TOBFAM3.detail.Count+2
  else
    ARow := TOBFAM3.detail.Count+1;

  Famille2 := GSFAM2.Cells [Colcode, GSFAM2.Row];

  NewLigne(GSFAM3, TOBFAM3, Niveau3, Famille1+Famille2);

  Fam3CellEnter(Self, ColCode, Arow, Cancel);

end;

Procedure TOF_BTFAMHIER.NewLigne(GS : THGRID; TOBFAM : TOB; Niveau, Libre : String);
Var TOBL  : TOB;
begin

  if Action = taConsult then Exit;

  TypeAction := 'C';

  GS.CacheEdit;
  GS.SynEnabled := False;

  TOBL := TOB.Create ('CHOIXCOD',TOBFAM,-1);

  TOBL.AddChampSupValeur('CC_TYPE',Niveau);
  TOBL.AddChampSupValeur('CC_CODE','');
  TOBL.AddChampSupValeur('CC_LIBELLE','');
  TOBL.AddChampSupValeur('CC_ABREGE','');
  TOBL.AddChampSupValeur('CC_LIBRE',Libre);

  GS.Options := GS.Options + [goEditing] - [goRowSelect];

  AfficheLagrille(GS, TOBFAM);

  GS.SynEnabled := True;

end;

procedure TOF_BTFAMHIER.EnableDisable(OK_Zone : Boolean; GS : THGrid);
Begin

  GS.Enabled        := OK_Zone;

  BInser1.Enabled   := OK_Zone;
  BDelete1.Enabled  := OK_Zone;
  //
  BInser2.Enabled   := OK_Zone;
  BDelete2.Enabled  := OK_Zone;
end;

Initialization
  registerclasses ( [ TOF_BTFamHier ] ) ;
end.

