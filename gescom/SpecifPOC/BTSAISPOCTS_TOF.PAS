{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/12/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAISPOCTS ()
Mots clefs ... : TOF;BTSAISPOCTS
*****************************************************************}
Unit BTSAISPOCTS_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Messages,
{$IFNDEF EAGLCLIENT}
     db,
     uDbxDataSet,
     mul,
     Fe_Main,
{$else}
     eMul, 
{$ENDIF}
     Vierge,
     uTob,
     HTB97, 
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     Hpanel,
     UTOF ;

const
  TheChamps = 'BLT_NUMDETAIL;BLT_CODEARTICLE;BLT_LIBELLE;BLT_QTEFACT;BLT_PUHTDEV;BLT_MONTANTHTDEV;';

Type
  TOF_BTSAISPOCTS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    GS : THGrid;
    TOBPiece : TOB;
    TOBDET,TheRecord,TOBPERE : TOB;
    LastDetail : Integer;
    stCell : string;
    SuiteWhere : string;
    procedure GSElipsisClick(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DefinieGrilleSaisie;
    procedure AfficheLaGrille;
    function AjouteDetail: TOB;
    procedure Reindice;
    function IsExitInterdit(Acol, Arow: integer): boolean;
    function  TraiteCodeArt(ACol, Arow: Integer; var Cancel: boolean) : boolean;
    function FindArticle(CodeArticle: string; TOBA: TOB): Boolean;
    procedure SetArticleToLigne(TOBL,TOBA : TOB);
    procedure CalculeLaLigne(TOBL : TOB);
    procedure AjouteLaLigne(Montant: double; Sens: string='+');
    procedure TraiteLibelle(Acol, ARow: Integer; Cancel: boolean);
    procedure TraiteQte(Acol, ARow: Integer; Cancel: boolean);
    procedure TraitePu(Acol, ARow: Integer; Cancel: boolean);
    function FindCodeArticle(CodeArticle: string; TOBA: TOB): Boolean;
    procedure AffichelaLigne(ARow : integer);
    function IsLigneComplete(Arow: Integer): Boolean;
    procedure NettoieLignesVide;
    procedure BdeleteClick (Sender : TObject);
  end ;

Implementation
uses FactUtil,UtilArticle,Windows,AglInit,FactTOB;


procedure TOF_BTSAISPOCTS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISPOCTS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISPOCTS.NettoieLignesVide;
var II : Integer;
begin
  for II := TOBdet.detail.count -1 downto 0 do
  begin
    if TOBDET.detail[II].GetString('BLT_CODEARTICLE')='' then TOBDET.detail[II].free;
  end;
end;

procedure TOF_BTSAISPOCTS.OnUpdate ;
var Indice : integer;
begin
  Inherited ;
  TOBDET.GetEcran(Ecran,THPanel(GetCOntrol('PENT')));
  if TOBDet.GetString('BLE_REFERENCETS')='' then
  begin
    PGIInfo('Vous devez renseigner la référence du TS/Avenant');
    Ecran.ModalResult := 0;
    Exit;
  end;
  if TOBDet.GetDateTime('BLE_DATE')=IDate1900 then
  begin
    PGIInfo('Vous devez renseigner une date pour le TS/Avenant');
    Ecran.ModalResult := 0;
    Exit;
  end;
  if TFVierge(Ecran).TypeAction = TaCreat then
  begin
    if TOBPERE.FindFirst(['BLE_REFERENCETS'],[TOBDET.getString('BLE_REFERENCETS')],true) <> nil then
    begin
      PGIInfo('Cette référence existe déjà.');
      Ecran.ModalResult := 0;
      Exit;
    end;
  end;
  NettoieLignesVide;
  for Indice := 0 to TOBDET.detail.count -1 do
  begin
    TOBDET.detail[Indice].SetString('BLT_REFERENCETS',TOBDet.GetString('BLE_REFERENCETS'));
  end;
  TheRecord.ClearDetail; TheRecord.Dupliquer(TOBDET,True,True,true); TheRecord.SetString('OK','X');
end ;

procedure TOF_BTSAISPOCTS.OnLoad ;
begin
  Inherited ;
end ;

procedure  TOF_BTSAISPOCTS.Reindice;
var II : Integer;
    Lig : integer;
begin
  for II := 0 to TOBdet.detail.count -1 do
  begin
    Lig := II +1;
    TOBDet.detail[II].SetInteger('BLT_NUMDETAIL',Lig);
    AffichelaLigne(Lig);
    LastDetail := Lig;
  end;
end;

function TOF_BTSAISPOCTS.AjouteDetail : TOB;
begin
  Inc(LastDetail);
  result := TOB.Create ('BLIGNETS',TOBDET,-1);
  result.AddChampSupValeur('NEW','X');
  result.SetInteger('BLT_NUMORDRE',TOBDET.GetInteger('BLE_NUMORDRE'));
  result.SetInteger('BLT_NUMDETAIL',LastDetail);
end;

procedure TOF_BTSAISPOCTS.OnArgument (S : String ) ;
var cancel : Boolean;
    Acol,Arow : Integer;
    critere,ST : string;
    X : Integer;
    ChampMul,ValMul : string;  
begin
  Inherited ;
  St := S;
  repeat
    Critere := Trim(ReadTokenSt(st));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'ACTION' then
        begin
          if ValMul = 'MODIFICATION' then TFVierge(Ecran).TypeAction := taModif
                                     else TFVierge(Ecran).TypeAction := taCreat;
        end;
      end;
    end;
  until Critere = '';

  TheRecord := LaTOB;
  TOBPERE := TOB(LaTOB.data);
  TOBPiece := TOB(TOBpere.data);
  cancel := false;
  SuiteWhere := FabricWhereNatArt(TOBPiece.getString('GP_NATUREPIECEG'), TOBPiece.getString('GP_DOMAINE'),'');
  TOBDET := TOB.Create ('BLIGNEETS',nil,-1); TOBDET.AddChampSupValeur('OK','-'); 
  TOBDET.dupliquer (LaTOB,True,True,true);
  TOBDET.PutEcran(Ecran,THPanel(GetCOntrol('PENT')));
  //
  lastDetail := 0;
  //
  if TOBDET.detail.count > 0 then LastDetail := TOBDET.Detail[TOBDET.detail.count-1].GetInteger('BLT_NUMDETAIL');
  //
  GS := THGrid(Getcontrol('GS'));
  DefinieGrilleSaisie;
  AfficheLaGrille;
  //
  if TFVierge(Ecran).TypeAction <> TaCreat then
  begin
    THPanel(GetControl('PENT')).Enabled := false;
  end;
  //
  GS.Row := 1; GS.col := 1; Acol := GS.Col; Arow := GS.Row;
  GS.OnRowEnter := GSRowEnter;
  GS.OnRowExit := GSRowExit;
  GS.OnCellEnter := GSCellEnter;
  GS.OnCellExit := GSCellExit;
  GS.OnElipsisClick := GSElipsisClick;
  GS.OnKeyDown := GSKeyDown;
  TToolbarButton97(GetControl('Bdelete')).OnClick := BdeleteClick;
  GSRowEnter(Self,GS.row,cancel,false);
  GSCellEnter(Self,Acol,Arow,cancel);
end ;

procedure TOF_BTSAISPOCTS.OnClose ;
begin
  TOBDet.free;
  Inherited ;
end ;

procedure TOF_BTSAISPOCTS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISPOCTS.OnCancel () ;
begin
  Inherited ;
end;

function TOF_BTSAISPOCTS.IsExitInterdit(Acol,Arow : integer): boolean;
begin
  result := false;
  if (acol=0) or (Arow=0) then exit;
  if (Acol= 1) and (GS.cells[Acol,Arow]='') and (GS.row >= Arow) then
  begin
    result := true;
    exit;
  end;
  if (GS.row > Arow) and (GS.Cells[GS.col,GS.row]='') and (GS.col<>1) then
  begin
    result := true;
  end;
end;

procedure TOF_BTSAISPOCTS.GSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  GS.ElipsisButton := (GS.Col = 1);
  stCell := GS.Cells[GS.Col,GS.Row];
  GS.RowCount := TOBDET.Detail.count +2;
end;

procedure TOF_BTSAISPOCTS.GSCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var OldMt : double;
    TOBL : TOB;
begin
  Cancel := IsExitInterdit(Acol,Arow);
  if cancel then exit;
  if GS.Cells[ACol,Arow] = stCell then exit;
  TOBL := TOBDET.detail[Arow-1];
  OldMt := TOBL.GetDouble('BLT_MONTANTHTDEV');
  if Acol = 1 then // Code Article
  begin
    TraiteCodeArt (ACol,Arow,Cancel);
  end else if Acol = 2 then
  begin
    TraiteLibelle (Acol,ARow,Cancel);
  end else if Acol = 3 then
  begin
    TraiteQte (Acol,ARow,Cancel);
  end else if Acol = 4 then
  begin
    TraitePu (Acol,ARow,Cancel);
  end;

  if (not Cancel) and (ACol <> 2) then
  begin
    AjouteLaLigne (OldMt,'-');
    CalculeLaLigne(TOBL);
    AjouteLaLigne (TOBL.GetDouble('BLT_MONTANTHTDEV'));
  end;
  AffichelaLigne(ARow);
end;

procedure TOF_BTSAISPOCTS.GSElipsisClick(Sender: TObject);
var RefArticle,StChamps : string;
    found : boolean;
    TOBA,TOBL : TOB;
begin
  if GS.col = 1 then
  begin
    TOBL := TOBDET.Detail[GS.row-1];  if TOBL.GetString('BLT_ARTICLE')<> '' then Exit;
    TOBA := TOB.Create ('ARTICLE',nil,-1);
    TRY
      if GS.Cells[GS.col,GS.Row] <> '' then StChamps := 'GA_CODEARTICLE=' + Trim(Copy(GS.Cells[GS.col,GS.Row], 1, 18));
      if SuiteWhere <> '' then
      begin
        if StChamps <> '' then StChamps := StChamps + ';XX_WHERE= AND (' + SuiteWhere+')'
                          else StChamps := 'XX_WHERE= AND (' + SuiteWhere+')';
      end;
      RefArticle := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps+';RECHERCHEARTICLE');
      if RefArticle <> '' then
      begin
        found := FindArticle(RefArticle,TOBA);
        if found then
        begin
          GS.Cells[GS.col,GS.row] := TOBA.GetString('GA_CODEARTICLE');
          SetArticleToLigne(TOBL,TOBA);
          CalculeLaLigne(TOBL);
          AjouteLaLigne (TOBL.GetDouble('BLT_MONTANTHTDEV'));
          AfficheLaLIgne(GS.row);
        end;
      end;
    FINALLY
      TOBA.free;
    END;
  end;
end;

function TOF_BTSAISPOCTS.IsLigneComplete (Arow : Integer) : Boolean;
var TOBL : TOB;
begin
  Result := false;
  if Arow <= 1  then begin Result := True; Exit; end;
  if Arow > TOBDET.detail.count +1 Then Exit;
  TOBL := TOBDET.detail[Arow-1];
  if TOBL.GetString('BLT_ARTICLE')<> '' then Result := True;
end;

procedure TOF_BTSAISPOCTS.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
  //if Ou > TOBDET.detail.count + 1 then BEGIN cancel := true; exit; END;
  if Ou >= TOBDET.Detail.count +1  then
  begin
    If (IsLigneComplete(Ou-1)) then AjouteDetail else Cancel := true;
  end;
end;

procedure TOF_BTSAISPOCTS.GSRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
end;

procedure TOF_BTSAISPOCTS.DefinieGrilleSaisie;
var FFM,FFQ,FFP : string;
    II : Integer;
begin
  FFM := '#';
  if V_PGI.OkDecV > 0 then
  begin
    FFM := '0.';
    for II := 1 to V_PGI.OkDecV - 1 do
    begin
      FFM := FFM + '#';
    end;
    FFM := FFM + '0';
  end;
  FFP := '#';
  if V_PGI.OkDecP > 0 then
  begin
    FFP := '0.';
    for II := 1 to V_PGI.OkDecP - 1 do
    begin
      FFP := FFP + '#';
    end;
    FFP := FFP + '0';
  end;
  FFQ := '#';
  if V_PGI.OkDecQ > 0 then
  begin
    FFQ := '0.';
    for II := 1 to V_PGI.OkDecQ - 1 do
    begin
      FFQ := FFQ + '#';
    end;
    FFQ := FFQ + '0';
  end;

  GS.VidePile(false);
  //
  GS.ColCount := 6;
  GS.Cells[0,0] := '';
  GS.ColWidths[0] := 20;
  GS.ColAligns[0] := taRightJustify;
  //
  GS.Cells[1,0] := 'Code';
  GS.ColWidths[1] := 30;
  GS.ColLengths[1] := 18;
  GS.ColAligns[1] := taLeftJustify;
  //
  GS.Cells[2,0] := 'Libellé';
  GS.ColWidths[2] := 70;
  GS.ColLengths[2] := 70;
  GS.ColAligns[2] := taLeftJustify;
  //
  GS.Cells[3,0] := 'Qté';
  GS.ColWidths[3] := 30;
  GS.ColLengths[3] := 14;
  GS.ColAligns[3] := taRightJustify;
  GS.ColTypes [3] := 'R';
  GS.ColFormats[3] := FFQ+';'+FFQ+'; ;';
  //
  GS.Cells[4,0] := 'Prix Unitaire';
  GS.ColWidths[4] := 30;
  GS.ColLengths[4] := 14;
  GS.ColTypes [4] := 'R';
  GS.ColAligns[4] := taRightJustify;
  GS.ColFormats[4] := FFP+';'+FFP+'; ;';
  //
  GS.Cells[5,0] := 'Montant';
  GS.ColWidths[5] := 40;
  GS.ColLengths[5] := 16;
  GS.ColTypes [5] := 'R';
  GS.ColAligns[5] := taRightJustify;
  GS.ColFormats[5] := FFM+';'+FFM+'; ;';
end;

procedure TOF_BTSAISPOCTS.AfficheLaGrille;
var II : Integer;
begin
  GS.RowCount := TOBDET.Detail.count +2;  if GS.rowCount = 2 then Gs.rowCount := Gs.rowCount +1;
  for II := 0 to TOBDET.Detail.count -1 do
  begin
    AffichelaLigne(II+1);
  end;
  TFVierge(Ecran).HMTrad.ResizeGridColumns(GS);
end;

function TOF_BTSAISPOCTS.TraiteCodeArt(ACol, Arow: Integer; var Cancel: boolean) : boolean;
var TOBL,TOBA : TOB;
    RefArticle,StChamps : string;
    found : boolean;
begin
  Result := true;
  if Arow > TOBDET.detail.Count then BEGIN cancel := True; exit; END;
  TOBL := TOBDET.Detail[Arow-1]; 
  if TOBL.GetString('BLT_ARTICLE')<> '' then
  begin
    GS.cells[Acol,Arow] := StCell;
    Exit;
  end;
  
  TOBA := TOB.Create ('ARTICLE',nil,-1);
  TRY
    RefArticle := GS.Cells[Acol,Arow];
    found := FindCodeArticle(RefArticle,TOBA);

    if not Found then
    begin
      if RefArticle <> '' then StChamps := 'GA_CODEARTICLE=' + Trim(Copy(RefArticle, 1, 18));
      //
      RefArticle := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps+';RECHERCHEARTICLE');
      if RefArticle <> '' then
      begin
        found := FindArticle(RefArticle,TOBA);
        if found then GS.Cells[Acol,Arow] := TOBA.GetString('GA_CODEARTICLE');
      end;
    end;
    if not Found then BEGIN cancel := True; Exit; end;
    SetArticleToLigne(TOBL,TOBA);
  FINALLY
    TOBA.free;
  END;
end;

procedure TOF_BTSAISPOCTS.TraiteLibelle (Acol,ARow : Integer; Cancel :boolean);
var TOBL : TOB;
begin
  if (Acol=0) or (Arow=0) then exit;
  TOBL := TOBDet.detail[Arow-1];
  TOBL.SetString('BLT_LIBELLE',Copy(GS.cells[Acol,Arow],1,70));
end;

procedure TOF_BTSAISPOCTS.TraitePu (Acol,ARow: Integer; Cancel :boolean);
var TOBL : TOB;
begin
  if (Acol=0) or (Arow=0) then exit;
  TOBL := TOBDet.detail[Arow-1];
  if Valeur(GS.cells[Acol,Arow]) = 0 then
  begin
    Cancel := true;
    PgiInfo('Vous ne pouvez indiquer une valorisation nulle');
    GS.cells[Acol,Arow] := stcell;
  end;
  TOBL.SetDouble('BLT_PUHTDEV',VALEUR(GS.cells[Acol,Arow]));
end;

procedure TOF_BTSAISPOCTS.TraiteQte (Acol,ARow: Integer; Cancel :boolean);
var TOBL : TOB;
begin
  if (Acol=0) or (Arow=0) then exit;
  TOBL := TOBDet.detail[Arow-1];
  if Valeur(GS.cells[Acol,Arow]) = 0 then
  begin
    Cancel := true;
    PgiInfo('Vous ne pouvez indiquer une quantité nulle');
    GS.cells[Acol,Arow] := stcell;
  end;
  TOBL.SetDouble('BLT_QTEFACT',VALEUR(GS.cells[Acol,Arow]));
end;

function TOF_BTSAISPOCTS.FindArticle(CodeArticle: string; TOBA : TOB): Boolean;
var QQ : TQuery;
begin
  QQ := OpenSQL('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+CodeArticle+'" AND GA_STATUTART IN ("GEN","UNI")',True,1,'',true);
  Result := not QQ.Eof;
  if not QQ.eof then
  begin
    TOBA.SelectDB('',QQ);
  end;
  Ferme(QQ);
end;

function TOF_BTSAISPOCTS.FindCodeArticle(CodeArticle: string; TOBA : TOB): Boolean;
var QQ : TQuery;
begin
  QQ := OpenSQL('SELECT * FROM ARTICLE WHERE GA_CODEARTICLE="'+CodeArticle+'" AND GA_STATUTART IN ("GEN","UNI")',True,1,'',true);
  Result := not QQ.Eof;
  if not QQ.eof then
  begin
    TOBA.SelectDB('',QQ);
  end;
  Ferme(QQ);
end;


procedure TOF_BTSAISPOCTS.SetArticleToLigne(TOBL, TOBA: TOB);
begin
  InitValoArtNomen (TOBA,'VEN');
  CalculePrixArticle (TOBA,'');
  TOBL.SetString('BLT_ARTICLE',TOBA.GetString('GA_ARTICLE'));
  TOBL.SetString('BLT_CODEARTICLE',TOBA.GetString('GA_CODEARTICLE'));
  TOBL.SetString('BLT_LIBELLE',TOBA.GetString('GA_LIBELLE'));
  TOBL.SetDouble('BLT_QTEFACT',1);
  TOBL.SetDouble('BLT_DPA',TOBA.GetDouble('GA_DPA'));
  TOBL.SetDouble('BLT_DPR',TOBA.GetDouble('GA_DPR'));
  TOBL.SetDouble('BLT_PUHTDEV',TOBA.GetDouble('GA_PVHT'));
end;

procedure TOF_BTSAISPOCTS.CalculeLaLigne(TOBL: TOB);
begin
  TOBL.SetDouble('BLT_MONTANTPA',ARRONDI(TOBL.GetDouble('BLT_QTEFACT')*TOBL.GetDouble('BLT_DPA'),V_PGI.OkDecV));
  TOBL.SetDouble('BLT_MONTANTPR',ARRONDI(TOBL.GetDouble('BLT_QTEFACT')*TOBL.GetDouble('BLT_DPR'),V_PGI.OkDecV));
  TOBL.SetDouble('BLT_MONTANTHTDEV',ARRONDI(TOBL.GetDouble('BLT_QTEFACT')*TOBL.GetDouble('BLT_PUHTDEV'),V_PGI.OkDecV));
end;

procedure TOF_BTSAISPOCTS.AjouteLaLigne(Montant : double; Sens: string);
begin
  if Sens = '+' then
  begin
    TOBDET.SetDouble('BLE_MONTANT',ARRONDI(TOBDET.GetDouble('BLE_MONTANT')+Montant,V_PGI.OkDecV));
  end else
  begin
    TOBDET.SetDouble('BLE_MONTANT',ARRONDI(TOBDET.GetDouble('BLE_MONTANT')-Montant,V_PGI.OkDecV));
  end;
  THNumEdit(GetControl('BLE_MONTANT')).Value := TOBDET.GetDouble('BLE_MONTANT');
end;

procedure TOF_BTSAISPOCTS.AffichelaLigne(ARow: integer);
begin
  TOBDET.Detail [Arow-1].PutLigneGrid(GS,Arow,false,False,TheChamps);
end;

procedure TOF_BTSAISPOCTS.BdeleteClick(Sender: TObject);
var TOBL : TOB;
    cancel : boolean;
    Arow,Acol : Integer;
begin
  if GS.row > TOBDET.detail.count Then Exit;
  if PGIAsk('Désire-vous réellement supprimer cette ligne') <> mrYes then Exit;
  TOBL := TOBDET.detail[GS.row-1];
  AjouteLaLigne(TOBL.GetDouble('BLT_MONTANTHTDEV'),'-');
  TOBL.free;
  GS.DeleteRow(GS.row);
  GS.SynEnabled := false;
  Reindice;
  GS.SynEnabled := True;
  StCell := GS.cells[GS.col,GS.Row];
  Arow := GS.row;
  Acol := GS.col;
  GSRowEnter(Self,ARow,Cancel,false);
  GSCellEnter(Self,Acol,Arow,cancel);
  StCell := GS.cells[GS.col,GS.Row];
end;

procedure TOF_BTSAISPOCTS.GSKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
begin
	if (Shift = []) and (Key = VK_RETURN)  then
  begin
    SendMessage(GS.Handle, WM_KeyDown, VK_TAB, 0);
    Key := 0;
  end;
	if (Shift = []) and (Key = VK_F5)  then
  begin
    Key := 0;
    GSElipsisClick(self);
  end;

end;

Initialization
  registerclasses ( [ TOF_BTSAISPOCTS ] ) ;
end.
