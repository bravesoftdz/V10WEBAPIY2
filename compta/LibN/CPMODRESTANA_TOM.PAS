{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 29/12/2005             
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CMODELRESTANA (CMODELRESTANA)
Mots clefs ... : TOM;CMODELRESTANA
*****************************************************************}
Unit CPMODRESTANA_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     AglInit,    //ActionToString
     Windows,
     Messages,
     Graphics,
     Grids,
     ParamSoc, //GetParamSoc
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,
     Fe_main,
{$else}
     eFiche,
     eFichList,
     Maineagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     UTob,
     Ent1,
     HTB97,                //TToolBarButton97
     ULibWindows,          //TestJoker
     Math,                 //Ceil
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     ULibAnalytique        //RecherchePremDerAxeVentil
     ;

  procedure CPLanceFiche_CPMODRESTANA_TOM(Axe, Modele: String; Comment: TActionFiche);
Type
  TOM_CMODELRESTANA = Class (TOM)
  private
    PCAxe:          TPageControl;
    TsAxes:         array[0..MaxAxe-1] of TTabSheet;
    GdComptes:      array[0..MaxAxe-1] of THGrid;
    Axe:            THValComboBox;
    Modele:         THEdit;
    Libelle:        THEdit;
    GdInterval:     THGrid;

    MsgBox:         THMsgBox;
    FAxeCrois:      HTStrings;
    FOnBValiderClick: TNotifyEvent;
    FOnBInsertClick:  TNotifyEvent;

    procedure InitGridInterval(InitRowCount: Boolean);
    procedure InitAxePageControl;
    procedure ChargeGridInterval(Axe, Modele: String);
    procedure InitGridCompte;
    procedure ChargeGridCompte;
    procedure NewLigne(ARow: Integer);
    function  QuickSearch(Grid: THGrid; ACol: Integer; Key: String; FirstRow, LastRow: Integer): Integer;
    procedure LocateCompte(Cpt: String);
    procedure DeleteRecordDetail(Axe, Modele: String; UpdateMode: Boolean);
    procedure ForceModeEdition;

    procedure AxeChange(Sender: TObject);
    procedure PCAxeChange(Sender: TObject);
    procedure GridIntervalKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GridIntervalKeyPress(Sender: TObject; var Key: Char);
    procedure GridIntervalSetEditText(Sender: TObject; ACol, ARow: Integer; const Value: HString);
    procedure GridIntervalDrawCell(Sender: TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState);
    procedure GridIntervalTopLeftChanged(Sender: TObject);
    procedure GridIntervalSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
    procedure GridIntervalEnter(Sender : TObject);
    procedure GridCptDblClick(Sender: TObject);
    procedure GridCompteColumnWidthsChanged(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
  public
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  end ;


  {Retorune le premier  axe croisé ou Axe}
  function GetFirstAxeAna(Axe: String): Integer;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  SAISUTIL;   //AffecteGrid


function GetFirstAxeAna(Axe: String): Integer;
var
  i: Integer;
begin
  Result := 1;
  if VH^.AnaCroisaxe then
    for i := 1 to MaxAxe do
      begin
      if GetParamSocSecur('SO_VENTILA' + IntToStr(i), False) then
        begin
        Result := i;
        break;
        end;
      end
  else if Length(Axe) > 0 then
    Result := StrToInt(Axe[Length(Axe)]);
end;

procedure CPLanceFiche_CPMODRESTANA_TOM(Axe, Modele: String; Comment: TActionFiche);
var
  Argument: String;
begin
  Argument := '';
  case Comment of
    taCreat, taCreatEnSerie:
      Argument := 'ACTION=CREATION;'+TAToStr(Comment);
    taModif,taModifEnSerie:
      Argument := 'ACTION=MODIFICATION;'+TAToStr(Comment)+';'+Axe+';'+Modele;
    taConsult :
      Argument := 'ACTION=CONSULTATION;'+TAToStr(Comment)+';'+Axe+';'+Modele;
  end;
  if Argument <> '' then
    AGLLanceFiche('CP','CPMODRESTANA_FIC','','',Argument);
end;


procedure TOM_CMODELRESTANA.OnNewRecord ;
begin
  Inherited ;
  Axe.Value     := 'A'+FAxeCrois[0];
  {b FP 13/06/2006}          
  if VH^.AnaCroisaxe then
    begin
    Axe.ItemIndex := 0;
    SetField('CRA_AXE', Axe.Value);
    end;
  {e FP 13/06/2006}

  ChargeGridInterval('', '');
end ;

procedure TOM_CMODELRESTANA.OnDeleteRecord ;
  function EstUtilise(Axe: String; Modele: String): Boolean;
  begin
    Result := ExisteSQL('SELECT CLA_GENERAL FROM CLIENGENEMODELA'+
                        ' WHERE CLA_CODE'+Axe[2]+'="'+Modele+'"');
  end;
var
  Ok: Boolean;
begin
  Inherited ;

  Ok := True;
  if EstUtilise(Axe.Value, Modele.Text) then begin
    if PGIAsk(MsgBox.Mess[0], Ecran.Caption) <> mrYes then begin
      Ok := False;
    end;
  end;

  if Ok then
    DeleteRecordDetail(Axe.Value, Modele.Text, False)
  else begin
    LastError := 1;
    LastErrorMsg := '';
  end;

end ;

procedure TOM_CMODELRESTANA.OnUpdateRecord ;
var
  iRow:       Integer;
  iCol:       Integer;
  iAxe:       Integer;
  NbInsert:   Integer;
  Insert:     Boolean;
  ValueDeb:   String;
  ValueFin:   String;

  vTobDetail:      TOB;
  vTobDetailFille: TOB;
begin
  {b FP 19/04/2006 FQ17825}
  if Trim(Modele.Text) = '' then      //Le modèle n'a pas de nom
    begin
    SetFocusControl('CAD_CODE');
    LastError := 1;
    LastErrorMsg := MsgBox.Mess[5];     
    Exit;
    end;
  {e FP 18/04/2006}
  Inherited ;

  vTobDetail := TOB.Create('$CMODELRESTANADET', nil, -1);
  try
    NbInsert := 0;
    for iRow:=GdInterval.FixedRows to GdInterval.RowCount-2 do begin
      {La dernière ligne de la grille est tj vide}
      Insert := False;
      for iCol:=0 to GdInterval.ColCount-1 do
        if Trim(GdInterval.Cells[iCol, iRow]) <> '' then
          Insert := True;
      if Insert then begin
        vTobDetailFille := TOB.Create('CMODELRESTANADET', vTobDetail, NbInsert);
        vTobDetailFille.PutValue('CAD_AXE',   Axe.Value);
        vTobDetailFille.PutValue('CAD_CODE',  Modele.Text);
        vTobDetailFille.PutValue('CAD_LIGNE', iRow-GdInterval.FixedRows+1);
        iCol := 0;
        for iAxe:=0 to FAxeCrois.Count-1 do begin
          ValueDeb := Trim(GdInterval.Cells[iCol,   iRow]);
          ValueFin := Trim(GdInterval.Cells[iCol+1, iRow]);
          if TestJoker(ValueDeb) then begin
            ValueDeb := TraduitJoker(ValueDeb);
            ValueFin := '';
          end
          else if ValueFin = '' then
            ValueFin := ValueDeb;
          vTobDetailFille.PutValue('CAD_DEBAXE'+FAxeCrois[iAxe], ValueDeb);
          vTobDetailFille.PutValue('CAD_FINAXE'+FAxeCrois[iAxe], ValueFin);
          Inc(iCol, 2);
        end;
        Inc(NbInsert);
      end;
    end;
    //Enregistrement
    try
      if vTobDetail.Detail.Count = 0 then begin
        LastError := 1;
        LastErrorMsg := MsgBox.Mess[1];
      end
      else begin
        begintrans;
        DeleteRecordDetail(Axe.Value, Modele.Text, True);
        vTobDetail.InsertDB(nil);
        committrans;
      end;
    except
      on e : exception do begin
        rollback;
        PGIInfo(e.message,'Erreur');
      end;
    end;
  finally
    FreeAndNil(vTobDetail);
  end;
end ;

procedure TOM_CMODELRESTANA.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CMODELRESTANA.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_CMODELRESTANA.OnLoadRecord ;
begin
  Inherited ;
  ChargeGridInterval(Axe.Value, Modele.Text);
  ChargeGridCompte;
end ;

procedure TOM_CMODELRESTANA.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_CMODELRESTANA.OnArgument ( S: String ) ;
var
  i:        Integer;
  StArgs:   String;
  Action:   TActionFiche;
  StAxe:    String;
  StModele: String;
  BValider: TToolbarButton97;
  BInsert:  TToolbarButton97;
begin
  Inherited ;
  Ecran.HelpContext := 1461050;//Lek 040706 1470500;  {FP 19/04/2006 FQ17839}

  StArgs     := S;
  ReadToKenPipe(StArgs, ';');
  Action     := StrToTa(ReadToKenPipe(StArgs, ';'));
  StAxe      := ReadToKenPipe(StArgs, ';');
  StModele   := ReadToKenPipe(StArgs, ';');

  Axe        := THValComboBox(Getcontrol('CRA_AXE', true));
  Modele     := THEdit(Getcontrol('CRA_CODE', true));
  Libelle    := THEdit(Getcontrol('CRA_LIBELLE', true));
  GdInterval := THGrid(GetControl('GDINTERVAL', True));
  BValider   := TToolbarButton97(GetControl('BVALIDER', True));
  BInsert    := TToolbarButton97(GetControl('BINSERT', True));
  PCAxe      := TPageControl(GetControl('PCAXE', True));
  for i:=0 to MaxAxe-1 do begin
    TsAxes[i]    := TTabSheet(GetControl('TSAXE'+IntToStr(i+1), True));
    GdComptes[i] := THGrid(GetControl('GDCPT'+IntToStr(i+1), True));
  end;

  MsgBox := THMsgBox.Create(Ecran);
  MsgBox.Mess.Add('Ce modèle est rattaché à un ou plusieurs comptes généraux, voulez-vous le supprimer?');
  MsgBox.Mess.Add('Il faut saisir au moins un intervalle.');
  MsgBox.Mess.Add('La colonne Fin ne doit pas contenir de caractère joker.');
  MsgBox.Mess.Add('Le code section Début doit être inférieur à celui de la section Fin.');       {FP 19/04/2006 FQ17844}
  MsgBox.Mess.Add('La colonne Fin n''est pas saisissable.');
  MsgBox.Mess.Add('Vous devez renseigner: Modèle');     {5}  {FP 19/04/2006 FQ17825}

  FAxeCrois := HTStringList.Create;

  if VH^.AnaCroisaxe then begin
    for i := 1 to MaxAxe do begin
      if GetParamSocSecur('SO_VENTILA' + IntToStr(i), False) then
        FAxeCrois.Add(IntToStr(i));
    end;
  end
  else if StAxe='' then
    FAxeCrois.Add('1')
  else
    FAxeCrois.Add(StAxe[2]);
{
  FAxeCrois.Clear;
  FAxeCrois.Add('1');
  FAxeCrois.Add('2');
  FAxeCrois.Add('3');
}
  Axe.Visible := (FAxeCrois.Count = 1);
  SetControlVisible('TCRA_AXE', Axe.Visible);

  Axe.OnChange          := AxeChange;
  PCAxe.OnChange        := PCAxeChange;
  GdInterval.OnKeyDown  := GridIntervalKeyDown;
  GdInterval.OnKeyPress := GridIntervalKeyPress;
  GdInterval.OnDrawCell := GridIntervalDrawCell;
  GdInterval.OnTopLeftChanged := GridIntervalTopLeftChanged;
  GdInterval.OnSelectCell     := GridIntervalSelectCell;
  GdInterval.OnSetEditText    := GridIntervalSetEditText;
  GdInterval.DefaultDrawing   := False;
  GdInterval.Options          := [goVertLine, goHorzLine]+ GdInterval.Options;
  GdInterval.OnEnter    := GridIntervalEnter;
  FOnBValiderClick      := BValider.OnClick;
  BValider.OnClick      := BValiderClick;
  FOnBInsertClick       := BInsert.OnClick;
  BInsert.OnClick       := BInsertClick;

  TFFiche(Ecran).TypeAction := Action;
  TFFiche(Ecran).FLequel    := StAxe+';'+StModele;

  SetControlProperty('PCAXE'     , 'TABORDER', 0);
  SetControlProperty('GDINTERVAL', 'TABORDER', 0);
  SetControlProperty('CRA_EXCLU' , 'TABORDER', 0);
  SetControlProperty('GBMODELE'  , 'TABORDER', 0);

  InitAxePageControl;
  InitGridInterval(True);
  InitGridCompte;
  AffecteGrid(GdInterval, Action);
  Focuscontrole(GdInterval, true);
end ;

procedure TOM_CMODELRESTANA.OnClose ;
begin
  Inherited ;
  FAxeCrois.Free;
end ;

procedure TOM_CMODELRESTANA.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_CMODELRESTANA.ChargeGridInterval(Axe, Modele: String);
var
  i:           Integer;
  iCol:        Integer;
  iRow:        Integer;
  vTobDetail:  TOB;
  ColName:     String;
begin
  if Modele = '' then begin       {Nouveau Modèle}
    NewLigne(GdInterval.FixedRows);
    Exit;
  end;

  vTobDetail := TOB.Create('$CMODELRESTANADET', nil, -1);
  try
    ColName := 'CAD_AXE, CAD_CODE, CAD_LIGNE,';
      for i:=0 to FAxeCrois.Count-1 do begin
        if i <> 0 then
          ColName:=ColName+',';
        ColName:=ColName+'CAD_DEBAXE'+FAxeCrois[i]+',CAD_FINAXE'+FAxeCrois[i];
      end;
    vTobDetail.LoadDetailDBFromSQL('CMODELRESTANADET',
      'SELECT '+ColName+' FROM CMODELRESTANADET'+
      ' WHERE CAD_AXE  = "'+Axe+'"'+
      '   AND CAD_CODE = "'+Modele+'"'+
      ' ORDER BY CAD_LIGNE');

    if vTobDetail.Detail.Count = 0 then begin
      NewLigne(GdInterval.FixedRows);
    end
    else begin
      ColName := '';
      for i:=0 to FAxeCrois.Count-1 do begin
        if i <> 0 then
          ColName:=ColName+';';
        ColName:=ColName+'CAD_DEBAXE'+FAxeCrois[i]+';CAD_FINAXE'+FAxeCrois[i];
      end;
      vTobDetail.PutGridDetail(GdInterval, False, False, ColName, True);
      {Si la cellule Debut = cellule Fin, il faut effacer la cellule Fin}
      for iRow:=GdInterval.FixedRows to GdInterval.RowCount-1 do begin
        for iCol:=0 to GdInterval.ColCount-1 do
          if ((iCol mod 2)=1) and (GdInterval.Cells[iCol, iRow]=GdInterval.Cells[iCol-1, iRow]) then
            GdInterval.Cells[iCol, iRow] := '';
      end;
      NewLigne(GdInterval.RowCount);
    end;
  finally
    FreeAndNil(vTobDetail);
  end;
end;

procedure TOM_CMODELRESTANA.NewLigne(ARow: Integer);
var
  i: Integer;
begin
  if ARow < GdInterval.RowCount then
    GdInterval.RowCount := ARow+1
  else if ARow = GdInterval.RowCount then
    GdInterval.RowCount := GdInterval.RowCount+1;

  for i:=0 to GdInterval.ColCount-1 do
    GdInterval.Cells[i, ARow] := '';
end;

procedure TOM_CMODELRESTANA.GridIntervalKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key in [VK_ESCAPE]) then
    TFFiche(Ecran).BFerme.Click
  else if (not(Key in [VK_PRIOR,VK_NEXT,VK_UP,VK_DOWN,VK_LEFT,VK_RIGHT,VK_TAB])) then begin
    ForceModeEdition;
  end;
end;

procedure TOM_CMODELRESTANA.GridIntervalKeyPress(Sender: TObject; var Key: Char);
begin
  Key := UpCase(Key);
  if IsCharAlphaNumeric(Key) and (GdInterval.Row = GdInterval.RowCount-1) then  {Ajoute tj une nouvelle ligne}
    NewLigne(GdInterval.RowCount);
end;

procedure TOM_CMODELRESTANA.GridIntervalSetEditText(Sender: TObject; ACol, ARow: Integer; const Value: HString);
begin
  if (GdInterval.Col = ACol) and (GdInterval.Row = ARow) then
    LocateCompte(Value);
end;

procedure TOM_CMODELRESTANA.GridIntervalDrawCell(Sender: TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  Text:       array[0..255] of Char;
  F:          TAlignment ;
  Grille:     THgrid ;
  OldBrush:   TBrush ;
  OldPen:     TPen ;
  NewRec:     TRect;
begin
  OldBrush := TBrush.Create ;
  OldPen   := TPen.Create ;

  Grille := GdInterval ;
  StrPCopy(Text,Grille.Cells[Col,Row]);

  OldBrush.assign(Grille.Canvas.Brush);
  OldPen.assign(Grille.Canvas.Pen);

  NewRec := Rect;
  if (gdfixed in State) and (Row = 0) then begin
    {Le premier titre est sur 2 colonnes}
    if (Col mod 2) = 0 then begin
      NewRec := Grille.CellRect(Col+1, Row);
      SetRect(NewRec, Rect.Left, Rect.Top, NewRec.Right, NewRec.Bottom);
    end
    else begin
      NewRec := Grille.CellRect(Col-1, Row);
      SetRect(NewRec, NewRec.Left, NewRec.Top, Rect.Right, Rect.Bottom);
    end;
  end;

  if (gdFixed in State) then begin
    Grille.Canvas.Brush.Color:=Grille.FixedColor ;
    Grille.Canvas.Font.Color:=Grille.Font.Color ;
    F:=taCenter
  end
  else begin
    F := taLeftJustify ;
    if (gdSelected in State) then begin
       Grille.Canvas.Brush.Color:=clHighlight ;
       Grille.Canvas.Font.Color:=clHighlightText ;
     end
     else begin
       Grille.Canvas.Brush.Color:=clWindow ;
       Grille.Canvas.Font.Color:=clWindowText ;
    end;
  end;

  Grille.Canvas.FillRect(NewRec);
  Grille.Canvas.FillRect(Rect);

  case F of
    taRightJustify : ExtTextOut(Grille.Canvas.Handle, NewRec.Right - Grille.Canvas.TextWidth(Grille.Cells[Col,Row])-3,
                               NewRec.Top + 2,ETO_OPAQUE or ETO_CLIPPED, @NewRec, Text, StrLen(Text), nil) ;
    taCenter       : ExtTextOut(Grille.Canvas.Handle,NewRec.Left + ((NewRec.Right-NewRec.Left-Grille.canvas.TextWidth(Grille.Cells[Col,Row])) div 2),
                         NewRec.Top + 2, ETO_OPAQUE or ETO_CLIPPED, @NewRec, Text, StrLen(Text), nil) ;
    else
      ExtTextOut(Grille.Canvas.Handle, NewRec.Left + 2, NewRec.Top + 2, ETO_OPAQUE or ETO_CLIPPED,
                 @NewRec, Text, StrLen(Text), nil) ;
  end;

  if (gdfixed in State) and Grille.Ctl3D then begin
    DrawEdge(Grille.Canvas.Handle, NewRec, BDR_RAISEDINNER, BF_BOTTOMRIGHT);
    DrawEdge(Grille.Canvas.Handle, NewRec, BDR_RAISEDINNER, BF_TOPLEFT);
  end;

  Grille.Canvas.Brush.Assign(OldBrush) ;
  Grille.Canvas.Pen.Assign(OldPen) ;
  OldBrush.Free;
  OldPen.Free;
end;

procedure TOM_CMODELRESTANA.GridIntervalTopLeftChanged(Sender: TObject);
begin
  GdInterval.Repaint;
end;

procedure TOM_CMODELRESTANA.InitGridInterval(InitRowCount: Boolean);
var
  i:     Integer;
begin
  GdInterval.SynEnabled := False;

  if InitRowCount then
    GdInterval.RowCount := GdInterval.FixedRows + 1;

  GdInterval.ColCount := FAxeCrois.Count*2;

  if FAxeCrois.Count = 1 then
    GdInterval.ColWidths[0] := (GdInterval.ClientWidth div 2)-10  {Pour la scrollbar}
  else begin
    {JP 16/11/07 : FQ 21873 : Boucle utile uniquement en Coise Axe}
    for i:=0 to FAxeCrois.Count-1 do
      GdInterval.Cells[(i * 2) + 1, 0] := RechDom('TTAXE', 'A' + FAxeCrois[i], False);
    GdInterval.ColWidths[0] := (GdInterval.ClientWidth div 4)+10; {Sinon pb d'affichage}
  end;

  i:=0;
  while i < GdInterval.ColCount-1 do begin
    GdInterval.Cells[i, 1]     := TraduireMemoire('Début');
    GdInterval.Cells[i+1, 1]   := TraduireMemoire('Fin');
    GdInterval.ColWidths[i]    := GdInterval.ColWidths[0];
    GdInterval.ColWidths[i+1]  := GdInterval.ColWidths[0];
    GdInterval.ColLengths[i]   := VH^.Cpta[TFichierBase(Ord(fbAxe1)+(i mod 2))].Lg;
    GdInterval.ColLengths[i+1] := VH^.Cpta[TFichierBase(Ord(fbAxe1)+(i mod 2))].Lg;
    Inc(i, 2);
  end;

  GdInterval.SynEnabled := True;
  GdInterval.Repaint;

  {JP 16/11/07 : FQ 21873 : gestion de la longueur des sections}
  if FAxeCrois.Count = 1 then begin
    GdInterval.ColLengths[0]:= VH^.Cpta[TFichierBase(StrToInt(FAxeCrois[0]) - 1)].Lg;
    GdInterval.ColLengths[1]:= VH^.Cpta[TFichierBase(StrToInt(FAxeCrois[0]) - 1)].Lg;
    GdInterval.Refresh;
  end;
end;

procedure TOM_CMODELRESTANA.InitAxePageControl;
var
  i:       Integer;
  Grid:    THGrid;
begin
  for i:=0 to MaxAxe-1 do begin
    TsAxes[i].Visible       := FAxeCrois.IndexOf(IntToStr(i+1)) <> -1;
    TsAxes[i].TabVisible    := TsAxes[i].Visible;
    GdComptes[i].Visible    := TsAxes[i].Visible;
    GdComptes[i].RowCount   := 2;
    GdComptes[i].Options    := GdComptes[0].Options;
    GdComptes[i].OnDblClick := GridCptDblClick;

    if FAxeCrois.Count = 1 then begin
      {Masque le page control}
      Grid          := GdComptes[i];
      Grid.Align    := alNone;
      Grid.Parent   := PCAxe.Parent;
      Grid.SetBounds(PCAxe.Left, PCAxe.Top, PCAxe.Width, PCAxe.Height);
      PCAxe.Visible := False;
    end;
  end;

  if FAxeCrois.Count > 1 then begin
    {Permet de visualiser la bonne grille}
    PCAxe.ActivePageIndex := StrToInt(FAxeCrois[1])-1;
    PCAxe.ActivePageIndex := StrToInt(FAxeCrois[0])-1;
  end;
end;


procedure TOM_CMODELRESTANA.InitGridCompte;
var
  i:       Integer;
begin
  for i:=0 to MaxAxe-1 do begin
    GdComptes[i].SynEnabled   := False;
    GdComptes[i].OnColumnWidthsChanged := GridCompteColumnWidthsChanged;
    GdComptes[i].ColWidths[0] := GdComptes[i].ClientWidth div 3;
    GdComptes[i].ColWidths[1] := (2*GdComptes[i].ClientWidth) div 3;
    GdComptes[i].ColAligns[0] := TaLeftJustify;
    GdComptes[i].ColAligns[1] := TaLeftJustify;
    GdComptes[i].SynEnabled   := True;
  end;
end;

procedure TOM_CMODELRESTANA.ChargeGridCompte;
var
  i:       Integer;
  Grid:    THGrid;
  Clause:  String;
  QCpt:    TQuery;
  iLigne:  Integer;
  icol:    Integer;
  RuptureAxe:    String;
  Axe:     String;        {FP 19/04/2006 FQ17840}
begin
  for i:=0 to MaxAxe-1 do begin
    GdComptes[i].RowCount := 2;
    for iCol:=0 to GdComptes[i].ColCount-1 do
      GdComptes[i].Cells[iCol, 1] := '';
  end;

  {Alimente les grilles contenant la liste des comptes sections}
  Clause := '';
  for i:=0 to FAxeCrois.Count-1 do begin
    if i <> 0 then
      Clause := Clause+' OR';
    {b FP 19/04/2006 FQ17840}
    Axe := 'A'+FAxeCrois[i];
    Clause := Clause+' (S_AXE="'+Axe+'"' +
      ' AND S_SECTION<>"'+GetInfoCpta(AxeToFb(Axe)).Attente+'"' +
      ')';
    {e FP 19/04/2006}
  end;
  QCpt:=OpenSQL('SELECT S_AXE, S_SECTION, S_LIBELLE FROM SECTION'+
                ' WHERE ('+Clause+')'+
                ' ORDER BY S_AXE, S_SECTION',True);
  try
    iLigne     := 0;
    Grid       := nil;                  
    RuptureAxe := '';
    while not QCpt.Eof do begin
      if (RuptureAxe='') or (RuptureAxe<>QCpt.FindField('S_AXE').AsString) then begin
        iLigne       := 1;
        RuptureAxe   := QCpt.FindField('S_AXE').AsString;
        Grid         := GdComptes[StrToInt(RuptureAxe[2])-1];
      end;
      if iLigne > 1 then
        Grid.RowCount := Grid.RowCount+1;
      Grid.Cells[0, iLigne] := QCpt.FindField('S_SECTION').AsString;
      Grid.Cells[1, iLigne] := QCpt.FindField('S_LIBELLE').AsString;
      Inc(iLigne);
      QCpt.Next;
    end;
  finally
    FreeAndNil(QCpt);
  end;
end;

procedure TOM_CMODELRESTANA.AxeChange(Sender: TObject);
begin
  if (FAxeCrois.Count = 1) and (Length(Axe.Value) > 0) then begin
    if FAxeCrois[0] <> Axe.Value[2] then begin
      GdComptes[StrToInt(FAxeCrois[0])-1].Visible := False;
      FAxeCrois.Clear;
      FAxeCrois.Add(Axe.Value[2]);
      GdComptes[StrToInt(FAxeCrois[0])-1].Visible := True;
      InitGridCompte;
      ChargeGridCompte;
    end;
  end;
  InitGridInterval(True);
  ChargeGridInterval('', '');
end;

procedure TOM_CMODELRESTANA.GridCptDblClick(Sender: TObject);
var
  Grid:   THGrid;
  i:      Integer;
  iAxe:   Integer;
  iCol:   Integer;
  Insert: Boolean;
begin
  Grid := (Sender as THGrid);

  iAxe := 0;
  for i:=0 to MaxAxe-1 do begin
    if GdComptes[i] = Grid then
      iAxe := i+1;
  end;

  if FAxeCrois.Count = 1 then
    iCol := 0
  else
    iCol := 2*(iAxe-1);
  if (iCol = GdInterval.Col) or ((iCol+1) = GdInterval.Col) then begin
    {La cellule active correspond à l'axe sélectionné}
    Insert := (Trim(GdInterval.Cells[GdInterval.Col, GdInterval.Row]) = '');
    GdInterval.Cells[GdInterval.Col, GdInterval.Row] := Grid.Cells[0, Grid.Row];
  end
  else if (Trim(GdInterval.Cells[iCol+1, GdInterval.Row]) = '') and
          (not TestJoker(GdInterval.Cells[iCol, GdInterval.Row])) then begin
    {Cellule fin vide et cellule début ne contient pas de caractère joker}
    Insert := (Trim(GdInterval.Cells[iCol+1, GdInterval.Row]) = '');
    GdInterval.Cells[iCol+1, GdInterval.Row] := Grid.Cells[0, Grid.Row];
  end
  else begin
    Insert := (Trim(GdInterval.Cells[iCol, GdInterval.Row]) = '');
    GdInterval.Cells[iCol, GdInterval.Row] := Grid.Cells[0, Grid.Row];
  end;

  if Insert and (GdInterval.Row = GdInterval.RowCount-1) then
    NewLigne(GdInterval.RowCount);
  ForceModeEdition;
end;

procedure TOM_CMODELRESTANA.GridIntervalSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
var
  iColDeb:   Integer;
  iColFin:   Integer;
  IsSrcColDeb:  Boolean;
  IsSrcColFin:  Boolean;
  IsDstColFin:  Boolean;
  Debut:     String;
  Fin:       String;
  ColJoker:  Boolean;
begin
  ColJoker    := False;
  IsSrcColDeb := False;
  IsSrcColFin := False;
  IsDstColFin := False;
  if (GdInterval.Col mod 2) = 0 then begin
    iColDeb     := GdInterval.Col;
    iColFin     := GdInterval.Col+1;
    IsSrcColDeb := True;
  end
  else begin
    iColDeb     := GdInterval.Col-1;
    iColFin     := GdInterval.Col;
    IsSrcColFin := True;
  end;
  if (ACol mod 2) = 1 then
    IsDstColFin := True;

  Debut := Trim(GdInterval.Cells[iColDeb, GdInterval.Row]);
  Fin   := Trim(GdInterval.Cells[iColFin, GdInterval.Row]);

  CanSelect := True;    {FP 22/06/2006}
  if CanSelect and IsSrcColDeb and TestJoker(GdInterval.Cells[iColDeb, GdInterval.Row]) then begin
    GdInterval.Cells[iColFin, GdInterval.Row]:='';
    ColJoker := True;
  end;

  if CanSelect and IsSrcColFin and TestJoker(GdInterval.Cells[iColFin, GdInterval.Row]) then begin
    CanSelect := False;
    LastError    := 1;
    LastErrorMsg := '';
    PGIError(MsgBox.Mess[2], Ecran.Caption);
  end;

  if CanSelect and (not ColJoker) and (Fin <> '') and (CompareText(Debut, Fin) > 0) then begin
    CanSelect := False;
    LastError    := 1;
    LastErrorMsg := '';
    PGIError(MsgBox.Mess[3], Ecran.Caption);
  end;

  if CanSelect and ColJoker and IsDstColFin then begin
    CanSelect := False;
    LastError    := 1;
    LastErrorMsg := '';
    PGIError(MsgBox.Mess[4], Ecran.Caption);
  end;

  if CanSelect and (FAxeCrois.Count > 1) then begin
    PCAxe.ActivePageIndex := Ceil((ACol+1)/2)-1;
  end;
  if CanSelect then
    LocateCompte(GdInterval.Cells[ACol, ARow]);
end;

procedure TOM_CMODELRESTANA.PCAxeChange(Sender: TObject);
begin
  GdInterval.Col := (PCAxe.ActivePageIndex*2);
end;

function TOM_CMODELRESTANA.QuickSearch(Grid: THGrid; ACol: Integer; Key: String; FirstRow, LastRow: Integer): Integer;
var
  Pivot: Integer;
begin
  if (FirstRow>=LastRow) then begin
    Result := CompareText(Key, Grid.Cells[ACol, FirstRow]);
    if Result <= 0 then
      Result := FirstRow
    else if Result > 0 then
      Result := FirstRow+1;
    if Result <> 0 then
      Result := -1*(Result+1);
  end
  else begin
    Pivot := ((LastRow-FirstRow) div 2)+FirstRow;
    Result := CompareText(Key, Grid.Cells[ACol, Pivot]);
    if Result = 0 then
      Result := Pivot
    else if Result < 0 then
      Result := QuickSearch(Grid, ACol, Key, FirstRow, Pivot-1)
    else if Result > 0 then
      Result := QuickSearch(Grid, ACol, Key, Pivot+1, LastRow);
  end;
end;

procedure TOM_CMODELRESTANA.LocateCompte(Cpt: String);
var
  Grid:  THGrid ;
  i:     Integer;
  iAxe:  Integer;
  Index: Integer;
begin
  {Recherche le texte saisi dans la grille des comptes}
  iAxe := 0;
  for i:=0 to MaxAxe-1 do
    if TsAxes[i] = PCAxe.ActivePage then
      iAxe := i+1;
  Grid  := GdComptes[iAxe-1];
  if Trim(Cpt) = '' then
    Index := Grid.FixedRows
  else
    Index := QuickSearch(Grid, 0, Cpt, Grid.FixedRows, Grid.RowCount-1);
  if Index < 0 then
    Index := (-1*Index)-1;
  Grid.Row := Min(Max(Grid.FixedRows, Index), Grid.RowCount-1);
end;

procedure TOM_CMODELRESTANA.DeleteRecordDetail(Axe, Modele: String; UpdateMode: Boolean);
begin
  BEGINTRANS;
  try
    ExecuteSQL('DELETE FROM CModelRestAnaDet'+
               ' WHERE CAD_AXE  = "'+Axe+'"' +
               '   AND CAD_CODE = "'+Modele+'"');
    if not UpdateMode then
      begin
      ExecuteSQL('UPDATE CLIENGENEMODELA SET CLA_CODE'+Axe[Length(Axe)]+'=""'+
                 ' WHERE CLA_CODE'+Axe[Length(Axe)]+'="'+Modele+'"');
      // FQ 21635 : suppression des liens modèle / comptes si aucun axe restreint...
      ExecuteSQL('DELETE FROM CLIENGENEMODELA WHERE (CLA_CODE1="" OR CLA_CODE1 IS NULL) AND '
                                                 + '(CLA_CODE2="" OR CLA_CODE2 IS NULL) AND '
                                                 + '(CLA_CODE3="" OR CLA_CODE3 IS NULL) AND '
                                                 + '(CLA_CODE4="" OR CLA_CODE4 IS NULL) AND '
                                                 + '(CLA_CODE5="" OR CLA_CODE5 IS NULL) ');
      end ;
    COMMITTRANS;
  except
    ROLLBACK;
    raise;
  end;
end;

procedure TOM_CMODELRESTANA.BValiderClick(Sender: TObject);
var
  CanSelect: Boolean;
  iRow:      Integer;
  iCol:      Integer;
begin
  if True {FP 19/04/2006 FQ17845 Ecran.ActiveControl = GdInterval} then begin
    {Les SpeedButtons ne prenant pas le Focus, on simule un changement de ligne}
    iRow := GdInterval.RowCount-1;
    if (GdInterval.Col = 0) and (GdInterval.Row = GdInterval.RowCount-1) then
      iCol := 1
    else
      iCol := 0;
    GridIntervalSelectCell(GdInterval, iCol, iRow, CanSelect);
  end;

  if (LastError = 0) and CanSelect then
    begin
    FOnBValiderClick(Sender);
    Ecran.ModalResult := mrOk;      {FP 22/06/2006}
    end
  else begin
    if Trim(LastErrorMsg) <> '' then
      PGIError(LastErrorMsg, Ecran.Caption);
    LastError    := 0;
    LastErrorMsg := '';
  end;
end;

procedure TOM_CMODELRESTANA.ForceModeEdition;
begin
  {Force le mode MAJ de la fiche}
{$IFNDEF EAGLCLIENT}
  if not (DS.State in [dsInsert, dsEdit]) then
    DS.Edit;
{$ENDIF}
  SetField('CRA_LIBELLE', Libelle.Text+'8');
  SetField('CRA_LIBELLE', Copy(Libelle.Text,1,Length(Libelle.Text)-1));
end;

procedure TOM_CMODELRESTANA.BInsertClick(Sender: TObject);
begin
  if VH^.JalCtrlBud<>'' then
    begin
    PGIInfo('Le contrôle en saisie sur le budget est incompatible avec la création d''un modèle de restrictions analytiques.'+Chr($D)+Chr($A)+
            'Si vous voulez créer un modèle vous devez décocher le paramètre de contrôle dans les paramètres société.');
    end
  else if Assigned(FOnBInsertClick) then
    begin
    FOnBInsertClick(Sender) ;
    end;
end;

procedure TOM_CMODELRESTANA.GridCompteColumnWidthsChanged(Sender: TObject);
begin
  InitGridCompte;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CMODELRESTANA.GridIntervalEnter(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {Pour le cas où l'on entre de la grille par un clik de la souris, pour que le focus soit pris correctement}
  GdInterval.Col := 0;
  PostMessage(GdInterval.Handle, WM_KEYDOWN, VK_TAB,  0) ;
  PostMessage(GdInterval.Handle, WM_KEYDOWN, VK_LEFT, 0) ;
end;

Initialization
  registerclasses ( [ TOM_CMODELRESTANA ] ) ;
end.

