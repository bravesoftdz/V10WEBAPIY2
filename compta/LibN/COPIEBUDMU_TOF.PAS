{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : COPIEBUDMU ()
Mots clefs ... : TOF;COPIEBUDMU
*****************************************************************}
Unit COPIEBUDMU_TOF ;

Interface

Uses StdCtrls, Controls, Classes, Windows, 
{$IFDEF EAGLCLIENT}
  eMul, MaineAGL,
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  db, mul, FE_Main,
{$ENDIF}
  uTob, Vierge, forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, HStatus, Graphics, Grids, Ent1,
  UTOF;

procedure RecopieBudgetMultiple;

Type TInfoBud = Class
        ExoDeb   : String ;
        ExoFin   : String ;
        PerDeb   : String ;
        PerFin   : String ;
        Axe      : String ;
        GeneAtt  : String ;
        SectAtt  : String ;
        CpteGen  : String ;
        CpteSec  : String ;
        Souche   : String ;
        NbPer    : Integer ;
        NbPiece  : Integer ;
        NumPiece : String ;
        NatPiece : String ;
        CoefCpy  : Double ;
        NatPieceCpy :String ;
      end ;

Type
  TOF_COPIEBUDMU = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    HM : THMsgBox;
    FListe : THGrid;
    BudJal, ExoDebD, ExoFinD, PerDebD, PerFinD, PerDebS, PerFinS : THValComboBox;
    RgCopi : THRadioGroup;
    TotalSelec : Integer;
    ListBud    : TStringList;
    NowFutur   : TDateTime;
    procedure BudJalChange(Sender: TObject);
    procedure BOptClick(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BValiderClick(Sender: TObject);

    procedure ChargeCbBudDestination;
    function  CombiendePeriode(Q : TQuery) : Integer;
    function  RempliUnTinfoBud(T : TQuery) : TInfoBud;
    procedure DetruitObjetCombo;
    procedure RempliFListe;
    procedure ChercheNbEcr(X : TInfoBud ; St : String);
    procedure CompteElemSelectionner;
    function  CpteAttOk(ARow : Integer) : Boolean;
    function  NbPeriodeOk(ARow : Integer) : Boolean;
    function  ListeCpteOk(ARow : Integer) : Boolean;
    function  NbPieceOk(ARow : Integer) : Boolean;
    function  ConTroleBudOk(ARow : Integer) : Boolean;
    function  ChercheDateCompta(D : TDateTime ; Exo : String ; ARow : Integer) : TDateTime;
    function  ChercheExo(D : TDateTime) : String;
    procedure RunLaCopie;
    procedure InsereLesEnreg(Info : TInfoBud ; ARow : Integer; TCpy : TQuery);
    procedure RequeteInsertion(Info : TInfoBud ; NextNumPiece : Longint ; Exo : String ; DCpta : TDateTime ; D,C : Double; TCpy : TQuery);
    procedure MajLesEnreg(Info : TInfoBud ;ARow : Integer; TCpy : TQuery);
    procedure MajNbEcr;
    procedure RempliPeriodeSource(Arow : Integer);
    procedure InverseSelection;
    procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState);
    function  GetCountBUDECR ( StCond : String ) : Integer;
  end ;

Implementation

uses SAISUTIL, // GetNum, SetIncNum
     OptCpyBu, // OptiondeCopideBudget
     eSaisBud; // VisuConsoBudget

procedure RecopieBudgetMultiple;
begin
  AGLLanceFiche('CP','CPCOPIEBUDMU','','','');
end;

procedure TOF_COPIEBUDMU.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_COPIEBUDMU.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_COPIEBUDMU.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_COPIEBUDMU.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_COPIEBUDMU.OnArgument (S : String ) ;
begin
  Inherited ;

  HM := THMsgBox.Create(Ecran);
  HM.Mess.Add('ligne sélectionnée');
  HM.Mess.Add('lignes sélectionnées');
  HM.Mess.Add('2;Copie multiple de budget;Aucun budget n''a été sélectionné pour la copie.;W;O;O;O;');
  HM.Mess.Add('3;Copie multiple de budget;Aucun budget sélectionné ne pourra être copié. Caractéristiques différentes ou aucune pièce à copier.;W;O;O;O;');
  HM.Mess.Add('4;Copie multiple de budget;Désirez-vous copier les budgets sélectionnés ?;Q;YN;N;N;');
  HM.Mess.Add('5;Copie multiple de budget;Désirez-vous visualiser les écritures budgétaires ?;Q;YN;N;N;');
  HM.Mess.Add('Recopie non effectuée.');
  HM.Mess.Add('7;Copie multiple de budget;La recopie de budget s''est correctement effectuée.;E;O;O;O;');

  FListe := THGrid(GetControl('FLISTE', True));
  BudJal := THValComboBox(GetControl('BUDJAL', True));
  BudJal.OnChange := BudJalChange;
  TButton(GetControl('BOPT', True)).OnClick := BOptClick;
  TButton(GetControl('BZOOM', True)).OnClick := BZoomClick;
  FListe.GetCellCanvas := GetCellCanvas;
// YMO FQ 16530 DblClick au lieu de Click
  FListe.OnDblClick := BOptClick;
  FListe.OnKeyDown := FListeKeyDown;
  FListe.OnMouseDown := FListeMouseDown;
  TFVierge(Ecran).BValider.OnClick := BValiderClick;

  ExoDebD := THValComboBox(GetControl('EXODEBD', True));
  ExoFinD := THValComboBox(GetControl('EXOFIND', True));
  PerDebD := THValComboBox(GetControl('PERDEBD', True));
  PerFinD := THValComboBox(GetControl('PERFIND', True));
  PerDebS := THValComboBox(GetControl('PERDEBS', True));
  PerFinS := THValComboBox(GetControl('PERFINS', True));
  RgCopi := THRadioGroup(GetControl('RGCOPI', True));

  // Combo destination
  ChargeCbBudDestination;

  // Liste origine
  RempliFliste;

  if BudJal.Values.Count>0 then BudJal.Value := BudJal.Values[0];
  ListBud := TStringList.Create;
end ;

procedure TOF_COPIEBUDMU.OnClose ;
begin
  Inherited ;
  FListe.VidePile(True);
  DetruitObjetCombo;
  ListBud.Free;
end ;

procedure TOF_COPIEBUDMU.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_COPIEBUDMU.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_COPIEBUDMU.CombiendePeriode(Q : TQuery): Integer;
var
  a,m,j,a1,m1,j1 : Word ;
begin
  if Q.FindField('BJ_EXODEB').AsString = Q.FindField('BJ_EXOFIN').AsString then begin
    DecodeDate(Q.FindField('BJ_PERDEB').AsDateTime, a, m, j);
    DecodeDate(Q.FindField('BJ_PERFIN').AsDateTime, a1, m1, j1);
    Result := (m1-m)+1;
    end
  else begin
    DecodeDate(Q.FindField('BJ_PERDEB').AsDateTime, a, m, j);
    DecodeDate(Q.FindField('BJ_PERFIN').AsDateTime, a1, m1, j1);
    Result := ((12-m)+1)+m1;
  end ;
end;

procedure TOF_COPIEBUDMU.ChargeCbBudDestination;
var
  X : TInfoBud ;
  QLoc : TQuery;
  i : Integer;
begin
  QLoc := OpenSQL('SELECT * FROM BUDJAL WHERE BJ_FERME="-"', True);
  try
    BudJal.Values.Clear;
    BudJal.Items.Clear;

    while not QLoc.EOF do begin
       X := RempliUnTinfoBud(QLoc);
      ChercheNbEcr(X, QLoc.FindField('BJ_BUDJAL').AsString);
      Budjal.Values.AddObject(QLoc.FindField('BJ_BUDJAL').AsString, X);
      Budjal.Items.Add(QLoc.FindField('BJ_LIBELLE').AsString);
      QLoc.Next;
    end;
  finally
    Ferme(QLoc);
  end;
end;

procedure TOF_COPIEBUDMU.RempliFListe;
var
  X : TInfoBud ;
  i : Integer ;
  QLoc : TQuery;
begin
  QLoc := OpenSql('SELECT * FROM BUDJAL B WHERE BJ_FERME="-" AND EXISTS(SELECT BE_BUDJAL FROM BUDECR WHERE BE_BUDJAL=B.BJ_BUDJAL)', True);
  try
    i := 2;
    SetControlEnabled('BOPT', Not QLoc.Eof);

    while not QLoc.EOF do begin
      FListe.RowCount := i;
      X := RempliUnTinfoBud(QLoc);
      FListe.Cells[0,FListe.RowCount-1] := QLoc.FindField('BJ_BUDJAL').AsString;
      ChercheNbEcr(X,FListe.Cells[0,FListe.RowCount-1]);
      FListe.Objects[0,FListe.RowCount-1] := X;
      FListe.Cells[1,FListe.RowCount-1] := QLoc.FindField('BJ_LIBELLE').AsString;
      FListe.Cells[2,FListe.RowCount-1] := IntToStr(X.NbPer);
      FListe.Cells[3,FListe.RowCount-1] := IntToStr(X.NbPiece);
      Inc(i);
      QLoc.Next;
    end;
  finally
    Ferme(QLoc);
  end;
end;

procedure TOF_COPIEBUDMU.MajNbEcr;
var
  i : Integer;
  St : String;
begin
  St := BudJal.Value;
  for i:=1 to FListe.RowCount-1 do begin
    if FListe.Cells[0,i]=St then begin
      ChercheNbEcr(TInfoBud(FListe.Objects[0,i]),St);
      FListe.Cells[3,i] := IntToStr(TInfoBud(FListe.Objects[0,i]).NbPiece);
      FListe.Invalidate;
      Break;
    end;
  end;
end;

procedure TOF_COPIEBUDMU.ChercheNbEcr(X: TInfoBud; St: String);
var
  i,j : Integer;
  QNbEcr : TQuery;
begin
  QNbEcr := OpenSQL('SELECT BE_NUMEROPIECE,BE_NATUREBUD FROM BUDECR WHERE BE_BUDJAL="'+St+'" GROUP BY BE_NUMEROPIECE,BE_NATUREBUD', True);
  try
    i:=0;
    X.NumPiece := '';
    X.NbPiece := 0;
    X.NatPiece := '';
    X.NatPieceCpy := '';
    while not QNbEcr.EOF do begin
      X.NumPiece := X.NumPiece + QNbEcr.FindField('BE_NUMEROPIECE').AsString + ':' + QNbEcr.FindField('BE_NATUREBUD').AsString + '@';
      if Pos(QNbEcr.FindField('BE_NATUREBUD').AsString, X.NatPiece)<=0 then
        X.NatPiece := X.NatPiece + QNbEcr.FindField('BE_NATUREBUD').AsString + ';';
      Inc(i);
      QNbEcr.Next
    end;
    X.NbPiece := i;
    X.NatPieceCpy := X.NatPiece;
  finally
    Ferme(QNbEcr);
  end;
end;

function TOF_COPIEBUDMU.RempliUnTinfoBud(T: TQuery): TInfoBud;
var
  X : TInfoBud ;
begin
  X := TInfoBud.Create;
  X.ExoDeb  := T.FindField('BJ_EXODEB').AsString;
  X.ExoFin  := T.FindField('BJ_EXOFIN').AsString;
  X.PerDeb  := T.FindField('BJ_PERDEB').AsString;
  X.PerFin  := T.FindField('BJ_PERFIN').AsString;
  X.Axe     := T.FindField('BJ_AXE').AsString;
  X.GeneAtt := T.FindField('BJ_GENEATTENTE').AsString;
  X.SectAtt := T.FindField('BJ_SECTATTENTE').AsString;
  X.CpteGen := T.FindField('BJ_BUDGENES').AsString;
  X.CpteSec := T.FindField('BJ_BUDSECTS').AsString;
  if T.FindField('BJ_BUDGENES2').AsString<>'' then X.CpteGen := Trim(T.FindField('BJ_BUDGENES').AsString) + Trim(T.FindField('BJ_BUDGENES2').AsString)
                                              else X.CpteGen := Trim(T.FindField('BJ_BUDGENES').AsString);
  if T.FindField('BJ_BUDSECTS2').AsString<>'' then X.CpteSec := Trim(T.FindField('BJ_BUDSECTS').AsString) + Trim(T.FindField('BJ_BUDSECTS2').AsString)
                                              else X.CpteSec := Trim(T.FindField('BJ_BUDSECTS').AsString);
  X.Souche := T.FindField('BJ_COMPTEURNORMAL').AsString;
  X.NbPer := CombiendePeriode(T);
  X.NumPiece := '';
  X.NbPiece := 0;
  X.NatPiece := '';
  X.CoefCpy := 0;
  Result := X;
end;

procedure TOF_COPIEBUDMU.DetruitObjetCombo;
var
  i : Integer ;
begin
  for i:=0 to BudJal.Values.Count-1 do
    TObject(BudJal.Values.Objects[i]).Free;
end;

procedure TOF_COPIEBUDMU.BudJalChange(Sender: TObject);
var
  i : Integer;
begin
  if BudJal.Value='' then Exit;
  ExoDebD.Value := TInfoBud(Budjal.Values.Objects[BudJal.Values.IndexOf(BudJal.Value)]).ExoDeb;
  ExoFinD.Value := TInfoBud(Budjal.Values.Objects[BudJal.Values.IndexOf(BudJal.Value)]).ExoFin;
  ListePeriode(ExoDebD.Value, PerDebD.Items, PerDebD.Values, True);
  ListePeriode(ExoFinD.Value, PerFinD.Items, PerFinD.Values, False);
  PerDebD.Value := TInfoBud(Budjal.Values.Objects[BudJal.Values.IndexOf(BudJal.Value)]).PerDeb;
  PerFinD.Value := TInfoBud(Budjal.Values.Objects[BudJal.Values.IndexOf(BudJal.Value)]).PerFin;
  for i:=1 to FListe.RowCount-1 do begin
    if FListe.Cells[0,i]=Budjal.Value then begin
      Fliste.Cells[FListe.ColCount-1,i] := '';
      FListe.Invalidate;
      CompteElemSelectionner;
      Break;
    end;
  end;
end;

procedure TOF_COPIEBUDMU.GetCellCanvas(Acol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if FListe.Cells[FListe.ColCount-1,ARow]='*' then FListe.Canvas.Font.Style := FListe.Canvas.Font.Style+[fsItalic]
                                              else FListe.Canvas.Font.Style := FListe.Canvas.Font.Style-[fsItalic];
end;

procedure TOF_COPIEBUDMU.InverseSelection;
begin
  if FListe.Cells[0,FListe.Row]=BudJal.Value then Exit;
  if FListe.Cells[0,FListe.Row]='' then Exit;
  if Fliste.Cells[FListe.ColCount-1,FListe.Row]='*' then Fliste.Cells[FListe.ColCount-1,FListe.Row]:=''
                                                    else Fliste.Cells[FListe.ColCount-1,FListe.Row]:='*';
  FListe.Invalidate;
  CompteElemSelectionner;
end;

procedure TOF_COPIEBUDMU.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (ssShift in Shift) And (Key=VK_DOWN) then InverseSelection else
  if (Shift=[]) And (Key=VK_SPACE) then begin
    InverseSelection;
    if ((FListe.Row<FListe.RowCount-1) and (Key<>VK_SPACE)) then FListe.Row := FListe.Row+1;
  end ;
end;

procedure TOF_COPIEBUDMU.FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if (ssCtrl in Shift) And (Button=mbLeft)then InverseSelection;
end;

procedure TOF_COPIEBUDMU.CompteElemSelectionner;
var
  i : Integer;
begin
  TotalSelec := 0;
  for i:=1 to FListe.RowCount-1 do
    if FListe.Cells[FListe.ColCount-1,i]='*' then Inc(TotalSelec);

  SetControlCaption('NB1', IntToStr(TotalSelec));
  if TotalSelec>1 then SetControlCaption('TEX1', HM.Mess[1])  // lignes sélectionnées
                  else SetControlCaption('TEX1', HM.Mess[0]); // ligne sélectionnée
end;

procedure TOF_COPIEBUDMU.BOptClick(Sender: TObject);
begin
  OptiondeCopideBudget(FListe.Cells[0,FListe.Row],
                       FListe.Cells[1,FListe.Row],
                       TInfoBud(FListe.Objects[0,FListe.Row]).NatPiece,
                       TInfoBud(FListe.Objects[0,FListe.Row]).NatPieceCpy,
                       TInfoBud(FListe.Objects[0,FListe.Row]).CoefCpy);
end;

procedure TOF_COPIEBUDMU.BValiderClick(Sender: TObject);
var
  i : Integer;
  io : TIOErr;
begin
  if FListe.Cells[0,1]='' then Exit;
  if TotalSelec=0 then begin
    HM.Execute(2,'',''); // Aucun budget n''a été sélectionné pour la copie.
    Exit;
  end;

  ListBud.Clear;
  for i:=1 to FListe.RowCount-1 do begin
    if FListe.Cells[FListe.ColCount-1,i]='*' then
      if ConTroleBudOk(i) then ListBud.Add(FListe.Cells[0,i]);
  end ;

  if ListBud.Count=0 then begin
    HM.Execute(3,'',''); // Aucun budget sélectionné ne pourra être copié. Caractéristiques différentes ou aucune pièce à copier.
    Exit;
  end;

  if HM.Execute(4,'','')<>mrYes then Exit; // Désirez-vous copier les budgets sélectionnés ?

  io := Transactions(RunLaCopie,2);

  if io<>oeOk then MessageAlerte(HM.Mess[6]) // Recopie non effectuée
              else HM.Execute(7,'','');      // La recopie de budget s''est correctement effectuée.
end;

function TOF_COPIEBUDMU.CpteAttOk(ARow: Integer): Boolean;
begin
  Result := False;
  if TInfoBud(FListe.Objects[0,ARow]).GeneAtt <> TInfoBud(Budjal.Values.Objects[BudJal.Values.IndexOf(BudJal.Value)]).GeneAtt then Exit;
  if TInfoBud(FListe.Objects[0,ARow]).SectAtt <> TInfoBud(Budjal.Values.Objects[BudJal.Values.IndexOf(BudJal.Value)]).SectAtt then Exit;
  Result := True;
end;

function TOF_COPIEBUDMU.NbPeriodeOk(ARow: Integer): Boolean;
begin
  Result := False;
  if TInfoBud(FListe.Objects[0,ARow]).NbPer <> TInfoBud(Budjal.Values.Objects[BudJal.Values.IndexOf(BudJal.Value)]).NbPer then Exit;
  Result := True;
end;

function TOF_COPIEBUDMU.ListeCpteOk(ARow: Integer): Boolean;
var
  StG,StS,St,StGD,StSD : String;
begin
  Result := False;
  if Length(TInfoBud(FListe.Objects[0,ARow]).CpteGen) <> Length(TInfoBud(Budjal.Values.Objects[BudJal.Values.IndexOf(BudJal.Value)]).CpteGen) then Exit;
  if Length(TInfoBud(FListe.Objects[0,ARow]).CpteSec) <> Length(TInfoBud(Budjal.Values.Objects[BudJal.Values.IndexOf(BudJal.Value)]).CpteSec) then Exit;

  StG := TInfoBud(FListe.Objects[0,ARow]).CpteGen;
  StS := TInfoBud(FListe.Objects[0,ARow]).CpteSec;
  StGD := TInfoBud(Budjal.Values.Objects[BudJal.Values.IndexOf(BudJal.Value)]).CpteGen;
  StSD := TInfoBud(Budjal.Values.Objects[BudJal.Values.IndexOf(BudJal.Value)]).CpteSec;
  While StG<>'' do begin
    St := ReadTokenSt(StG);
    if Pos(St,StGD)<=0 then Exit;
  end;

  While StS<>'' do begin
    St := ReadTokenSt(StS);
    if Pos(St,StSD)<=0 then Exit;
  end;
  Result := True;
end;

function TOF_COPIEBUDMU.NbPieceOk(ARow: Integer): Boolean;
begin
  Result := False;
  if TInfoBud(FListe.Objects[0,ARow]).NbPiece=0 then Exit;
  Result := True;
end;

function TOF_COPIEBUDMU.ConTroleBudOk(ARow: Integer): Boolean;
begin
  Result := False;
  if Not CpteAttOk(ARow)   then Exit;
  if Not NbPeriodeOk(ARow) then Exit;
  if Not ListeCpteOk(ARow) then Exit;
  if Not NbPieceOk(ARow)   then Exit;
  Result := True;
end;

procedure TOF_COPIEBUDMU.RempliPeriodeSource(Arow: Integer);
begin
  ListePeriode(TInfoBud(FListe.Objects[0,ARow]).ExoDeb, PerDebS.Items,PerDebS.Values,True);
  ListePeriode(TInfoBud(FListe.Objects[0,ARow]).ExoFin, PerFinS.Items,PerFinS.Values,False);
  PerDebS.Value := TInfoBud(FListe.Objects[0,ARow]).PerDeb;
  PerFinS.Value := TInfoBud(FListe.Objects[0,ARow]).PerFin;
end;

function TOF_COPIEBUDMU.ChercheDateCompta(D: TDateTime; Exo: String; ARow: Integer): TDateTime;
var
  Ind : Integer;
begin
  if Exo = TInfoBud(FListe.Objects[0,ARow]).ExoDeb then Ind := PerDebS.Values.IndexOf(DateToStr(D))-PerDebS.ItemIndex
                                                   else Ind := PerDebS.Values.Count + PerFinS.Values.IndexOf(DateToStr(FinDeMois(D))) - PerDebS.ItemIndex;
  if (PerDebD.ItemIndex+Ind) <= (PerDebD.Values.Count-1) then Result := StrToDate(PerDebD.Values[Ind+PerDebD.ItemIndex])
                                                         else Result := DebutdeMois(StrToDate(PerFinD.Values[(Ind+PerFinD.ItemIndex) - (PerFinD.Values.Count-1)]));
end;

function TOF_COPIEBUDMU.ChercheExo(D: TDateTime): String;
begin
  if PerDebD.Values.IndexOf(DateToStr(D))>=0 then Result := ExoDebD.Value
                                             else Result := ExoFinD.Value;
end;

procedure TOF_COPIEBUDMU.RunLaCopie;
var
  Sql,St,Nature : String;
  i,j,k,ARow : Integer;
  TLoc, TCpy : TQuery;
begin
  NowFutur := NowH;
  if RgCopi.ItemIndex=0 then ExecuteSql('DELETE FROM BUDECR WHERE BE_BUDJAL="'+BudJal.Value+'"');

  for i:=0 to ListBud.Count-1 do begin
    ARow := 0;
    for j:=1 to FListe.RowCount-1 do
      if FListe.Cells[0,j]=ListBud.Strings[i] then begin
        ARow := j;
        Break;
      end;
    if ARow=0 then Continue;
    RempliPeriodeSource(ARow);
    Nature := TInfoBud(FListe.Objects[0,ARow]).NatPieceCpy;

    While Nature <>'' do begin
      St := ReadTokenSt(Nature);
      TLoc := OpenSql('SELECT DISTINCT BE_NUMEROPIECE FROM BUDECR WHERE BE_BUDJAL="'+FListe.Cells[0,ARow]+'" '+
                      'AND BE_NATUREBUD="'+St+'" AND BE_QUALIFPIECE="N" ORDER BY BE_NUMEROPIECE', True);
      try
        while not TLoc.EOF do begin
          Sql := 'BE_BUDJAL="'+FListe.Cells[0,ARow]
               + '" AND BE_NATUREBUD="'  + St
               + '" AND BE_NUMEROPIECE=' + TLoc.FindField('BE_NUMEROPIECE').AsString
               + ' AND BE_QUALIFPIECE="N"';
          InitMove(GetCountBUDECR(Sql),'');
          Sql := 'SELECT * FROM BUDECR WHERE ' + Sql + ' ORDER BY BE_DATECOMPTABLE';
          TCpy := OpenSQL(SQL, True); // VL 14012005 FQ 15182
          try
            if Not TCpy.Eof then begin
              Case RgCopi.ItemIndex of
                0,1 : InsereLesEnreg(TInfoBud(FListe.Objects[0,ARow]),ARow, TCpy);
                2   : MajLesEnreg(TInfoBud(FListe.Objects[0,ARow]),ARow, TCpy);
              end;
            end;
          finally
            Ferme(TCpy);
          end;
          TLoc.Next;
          FiniMove;
        end;
      finally
        Ferme(TLoc);
      end;
    end;
  end;

  MajNbEcr;

  if HM.Execute(5,'','')=mrYes then // Désirez-vous visualiser les écritures budgétaires ?
    VisuConsoBudget('G', BudJal.Value);
end;

procedure TOF_COPIEBUDMU.InsereLesEnreg(Info: TInfoBud; ARow: Integer; TCpy : TQuery);
var
  MM : String17;
  NextNumPiece : Longint;
  Exo : String;
  DCpta : TDateTime;
  D,C : Double;
  i : Integer;
begin
  NextNumPiece := GetNum(EcrBud,TInfoBud(BudJal.Values.Objects[BudJal.Values.IndexOf(BudJal.Value)]).Souche,MM,0);
  while not TCpy.EOF do begin
    DCpta := ChercheDateCompta(TCpy.FindField('BE_DATECOMPTABLE').AsDateTime, TCpy.FindField('BE_EXERCICE').AsString, ARow);
    Exo := ChercheExo(DCpta);
    D := TCpy.FindField('BE_DEBIT').AsFloat + ((TCpy.FindField('BE_DEBIT').AsFloat*Info.CoefCpy)/100);
    C := TCpy.FindField('BE_CREDIT').AsFloat + ((TCpy.FindField('BE_CREDIT').AsFloat*Info.CoefCpy)/100);
    MoveCur(False);
    RequeteInsertion(Info, NextNumPiece, Exo, DcPta, D, C, TCpy);
    TCpy.Next;
  end;
  SetIncNum(EcrBud, TInfoBud(BudJal.Values.Objects[BudJal.Values.IndexOf(BudJal.Value)]).Souche, NextNumPiece, 0);
end;

procedure TOF_COPIEBUDMU.RequeteInsertion(Info: TInfoBud; NextNumPiece: Integer; Exo: String; DCpta: TDateTime; D, C: Double; TCpy : TQuery);
var
  T : Tob;
  {$IFNDEF EAGLCLIENT}
  i : Integer;
  {$ENDIF EAGLCLIENT}
begin
  T := Tob.Create('BUDECR', nil, -1);
  {$IFDEF EAGLCLIENT}
  T.InitValeurs;
  T.Dupliquer(TCpy, False, True);
  T.VirtuelleToReelle('BUDECR');

  T.SetString('BE_BUDJAL', GetControlText('BUDJAL'));
  T.SetString('BE_EXERCICE', Exo);
  T.SetDateTime('BE_DATECOMPTABLE', DCpta);
  T.SetDateTime('BE_DATEMODIF', NowFutur);
  T.SetString('BE_AXE', Info.Axe);
  T.SetInteger('BE_NUMEROPIECE', NextNumPiece);
  T.SetDouble('BE_DEBIT', D);
  T.SetDouble('BE_CREDIT', C);
  T.InsertOrUpdateDB;
  {$ELSE}

   for i := 0 to TCpy.FieldCount - 1 do begin
     if TCpy.Fields[i].FieldName='BE_BUDJAL'        then T.PutValue('BE_BUDJAL', GetControlText('BUDJAL'))   else
     if TCpy.Fields[i].FieldName='BE_EXERCICE'      then T.PutValue('BE_EXERCICE', Exo) else
     if TCpy.Fields[i].FieldName='BE_DATECOMPTABLE' then T.PutValue('BE_DATECOMPTABLE', DCpta) else
     if TCpy.Fields[i].FieldName='BE_DATEMODIF'     then T.PutValue('BE_DATEMODIF', NowFutur) else
     if TCpy.Fields[i].FieldName='BE_AXE'           then T.PutValue('BE_AXE', Info.Axe) else
     if TCpy.Fields[i].FieldName='BE_NUMEROPIECE'   then T.PutValue('BE_NUMEROPIECE', NextNumPiece) else
     if TCpy.Fields[i].FieldName='BE_DEBIT'         then T.PutValue('BE_DEBIT', D) else
     if TCpy.Fields[i].FieldName='BE_CREDIT'        then T.PutValue('BE_CREDIT', C) else
     if TCpy.Fields[i].FieldName='BE_BLOCNOTE'      then begin
       if not TMemoField(TCpy.FindField('BE_BLOCNOTE')).IsNull then
         T.PutValue('BE_BLOCNOTE', TCpy.FindField('BE_BLOCNOTE').AsVariant);
     end
     else
       T.PutValue(TCpy.Fields[i].FieldName, TCpy.Fields[i].AsVariant )
   end;
  {$ENDIF EAGLCLIENT}

  T.Free;
end;

procedure TOF_COPIEBUDMU.MajLesEnreg(Info: TInfoBud; ARow: Integer; TCpy : TQuery);
var
  Exo    : String;
  lStSql : String;
  DCpta  : TDateTime;
  D,C    : Double;
  QMaj   : TQuery;
  i : Integer;
begin
  while not TCpy.EOF do begin
    MoveCur(False);
    DCpta := ChercheDateCompta(TCpy.FindField('BE_DATECOMPTABLE').AsDateTime, TCpy.FindField('BE_EXERCICE').AsString, ARow);
    Exo := ChercheExo(DCpta);
    D := TCpy.FindField('BE_DEBIT').AsFloat + ((TCpy.FindField('BE_DEBIT').AsFloat*Info.CoefCpy)/100);
    C := TCpy.FindField('BE_CREDIT').AsFloat + ((TCpy.FindField('BE_CREDIT').AsFloat*Info.CoefCpy)/100);

    // Requête paramétrée remplacée par un openSql
    lStSql := 'SELECT * FROM BUDECR '
            + ' WHERE BE_BUDGENE="'       + TCpy.FindField('BE_BUDGENE').AsString
            + '" AND BE_BUDJAL="'         + BudJal.Value
            + '" AND BE_EXERCICE="'       + Exo
            + '" AND BE_DATECOMPTABLE="'  + USDateTime(DCpta)
            + '" AND BE_BUDSECT="'        + TCpy.FindField('BE_BUDSECT').AsString
            + '" AND BE_AXE="'            + TCpy.FindField('BE_AXE').AsString
            + '" AND BE_NATUREBUD="'      + TCpy.FindField('BE_NATUREBUD').AsString
            + '" AND BE_QUALIFPIECE="'    + TCpy.FindField('BE_QUALIFPIECE').AsString
            + '" ';
    QMaj := OpenSQL(lstSQL, False);
    if Not QMaj.Eof then begin
      QMaj.Edit ;
      QMaj.FindField('BE_DEBIT').AsFloat        := QMaj.FindField('BE_DEBIT').AsFloat+D ;
      QMaj.FindField('BE_CREDIT').AsFloat       := QMaj.FindField('BE_CREDIT').AsFloat+C ;
      QMaj.FindField('BE_DATEMODIF').AsDateTime := NowFutur;
      QMaj.Post ;
      end
    else begin
      RequeteInsertion(Info, TCpy.FindField('BE_NUMEROPIECE').AsInteger, Exo, DCpta, D, C, TCpy);
    end ;
    Ferme(QMaj);
    TCpy.Next;
  end ;
end;

procedure TOF_COPIEBUDMU.BZoomClick(Sender: TObject);
begin
  VisuConsoBudget('G', Fliste.Cells[0,FListe.Row]);
end;

function TOF_COPIEBUDMU.GetCountBUDECR(StCond: String): Integer;
var
  QCount : TQuery;
begin
  Result := 0 ;
  QCount := OpenSQL('SELECT COUNT(*) TOTAL FROM BUDECR WHERE ' + StCond, True);
  if not QCount.Eof then Result := QCount.FindField('TOTAL').AsInteger;
  Ferme(QCount);
end;

Initialization
  registerclasses ( [ TOF_COPIEBUDMU ] ) ;
end.
