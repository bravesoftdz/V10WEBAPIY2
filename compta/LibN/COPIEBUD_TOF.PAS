{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 28/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : COPIEBUD ()
Mots clefs ... : TOF;COPIEBUD
*****************************************************************}
Unit COPIEBUD_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
  eMul, MaineAGL,
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  db, mul, FE_Main,
{$ENDIF}
  Vierge, forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, HStatus,
  uTob, UTOF;

procedure RecopieBudgetSimple;

Type
  TOF_COPIEBUD = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    NumPiece : String;
    NextNumPiece : Longint;
    CbS, CbD : THValComboBox;
    PerFinS, PerDebS : THValComboBox;
    PerFinD, PerDebD : THValComboBox;
    HM : THMsgBox;

    procedure BudSChange(Sender: TObject);
    procedure NatSChange(Sender: TObject);
    procedure BValiderClick(Sender: TObject);

    procedure ChargeInfosJournal(Sender : TObject);
    procedure ChageInfoPeriode(Pref : String ; Q : TQuery);
    procedure ChercheNbEcr;
    function  ChargeLesPeriodes : Boolean;
    function  ControleCopieOk : Boolean;
    function  ControleJalOk : Boolean;
    function  NumPieceOk : Boolean;
    function  ControleNatOk : Boolean;
    function  ControleComptesOk : Boolean;
    function  ControleAxeOk : Boolean;
    function  ControleCptAttOk : Boolean;
    function  ChercheLaDateCompta(D : TDateTime) : TDateTime;
    procedure RunLaCopie;
  end ;

Implementation

uses SAISUTIL; // GetNum

procedure RecopieBudgetSimple;
begin
  AGLLanceFiche('CP','CPCOPIEBUD','','','');
end;

procedure TOF_COPIEBUD.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_COPIEBUD.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_COPIEBUD.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_COPIEBUD.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_COPIEBUD.OnArgument (S : String ) ;
begin
  Inherited ;

  HM := THMsgBox.Create(Ecran);
  HM.Mess.Add('0;Copie de budget;Désirez-vous copier le budget sélectionné ?;Q;YN;N;N;');
  HM.Mess.Add('1;Copie de budget;Le journal source comporte aucune ligne d''écriture.;W;O;O;O;');
  HM.Mess.Add('2;Copie de budget;Aucun journal a été sélectionné.;W;O;O;O;');
  HM.Mess.Add('3;Copie de budget;Aucune nature de budget a été sélectionnée.;W;O;O;O;');
  HM.Mess.Add('4;Copie de budget;Le nombre de période du journal de destination est différent du journal source.;W;O;O;O;');
  HM.Mess.Add('5;Copie de budget;Les comptes généraux budgétaires sont différents. Vous devez choisir un autre journal de destination.;W;O;O;O;');
  HM.Mess.Add('6;Copie de budget;Les sections budgétaires sont différentes. Vous devez choisir un autre journal de destination.;W;O;O;O;');
  HM.Mess.Add('7;Copie de budget;L''axe du journal de destination est différent de l''axe du journal source. Vous devez choisir un autre journal de destination.;W;O;O;O;');
  HM.Mess.Add('Insertion des enregistrements en cours');
  HM.Mess.Add('Nombre de pièce trouvée');
  HM.Mess.Add('Nombre de pièces trouvées');
  HM.Mess.Add('11;Copie de budget;Le compte général d''attente du journal de destination est différent du compte général d''attente du journal source. Vous devez choisir un autre journal de destination.;W;O;O;O;');
  HM.Mess.Add('12;Copie de budget;La section budgétaire d''attente du journal de destination est différent de la section budgétaire d''attente du journal source. Vous devez choisir un autre journal de destination.;W;O;O;O;');
  HM.Mess.Add('13;Copie de budget;Le coefficient à appliquer pour la copie doit être une valeur numérique.;W;O;O;O;');
  HM.Mess.Add('14;Copie de budget;Incohérence dans le choix des périodes du budget source.;W;O;O;O;');
  HM.Mess.Add('15;Copie de budget;Incohérence dans le choix des périodes du budget destination.;W;O;O;O;');
  HM.Mess.Add('16;Copie de budget;Le nombre de période du budget destination est différente du nombre de période du budget source.;W;O;O;O;');
  HM.Mess.Add('Recopie non effectuée;');
  HM.Mess.Add('18;Copie de budget;La recopie de budget s''est correctement effectuée.;E;O;O;O;');

  CbS := THValComboBox(GetControl('CBS', True));
  CbD := THValComboBox(GetControl('CBD', True));

  PerFinS := THValComboBox(GetControl('PERFINS', True));
  PerDebS := THValComboBox(GetControl('PERDEBS', True));
  PerFinD := THValComboBox(GetControl('PERFIND', True));
  PerDebD := THValComboBox(GetControl('PERDEBD', True));

  THValComboBox(GetControl('BUDS', True)).OnChange := BudSChange;
  THValComboBox(GetControl('BUDD', True)).OnChange := BudSChange;
  THValComboBox(GetControl('NATS', True)).OnChange := NatSChange;
  TFVierge(Ecran).BValider.OnClick := BValiderClick;

  NumPiece := '';
  SetControlCaption('NBECR', '0');
  SetControlText('NATS', 'INI');
  SetControlText('COEF', '0');
end ;

procedure TOF_COPIEBUD.OnClose ;
begin
  Inherited ;

  HM.Free;
end ;

procedure TOF_COPIEBUD.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_COPIEBUD.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_COPIEBUD.ChargeInfosJournal(Sender: TObject);
var
  QLoc  : TQuery;
  Cb : THValComboBox;
  Pref : String;
begin
  Cb := THValComboBox(Sender);
  Pref := Copy(Cb.Name, Length(Cb.Name), 1);
  QLoc := OpenSQL('SELECT * FROM BUDJAL WHERE BJ_BUDJAL="'+Cb.Value+'"',True);

  if not QLoc.EOF then begin
    THValComboBox(GetControl('ExoDeb'+Pref, True)).Value := QLoc.FindField('BJ_EXODEB').AsString;
    THValComboBox(GetControl('ExoFin'+Pref, True)).Value := QLoc.FindField('BJ_EXOFIN').AsString;
    THValComboBox(GetControl('Ax'+Pref, True)).Value := QLoc.FindField('BJ_AXE').AsString;
    THEdit(GetControl('Gatt'+Pref, True)).Text := QLoc.FindField('BJ_GENEATTENTE').AsString;
    THEdit(GetControl('Satt'+Pref, True)).Text := QLoc.FindField('BJ_SECTATTENTE').AsString;
    if QLoc.FindField('BJ_BUDGENES2').AsString<>'' then THEdit(GetControl('Gen'+Pref, True)).Text := QLoc.FindField('BJ_BUDGENES').AsString+QLoc.FindField('BJ_BUDGENES2').AsString
                                                   else THEdit(GetControl('Gen'+Pref, True)).Text := QLoc.FindField('BJ_BUDGENES').AsString ;
    if QLoc.FindField('BJ_BUDSECTS2').AsString<>'' then THEdit(GetControl('Sec'+Pref, True)).Text := QLoc.FindField('BJ_BUDSECTS').AsString+QLoc.FindField('BJ_BUDSECTS2').AsString
                                                   else ThEdit(GetControl('Sec'+Pref, True)).Text := QLoc.FindField('BJ_BUDSECTS').AsString ;
    THEdit(GetControl('Sh'+Pref, True)).Text := QLoc.FindField('BJ_COMPTEURNORMAL').AsString;
    ChageInfoPeriode(Pref,QLoc);
  end;
  Ferme(QLoc);
end;

procedure TOF_COPIEBUD.ChageInfoPeriode(Pref: String; Q: TQuery);
var
  DDeb,DFin,DTemp : TDateTime;
  LDat : String;
begin
  DDeb := Q.FindField('BJ_PERDEB').AsDateTime;
  DFin := Q.FindField('BJ_PERFIN').AsDateTime;

  THValComboBox(GetControl('PerDeb'+Pref, True)).Values.Clear;
  THValComboBox(GetControl('PerDeb'+Pref, True)).Items.Clear;
  THValComboBox(GetControl('PerFin'+Pref, True)).Values.Clear;
  THValComboBox(GetControl('PerFin'+Pref, True)).Items.Clear;

  Repeat
    DTemp := FindeMois(DDeb);
    LDat := FormatDateTime('mmmm yyyy',DDeb);
    LDat := FirstMajuscule(LDat);
    THValComboBox(GetControl('PerDeb'+Pref, True)).Items.Add(LDat);
    THValComboBox(GetControl('PerDeb'+Pref, True)).Values.Add(DateToStr(DDeb));
    LDat := FormatDateTime('mmmm yyyy',DTemp) ; LDat:=FirstMajuscule(LDat);
    THValComboBox(GetControl('PerFin'+Pref, True)).Items.Add(LDat);
    THValComboBox(GetControl('PerFin'+Pref, True)).Values.Add(DateToStr(DTemp));
    DDeb := PlusMois(DDeb,1);
  Until DTemp>=DFin ;

  if THValComboBox(GetControl('PerDeb'+Pref, True)).Values.Count>0 then
     THValComboBox(GetControl('PerDeb'+Pref, True)).Value := THValComboBox(GetControl('PerDeb'+Pref, True)).Values[0];
  if THValComboBox(GetControl('PerFin'+Pref, True)).Values.Count>0 then
     THValComboBox(GetControl('PerFin'+Pref, True)).Value := THValComboBox(GetControl('PerFin'+Pref, True)).Values[THValComboBox(GetControl('PerFin'+Pref, True)).Values.Count-1];
end;

procedure TOF_COPIEBUD.BudSChange(Sender: TObject);
begin
  ChargeInfosJournal(Sender);
  if THValComboBox(Sender).Name = 'BUDS' then ChercheNbEcr;
end;

procedure TOF_COPIEBUD.ChercheNbEcr;
var
  i : Integer;
  Q : TQuery;
begin
  Q := OpenSQL('SELECT DISTINCT BE_NUMEROPIECE FROM BUDECR WHERE BE_BUDJAL="'+GetControlText('BUDS')+'" AND BE_NATUREBUD="'+GetControlText('NATS')+'"', True);

  i := 0;
  NumPiece := '';
  While Not Q.Eof do begin
    NumPiece := NumPiece + IntToStr(Q.Fields[0].AsInteger)+';';
    Inc(i);
    Q.Next;
  end;
  SetControlCaption('NBECR', IntToStr(i));
  Ferme(Q);
  if i>1 then SetControlCaption('NBE', HM.Mess[10]) // Nombre de pièces trouvées
         else SetControlCaption('NBE', HM.Mess[9]);  // Nombre de pièce trouvée
end;

procedure TOF_COPIEBUD.NatSChange(Sender: TObject);
begin
  if GetControlText('BUDS')='' then Exit;
  ChercheNbEcr ;
end;

procedure TOF_COPIEBUD.BValiderClick(Sender: TObject);
var
  io : TIoErr;
begin
  if HM.Execute(0,'','')<>mrYes then Exit; // Désirez-vous copier le budget sélectionné ?
  if Not ControleCopieOk then Exit ;

  io:=Transactions(RunLaCopie,2) ;
  if io<>oeOk then PGIError(HM.Mess[17]) // Recopie non effectuée
              else HM.Execute(18,'',''); // La recopie de budget s''est correctement effectuée.
end;

function TOF_COPIEBUD.ChargeLesPeriodes: Boolean;
var
  i : Integer;
begin
  Result := False;
  CbS.Values.Clear;
  CbS.Items.Clear;
  CbD.Values.Clear;
  CbD.Items.Clear;

  for i := PerDebS.ItemIndex to PerDebS.Values.Count-1 do
    if DebutdeMois(StrToDate(PerFinS.Values[PerFinS.ItemIndex]))>=StrToDate(PerDebS.Values[i]) then begin
      CbS.Values.Add(PerDebS.Values[i]);
      CbS.Items.Add(PerDebS.Items[i]);
    end;

  if StrToDate(CbS.Values[CbS.Values.Count-1])<DebutdeMois(StrToDate(PerFinS.Values[0])) then
    for i:=0 to PerFinS.ItemIndex do begin
      CbS.Values.Add(DateToStr(DebutDeMois(StrToDate(PerFinS.Values[i]))));
      CbS.Items.Add(PerFinS.Items[i]);
    end;

  for i:=PerDebD.ItemIndex to PerDebD.Values.Count-1 do
    if DebutdeMois(StrToDate(PerFinD.Values[PerFinD.ItemIndex]))>=StrToDate(PerDebD.Values[i]) then begin
      CbD.Values.Add(PerDebD.Values[i]);
      CbD.Items.Add(PerDebD.Items[i]);
    end;

  if StrToDate(CbD.Values[CbD.Values.Count-1])<DebutdeMois(StrToDate(PerFinD.Values[0])) then
    for i:=0 to PerFinD.ItemIndex do begin
      CbD.Values.Add(DateToStr(DebutDeMois(StrToDate(PerFinD.Values[i]))));
      CbD.Items.Add(PerFinD.Items[i]);
    end;

  if CbS.Values.Count<>CbD.Values.Count then begin
    HM.Execute(16,'',''); // Le nombre de période du budget destination est différente du nombre de période du budget source.
    PerFinD.SetFocus;
    Exit;
  end;
  Result := True;
end;

function TOF_COPIEBUD.ControleCopieOk: Boolean;
begin
  Result := False;
  if Not ControleJalOk     then Exit;
  if Not NumPieceOk        then Exit;
  if Not ControleNatOk     then Exit;
  if Not ControleComptesOk then Exit;
  if Not ControleAxeOk     then Exit;
  if Not ControleCptAttOk  then Exit;
  if Not ChargeLesPeriodes then Exit;
  Result:=True; 
end;

function TOF_COPIEBUD.ControleJalOk: Boolean;
begin
  Result := False;
  if GetControlText('BUDS')='' then begin
    HM.Execute(2,'','');  // Aucun journal a été sélectionné.
    SetFocusControl('BUDS');
    Exit;
  end;
  if GetControlText('BUDD')='' then begin
    HM.Execute(2,'','');  // Aucun journal a été sélectionné.
    SetFocusControl('BUDD');
    Exit;
  end;
  Result := True;
end;

function TOF_COPIEBUD.NumPieceOk: Boolean;
begin
  Result := False;
  if NumPiece='' then begin
    HM.Execute(1,'',''); // Le journal source comporte aucune ligne d''écriture.
    SetFocusControl('BUDS');
    Exit ;
  end;
  Result := True;
end;

function TOF_COPIEBUD.ControleNatOk: Boolean;
begin
  Result := False;
  if GetControlText('NATS')='' then begin
    HM.Execute(3,'',''); // Aucune nature de budget a été sélectionnée.
    SetFocusControl('NATS');
    Exit;
  end;
  if GetControlText('NATD')='' then begin
    HM.Execute(3,'',''); // Aucune nature de budget a été sélectionnée.
    SetFocusControl('NATD');
    Exit;
  end;
  Result := True;
end;

function TOF_COPIEBUD.ControleComptesOk: Boolean;
var
  StGS, StSS, St, GenS, GenD, SecS, SecD : String;
  i : Integer;
begin
  Result := False;
  GenS := GetControlText('GENS');
  GenD := GetControlText('GEND');
  if Length(GenS)<>Length(GenD) then begin
    HM.Execute(5,'',''); // Les comptes généraux budgétaires sont différents. Vous devez choisir un autre journal de destination.
    SetFocusControl('BUDD');
    Exit;
  end;

  SecS := GetControlText('SECS');
  SecD := GetControlText('SECD');
  if Length(SecS)<>Length(SecD) then begin
    HM.Execute(6,'',''); // Les sections budgétaires sont différentes. Vous devez choisir un autre journal de destination.
    SetFocusControl('BUDD');
    Exit;
  end;
  StGS := Trim(GenS);
  StSS := Trim(SecS);

  While StGS<>'' do begin
    St := ReadTokenSt(StGS)+';';
    i := Pos(St,GenD);
    if i<=0 then begin
      HM.Execute(5,'',''); // Les comptes généraux budgétaires sont différents. Vous devez choisir un autre journal de destination.
      SetFocusControl('BUDD');
      Exit;
    end;
  end;

  While StSS<>'' do begin
    St := ReadTokenSt(StSS)+';';
    i := Pos(St,SecD);
    if i<=0 then begin
      HM.Execute(6,'',''); // Les sections budgétaires sont différentes. Vous devez choisir un autre journal de destination.
      SetFocusControl('BUDD');
      Exit;
    end;
  end;
  Result := True;
end;

function TOF_COPIEBUD.ControleAxeOk: Boolean;
begin
  Result := False;
  if GetControlText('AXS')<>GetControlText('AXD') then begin
    HM.Execute(7,'',''); // L''axe du journal de destination est différent de l''axe du journal source. Vous devez choisir un autre journal de destination.
    SetFocusControl('BUDD');
    Exit;
  end;
  Result := True;
end;

function TOF_COPIEBUD.ControleCptAttOk: Boolean;
var
  GattS, GattD, SattS, SattD : String;
begin
  Result := False;

  GattS := GetControlText('GATTS');
  GattD := GetControlText('GATTD');
  if GattS<>GattD then begin
    HM.Execute(11,'','');    // Le compte général d''attente du journal de destination est différent du compte général d''attente du journal source.
    SetFocusControl('BUDD'); // Vous devez choisir un autre journal de destination.
    Exit;
  end;

  SattS := GetControlText('SATTS');
  SattD := GetControlText('SATTD');
  if SattS<>SattD then begin
    HM.Execute(12,'','');     // La section budgétaire d''attente du journal de destination est différent de la section budgétaire d''attente du journal source.
    SetFocusControl('BUDD'); // Vous devez choisir un autre journal de destination.
    Exit;
  end;
  Result := True;
end;

function TOF_COPIEBUD.ChercheLaDateCompta(D: TDateTime): TDateTime;
begin
  Result := StrToDate(CbD.Values[CbS.Values.IndexOf(DateToStr(D))]);
end;

procedure TOF_COPIEBUD.RunLaCopie;
var
  MM : String17;
  i : Integer;
  Exo,MemoNumPiece : String;
  DCpta,NowFutur : TDateTime;
  D,C,UnCoef : Double;
  T  : Tob;
  TLoc : TQuery;
begin
  MM := '';
  MemoNumPiece := NumPiece;
  NowFutur := NowH;

  T := Tob.Create('BUDECR', nil, -1);
  While NumPiece<>'' do begin
    i := 0;
    NextNumPiece := GetNum(EcrBud, GetControlText('SHD'), MM, 0);
    TLoc := OpenSql('SELECT * FROM BUDECR WHERE BE_BUDJAL="'+GetControlText('BUDS')+'" '+
                    'AND BE_NATUREBUD="'+GetControlText('NATS')+'" '+
                    'AND BE_QUALIFPIECE="N" AND BE_NUMEROPIECE="'+ReadTokenSt(NumPiece)+'" '+
                    'AND BE_DATECOMPTABLE>="'+USDateTime(StrToDate(PerDebS.Values[PerDebS.ItemIndex]))+'" AND '+
                    'BE_DATECOMPTABLE<="'+UsDateTime(DebutdeMois(StrToDate(PerFinS.Values[PerFinS.ItemIndex])))+'" ORDER BY BE_DATECOMPTABLE', True);
    InitMove(RecordsCount(TLoc),HM.Mess[8]); // Insertion des enregistrements en cours

    While Not TLoc.Eof do begin
      DCpta := ChercheLaDateCompta(TLoc.FindField('BE_DATECOMPTABLE').AsDateTime);
      Exo := QuelExoDt(DCpta);
      UnCoef := StrToFloat(GetControlText('COEF'));
      D := TLoc.FindField('BE_DEBIT').AsFloat+((TLoc.FindField('BE_DEBIT').AsFloat*UnCoef)/100);
      C := TLoc.FindField('BE_CREDIT').AsFloat+((TLoc.FindField('BE_CREDIT').AsFloat*UnCoef)/100);

      T.InitValeurs;
      {$IFDEF EAGLCLIENT}
      T.Dupliquer(TLoc.Detail[i], False, True);
      T.VirtuelleToReelle('BUDECR');

      T.SetString('BE_BUDJAL', GetControlText('BUDD'));
      T.SetString('BE_EXERCICE', Exo);
      T.SetDateTime('BE_DATECOMPTABLE', DCpta);
      T.SetString('BE_AXE', GetControlText('AXD'));
      T.SetInteger('BE_NUMEROPIECE', NextNumPiece);
      T.SetDouble('BE_DEBIT', D);
      T.SetDouble('BE_CREDIT', C);
      T.SetDateTime('BE_DATEMODIF', NowFutur);
      T.SetString('BE_NATUREBUD', GetControlText('NATD'));
      inc(i);
      {$ELSE}
       for i := 0 to TLoc.FieldCount - 1 do begin
         if TLoc.Fields[i].FieldName='BE_BUDJAL'        then T.PutValue('BE_BUDJAL', GetControlText('BUDD'))   else
         if TLoc.Fields[i].FieldName='BE_EXERCICE'      then T.PutValue('BE_EXERCICE', Exo) else
         if TLoc.Fields[i].FieldName='BE_DATECOMPTABLE' then T.PutValue('BE_DATECOMPTABLE', DCpta) else
         if TLoc.Fields[i].FieldName='BE_DATEMODIF'     then T.PutValue('BE_DATEMODIF', NowFutur) else
         if TLoc.Fields[i].FieldName='BE_AXE'           then T.PutValue('BE_AXE', GetControlText('AXD')) else
         if TLoc.Fields[i].FieldName='BE_NUMEROPIECE'   then T.PutValue('BE_NUMEROPIECE', NextNumPiece) else
         if TLoc.Fields[i].FieldName='BE_DEBIT'         then T.PutValue('BE_DEBIT', D) else
         if TLoc.Fields[i].FieldName='BE_CREDIT'        then T.PutValue('BE_CREDIT', C) else
         if TLoc.Fields[i].FieldName='BE_NATUREBUD'     then T.PutValue('BE_NATUREBUD', GetControlText('NATD')) else
         if TLoc.Fields[i].FieldName='BE_BLOCNOTE'      then begin
           if not TMemoField(TLoc.FindField('BE_BLOCNOTE')).IsNull then
             T.PutValue('BE_BLOCNOTE', TLoc.FindField('BE_BLOCNOTE').AsVariant);
         end
         else
           T.PutValue(TLoc.Fields[i].FieldName, TLoc.Fields[i].AsVariant )
       end;
      {$ENDIF EAGLCLIENT}
      T.InsertOrUpdateDB;
      MoveCur(False);
      TLoc.Next;
    end;
    SetIncNum(EcrBud, GetControlText('SHD'), NextNumPiece, 0);
    FiniMove;
    Ferme(TLoc);
  end;
  NumPiece := MemoNumPiece;
  T.Free;
end;

Initialization
  registerclasses ( [ TOF_COPIEBUD ] ) ;
end.
