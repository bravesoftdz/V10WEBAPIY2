{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 21/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SUPPRBUDS ()
Mots clefs ... : TOF;SUPPRBUDS
*****************************************************************}
Unit SUPPRBUDS_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     eMul, uTob, Maineagl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     db, mul, FE_Main,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, utilPGI, Rapsuppr,
     BUDSECT_TOM, // FicheBudsect 
     UTOF ;

procedure SuppressionCpteBudS;

Type
  TOF_SUPPRBUDS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Mul : TFMul;
    HM, MsgDel : THMsgBox;
    Nblig     : Integer;
    TDelBudg, TNotDel : TList;
    Effacer   : Boolean;
    NotEffacer : Boolean;
    BudgSect, BudgAxe : String;
    CbAtt : TStringList;
    procedure BOuvrirClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);

    function  Detruit(St,StAx : String):Byte;
    procedure Degage;
    procedure RempliCbAtt;
    procedure MajListeCompte(St,Stax : String);
    procedure MajCroisementCompte(St,Stax : String);
    function  CoupeA2000(St : String) : String;
  end ;
Implementation

procedure SuppressionCpteBudS;
begin
  if not _BlocageMonoPoste(False, '', True) then Exit;
  AGLLanceFiche('CP', 'CPSUPPRBUDS', '', '', '');
end;

procedure TOF_SUPPRBUDS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDS.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDS.OnArgument (S : String ) ;
begin
  Inherited ;

  HM := THMsgBox.Create(Ecran);
  HM.Mess.Add('0;Suppression des sections budgétaires;Confirmez-vous la suppression des enregistrements sélectionnés ?;Q;YNC;N;C;');
  HM.Mess.Add('Cette section budgétaire comporte des écritures budgétaires');
  HM.Mess.Add('2;Suppression des sections budgétaires;Vous n''avez rien sélectionné.;E;O;O;O;');
  HM.Mess.Add('3;Suppression des sections budgétaires;Désirez-vous un compte rendu des sections détruites ?;Q;YNC;N;C;');
  HM.Mess.Add('4;Suppression des sections budgétaires;Désirez-vous un compte rendu des sections non détruites ?;Q;YNC;N;C;');
  HM.Mess.Add('Section supprimé');
  HM.Mess.Add('Section bugétaire en cours d''utilisation.');
  HM.Mess.Add('Aucun');
  HM.Mess.Add('élément sélectionné');
  HM.Mess.Add('éléments sélectionnés');
  HM.Mess.Add('Cette section budgétaire est une section d''attente.');
  HM.Mess.Add('11;');

  MsgDel := THMsgBox.Create(Ecran);
  MsgDel.Mess.Add('0;Suppression des sections budgétaires;Cette section comporte des écritures budgétaires. Les écritures de cette section seront aussi supprimées. Désirez-vous la supprimer ?;Q;YNC;N;C;');
  MsgDel.Mess.Add('1;Suppression des sections budgétaires;Vous ne pouvez pas supprimer cette section : elle est en cours d''utilisation.;W;O;O;O;');
  MsgDel.Mess.Add('2;Suppression des sections budgétaires;Vous ne pouvez pas supprimer cette section : section budgétaire d''attente.;W;O;O;O;');

  CbAtt := TStringList.Create;

  Mul := TFMul(Ecran);
  Mul.BOuvrir.OnClick := BOuvrirClick;
  Mul.FListe.OnDblClick := FListeDblClick;

  TDelBudg := TList.Create;
  TNotDel := TList.Create;
  RempliCbAtt;

  Ecran.HelpContext := 15139000;
end ;

procedure TOF_SUPPRBUDS.OnClose ;
begin
  Inherited ;

  HM.Free;
  MsgDel.Free;
  CbAtt.Free;

  TDelBudg.Clear;
  TDelBudg.Free;
  TNotDel.Clear;
  TNotDel.Free;
  _DeblocageMonoPoste(False, '', True);
end ;

procedure TOF_SUPPRBUDS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDS.BOuvrirClick(Sender: TObject);
var
  i : Integer;
  j : Byte;
  X : DelInfo;
  Code,Ax,Lib : String;
begin
  NbLig := Mul.Fliste.NbSelected;
  TDelBudg.Clear;
  TNotDel.Clear;
  if NbLig<=0 then begin
    HM.execute(2,'',''); // Vous n''avez rien sélectionné.
    Exit;
  end;
  if HM.Execute(0,'','')<>mrYes then Exit; // Confirmez-vous la suppression des enregistrements sélectionnés ?
  Effacer := False;
  NotEffacer := False;
  if NbLig>1 then begin
    for i:=0 to NbLig-1 do begin
      Mul.Fliste.GotoLeBookMark(i);
      Code := GetField('BS_BUDSECT');
      Ax := GetField('BS_AXE');
      Lib := GetField('BS_LIBELLE');
      j := Detruit(Code,Ax);
      if j<=0 then begin
        X := DelInfo.Create;
        X.LeCod := Code;
        X.LeLib := Lib;
        X.LeMess := HM.Mess[5]; // Compte supprimé
        TDelBudg.Add(X);
        Effacer := True;
        end
      else begin
        X := DelInfo.Create;
        X.LeCod := Code;
        X.LeLib := Lib;
        X.LeMess := HM.Mess[j];
        TNotDel.Add(X);
        NotEffacer := True;
      end
    end;
    end
  else begin
    Mul.Fliste.GotoLeBookMark(0);
    Code := GetField('BS_BUDSECT');
    Ax := GetField('BS_AXE');
    j := Detruit(Code,Ax);
    if j=6 then MsgDel.Execute(1,'',''); // 
    if j=10 then MsgDel.Execute(2,'','');
  end;
  if Effacer    then if HM.Execute(3,'','')=mrYes then RapportDeSuppression(TDelBudg,1);
  if NotEffacer then if HM.Execute(4,'','')=mrYes then RapportDeSuppression(TNotDel,1);
  Mul.BChercheClick(Nil);
end;

procedure TOF_SUPPRBUDS.FListeDblClick(Sender: TObject);
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit;
  FicheBudsect(nil, GetField('BS_AXE'), GetField('BS_BUDSECT'), taConsult, 0);
end;

function TOF_SUPPRBUDS.CoupeA2000(St: String): String;
var
  St1,St2 : String;
begin
  if Length(St)<=2000 then begin
    Result := St;
    Exit;
  end;
  St1 := '';
  St2 := '';
  while St<>'' do begin
    St1 := ReadTokenSt(St);
    if Length(St2+St1+';')>2000 then begin
      Result := St2 + '<<>>' + St1 + ';' + St;
      Exit;
      end
    else begin
      if St2<>'' then St2 := St2 + St1 + ';'
                 else St2 := St1 + ';';
    end;
  end;
end;

procedure TOF_SUPPRBUDS.Degage;
begin
  if ExecuteSQL('DELETE FROM BUDSECT WHERE BS_BUDSECT="'+BudgSect+'" AND BS_AXE="'+BudgAxe+'"')<>1 then V_PGI.IoError := oeUnknown;
end;

function TOF_SUPPRBUDS.Detruit(St,StAx : String):Byte;
begin
  Result := 0;
  if CbAtt.IndexOf(St+';'+StAx)<>-1 then begin
    Result := 10;
    Exit;
  end;
  if EstMouvementeBudsect(St,StAx) then begin
    if MsgDel.Execute(0,'','')<>mrYes then begin
      Result := 1;
      Exit;
    end;
  end;
  BudgSect := St;
  BudgAxe := StAx;
  if Transactions(Degage,5)<>oeOK then begin
    MessageAlerte(HM.Mess[6]);
    Result := 6;
    Exit;
  end;
  ExecuteSQL('DELETE FROM BUDECR WHERE BE_BUDSECT="'+BudgSect+'" And BE_AXE="'+BudgAxe+'"');
  MajCroisementCompte(BudgSect,BudgAxe);
  MajListeCompte(BudgSect,BudgAxe);
end;

procedure TOF_SUPPRBUDS.MajListeCompte(St,Stax : String);
var
  QLoc : TQuery;
  StC,St1,StTemp : String;
  Trouver : Boolean;
  nombudget : string; //fb 28/09/2005 FQ 16046
begin
{b fb 28/09/2005 FQ 16046}
(*  QLoc := OpenSQL('SELECT BJ_BUDSECTS,BJ_BUDSECTS2 FROM BUDJAL WHERE BJ_AXE="'+Stax+'"', False);
  while not QLoc.Eof do begin
    Trouver := False;
    St1 := '';
    if QLoc.Fields[1].AsString<>'' then StC := QLoc.Fields[0].AsString + QLoc.Fields[1].AsString
                                   else StC := QLoc.Fields[0].AsString;
    while StC<>'' do begin
      StTemp := ReadTokenSt(StC);
      if StTemp=St then Trouver := True
                   else St1 := St1 + StTemp + ';';
    end;

    if Trouver then begin
      StTemp := CoupeA2000(St1);
      QLoc.Edit;
      if Pos('<<>>',StTemp)<=0 then QLoc.Fields[0].AsString:=StTemp
      else begin
        QLoc.Fields[0].AsString := Copy(StTemp,1,Pos('<<>>',StTemp)-1);
        QLoc.Fields[1].AsString := Copy(StTemp,Pos('<<>>',StTemp)+1,Length(StTemp));
      end;
      QLoc.Post;
    end;
    QLoc.Next;
  end;
  Ferme(QLoc);*)

  QLoc := OpenSQL('SELECT BJ_BUDJAL, BJ_BUDSECTS,BJ_BUDSECTS2 FROM BUDJAL WHERE BJ_AXE="'+Stax+'"', False);

  while not QLoc.Eof do begin
    Trouver := False;
    St1 := '';

    nombudget:=QLoc.findfield('BJ_BUDJAL').AsString;
    if QLoc.findfield('BJ_BUDSECTS2').AsString<>'' then
      StC := QLoc.findfield('BJ_BUDSECTS').AsString + QLoc.findfield('BJ_BUDSECTS2').AsString
    else
      StC := QLoc.findfield('BJ_BUDSECTS').AsString;

    while StC<>'' do begin
      StTemp := ReadTokenSt(StC);
      if StTemp<>' ' then
        if StTemp=St then Trouver := True
                     else St1 := St1 + StTemp + ';';
    end;

    if Trouver then begin
      StTemp := CoupeA2000(St1);
      if Pos('<<>>',StTemp)<=0 then
         EXECUTESQL('UPDATE BUDJAL SET BJ_BUDSECTS="' + sttemp +
         '" WHERE BJ_AXE="' + Stax + '" AND BJ_BUDJAL="' + nombudget + '"')
      else begin
         EXECUTESQL('UPDATE BUDJAL SET BJ_BUDSECTS="' + Copy(StTemp,1,Pos('<<>>',StTemp)-1) +
         '", BJ_BUDSECTS2="' + Copy(StTemp,Pos('<<>>',StTemp)+1,Length(StTemp))+
         '" WHERE BJ_AXE="' + Stax + '" AND BJ_BUDJAL="' + nombudget + '"');
      end;
    end;
    QLoc.Next;
  end;
  Ferme(QLoc);
{e fb 28/09/2005 FQ 16046}
end;

procedure TOF_SUPPRBUDS.RempliCbAtt;
var
  QLoc : TQuery;
begin
  CbAtt.Clear;
  QLoc := OpenSQL('SELECT BS_BUDSECT,BS_AXE FROM BUDSECT WHERE BS_ATTENTE="X"', True);
  while not QLoc.Eof do begin
    CbAtt.Add(QLoc.Fields[0].AsString + ';' + QLoc.Fields[1].AsString);
    QLoc.Next;
  end;
  Ferme(QLoc);
end;

procedure TOF_SUPPRBUDS.MajCroisementCompte(St, Stax: String);
var
  QLoc : TQuery;
  StC,StTemp : String;
begin
  QLoc := OpenSQL('SELECT BJ_BUDJAL,BJ_BUDSECTS FROM BUDJAL WHERE BJ_AXE="'+Stax+'"', True);
  while not QLoc.Eof do begin
    StC := QLoc.Fields[1].AsString;
    while StC<>'' do begin
      StTemp := ReadTokenSt(StC);
      if StTemp=St then begin
        ExecuteSql('DELETE FROM CROISCPT WHERE CX_SECTION="'+St+'" AND CX_JAL="'+QLoc.Fields[0].AsString+'" AND CX_TYPE="BUD"');
        Break;
      end;
    end;
    QLoc.Next;
  end;
  Ferme(QLoc);
end;

Initialization
  registerclasses ( [ TOF_SUPPRBUDS ] ) ; 
end.
