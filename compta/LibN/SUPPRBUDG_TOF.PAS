{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 19/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SUPPRBUDG ()
Mots clefs ... : TOF;SUPPRBUDG
*****************************************************************}
Unit SUPPRBUDG_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     eMul, uTob, Maineagl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     db, mul, FE_Main,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, utilPGI, Rapsuppr,
     BudGene_TOM, // FicheBudgene 
     UTOF ;

procedure SuppressionCpteBudG;

Type
  TOF_SUPPRBUDG = Class (TOF)
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
    BudgCode  : String;
    CbAtt : TStringList;
    procedure BOuvrirClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);

    function  Detruit(St : String): Byte;
    procedure Degage;
    procedure RempliCbAtt;
    procedure MajListeCompte(St : String);
    function  CoupeA2000(St : String) : String;
  end ;

Implementation

procedure SuppressionCpteBudG;
begin
  if not _BlocageMonoPoste(False) then Exit;
  AGLLanceFiche('CP', 'CPSUPPRBUDG', '', '', '');
end;

procedure TOF_SUPPRBUDG.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDG.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDG.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDG.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDG.OnArgument (S : String ) ;
begin
  Inherited ;

  HM := THMsgBox.Create(Ecran);
  HM.Mess.Add('0;Suppression des comptes budgétaires;Confirmez-vous la suppression des enregistrements sélectionnés ?;Q;YN;N;N;');
  HM.Mess.Add('Ce compte budgétaire comporte des écritures budgétaires');
  HM.Mess.Add('2;Suppression des comptes budgétaires;Vous n''avez rien sélectionné.;E;O;O;O;');
  HM.Mess.Add('3;Suppression des comptes budgétaires;Désirez-vous un compte rendu des comptes détruits ?;Q;YNC;N;C;');
  HM.Mess.Add('4;Suppression des comptes budgétaires;Désirez-vous un compte rendu des comptes non détruits ?;Q;YNC;N;C;');
  HM.Mess.Add('Compte supprimé');
  HM.Mess.Add('Compte bugétaire en cours d''utilisation.');
  HM.Mess.Add('Aucun');
  HM.Mess.Add('élément sélectionné');
  HM.Mess.Add('éléments sélectionnés');
  HM.Mess.Add('Ce compte budgétaire est un compte d''attente');
  HM.Mess.Add('11;');

  MsgDel := THMsgBox.Create(Ecran);
  MsgDel.Mess.Add('0;Suppression des comptes budgétaires;Ce compte comporte des écritures budgétaires. Les écritures de ce compte seront aussi supprimées. Désirez-vous le supprimer ?;Q;YNC;N;C;');
  MsgDel.Mess.Add('1;Suppression des comptes budgétaires;Vous ne pouvez pas supprimer ce compte : il est en cours d''utilisation.;W;O;O;O;');
  MsgDel.Mess.Add('2;Suppression des comptes budgétaires;Vous ne pouvez pas supprimer ce compte : compte budgétaire d''attente.;W;O;O;O;');

  CbAtt := TStringList.Create;

  Mul := TFMul(Ecran);
  Mul.BOuvrir.OnClick := BOuvrirClick;
  Mul.FListe.OnDblClick := FListeDblClick;

  TDelBudg := TList.Create;
  TNotDel := TList.Create;
  RempliCbAtt;

  Ecran.HelpContext := 15119000;
end ;

procedure TOF_SUPPRBUDG.OnClose ;
begin
  Inherited ;

  HM.Free;
  MsgDel.Free;
  CbAtt.Free;

  TDelBudg.Clear;
  TDelBudg.Free;
  TNotDel.Clear;
  TNotDel.Free;
  _DeblocageMonoPoste(True);
end ;

procedure TOF_SUPPRBUDG.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDG.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDG.BOuvrirClick(Sender: TObject);
var
  i : Integer;
  j : Byte;
  X : DelInfo;
  Code,Lib : String;
begin
  NbLig := Mul.FListe.NbSelected;
  TDelBudg.Clear;
  TNotDel.Clear;
  if NbLig<=0 then begin
    HM.Execute(2,'',''); // Vous n''avez rien sélectionné.
    Exit;
  end;
  if HM.Execute(0,'','')<>mrYes then Exit; // Confirmez-vous la suppression des enregistrements sélectionnés ?
  Effacer := False;
  NotEffacer := False;
  if NbLig>1 then begin
    for i:=0 to NbLig-1 do begin
      Mul.Fliste.GotoLeBookMark(i);
      Code := GetField('BG_BUDGENE');
      Lib := GetField('BG_LIBELLE');
      j := Detruit(Code);
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
    Code := GetField('BG_BUDGENE');
    j := Detruit(Code);
    if j=6 then MsgDel.Execute(1,'','');
    if j=10 then MsgDel.Execute(2,'','');
  end;

  if Effacer    then if HM.Execute(3,'','')=mrYes then RapportDeSuppression(TDelBudg,1);
  if NotEffacer then if HM.Execute(4,'','')=mrYes then RapportDeSuppression(TNotDel,1);
  Mul.BChercheClick(Nil);
end;

procedure TOF_SUPPRBUDG.FListeDblClick(Sender: TObject);
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit;
  FicheBudgene(nil, '', GetField('BG_BUDGENE'), taConsult, 0);
end;

function TOF_SUPPRBUDG.CoupeA2000(St: String): String;
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
      Result := St2 + '<<>>' + St1 + ';' +St;
      Exit;
      end
    else begin
      if St2<>'' then St2 := St2 + St1 + ';'
                 else St2 := St1 + ';';
    end;
  end;
end;

procedure TOF_SUPPRBUDG.Degage;
begin
  if ExecuteSQL('DELETE FROM BUDGENE WHERE BG_BUDGENE="'+BudgCode+'"')<>1 then V_PGI.IoError:=oeUnknown;
end;

function TOF_SUPPRBUDG.Detruit(St: String): Byte;
begin
  Result:=0 ;
  if CbAtt.IndexOf(St)<>-1 then begin
    Result := 10;
    Exit;
  end;
  if EstMouvementeBudgen(St)then begin
    if MsgDel.Execute(0,'','')<>mrYes then begin
      Result := 1;
      Exit;
    end;
  end;
  BudgCode := St;
  if Transactions(Degage,5)<>oeOK then begin
    MessageAlerte(HM.Mess[6]);
    Result := 6;
    Exit;
  end;
  ExecuteSQL('DELETE FROM BUDECR WHERE BE_BUDGENE="'+BudgCode+'"');
  ExecuteSql('Delete From CROISCPT Where CX_COMPTE="'+BudgCode+'" And CX_TYPE="BUD"');
  MajListeCompte(BudgCode);
end;

procedure TOF_SUPPRBUDG.MajListeCompte(St: String);
Var
  QLoc : TQuery;
  StC,St1,StTemp : String;
  Trouver : Boolean;
  nombudget : string; //fb 28/09/2005 FQ 16043
begin
{b fb 28/09/2005 FQ 16043}
(*  QLoc := OpenSql('SELECT BJ_BUDGENES,BJ_BUDGENES2 FROM BUDJAL',False);

  while not QLoc.Eof do begin
    Trouver := False;
    St1 := '';
    if QLoc.Fields[1].AsString<>'' then StC := QLoc.Fields[0].AsString + QLoc.Fields[1].AsString
                                   else StC := QLoc.Fields[0].AsString;
    while StC<>'' do begin
      StTemp := ReadTokenSt(StC);
        if StTemp=St then Trouver:=True
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
  Ferme(QLoc);                           *)
  
  QLoc := OpenSql('SELECT BJ_BUDJAL, BJ_BUDGENES,BJ_BUDGENES2 FROM BUDJAL',true);

  while not QLoc.Eof do begin
    Trouver := False;
    St1 := '';

    nombudget:=QLoc.findfield('BJ_BUDJAL').AsString;
    if QLoc.findfield('BJ_BUDGENES2').AsString<>'' then
      StC := QLoc.findfield('BJ_BUDGENES').AsString + QLoc.findfield('BJ_BUDGENES2').AsString
    else
      StC := QLoc.findfield('BJ_BUDGENES').AsString;
    while StC<>'' do begin
      StTemp := ReadTokenSt(StC);
      if StTemp<>' ' then
        if StTemp=St then Trouver:=True
                   else St1 := St1 + StTemp + ';';
    end;

    if Trouver then begin
      StTemp := CoupeA2000(St1);
      if Pos('<<>>',StTemp)<=0 then
         EXECUTESQL('UPDATE BUDJAL SET BJ_BUDGENES="' + sttemp + '" WHERE BJ_BUDJAL="' + nombudget + '"')
      else begin
         EXECUTESQL('UPDATE BUDJAL SET BJ_BUDGENES="' + Copy(StTemp,1,Pos('<<>>',StTemp)-1) +
         '", BJ_BUDGENES2="' + Copy(StTemp,Pos('<<>>',StTemp)+1,Length(StTemp))+ '" WHERE BJ_BUDJAL="' + nombudget + '"');
      end;
    end;
    QLoc.Next;
  end;
  Ferme(QLoc);
{e fb 28/09/2005 FQ 16043}
end;

procedure TOF_SUPPRBUDG.RempliCbAtt;
var
  QLoc : TQuery;
begin
  CbAtt.Clear;
  QLoc := OpenSQL('SELECT BG_BUDGENE FROM BUDGENE WHERE BG_ATTENTE="X"',True);
  while Not QLoc.Eof do begin
    CbAtt.Add(QLoc.Fields[0].AsString);
    QLoc.Next;
  END;
  Ferme(QLoc);
end;

Initialization
  registerclasses ( [ TOF_SUPPRBUDG ] ) ; 
end.
