{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 21/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SUPPRBUDJ ()
Mots clefs ... : TOF;SUPPRBUDJ
*****************************************************************}
Unit SUPPRBUDJ_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     eMul, uTob, Maineagl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     db, mul, FE_Main,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, utilPGI, Rapsuppr, Ent1,
     UTOF ;

procedure SuppressionJournauxBud;

Type
  TOF_SUPPRBUDJ = Class (TOF)
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
    Nblig : Integer;
    TDelJal, TNotDel : TList;
    Effacer, NotEffacer : Boolean;
    JalCode   : String ;
    procedure BOuvrirClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);

    function  Detruit(Stc : String): Byte;
    procedure Degage;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  BUDJAL_TOM; // FicheBudjal


procedure SuppressionJournauxBud;
begin
  if Not _BlocageMonoPoste(False, '', True) then Exit;
  AGLLanceFiche('CP', 'CPSUPPRBUDJ', '', '', '');
end;

procedure TOF_SUPPRBUDJ.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDJ.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDJ.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDJ.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDJ.OnArgument (S : String ) ;
begin
  Inherited ;

  HM := THMsgBox.Create(Ecran);
  HM.Mess.Add('Journal supprimé');
  HM.Mess.Add('Ce journal à des écritures budgétaires');
  HM.Mess.Add('Journal en cours d''utilisation.');
  HM.Mess.Add('Aucun');
  HM.Mess.Add('élément sélectionné');
  HM.Mess.Add('éléments sélectionnés');
  HM.Mess.Add('6;');

  MsgDel := THMsgBox.Create(Ecran);
  MsgDel.Mess.Add('0;Suppression des journaux budgétaires;Vous n''avez rien sélectionné.;E;O;O;O;');
  MsgDel.Mess.Add('1;Suppression des journaux budgétaires;Ce journal est mouvementé par des écritures budgétaires. Désirez-vous le supprimer ?;Q;YNC;N;C;');
  MsgDel.Mess.Add('2;Suppression des journaux budgétaires;Confirmez-vous la suppression des enregistrements sélectionnés ?;Q;YNC;N;C;');
  MsgDel.Mess.Add('3;Suppression des journaux budgétaires;Vous ne pouvez pas supprimer ce journal : il est en cours d''utilisation.;W;O;O;O;');
  MsgDel.Mess.Add('4;Suppression des journaux budgétaires;Désirez-vous un compte-rendu des comptes détruits ?;Q;YNC;N;C;');
  MsgDel.Mess.Add('5;Suppression des journaux budgétaires;Désirez-vous un compte-rendu des comptes non détruits ?;Q;YNC;N;C;');

  Mul := TFMul(Ecran);
  Mul.BOuvrir.OnClick := BOuvrirClick;
  Mul.FListe.OnDblClick := FListeDblClick;

  TDelJal := TList.Create;
  TNotDel := TList.Create;
  Ecran.HelpContext := 15159000;
end ;

procedure TOF_SUPPRBUDJ.OnClose ;
begin
  Inherited ;

  HM.Free;
  MsgDel.Free;

  AvertirMultiTable('ttBudJal') ;
  TDelJal.Clear;
  TDelJal.Free;
  TNotDel.Clear;
  TNotDel.Free;

  _DeblocageMonoPoste(False, '', True);
end ;

procedure TOF_SUPPRBUDJ.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDJ.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_SUPPRBUDJ.BOuvrirClick(Sender: TObject);
var
  i : Integer;
  j : Byte;
  X : DelInfo;
  Code,Lib : String;
begin
  NbLig := Mul.FListe.NbSelected;
  TDelJal.Clear;
  TNotDel.Clear;
  if NbLig<=0 then begin
    MsgDel.Execute(0,'',''); // Vous n''avez rien sélectionné.
    Exit;
  end;
  if MsgDel.Execute(2,'','')<>mrYes then Exit; // Confirmez-vous la suppression des enregistrements sélectionnés ?
  Effacer := False;
  NotEffacer := False;

  if NbLig>1 then begin
    for i:=0 to NbLig-1 do begin
      Mul.FListe.GotoLeBookMark(i);
      Code := GetField('BJ_BUDJAL');
      Lib := GetField('BJ_LIBELLE');
      j := Detruit(Code);
      if j<=0 then begin
        X := DelInfo.Create;
        X.LeCod := Code;
        X.LeLib := Lib;
        X.LeMess := HM.Mess[0];
        TDelJal.Add(X);
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
    Code := GetField('BJ_BUDJAL');
    j := Detruit(Code);
    if j=2 then MsgDel.Execute(3,'','');
  end;

  if Effacer then if MsgDel.Execute(4,'','')=mrYes then RapportDeSuppression(TDelJal,1);
  if NotEffacer then if MsgDel.Execute(5,'','')=mrYes then RapportDeSuppression(TNotDel,1);
  Mul.BChercheClick(Nil);
end;

procedure TOF_SUPPRBUDJ.Degage;
begin
  if ExecuteSQL('DELETE FROM BUDJAL WHERE BJ_BUDJAL="'+JalCode+'"')<>1 then V_PGI.IoError := oeUnknown;
end;

function TOF_SUPPRBUDJ.Detruit(Stc: String): Byte;
var
  Qloc : TQuery;
begin
  Result := 0;
  if EstMouvementeBudjal(Stc) then begin
    if MsgDel.Execute(1,'','')<>mrYes then begin
      Result := 1;
      Exit;
    end;
  end;
  JalCode := Stc;
  if Transactions(Degage,5)<>oeOK then begin
    MessageAlerte(HM.Mess[2]);
    Result := 2;
    Exit;
  end;
  ExecuteSQL('DELETE FROM BUDECR WHERE BE_BUDJAL="'+JalCode+'"');
  QLoc := OpenSQL('SELECT BJ_CATEGORIE FROM BUDJAL WHERE BJ_BUDJAL="'+JalCode+'"',True);
  if QLoc.Fields[0].AsString<>'' then JalCode := QLoc.Fields[0].AsString;
  Ferme(QLoc);
  ExecuteSQL('DELETE FROM CROISCPT WHERE CX_JAL="'+JalCode+'" And CX_TYPE="BUD"');
end;

procedure TOF_SUPPRBUDJ.FListeDblClick(Sender: TObject);
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit;
  FicheBudjal(nil, '', GetField('BJ_BUDJAL'), taConsult, 0);
end;

Initialization
  registerclasses ( [ TOF_SUPPRBUDJ ] ) ; 
end.
