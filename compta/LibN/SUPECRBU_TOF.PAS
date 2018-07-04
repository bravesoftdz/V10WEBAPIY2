{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 06/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SUPECRBU ()
Mots clefs ... : TOF;SUPECRBU
*****************************************************************}
Unit SUPECRBU_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
     eSaisBud, // SaisieBudget
{$IFDEF EAGLCLIENT}
     eMul, uTob, MaineAgl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     db, mul, FE_Main,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox,
     HStatus, Ent1, UTOF;

Procedure DetruitBudgets;

Type
  TOF_SUPECRBU = Class (TOF)
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
    HM : THMsgBox;
    procedure FListeDblClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure DetruitLaPiece;
  end ;

Implementation

Procedure DetruitBudgets;
begin
  AGLLanceFiche('CP', 'CPSUPECRBU', '', '', '');
end;

procedure TOF_SUPECRBU.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_SUPECRBU.OnDelete ;
begin
  Inherited;
end ;

procedure TOF_SUPECRBU.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_SUPECRBU.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_SUPECRBU.OnArgument (S : String ) ;
begin
  Inherited ;

  HM := THMsgBox.Create(Ecran);
  // Suppression des budgets
  HM.Mess.Add('0;?Caption?;Désirez-vous supprimer ces budgets ?;Q;YN;N;N;');
  HM.Mess.Add('1;?Caption?;Vous devez sélectionner des budgets.;E;O;O;O;');
  HM.Mess.Add('Destruction non effectuée.');

  // FormShow
  PositionneEtabUser(GetControl('BE_ETABLISSEMENT', True));

  // Evènements
  Mul := TFMul(Ecran);
  Mul.FListe.OnDblClick := FListeDblClick;
  Mul.BOuvrir.OnClick := BOuvrirClick;
end ;

procedure TOF_SUPECRBU.OnClose ;
begin
  Inherited ;

  HM.Free;
end ;

procedure TOF_SUPECRBU.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_SUPECRBU.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_SUPECRBU.FListeDblClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
var
  TOBListe, TOBLigne : TOB;
{$ENDIF}
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit ;

{$IFDEF EAGLCLIENT}
  TOBListe := TOB.Create('Liste Ecriture',nil,-1) ;
  TOBLigne := Mul.Q.TQ.Detail[Mul.FListe.row-1] ;
  TOB.Create('BUDECR',TOBListe,-1) ;
  TOBListe.Detail[0].Dupliquer(TOBLigne, False, True);
  TrouveEtLanceSaisBud(TOBListe, taConsult);
  FreeAndNil(TOBListe);
{$ELSE}
  TrouveEtLanceSaisBud(Mul.Q, taConsult);
{$ENDIF}
end;

procedure TOF_SUPECRBU.DetruitLaPiece;
var
  Q : TQuery;
begin
{$IFDEF EAGLCLIENT}
  Q := Mul.Q.TQ;
{$ELSE}
  Q := Mul.Q;
{$ENDIF}

  if ExecuteSql('DELETE FROM BUDECR WHERE BE_BUDJAL="'+Q.FindField('BE_BUDJAL').AsString+'" '+
                'AND BE_NATUREBUD="'+Q.FindField('BE_NATUREBUD').AsString+'" '+
                'AND BE_NUMEROPIECE='+IntToStr(Q.FindField('BE_NUMEROPIECE').AsInteger)+' '+
                'AND BE_QUALIFPIECE="'+Q.FindField('BE_QUALIFPIECE').AsString+'" '+
                'AND BE_VALIDE="-"')<=0 then V_PGI.IoError:=oeUnknown;
end;

procedure TOF_SUPECRBU.BOuvrirClick(Sender: TObject);
var
  i,NbD : integer;
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit ;

  NbD := Mul.FListe.NbSelected;
  if NbD<=0 then begin
    HM.Execute(1,Ecran.caption,'');
    Exit;
  end;
  if HM.Execute(0,Ecran.Caption,'')<>mrYes then Exit;
  Application.ProcessMessages;
  InitMove(NbD,'');
  for i:=0 to NbD-1 do begin
    Mul.FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
    Mul.Q.TQ.Seek (Mul.FListe.Row-1);
{$ENDIF}
    MoveCur(FALSE);
    if Transactions(DetruitLaPiece,3)<>oeOK then begin
      MessageAlerte(HM.Mess[2]);
      FiniMove;
      Exit;
    end;
  end;
  Mul.FListe.ClearSelected;
  FiniMove;
  Mul.BChercheClick(Nil);
end;

Initialization
  registerclasses ( [ TOF_SUPECRBU ] ) ; 
end.
