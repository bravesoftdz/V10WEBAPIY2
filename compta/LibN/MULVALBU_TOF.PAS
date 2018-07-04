{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 27/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MULVALBU ()
Mots clefs ... : TOF;MULVALBU
*****************************************************************}
Unit MULVALBU_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     eMul, uTob, Maineagl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     db, mul, FE_Main,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, AglInit, Ent1,
     UTOF ;

procedure MultiCritereValBud(AValider : Boolean);

Type
  TOF_MULVALBU = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    AValider : Boolean ;
    NbLig : Integer ;
    Mul : TFmul;
    HM : THMsgBox;
    CbJal : TStringList;

    procedure FListeDblClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);

    procedure RempliCbJal ;
    procedure LanceTraitement ;
  end ;

Implementation

uses eSaisbud;  // TrouveEtLanceSaisBud

procedure MultiCritereValBud(AValider : Boolean);
var
  szArg : String;
begin
  if AValider then szArg := 'X'
              else szArg := '-';

  AGLLanceFiche('CP', 'CPMULVALBU', '', '', szArg);
end;

procedure TOF_MULVALBU.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MULVALBU.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MULVALBU.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_MULVALBU.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_MULVALBU.OnArgument (S : String ) ;
begin
  Inherited ;
  AValider := (S = 'X');

  Mul := TFmul(Ecran);
  CbJal := TStringList.Create;

  HM := THMsgBox.Create(Ecran);
  HM.Mess.Add('Validation des écritures budgétaires');
  HM.Mess.Add('Dévalidation des écritures budgétaires');
  HM.Mess.Add('Valider');
  HM.Mess.Add('Dévalider');
  HM.Mess.Add('4;Dévalidation des écritures budgétaires;Confirmez vous la dévalidation de l''écriture.;Q;YN;N;N;');
  HM.Mess.Add('5;Validation des écritures budgétaires;Confirmez vous la validation de l''écriture.;Q;YN;N;N;');
  HM.Mess.Add('Aucun');
  HM.Mess.Add('élément sélectionné');
  HM.Mess.Add('éléments sélectionnés');
  HM.Mess.Add('9;Dévalidation des écritures budgétaires;Vous n''avez rien sélectionné.;E;O;O;O;');
  HM.Mess.Add('10;Validation des écritures budgétaires;Vous n''avez rien sélectionné.;E;O;O;O;');

  if AValider then begin
    Ecran.Caption := HM.Mess[0];
    Mul.BOuvrir.Hint := HM.Mess[2];
    SetControlChecked('BE_VALIDE', False);
    Ecran.HelpContext := 15250000;
    end
  else begin
    Ecran.Caption := HM.Mess[1] ;
    Mul.BOuvrir.Hint := HM.Mess[3] ;
    SetControlChecked('BE_VALIDE', True);
    Ecran.HelpContext := 15260000 ;
  end;
  UpdateCaption(Ecran);

  TComboBox(GetControl('BE_BUDJAL', True)).ItemIndex := 0;
  TComboBox(GetControl('BE_NATUREBUD', True)).ItemIndex := 0;
  RempliCbJal;
  PositionneEtabUser(GetControl('BE_ETABLISSEMENT', True));

  Mul.FListe.OnDblClick := FListeDblClick;
  Mul.BOuvrir.OnClick := BOuvrirClick;
end ;

procedure TOF_MULVALBU.OnClose ;
begin
  Inherited ;

  HM.Free;
  CbJal.Free;
end ;

procedure TOF_MULVALBU.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_MULVALBU.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_MULVALBU.LanceTraitement;
var
  Flag : Char;
  i : Integer;
  Sql : String;
  Q : TQuery;
begin
{$IFDEF EAGLCLIENT}
  Q := TFMul(Ecran).Q.TQ;
{$ELSE}
  Q := TFMul(Ecran).Q;
{$ENDIF}

  if AValider then Flag:='X'
              else Flag:='-' ;

  for i:=0 to NbLig-1 do begin
    Mul.FListe.GotoLeBookMark(i);
    {$IFDEF EAGLCLIENT}
    TFMul(Ecran).Q.TQ.Seek(Mul.FListe.Row-1);
    {$ENDIF}

    if CbJal.IndexOf(Q.FindField('BE_BUDJAL').AsString)<>-1 then Continue;

    Sql:='UPDATE BUDECR SET BE_VALIDE="'+Flag+'" '+
         'WHERE BE_BUDJAL="'+Q.FindField('BE_BUDJAL').AsString+'" '+
         'AND BE_NATUREBUD="'+Q.FindField('BE_NATUREBUD').AsString+'" '+
         'AND BE_NUMEROPIECE='+Q.FindField('BE_NUMEROPIECE').AsString+' '+
//         'AND BE_DATECOMPTABLE="'+UsDateTime(Q.FindField('BE_DATECOMPTABLE').AsString)+'" '+
         'AND BE_QUALIFPIECE="'+Q.FindField('BE_QUALIFPIECE').AsString+'" ' ;
    ExecuteSQL(Sql);
  end;
end;

procedure TOF_MULVALBU.RempliCbJal;
var
  QLoc : TQuery ;
begin
  CbJal.Clear;
  QLoc := OpenSQL('SELECT BJ_BUDJAL FROM BUDJAL WHERE BJ_FERME="X"', True);
  while not QLoc.Eof do begin
    CbJal.Add(QLoc.Fields[0].AsString);
    QLoc.Next ;
  end;
  Ferme(QLoc);
end;

procedure TOF_MULVALBU.FListeDblClick(Sender: TObject);
begin

  if (GetDataSet.Eof) And (GetDataSet.Bof) then Exit;
  {$IFDEF EAGLCLIENT}
  TrouveEtLanceSaisBud(GetDataSet, taConsult);
  {$ELSE}
  TrouveEtLanceSaisBud(Mul.Q, taConsult);
  {$ENDIF EAGLCLIENT}
end;

procedure TOF_MULVALBU.BOuvrirClick(Sender: TObject);
begin
  NbLig := Mul.Fliste.NbSelected;
  if NbLig<=0 then begin
    HM.Execute(Ord(AValider)+9,'','');
    Exit;
  end;
  if HM.Execute(Ord(AValider)+4,'','')=mrYes then begin
    LanceTraitement;
    Mul.BChercheClick(Nil);
  end;
end;

Initialization
  registerclasses ( [ TOF_MULVALBU ] ) ; 
end.
