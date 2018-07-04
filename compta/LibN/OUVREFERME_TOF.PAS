{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : OUVREFERME ()
Mots clefs ... : TOF;OUVREFERME
*****************************************************************}
Unit OUVREFERME_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     eMul, uTob, Maineagl,
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     db, mul, FE_Main, HDB, 
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, AglInit, Ent1, utilPGI,
     UTOF ;

procedure OuvreFermeCpte(Fermer:Boolean);

Type
  TOF_OUVREFERME = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Fermer : Boolean;
    HM : THMsgBox;
    NbLig : Integer;
    {$IFDEF EAGLCLIENT}
    FListe : THGrid;
    {$ELSE}
    FListe : THDBGrid;
    {$ENDIF}
    procedure FListeDblClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure OuvrirOuFermer;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  BUDJAL_TOM; // FicheBudjal

procedure OuvreFermeCpte(Fermer:Boolean);
var
  pszArg : String;
begin
  if _Blocage(['nrCloture','nrBatch'],False,'nrBatch') then Exit;

  if Fermer then pszArg := 'X'
            else pszArg := '-';
    AGLLanceFiche('CP','CPOUVREFERME','','',pszArg);
end;

procedure TOF_OUVREFERME.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_OUVREFERME.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_OUVREFERME.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_OUVREFERME.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_OUVREFERME.OnArgument (S : String ) ;
begin
  Inherited ;
  Fermer := (S='X');

  HM := THMsgBox.Create(Ecran);
  HM.Mess.Add('Fermeture des budgets');
  HM.Mess.Add('Ré-ouverture des budgets');
  HM.Mess.Add('Fermer les budgets');
  HM.Mess.Add('Ré-ouvrir les budgets');
  HM.Mess.Add('4;?caption?;Désirez vous ré-ouvrir les journaux sélectionnés ?;Q;YNC;N;C;');
  HM.Mess.Add('5;?caption?;Désirez vous fermer les journaux sélectionnés ?;Q;YNC;N;C;');
  HM.Mess.Add('Certains comptes sont déjà en traitement. Modification non effectuée.');
  HM.Mess.Add('7;?caption?;Vous n''avez rien sélectionné.;E;O;O;O;');

  if Fermer then begin
    Ecran.HelpContext := 15161000;
    Ecran.Caption:=HM.Mess[1];       // Ré-ouverture des budgets
    TFMul(Ecran).BOuvrir.Hint:=HM.Mess[3]; // Ré-ouvrir les budgets
    SetControlChecked('BJ_FERME', True);
    end
  else begin
    Ecran.HelpContext := 15163000;
    Ecran.Caption:=HM.Mess[0];       // Fermeture des budgets
    TFMul(Ecran).BOuvrir.Hint:=HM.Mess[2]; // Fermer les budgets
    SetControlChecked('BJ_FERME', False);
  end;
  UpdateCaption(Ecran);
  FListe := TFMul(Ecran).FListe;
  FListe.OnDblClick := FListeDblClick;
  TFMul(Ecran).BOuvrir.OnClick := BOuvrirClick;
  TButton(GetControl('BZOOM', True)).OnClick := FListeDblClick;
end ;

procedure TOF_OUVREFERME.OnClose ;
begin
  Inherited ;
  HM.Free;
  AvertirMultiTable('ttBudjal');
  AvertirMultiTable('ttBudjalSais');

  _Bloqueur('nrBatch',False);
end ;

procedure TOF_OUVREFERME.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_OUVREFERME.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_OUVREFERME.BOuvrirClick(Sender: TObject);
var
  Msg,i : Integer;
begin
  NbLig := FListe.NbSelected;
  if NbLig<=0 then begin
    HM.Execute(7,Ecran.Caption,'');
    Exit;
  end;
  if Fermer then Msg:=4  // Désirez vous ré-ouvrir les journaux sélectionnés ?
            else Msg:=5; // Désirez vous fermer les journaux sélectionnés ?
  if HM.Execute(Msg, Ecran.Caption,'')<>mrYes then Exit;

  for i:=0 to NbLig-1 do begin
    Fliste.GotoLeBookMark(i);
    if Transactions(OuvrirOuFermer,5)<>oeOK then begin
      MessageAlerte(HM.Mess[6]); // Certains comptes sont déjà en traitement. Modification non effectuée.
      Break;
    end;
  end;
  FListe.ClearSelected;
  TFMul(Ecran).BChercheClick(Nil);
end;

procedure TOF_OUVREFERME.OuvrirOuFermer;
var
  StC : String;
begin
 StC := GetField('BJ_BUDJAL');
  if Fermer then begin
    if ExecuteSql('UPDATE BUDJAL SET BJ_FERME="-", BJ_DATEOUVERTURE="'+UsDateTime(Date)+'", '+
                  'BJ_DATEMODIF="'+UsDateTime(Date)+'" WHERE BJ_BUDJAL="'+Stc+'"')<>1 then V_PGI.IoError:=oeUnknown;
    end
  else begin
    if ExecuteSql('UPDATE BUDJAL SET BJ_FERME="X", BJ_DATEFERMETURE="'+UsDateTime(Date)+'", '+
                  'BJ_DATEMODIF="'+UsDateTime(Date)+'" WHERE BJ_BUDJAL="'+Stc+'"')<>1 then V_PGI.IoError:=oeUnknown;
  end;
end;

procedure TOF_OUVREFERME.FListeDblClick(Sender: TObject);
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit;
  FicheBudjal(nil, '', GetField('BJ_BUDJAL'), taConsult, 0);
end;

Initialization
  registerclasses ( [ TOF_OUVREFERME ] ) ;
end.
