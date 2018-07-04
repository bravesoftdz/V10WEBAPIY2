{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 10/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MULBUDG ()
Mots clefs ... : TOF;MULBUDG
*****************************************************************}
Unit MULBUDG_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     eMul, uTob, Maineagl,
{$ELSE}
     db, mul, FE_Main,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, AglInit, Ent1,
     UTOF;

procedure MulticritereBudgene(Comment : TActionFiche);

Type
  TOF_MULBUDG = Class (TOF)
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
    Mul : TFMul;
    Comment : TActionFiche;
    procedure FListeDblClick(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
  end;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  BUDGENE_TOM, // FicheBudgene
  MZSUtil     // ModifieEnSerie
  ;

procedure MulticritereBudgene(Comment : TActionFiche);
begin
  AGLLanceFiche('CP', 'CPMULBUDG', '', '', TAToStr(Comment));
end;

procedure TOF_MULBUDG.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDG.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDG.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDG.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDG.OnArgument (S : String ) ;
var
  shtLibre : TTabSheet;
begin
  Inherited;

  HM := THMsgBox.Create(Ecran);
  HM.Mess.Add('Visualisation des comptes budgétaires');
  HM.Mess.Add('Modification des comptes budgétaires');
  HM.Mess.Add('Modification de zones en série des comptes budgétaires');
  HM.Mess.Add('Modifier en série');
  HM.Mess.Add('Comptes budgétaires');

  shtLibre := TTabSheet(GetControl('PZLIBRE', True));
  if (shtLibre <> nil) then LibellesTableLibre(shtLibre,'TBG_TABLE','BG_TABLE','B');

  // Evènements
  Mul := TFMul(Ecran);
  Mul.FListe.OnDblClick := FListeDblClick;
  Mul.Binsert.OnClick := BinsertClick;
  Mul.BOuvrir.OnClick := FListeDblClick;

  Comment := StrToTA(S);
  Case Comment Of
    taConsult : begin
      Ecran.Caption := HM.Mess[0];
      Ecran.HelpContext := 15111000;
      SetControlProperty('FListe','ListeParam','MULVBUDG') ;
  		Mul.Q.Liste := 'MULVBUDG';
    end;
    taModif : begin
      Ecran.Caption := HM.Mess[1];
      Ecran.HelpContext := 15115000;
      SetControlProperty('FListe','ListeParam','MULMBUDG') ;
  		Mul.Q.Liste := 'MULMBUDG';
    end;
    taModifEnSerie : begin
      Ecran.Caption := HM.Mess[2];
      Ecran.HelpContext := 15117000;
      SetControlProperty('FListe','ListeParam','MULMBUDG') ;
  		Mul.Q.Liste := 'MULMBUDG';
    end;
  end ;
  if ((EstSerie(S5)) or (EstSerie(S3))) then Ecran.Caption := HM.Mess[4]; // Comptes budgétaires
  UpdateCaption(Ecran);

  // FormShow
  if Comment=taModifEnSerie then begin
{$IFDEF EAGLCLIENT}
    Mul.FListe.MultiSelect := True;
{$ELSE}
    Mul.FListe.MultiSelection := True;
{$ENDIF}
    Mul.BOuvrir.Hint := HM.Mess[3]; // Modifier en série
    Mul.bSelectAll.Visible:=True ;
    end
  else begin
{$IFDEF EAGLCLIENT}
   Mul.FListe.MultiSelect := False;
{$ELSE}
   Mul.FListe.MultiSelection := False;
{$ENDIF}
  end;

  if ((Comment<>taConsult)) then Mul.BInsert.Visible := True;
end ;

procedure TOF_MULBUDG.OnClose ;
begin
  Inherited ;

  HM.Free;
end ;

procedure TOF_MULBUDG.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDG.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDG.FListeDblClick(Sender: TObject);
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit;

  if Comment<>taModifEnSerie then begin
    FicheBudgene(nil, '', GetField('BG_BUDGENE'), Comment, 0);
    if Comment<>taConsult then Mul.BChercheClick(Nil);
    end
  else if TControl(Sender).Name='FListe' then begin
    FicheBudgene(nil, '', GetField('BG_BUDGENE'), Comment, 0);
    Mul.Fliste.ClearSelected;
    end
  else begin
    if (Mul.Fliste.NbSelected>0) or (Mul.FListe.AllSelected) then begin
      ModifieEnSerie('BUDGENE', '', Mul.FListe, Mul.Q);
      Mul.ChercheClick;
    end;
  end;
  Ecran.Cursor := SyncrDefault;
end;

procedure TOF_MULBUDG.BinsertClick(Sender: TObject);
begin
  FicheBudgene(Nil,'','',taCreatEnSerie,0);
end;

Initialization
  registerclasses ( [ TOF_MULBUDG ] ) ; 
end.
