{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 11/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MULBUDJ ()
Mots clefs ... : TOF;MULBUDJ
*****************************************************************}
Unit MULBUDJ_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     eMul, uTob, Maineagl,
{$ELSE}
     db, dbtables, mul, FE_Main,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, AglInit, Ent1,
     UTOF ;

procedure MulticritereBudjal(Comment : TActionFiche);

Type
  TOF_MULBUDJ = Class (TOF)
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
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  BUDJAL_TOM, // FicheBudjal
  MZSUtil;     // ModifieEnSerie



procedure MulticritereBudjal(Comment : TActionFiche);
begin
  AGLLanceFiche('CP', 'CPMULBUDJ', '', '', TAToStr(Comment));
end;

procedure TOF_MULBUDJ.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDJ.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDJ.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDJ.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDJ.OnArgument (S : String ) ;
begin
  Inherited;

  HM := THMsgBox.Create(Ecran);
  HM.Mess.Add('Visualisation des budgets');
  HM.Mess.Add('Modification des budgets');
  HM.Mess.Add('Modification de zones en série des budgets');
  HM.Mess.Add('Modifier en série');
  HM.Mess.Add('Budgets');

  // Evènements
  Mul := TFMul(Ecran);
  Mul.FListe.OnDblClick := FListeDblClick;
  Mul.Binsert.OnClick := BinsertClick;
  Mul.BOuvrir.OnClick := FListeDblClick;

  Comment := StrToTA(S);
  Case Comment Of
    taConsult : begin
      Ecran.Caption := HM.Mess[0];
      Ecran.HelpContext := 15151000;
      SetControlProperty('FListe','ListeParam','MULVBUDJ') ;
  		Mul.Q.Liste := 'MULVBUDJ';
    end;
    taModif : begin
      Ecran.Caption := HM.Mess[1];
      Ecran.HelpContext := 15155000;
      SetControlProperty('FListe','ListeParam','MULMBUDJ') ;
  		Mul.Q.Liste := 'MULMBUDJ';
    end;
    taModifEnSerie : begin
      Ecran.Caption := HM.Mess[2];
      Ecran.HelpContext := 15157000;
      SetControlProperty('FListe','ListeParam','MULMBUDJ') ;
  		Mul.Q.Liste := 'MULMBUDJ';
    end;
  end ;
  if ((EstSerie(S5)) or (EstSerie(S3))) then Ecran.Caption := HM.Mess[4]; // Budgets
  UpdateCaption(Ecran);

  // FormShow
  TComboBox(GetControl('BJ_AXE', True)).ItemIndex := 0;
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

procedure TOF_MULBUDJ.OnClose ;
begin
  Inherited ;

  HM.Free;
end ;

procedure TOF_MULBUDJ.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDJ.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDJ.FListeDblClick(Sender: TObject);
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit;

  if Comment<>taModifEnSerie then begin
    FicheBudjal(nil, GetField('BJ_AXE'), GetField('BJ_BUDJAL'), Comment, 0);
    if Comment<>taConsult then Mul.BChercheClick(Nil);
    end
  else if TControl(Sender).Name='FListe' then begin
    FicheBudjal(nil, GetField('BJ_AXE'), GetField('BJ_BUDJAL'), Comment, 0);
    Mul.Fliste.ClearSelected;
    end
  else begin
    if (Mul.Fliste.NbSelected>0) or (Mul.FListe.AllSelected) then begin
      ModifieEnSerie('BUDJAL', GetField('BJ_AXE'), Mul.FListe, Mul.Q);
      Mul.ChercheClick;
    end;
  end;
  Ecran.Cursor := SyncrDefault;
end;

procedure TOF_MULBUDJ.BinsertClick(Sender: TObject);
begin
  FicheBudjal(Nil,'','',taCreatEnSerie,0);
end;

Initialization
  registerclasses ( [ TOF_MULBUDJ ] ) ; 
end.
