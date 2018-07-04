{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 11/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MULBUDS ()
Mots clefs ... : TOF;MULBUDS
*****************************************************************}
Unit MULBUDS_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     eMul, uTob, Maineagl,
{$ELSE}
     db, dbtables, mul, FE_Main,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, AglInit, Ent1,
     UTOF ;

procedure MulticritereBudsect(Comment : TActionFiche);

Type
  TOF_MULBUDS = Class (TOF)
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
  BUDSECT_TOM, // FicheBudsect
  MZSUtil;     // ModifieEnSerie


procedure MulticritereBudsect(Comment : TActionFiche);
begin
  AGLLanceFiche('CP', 'CPMULBUDS', '', '', TAToStr(Comment));
end;

procedure TOF_MULBUDS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDS.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDS.OnArgument (S : String ) ;
var
  shtLibre : TTabSheet;
begin
  Inherited;

  HM := THMsgBox.Create(Ecran);
  HM.Mess.Add('Visualisation des sections budgétaires');
  HM.Mess.Add('Modification des sections budgétaires');
  HM.Mess.Add('Modification de zones en série des sections budgétaires');
  HM.Mess.Add('Modifier en série');
  HM.Mess.Add('Sections budgétaires');

  shtLibre := TTabSheet(GetControl('PZLIBRE', True));
  if (shtLibre <> nil) then LibellesTableLibre(shtLibre,'TBS_TABLE','BS_TABLE','D');

  // Evènements
  Mul := TFMul(Ecran);
  Mul.FListe.OnDblClick := FListeDblClick;
  Mul.Binsert.OnClick := BinsertClick;
  Mul.BOuvrir.OnClick := FListeDblClick;

  Comment := StrToTA(S);
  Case Comment Of
    taConsult : begin
      Ecran.Caption := HM.Mess[0];
      Ecran.HelpContext := 15131000;
      SetControlProperty('FListe','ListeParam','MULVBUDS') ;
  		Mul.Q.Liste := 'MULVBUDS';
    end;
    taModif : begin
      Ecran.Caption := HM.Mess[1];
      Ecran.HelpContext := 15135000;
      SetControlProperty('FListe','ListeParam','MULMBUDS') ;
  		Mul.Q.Liste := 'MULMBUDS';
    end;
    taModifEnSerie : begin
      Ecran.Caption := HM.Mess[2];
      Ecran.HelpContext := 15137000;
      SetControlProperty('FListe','ListeParam','MULMBUDS') ;
  		Mul.Q.Liste := 'MULMBUDS';
    end;
  end ;
  if ((EstSerie(S5)) or (EstSerie(S3))) then Ecran.Caption := HM.Mess[4]; // Sections budgétaires
  UpdateCaption(Ecran);

  // FormShow
  TComboBox(GetControl('BS_AXE', True)).ItemIndex := 0;
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

procedure TOF_MULBUDS.OnClose ;
begin
  Inherited ;

  HM.Free;
end ;

procedure TOF_MULBUDS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_MULBUDS.FListeDblClick(Sender: TObject);
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit;

  if Comment<>taModifEnSerie then begin
    FicheBudsect(nil, GetField('BS_AXE'), GetField('BS_BUDSECT'), Comment, 0);
    if Comment<>taConsult then Mul.BChercheClick(Nil);
    end
  else if TControl(Sender).Name='FListe' then begin
    FicheBudsect(nil, GetField('BS_AXE'), GetField('BS_BUDSECT'), Comment, 0);
    Mul.Fliste.ClearSelected;
    end
  else begin
    if (Mul.Fliste.NbSelected>0) or (Mul.FListe.AllSelected) then begin
      ModifieEnSerie('BUDSECT', GetField('BS_AXE'), Mul.FListe, Mul.Q);
      Mul.ChercheClick;
    end;
  end;
  Ecran.Cursor := SyncrDefault;
end;

procedure TOF_MULBUDS.BinsertClick(Sender: TObject);
begin
  FicheBudsect(Nil,'','',taCreatEnSerie,0);
end;

Initialization
  registerclasses ( [ TOF_MULBUDS ] ) ; 
end.
