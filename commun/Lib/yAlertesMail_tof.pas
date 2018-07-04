Unit yAlertesMail_tof ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,Fe_Main,
{$else}
     MaineAgl,
     eMul,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uTob,htb97 ;

Type
  TOF_YALERTES_MAIL = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  Private
    CodeAlerte : string;
    procedure AjouterClick(Sender: TObject);
    procedure EnleverClick(Sender: TObject);
    procedure OnAdresseElipsisClick(Sender: TObject);
  end ;

Implementation

procedure TOF_YALERTES_MAIL.OnArgument (S : String ) ;
var TobAlertes : tob;
    Q : TQuery;
    MC:TListBox;
    i : integer;
begin
  Inherited ;
  CodeAlerte:=S;
  TobAlertes:=TOB.create ('les alertes',NIL,-1);
  Q := OpenSql ('Select * from YALERTESMAIL where YAM_ALERTE="'+CodeAlerte+'"',TRUE);
  TobAlertes.LoadDetailDB ('YALERTESMAIL', '', '', Q, False);
  ferme(Q) ;
  if TobAlertes.detail.count <> 0 then
    begin
    MC:=TListBox(GetControl('LISTMAIL'));
    for i:=0 to TobAlertes.detail.count-1 do
        MC.Items.Add(TobAlertes.detail[i].GetValue('YAM_MAIL'));
    end;
  TobAlertes.free;
  TToolBarButton97(GetControl('BADD')).OnClick:=AjouterClick;
  TToolBarButton97(GetControl('BSUPP')).OnClick:=EnleverClick;
  if Assigned(GetControl('ADRESSE')) then
    THEdit(GetControl('ADRESSE')).OnElipsisClick := OnAdresseElipsisClick ;
end ;

procedure TOF_YALERTES_MAIL.OnAdresseElipsisClick(Sender: TObject);
var stListeEmail,eMail : String;
    MC:TListBox;
begin
  stListeEmail:=AGLLanceFiche('Y','YALERTES_RESSOURC','','','ALERTE');
  MC:=TListBox(GetControl('LISTMAIL'));
  if (stListeEmail = '') or (not Assigned(MC)) then exit;
  Repeat
    eMail:=ReadToKenPipe(stListeEmail,'|');
    if eMail <> '' then MC.Items.add(eMail);
  until eMail='';
  THEdit(GetControl('ADRESSE')).text:='';
end;

procedure TOF_YALERTES_MAIL.AjouterClick(Sender: TObject);
var MC:TListBox;
begin
  if THEdit(GetControl('ADRESSE')).text <> '' then
    begin
    MC:=TListBox(GetControl('LISTMAIL'));
    if Assigned(MC) then
      MC.Items.Add(THEdit(GetControl('ADRESSE')).text);
    THEdit(GetControl('ADRESSE')).text:='';
    end;
end;

procedure TOF_YALERTES_MAIL.EnleverClick(Sender: TObject);
var MC:TListBox;
begin
  MC:=TListBox(GetControl('LISTMAIL'));
  if (MC.ItemIndex = -1 ) or ( (MC.ItemIndex >= 0) and
      ( MC.Selected[MC.ItemIndex] = False) ) then
    begin
    PGIBox(TraduireMemoire('Vous devez sélectionner une ligne'),'Enlever une ligne');
    exit;
    end;
  MC.Items.Delete ( MC.ItemIndex);
end;

procedure TOF_YALERTES_MAIL.OnClose ;
var MC:TListBox;
    i : integer;
begin
  Inherited ;
  ExecuteSql ('delete from YALERTESMAIL where YAM_ALERTE="'+CodeAlerte+'"');
  MC:=TListBox(GetControl('LISTMAIL'));
  if MC.Items.count <> 0 then
    for i:=0 to MC.Items.count-1 do
      begin
      ExecuteSQL ('INSERT INTO YALERTESMAIL (YAM_ALERTE,YAM_NUMERO,YAM_MAIL)'
      +' values ("'+CodeAlerte+'",'+IntToStr(i+1)+','
      +'"'+MC.Items[i]+'")');
      end;
end ;
Initialization
  registerclasses ( [ TOF_YALERTES_MAIL ] ) ;
end.

