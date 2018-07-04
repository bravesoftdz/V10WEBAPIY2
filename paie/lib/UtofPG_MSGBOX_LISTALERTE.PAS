{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMSGBOX_LISTALERTE ()
Mots clefs ... : TOF;PGMSGBOX_LISTALERTE
*****************************************************************}
Unit UTofPG_MSGBOX_LISTALERTE;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     Vierge,
     HMsgBox, 
     UTOF ; 

Type
  TOF_PGMSGBOX_LISTALERTE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
  private
    procedure ClickYes (Sender : TObject);
    procedure ClickNo  (Sender : TObject);
  end ;

Implementation

procedure TOF_PGMSGBOX_LISTALERTE.ClickNo(Sender: TObject);
begin
  TFVierge(Ecran).Retour := 'Non';
  Ecran.Close;
end;

procedure TOF_PGMSGBOX_LISTALERTE.ClickYes(Sender: TObject);
begin
  TFVierge(Ecran).Retour := 'Oui';
  Ecran.Close;
end;

procedure TOF_PGMSGBOX_LISTALERTE.OnArgument (S : String ) ;
var
  stTitre, stMessage, stListeAlertes : String;
  ListeAlertes : TStringList;
  tempAlerte : String;
begin
  Inherited ;
  stTitre := Trim(ReadTokenPipe(S, ';'));
  SetControlText('LBLTITRE',stTitre);
  stMessage := Trim(ReadTokenPipe(S, ';'));
  SetControlText('LBLMESSAGE1',stMessage);
  stMessage := Trim(ReadTokenPipe(S, ';'));
  SetControlText('LBLMESSAGE2',stMessage);
  stMessage := Trim(ReadTokenPipe(S, ';'));
  SetControlText('LBLMESSAGE3',stMessage);
  stListeAlertes := Trim(ReadTokenPipe(S, ';'));
  tempAlerte := Trim(ReadTokenPipe(stListeAlertes, '|'));
  ListeAlertes := TStringList.Create;
  while (stListeAlertes <> '') or (tempAlerte <> '') do
  begin
    ListeAlertes.Add(tempAlerte);
    tempAlerte := Trim(ReadTokenPipe(stListeAlertes, '|'));
  end;
  (GetControl('LISTALERTES') as THListBox).Items.AddStrings(ListeAlertes);
  (GetControl('BTNYES') as THBitBtn).OnClick := ClickYes;
  (GetControl('BTNNO')  as THBitBtn).OnClick := ClickNo;
  ListeAlertes.Free;
end ;

procedure TOF_PGMSGBOX_LISTALERTE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGMSGBOX_LISTALERTE.OnDisplay () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_PGMSGBOX_LISTALERTE ] ) ; 
end.
