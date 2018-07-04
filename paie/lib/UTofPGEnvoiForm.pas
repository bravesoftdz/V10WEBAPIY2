{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 19/02/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGENVOIFORM ()
Mots clefs ... : TOF;PGENVOIFORM
*****************************************************************}
unit UTofPGEnvoiForm;

interface

uses StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF,
  Entpaie, HTB97, UTob, aglMail, MailOL;

type
  TOF_PGENVOIFORM = class(TOF)
    procedure OnArgument(S: string); override;
  private
    NumeroEnvoi, Support, Reel, EmisPar: string;
    procedure LanceEnvoi(Sender: TOBject);
  end;

implementation

procedure TOF_PGENVOIFORM.OnArgument(S: string);
var
  OPCA: string;
  Fichier, Libelle, Millesime: string;
  Q: TQuery;
  BLance: TToolBarButton97;
begin
  inherited;
  NumeroEnvoi := ReadTokenPipe(S, ';');
  OPCA := ReadTokenPipe(S, ';');
  Support := ReadTokenPipe(S, ';');
  Reel := ReadTokenPipe(S, ';');
  EmisPar := ReadTokenPipe(S, ';');
  Q := OpenSQL('SELECT PVF_NOMFICHIER,PVF_LIBELLE,PVF_MILLESSOC FROM ENVOIFORMATION WHERE PVF_NUMENVOI=' + NumeroEnvoi + '', True);
  if not Q.Eof then
  begin
    Fichier := Q.FindField('PVF_NOMFICHIER').AsString;
    Libelle := Q.FindField('PVF_LIBELLE').AsString;
    Millesime := Q.FindField('PVF_MILLESSOC').AsString;
  end;
  Ferme(Q);
  SetControlCaption('LIBELLE', Libelle);
  SetControlCaption('LBLMILLESIME', Millesime);
  SetControlCaption('LBLEMETTEUR', RechDom('PGEMETTEURSOC', EmisPar, False));
  SetControlCaption('LBLSUPPORT', RechDom('PGSUPPORTEDI', Support, False));
  SetControlText('LBLREEL', Reel);
  SetControlText('NOMFIC', Fichier);
  BLance := TToolBarButton97(GetControl('BLANCE'));
  if BLance <> nil then BLance.OnClick := LanceEnvoi;
end;

procedure TOF_PGENVOIFORM.LanceEnvoi(Sender: TObject);
var
  Liste: HTStringList;
  Q: TQuery;
  TobEnvoi, T: Tob;
begin
  Q := OpenSQL('SELECT * FROM ENVOIFORMATION WHERE PVF_NUMENVOI=' + NumeroEnvoi + '', True);
  TobEnvoi := Tob.Create('ENVOIFORMATION', nil, -1);
  TobEnvoi.LoadDetailDB('ENVOIFORMATION', '', '', Q, False);
  Ferme(Q);
  T := TobEnvoi.FindFirst(['PVF_NUMENVOI'], [StrToInt(NumeroEnvoi)], False);
  if T <> nil then
  begin
    T.PutValue('PVF_ENVOIREEL', Reel);
    T.PutValue('PVF_EMETSOC', EmisPar);
    T.PutValue('PVF_SUPPORTEMIS', Support);
    T.PutValue('PVF_STATUTENVOI', '010');
    T.PutValue('PVF_DATEENVOI', Date);
    T.UpdateDB(False);
  end;
  if Support = 'TEL' then
  begin
    Liste := HTStringList.Create;
    Liste.Add('Veuillez trouver ci-joint le fichier passerelle');
    {$IFDEF EAGLCLIENT}
    SendMail('Envoi formation', '', '', Liste, VH_PAIE.PGCheminEagl + '\' + GetControlText('NOMFIC') + '.xls', FALSE);
    {$ELSE}
    SendMail('Envoi formation', '', '', Liste, V_PGI.DatPath + '\' + GetControlText('NOMFIC') + '.xls', FALSE);
    {$ENDIF}
    Liste.Clear;
    Liste.Free;
  end;
end;

initialization
  registerclasses([TOF_PGENVOIFORM]);
end.

