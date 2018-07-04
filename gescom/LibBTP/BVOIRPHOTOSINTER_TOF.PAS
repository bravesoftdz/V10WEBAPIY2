{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/02/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BVOIRPHOTOSINTER ()
Mots clefs ... : TOF;BVOIRPHOTOSINTER
*****************************************************************}
Unit BVOIRPHOTOSINTER_TOF ;

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
     HMsgBox,
     HRichOLE,
     UTOF ;

Type
  TOF_BVOIRPHOTOSINTER = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TL : TlistBox;
    PHOTO : THImage;
    procedure ChargeListe;
    procedure ListImageClick (Sender : TObject);
    procedure LoadImage(Index: integer);
  end ;

Implementation
uses utilPGI;

procedure TOF_BVOIRPHOTOSINTER.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BVOIRPHOTOSINTER.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BVOIRPHOTOSINTER.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BVOIRPHOTOSINTER.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BVOIRPHOTOSINTER.OnArgument (S : String ) ;
begin
  Inherited ;
  TL := TListBox(getControl('LISTIMAGE'));
  PHOTO := THImage(GetControl('PHOTO'));
  TL.OnClick := ListImageClick;
  ChargeListe;
  LoadImage (0);
end ;

procedure TOF_BVOIRPHOTOSINTER.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BVOIRPHOTOSINTER.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BVOIRPHOTOSINTER.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BVOIRPHOTOSINTER.ChargeListe;
var II : Integer;
begin
  if TL = nil then Exit;
  //
  For II := 0 to LaTOB.detail.count -1 do
  begin
    TL.AddItem(LaTOB.detail[II].getString('LO_LIBELLE'),LaTOB.detail[II]);
  end;
  //
end;

procedure TOF_BVOIRPHOTOSINTER.LoadImage(Index : integer);
var QQ : TQuery;
    Codeappel : string;
begin
  if Index < 0 then exit;
  if LaTOB.detail.count = 0 then exit; 
  if Index >= LaTOB.detail.count then Exit;
  Codeappel := LaTOB.detail[Index].GetString('LO_IDENTIFIANT');
  THRichEditOLE(GetControl('ROANNOTATION')).Text := '';
  StringtoRich(THRichEditOLE(GetControl('ROANNOTATION')),LaTOB.detail[Index].getValue('ANNOT'));
  index := LaTOB.detail[Index].GetInteger('LO_RANGBLOB');
  QQ := OpenSQL('SELECT LO_OBJET FROM LIENSOLE WHERE '+
                'LO_TABLEBLOB="APP" AND '+
                'LO_RANGBLOB='+InttoStr(Index)+' AND '+
                'LO_IDENTIFIANT="'+CodeAppel+'"',True,1,'',true);
  if not QQ.eof then
  begin
    LoadBitMapFromChamp(QQ, 'LO_OBJET', PHOTO, True);
  end;
  Ferme(QQ);
  //
  //
end;

procedure TOF_BVOIRPHOTOSINTER.ListImageClick(Sender: TObject);
begin
  LoadImage(TL.ItemIndex);
end;

Initialization
  registerclasses ( [ TOF_BVOIRPHOTOSINTER ] ) ;
end.
