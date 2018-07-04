unit dpBlocNotes;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, StdCtrls, ComCtrls, HRichEdt, HRichOLE, HSysMenu, HTB97, UTob,
  HMsgBox, hpanel, uiUtil;

type
  TFBlocNotes = class(TFVierge)
    OleBlocNotes: THRichEditOLE;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure OleBlocNotesChange(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BValiderClick(Sender: TObject);
  private
    { Déclarations privées }
    bDuringLoad, bChange : Boolean;
    FGuidPer : String;
    TobBlob : TOB;
    function SauveEnregBlocNotes : Boolean;
  public
    { Déclarations publiques }
  end;

procedure ShowBlocNotes(GuidPer: String; Inside:ThPanel=nil);  //LMO

////////// IMPLEMENTATION //////////
implementation

uses galOutil;

procedure ShowBlocNotes(GuidPer: String ; Inside:ThPanel=nil);//LMO
var
  FBlocNotes : TFBlocNotes;
begin
  FBlocNotes := TFBlocNotes.Create(Application);
  FBlocNotes.FGuidPer := GuidPer;
  if Inside = nil then    //+LMO
  begin
    try
      FBlocNotes.ShowModal ;
    finally
      FBlocNotes.Free ;
    end ;
  end
  else                     //-LMO
  begin
    try
      InitInside(FBlocNotes,Inside) ;
      FBlocNotes.Show ;
    except
    end ;
  end ;
end;


{$R *.DFM}

procedure TFBlocNotes.FormShow(Sender: TObject);
begin
  inherited;
  bDuringLoad := True;
  bChange := False;

  TobBlob := Tob.Create('LIENSOLE', nil, -1);

  // Renseigne la clé primaire
  TobBlob.PutValue('LO_TABLEBLOB', 'ANN');
  //TobBlob.PutValue('LO_IDENTIFIANT', 'OLE_BLOCNOTES');
  TobBlob.PutValue('LO_IDENTIFIANT', FGuidPer);
  // 1 = Activité
  // 2 = Bloc-notes
  // 3 = Régime matrimonial
  // 4 = Objet soc
  TobBlob.PutValue('LO_RANGBLOB', 2);

  if TobBlob.LoadDB then
    StringToRich(TRichEdit(OleBlocNotes),TobBlob.GetValue('LO_OBJET'))
  else
    StringToRich(TRichEdit(OleBlocNotes),'') ;

  bDuringLoad := False;
end;

procedure TFBlocNotes.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
  if TobBlob<>Nil then TobBlob.Free;
  InsideSynthese (self) ;
end;

procedure TFBlocNotes.OleBlocNotesChange(Sender: TObject);
begin
  inherited;
  if not bDuringLoad then bChange := True;
end;

procedure TFBlocNotes.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
var rep: Integer;
begin
  inherited;
  // Changement demandé
  if bChange then
    begin
    rep := PGIAskCancel('Voulez-vous enregistrer les modifications ?', Self.Caption);
    if rep=mrCancel then
      CanClose := False
    else if rep=mrNo then
      bChange := False
    else if rep=mrYes then
      CanClose := SauveEnregBlocNotes;
    end;
end;

function TFBlocNotes.SauveEnregBlocNotes: Boolean;
var letexte : String;
begin
  Result := False;
  leTexte := RichToString(OleBlocNotes);
  TobBlob.PutValue('LO_OBJET',leTexte);
  TobBlob.InsertOrUpdateDB;
  Result := True;
end;

procedure TFBlocNotes.BValiderClick(Sender: TObject);
begin
  // Par défaut, le bouton BValider referme la fiche
  if Not SauveEnregBlocNotes then
    // donc on l'en empêche si SauveEnreg n'a pas pu valider
    ModalResult := mrNone;
end;

end.
