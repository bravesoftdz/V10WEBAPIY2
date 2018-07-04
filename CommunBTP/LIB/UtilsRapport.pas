{***********UNITE*************************************************
Auteur  ...... : Franck VAUTRAIN
Créé le ...... : 19/07/2011
Modifié le ... :   /  /    
Description .. : Classes de gestion des rapports d'erreur ou d'intégration.
Suite ........ : création de mémo
Suite ........ : affichage du mémo
Suite ........ : edition du mémo
Mots clefs ... : RAPPORT; ERREUR; MEMO
*****************************************************************}
unit UtilsRapport;

interface

uses  StdCtrls,
      Controls,
      Classes,
      forms,
      sysutils,
      ComCtrls,
      HCtrls,
      HEnt1,
      {$IFNDEF EAGLCLIENT}
      db,
       {$IFNDEF DBXPRESS}
       dbTables,
       {$ELSE}
       uDbxDataSet,
       {$ENDIF}
      {$ELSE}
      HDB,
      {$ENDIF}
      HMsgBox,
      HRichOLE,
      CBPPath,
      Printers,
      uTOB,
      HTB97;
Type
    TGestionRapport = class(Tobject)
    private
      fToolWin  : TToolWindow97;
      fMemo     : TMemo;
      fTitre    : String;
      fFileName : String;
      fTableID  : String;
      fQualif   : String;
      fLibelle  : String;
      fIDLienOLE: String;
      fAffiche  : Boolean;
      fPrint    : Boolean;
      fSauve    : Boolean;
      fClose    : Boolean;
      fPosTop   : Integer;
      fPosLeft  : Integer;
      fWidth    : Integer;
      fHeight   : Integer;
      //
      function  isVisible: boolean;
      procedure SetVisible(const Value: boolean);

    public
      constructor Create(Ecran: Tform);
      destructor  Destroy; override;
      //
      property Memo     : TMemo   read fMemo      write fMemo;
      property Titre    : String  read fTitre     write fTitre;
      property FileName : String  read fFileName  write fFileName;
      property TableID  : String  read fTableID   write fTableID;
      property Qualif   : String  read fQualif    write fQualif;
      property IDLienOLE: String  read fIDLienOLE write fIDLienOLE;
      property Libelle  : String  read fLibelle   write fLibelle;
      property Affiche  : Boolean read fAffiche   write fAffiche;
      property Print    : Boolean read fPrint     write fPrint;
      property Close    : Boolean read fClose     write fClose;
      property Sauve    : Boolean read fSauve     write fSauve;
      property Visible  : boolean read isVisible  write SetVisible;
      property PosTop   : Integer read fPosTop    write fPosTop;
      property PosLeft  : Integer read fPosLeft   write fPosLeft;
      property Width    : Integer read fWidth     write fWidth;
      property Height   : Integer read fHeight    write fHeight;

      //
      Procedure AfficheRapport;
      Procedure InitRapport;
      Procedure PrintRapport;
      Procedure SauveLigMemo(TxtLigne: string);
      Procedure SauveRapport;
      //
      Procedure SauveRapportLo;
      Procedure ChargeRapportLO;
      Procedure DeleteRapportLo;
      //
    end;

implementation


//création de la classe et création  pbjet Ttoolwindows container et tMemo
constructor TGestionRapport.Create(Ecran : Tform);
begin

  fTitre    := Titre;
  fClose    := Close;
  fAffiche  := Affiche;
  fPrint    := Print;
  fSauve    := Sauve;

  fAffiche  := True;
  fClose    := True;
  fFileName := '';
  fTableID  := TableID;
  fQualif   := Qualif;
  fIDLienOLE:= IDLienOLE;
  fLibelle  := Libelle;

  PosTop  := Round(Screen.Height / 2);
  PosLeft := Round(Screen.Width / 2);
  Height  := round(screen.height / 2);
  Width   := round(screen.width / 4);
  //creation de ttoolwindows container du memo
  fToolWin := TToolWindow97.Create (Ecran);
  fToolWin.parent       := Ecran;
  fToolWin.visible      := False;
  fToolWin.CloseButton  := fClose;
  fToolWin.height       := Height;
  fToolWin.width        := Width;

  //création du mémo dans container ttoolwindows
  fmemo := Tmemo.create (fToolWin);
  fmemo.parent  := fToolWin;
  fmemo.visible := False;
  fMemo.Height  := fToolWin.Height - 120;
  fMemo.Width   := fToolWin.Width - 120;
  fMemo.Align   := alClient;
  fMemo.ScrollBars := ssBoth ;
  fMemo.Enabled    := True;
  fMemo.ReadOnly   := False;

end;

//initialisation du mémo du rapport
procedure TGestionRapport.InitRapport;
begin
  if Assigned(fmemo) then fMemo.clear;
end;

//impression du rapport
procedure TGestionRapport.PrintRapport;
var P: TextFile;
    i: integer;
begin

  AssignPrn(P);

  try
    Rewrite(P);
    for i := 0 to fMemo.Lines.Count - 1 do
      WriteLn(P, fMemo.Lines[i]);
  finally
    CloseFile(P);
  end;

end;

//Sauvegarde du rapport sous format Log
procedure TGestionRapport.SauveRapport;
Var Rep : string;
begin

  Rep := IncludeTrailingBackSlash(TCBPPath.GetCegidDataDistri);

  if fFileName = '' then
    fFileName := Rep + fTitre + '.Log'
  else
    fFileName := Rep + fFileName;

 if FileExists(fFileName) then DeleteFile(fFileName);

 fMemo.Lines.SaveToFile(fFileName);

end;

Procedure TGestionRapport.SauveLigMemo(TxtLigne : string);
begin

  fMemo.Lines.Add(TxtLigne);

end;

//affichage de la Ttoolwindows et du mémo associé... gestion des option d'affichage, impression et sauvaegarde...
Procedure TGestionRapport.AfficheRapport;
begin

  fToolWin.Caption     := fTitre;
  fToolWin.CloseButton := fClose;
  fToolWin.Left        := PosLeft;
  fToolWin.Top         := PosTop;
  fToolWin.Width       := Width;
  fToolWin.Height      := Height;

  if fSauve then SauveRapport;

  fToolWin.visible     := fAffiche;
  fmemo.visible        := fAffiche;

  if fprint then PrintRapport;

end;


destructor TGestionRapport.Destroy;
begin

  fMemo.Free;
  fToolWin.Free;

inherited;
end;


Procedure TGestionRapport.SauveRapportLo;
Var Req     : String;
    QQ      : TQuery;
    TobOLE  : Tob;
begin

	Req := '';

  //chargement du memo à partir du fichier LienOLE (LO)
  req := 'SELECT * FROM LIENSOLE WHERE ';
  Req := Req + ' LO_TABLEBLOB = "' + fTableID + '" AND ';
  Req := Req + ' LO_QUALIFIANTBLOB = "' + fQualif + '" AND ';
  Req := Req + ' LO_IDENTIFIANT ="' + fIDLienOLE + '"';
  Req := Req + ' AND LO_RANGBLOB = 1';

  QQ := OpenSql (Req,true);

  TobOLE := Tob.Create('LIENSOLE' ,Nil, -1);
  TobOLE.selectDB('',QQ);

  if QQ.eof then
  begin
    TobOLE.PutVALUE('LO_TABLEBLOB', fTableID);
    TobOLE.PutVALUE('LO_QUALIFIANTBLOB', fQualif);
    TobOLE.PutVALUE('LO_EMPLOIBLOB', '');
    TobOLE.PutVALUE('LO_IDENTIFIANT', fIDLienOLE);
    TobOLE.PutVALUE('LO_RANGBLOB',1);
  end;

  ferme(QQ);

  TobOLE.PutVALUE('LO_LIBELLE', fLibelle);
  TobOLE.PutVALUE('LO_OBJET', fMemo.Text);

  TobOLE.InsertOrUpdateDB;

  TobOLE.Free;

end;

Procedure TGestionRapport.ChargeRapportLO;
Var Req     : String;
    QQ      : TQuery;
    TobOLE  : Tob;
Begin

	Req := '';

  //chargement du memo à partir du fichier LienOLE (LO)
  req := 'SELECT * FROM LIENSOLE WHERE ';
  Req := Req + ' LO_TABLEBLOB = "' + fTableID + '" AND ';
  Req := Req + ' LO_QUALIFIANTBLOB = "' + fQualif + '" AND ';
  Req := Req + ' LO_IDENTIFIANT ="' + fIDLienOLE + '"';
  Req := Req + ' AND LO_RANGBLOB = 1';

  QQ := OpenSql (Req,true);

  if not QQ.eof then
  begin
    TobOLE := Tob.Create('LIENSOLE' ,Nil, -1);
    TobOLE.selectDB('',QQ);
    fLibelle := TobOLE.GetValue('LO_LIBELLE');
    fMemo.Text := TobOLE.GetVALUE('LO_OBJET');
    TobOLE.Free;
  end;

  ferme (QQ);

end;

Procedure TGestionRapport.DeleteRapportLo;
Var Req     : String;
begin

	Req := '';

  //suppression du Mémo dans fichier LienOLE (LO)
  req := 'DELETE FROM LIENSOLE WHERE ';
  Req := Req + ' LO_TABLEBLOB = "' + fTableID + '" AND ';
  Req := Req + ' LO_QUALIFIANTBLOB = "' + fQualif + '" AND ';
  Req := Req + ' LO_IDENTIFIANT ="' + fIDLienOLE + '"';
  Req := Req + ' AND LO_RANGBLOB = 1';

  ExecuteSQL(Req);

  fMemo.Clear;

end;


function TGestionRapport.isVisible: boolean;
begin
  result := fToolWin.Visible;
end;

procedure TGestionRapport.SetVisible(const Value: boolean);
begin
  fToolWin.Visible := value;
end;

end.
