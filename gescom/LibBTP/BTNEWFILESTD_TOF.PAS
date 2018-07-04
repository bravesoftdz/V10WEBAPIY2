{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTNEWFILESTD ()
Mots clefs ... : TOF;BTNEWFILESTD
*****************************************************************}
Unit BTNEWFILESTD_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     uDbxDataSet, 
     mul, 
{$else}
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     ParamSoc,
     HTB97,
     Windows,
     UTOF ;

Type
  TOF_BTNEWFILESTD = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    NomFicDest : string;
    EmplacementStockage : string;
    function StockeFichierDansReferentiel: boolean;
    procedure BVoirDocClick(Sender: TObject);
  end ;



Implementation
uses UtilFichiers,UFileAssoc;

procedure TOF_BTNEWFILESTD.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTNEWFILESTD.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTNEWFILESTD.OnUpdate ;
begin
  Inherited ;
  if (Thedit(GetControl('FILENAME')).Text = '') or (Thedit(GetControl('DESCRIPTIF')).Text='') then
  begin
    PGIInfo('Vous devez renseigner toutes les informations');
    Ecran.ModalResult := 0;
    Exit;
  end;
  if (THValComboBox(GetControl('CRITSTOCK')).Value = '') then
  begin
    PGIInfo('Vous devez renseigner le critère de classement ');
    Ecran.ModalResult := 0;
    Exit;
  end;
  //
  if not FileExists(Thedit(GetControl('FILENAME')).Text) then
  begin
    PGIInfo('ATTENTION : Ce fichier n''existe pas');
    Ecran.ModalResult := 0;
    Exit;
  end;
  //
  if StockeFichierDansReferentiel then
  begin
    LaTOB.SetString('FILENAME',ExtractFileName(NomFicDest));
    LaTOB.SetString('FILESTOCK',THValComboBox(GetControl('CRITSTOCK')).Value);
    LaTOB.SetString('FILEDESC',Thedit(GetControl('DESCRIPTIF')).Text);
    LaTOB.SetString('OKOK','X');
  end;
end;

procedure TOF_BTNEWFILESTD.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTNEWFILESTD.OnArgument (S : String ) ;
begin
  Inherited ;
  //
  if LaTOB.GetString('CODE')='REFERENCE' then
  begin
    Ecran.Caption := 'Nouveau fichier à intégrer dans référentiel';
  end else
  begin
    Ecran.Caption := 'Sélection d''un fichier';
  end;
  //
  EmplacementStockage := GetParamSocSecur ('SO_BTEMPLFILEREF','');
  if EmplacementStockage ='' then
  begin
    TToolbarButton97 (GetControl('BValider')).Visible := false;
  end;
  IF not DirectoryExists(EmplacementStockage) then
  begin
    PgiInfo ('Merci de créer l''emplacement de stockage');
    TToolbarButton97 (GetControl('BValider')).Visible := false;
  end;
  TToolbarButton97 (GetControl('BVOIR')).Onclick := BVoirDocClick;
end ;

procedure TOF_BTNEWFILESTD.BVoirDocClick(Sender: TObject);
var FileName : string;
begin
  FileName := Thedit(GetControl('FILENAME')).Text;
  OuvreDocument (FileName);
end;

procedure TOF_BTNEWFILESTD.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTNEWFILESTD.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTNEWFILESTD.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_BTNEWFILESTD.StockeFichierDansReferentiel : boolean;
var SrcFile : string;
begin
  result := false;
  SrcFile := Thedit(GetControl('FILENAME')).Text;
  NomFicDest := StoreFileInRef (EmplacementStockage,SrcFile);
  if NomFicDest <> '' then result := true;
end;

Initialization
  registerclasses ( [ TOF_BTNEWFILESTD ] ) ; 
end.
