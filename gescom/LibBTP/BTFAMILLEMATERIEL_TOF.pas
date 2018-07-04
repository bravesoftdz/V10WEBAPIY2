{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 03/06/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTFAMILLEMATERIEL ()
Mots clefs ... : TOF;BTFAMILLEMATERIEL
*****************************************************************}
Unit BTFAMILLEMATERIEL_TOF ;

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
     uTOB,
     HTB97,
     UTOF ; 

Type
  TOF_BTFAMILLEMATERIEL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Action      : TActionFiche;
    //
    CodeFamille : THEdit;
    LibFamille  : THEdit;
    Cadencement : THValComboBox;
    //
    LibAlpha1   : THEdit;
    LibAlpha2   : THEdit;
    LibAlpha3   : THEdit;
    LibAlpha4   : THEdit;
    LibAlpha5   : THEdit;
    LibAlpha6   : THEdit;
    LibAlpha7   : THEdit;
    LibAlpha8   : THEdit;
    LibAlpha9   : THEdit;
    LibAlpha10  : THEdit;
    //
    Libdate1    : THEdit;
    Libdate2    : THEdit;
    Libdate3    : THEdit;
    Libdate4    : THEdit;
    //
    LibBoolean1 : THEdit;
    LibBoolean2 : THEdit;
    LibBoolean3 : THEdit;
    LibBoolean4 : THEdit;
    //
    NonGereplanning : THCheckbox;
    //
    BTDuplication : TToolbarButton97;
    BTSupprime    : TToolbarButton97;
    //
    TobFamMat     : Tob;
    //
    StSQL         : String;
    //
    QQ            : TQuery;
    //
    procedure ChargeZoneEcran;
    procedure ChargeZoneTOB;
    procedure Controlechamp(Champ, Valeur: String);
    procedure CreateTOB;
    procedure DestroyTOB;
    procedure GetObjects;
    procedure SetScreenEvents;
    procedure OnDupliqueFiche(Sender: TObject);
  end ;

const
	// libellés des messages
  TexteMessage: array[1..3] of string  = (
          {1}  'Le Code Famille est obligatoire'
          {2}, 'Suppression impossible ce code famille est déjà utilisé sur un Matériel'
          {3}, 'La suppression a échoué'
                                         );

Implementation

procedure TOF_BTFAMILLEMATERIEL.OnNew ;
begin
  Inherited ;

  CodeFamille.Text  := '';
  LibFamille.Text   := '';
  Cadencement.value := '0';
  //
  LibAlpha1.Text    := '';
  LibAlpha2.Text    := '';
  LibAlpha3.Text    := '';
  LibAlpha4.Text    := '';
  LibAlpha5.Text    := '';
  LibAlpha6.Text    := '';
  LibAlpha7.Text    := '';
  LibAlpha8.Text    := '';
  LibAlpha9.Text    := '';
  LibAlpha10.Text   := '';
  //
  Libdate1.Text     := '';
  Libdate2.Text     := '';
  Libdate3.Text     := '';
  Libdate4.Text     := '';
  //
  LibBoolean1.Text  := '';
  LibBoolean2.Text  := '';
  LibBoolean3.Text  := '';
  LibBoolean4.Text  := '';
  //
end ;

procedure TOF_BTFAMILLEMATERIEL.OnDelete ;
begin
  Inherited ;

  Ecran.ModalResult := 0;

  if PGIAsk('Confirmez-vous la suppression de cette Famille Matériel : ' + CodeFamille.Text + ' ?', 'Famille Matériel') = Mryes then
  begin
    //contrôle si Famille non présente sur Matériel...
    StSQL := 'SELECT BMA_CODEMATERIEL FROM BTMATERIEL WHERE BMA_CODEFAMILLE="' + CodeFamille.Text + '"';
    if existeSQL(STSQL) then
      PGIError(TexteMessage[2], 'Famille Matériel')
    else
    Begin
      //suppression pure et simple de l'enregistrement avec confirmation
      StSQL := 'DELETE BTFAMILLEMAT WHERE BFM_CODEFAMILLE="' + CodeFamille.Text + '"';
      if ExecuteSQL(StSQL)=0 then PGIError(TexteMessage[3], 'Famille Matériel');;
      Ecran.ModalResult := 1;
    end;
  end;

end ;

procedure TOF_BTFAMILLEMATERIEL.OnUpdate ;
begin
  Inherited ;

  ChargeZoneTOB;

  //mise à jour de la table Famille Matériel
  TobFamMat.InsertOrUpdateDB(true);

end ;

procedure TOF_BTFAMILLEMATERIEL.OnLoad ;
begin
  Inherited ;

  StSQL := 'SELECT * FROM BTFAMILLEMAT WHERE BFM_CODEFAMILLE="' + CodeFamille.Text + '"';
  QQ := OpenSQL(StSQL, True);

  If not QQ.Eof then
  begin
      TobFamMat.SelectDB('BTFAMILLEMAT', QQ);
  end;

  ChargeZoneEcran;

  ferme(QQ); 

end ;

procedure TOF_BTFAMILLEMATERIEL.OnArgument (S : String ) ;
var x       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin
  Inherited ;
  //
  //Chargement des zones ecran dans des zones programme
  GetObjects;
  //
  CreateTOB;
  //
  Critere := uppercase(Trim(ReadTokenSt(S)));
  while Critere <> '' do
  begin
     x := pos('=', Critere);
     if x <> 0 then
        begin
        Champ  := copy(Critere, 1, x - 1);
        Valeur := copy(Critere, x + 1, length(Critere));
        end
     else
        Champ := Critere;
     ControleChamp(Champ, Valeur);
     Critere:= uppercase(Trim(ReadTokenSt(S)));
  end;

  //Gestion des évènement des zones écran
  SetScreenEvents;

  Cadencement.Plus := ' AND CO_LIBRE="PMA"';

  if Action <> taCreat then
    CodeFamille.enabled := false
  else
  begin
    BTDuplication.Visible := False;
    BTSupprime.Visible    := False;
  end;


end ;

procedure TOF_BTFAMILLEMATERIEL.OnClose ;
begin
  Inherited ;

  DestroyTOB;

end ;

procedure TOF_BTFAMILLEMATERIEL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTFAMILLEMATERIEL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTFAMILLEMATERIEL.GetObjects;
begin

  CodeFamille := THEdit(Getcontrol('BFM_CODEFAMILLE'));
  LibFamille  := THEdit(Getcontrol('BFM_LIBELLE'));
  Cadencement := THValCombobox(GetControl('BFM_RYTHMEGESTION'));
  //
  LibAlpha1   := THEdit(Getcontrol('BFM_LIBALPHA1'));
  LibAlpha2   := THEdit(Getcontrol('BFM_LIBALPHA2'));
  LibAlpha3   := THEdit(Getcontrol('BFM_LIBALPHA3'));
  LibAlpha4   := THEdit(Getcontrol('BFM_LIBALPHA4'));
  LibAlpha5   := THEdit(Getcontrol('BFM_LIBALPHA5'));
  LibAlpha6   := THEdit(Getcontrol('BFM_LIBALPHA6'));
  LibAlpha7   := THEdit(Getcontrol('BFM_LIBALPHA7'));
  LibAlpha8   := THEdit(Getcontrol('BFM_LIBALPHA8'));
  LibAlpha9   := THEdit(Getcontrol('BFM_LIBALPHA9'));
  LibAlpha10  := THEdit(Getcontrol('BFM_LIBALPHA10'));
  //
  Libdate1    := THEdit(Getcontrol('BFM_LIBDATE1'));
  Libdate2    := THEdit(Getcontrol('BFM_LIBDATE2'));
  Libdate3    := THEdit(Getcontrol('BFM_LIBDATE3'));
  Libdate4    := THEdit(Getcontrol('BFM_LIBDATE4'));
  //
  LibBoolean1 := THEdit(Getcontrol('BFM_LIBBOOLEAN1'));
  LibBoolean2 := THEdit(Getcontrol('BFM_LIBBOOLEAN2'));
  LibBoolean3 := THEdit(Getcontrol('BFM_LIBBOOLEAN3'));
  LibBoolean4 := THEdit(Getcontrol('BFM_LIBBOOLEAN4'));
  //
  NonGereplanning := THCheckbox(Getcontrol('NONGEREPLANNING'));
  //
  BTDuplication := TToolbarButton97(GetControl('BDUPLICATION'));
  BTSupprime    := TToolbarButton97(GetControl('BDELETE'));

end;

procedure TOF_BTFAMILLEMATERIEL.SetScreenEvents;
begin

  BtDuplication.OnClick := OnDupliqueFiche;

end;

Procedure TOF_BTFAMILLEMATERIEL.OnDupliqueFiche(Sender : TObject);
begin

  if action = TaCreat then Exit;

  Action                := TaCreat;

  BTDuplication.Visible := False;
  BTSupprime.Visible    := False;
  //
  CodeFamille.Text      := '';
  LibFamille.Text       := '';
  codeFamille.enabled   := True;
  //
  CodeFamille.SetFocus;
end;


Procedure TOF_BTFAMILLEMATERIEL.Controlechamp(Champ, Valeur : String);
begin

  if Champ='ACTION' then
  begin
    if      Valeur='CREATION'     then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end;

  if Champ='CODEFAMILLE' then CodeFamille.text := Valeur;

end;

Procedure TOF_BTFAMILLEMATERIEL.CreateTOB;
begin

  TobFamMat     := Tob.Create('BTFAMILLEMAT', nil, -1);

end;

procedure TOF_BTFAMILLEMATERIEL.DestroyTOB;
begin

  FreeAndNil(TobFamMat);

end;

Procedure TOF_BTFAMILLEMATERIEL.ChargeZoneEcran;
begin

  CodeFamille.Text  := TOBFamMat.GetString('BFM_CODEFAMILLE');
  LibFamille.Text   := TOBFamMat.GetString('BFM_LIBELLE');
  Cadencement.Value := TobFamMat.GetString('BFM_RYTHMEGESTION');
  //
  LibAlpha1.Text    := TOBFamMat.GetString('BFM_LIBALPHA1');
  LibAlpha2.Text    := TOBFamMat.GetString('BFM_LIBALPHA2');
  LibAlpha3.Text    := TOBFamMat.GetString('BFM_LIBALPHA3');
  LibAlpha4.Text    := TOBFamMat.GetString('BFM_LIBALPHA4');
  LibAlpha5.Text    := TOBFamMat.GetString('BFM_LIBALPHA5');
  LibAlpha6.Text    := TOBFamMat.GetString('BFM_LIBALPHA6');
  LibAlpha7.Text    := TOBFamMat.GetString('BFM_LIBALPHA7');
  LibAlpha8.Text    := TOBFamMat.GetString('BFM_LIBALPHA8');
  LibAlpha9.Text    := TOBFamMat.GetString('BFM_LIBALPHA9');
  LibAlpha10.Text   := TOBFamMat.GetString('BFM_LIBALPHA10');
  //
  Libdate1.Text     := TOBFamMat.GetString('BFM_LIBDATE1');
  Libdate2.Text     := TOBFamMat.GetString('BFM_LIBDATE2');
  Libdate3.Text     := TOBFamMat.GetString('BFM_LIBDATE3');
  Libdate4.Text     := TOBFamMat.GetString('BFM_LIBDATE4');
  //
  LibBoolean1.Text  := TOBFamMat.GetString('BFM_LIBBOOLEAN1');
  LibBoolean2.Text  := TOBFamMat.GetString('BFM_LIBBOOLEAN2');
  LibBoolean3.Text  := TOBFamMat.GetString('BFM_LIBBOOLEAN3');
  LibBoolean4.Text  := TOBFamMat.GetString('BFM_LIBBOOLEAN4');
  //
  NonGereplanning.Checked := TobFamMat.GetBoolean('BFM_NONGEREPLANNING');
end;

Procedure TOF_BTFAMILLEMATERIEL.ChargeZoneTOB;
begin

  TOBFamMat.PutValue('BFM_CODEFAMILLE',CodeFamille.Text);
  TOBFamMat.PutValue('BFM_LIBELLE',LibFamille.Text);
  TobFamMat.PutValue('BFM_RYTHMEGESTION',Cadencement.value);
  //
  TOBFamMat.PutValue('BFM_LIBALPHA1',LibAlpha1.Text);
  TOBFamMat.PutValue('BFM_LIBALPHA2',LibAlpha2.Text);
  TOBFamMat.PutValue('BFM_LIBALPHA3',LibAlpha3.Text);
  TOBFamMat.PutValue('BFM_LIBALPHA4',LibAlpha4.Text);
  TOBFamMat.PutValue('BFM_LIBALPHA5',LibAlpha5.Text);
  TOBFamMat.PutValue('BFM_LIBALPHA6',LibAlpha6.Text);
  TOBFamMat.PutValue('BFM_LIBALPHA7',LibAlpha7.Text);
  TOBFamMat.PutValue('BFM_LIBALPHA8',LibAlpha8.Text);
  TOBFamMat.PutValue('BFM_LIBALPHA9',LibAlpha9.Text);
  TOBFamMat.PutValue('BFM_LIBALPHA10',LibAlpha10.Text);
  //
  TOBFamMat.PutValue('BFM_LIBDATE1',Libdate1.Text);
  TOBFamMat.PutValue('BFM_LIBDATE2',Libdate2.Text);
  TOBFamMat.PutValue('BFM_LIBDATE3',Libdate3.Text);
  TOBFamMat.PutValue('BFM_LIBDATE4',Libdate4.Text);
  //
  TOBFamMat.PutValue('BFM_LIBBOOLEAN1',LibBoolean1.Text);
  TOBFamMat.PutValue('BFM_LIBBOOLEAN2',LibBoolean2.Text);
  TOBFamMat.PutValue('BFM_LIBBOOLEAN3',LibBoolean3.Text);
  TOBFamMat.PutValue('BFM_LIBBOOLEAN4',LibBoolean4.Text);
  //
  if NonGereplanning.checked then
  TobFamMat.PutValue('BFM_NONGEREPLANNING', 'X')
  else
  TobFamMat.PutValue('BFM_NONGEREPLANNING', '-');

end;

Initialization
  registerclasses ( [ TOF_BTFAMILLEMATERIEL ] ) ;
end.

