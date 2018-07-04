{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/01/2012
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CREATEDEMPRIX ()
Mots clefs ... : TOF;CREATEDEMPRIX
*****************************************************************}
Unit CREATEDEMPRIX_TOF ;

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
     UTOF,
     Utob,
     variants,
     uentcommun;

Type
  TOF_CREATEDEMPRIX = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  Private

    NaturePiece : THEdit;
    Souche      : THEdit;
    NumeroPiece : THEdit;
    Indice      : THEdit;
    Unique      : THEdit;
    Libelle     : THEdit;

    LIBPiece    : THLabel;
    LIBDdePrix  : THLabel;

    Traite      : TCheckBox;

    TheAction   : String;

    fFromTOB    : Boolean;

    Cledoc      : R_Cledoc;

    //
    TOBparam    : TOB;
    TobPiece    : TOB;
    TobDemPrix  : TOB;
    TOBArtDemPx : TOB;
    TobFrsDemPx : TOB;
    TobDetDemPx : TOB;
    TobDdePrix  : Tob;
    //
    procedure CalculNumUnique;
    procedure ChargeCleDoc;
    procedure ChargeTob;
    procedure ControleChamp(Champ, Valeur: String);
    procedure ControleCritere(Critere: String);
    procedure AffichageEcran;

  Public

  end ;

const
  // libellés des messages de la TOM Affaire
  TexteMessage : array [1..6] of string = (
    {1} 'Confirmez-vous la suppression de cette demande de prix ?',
    {2} 'Erreur de suppression des demandes de Prix',
    {3} 'Erreur de suppression des entêtes de demande de prix',
    {4} 'Erreur de suppression des Articles en demande de prix',
    {5} 'Erreur de suppression des fournisseurs demande de prix',
    {6} 'Erreur de suppression des détails de demande de prix'
    ) ;

Implementation
uses UtilTOBPiece,
     UDemandePrix;


procedure TOF_CREATEDEMPRIX.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CREATEDEMPRIX.OnDelete ;
Var NbError : integer;
begin
  Inherited ;

  NbError := 0;

  if (PGIAsk(TexteMessage[1], ecran.Caption) = mrYes) then
  begin
    if not FfromTob then
    begin
      if (TobDdePrix <> nil) then
      begin
        TobDdePrix.DeleteDB();
        //suppression des tables associées à l'entête
        NbError := SupprimeDdePrix(Cledoc, StrToInt(Unique.text),'ARTICLEDEMPRIX', 'BDP', TTdArticleDemPrix);
        NbError := SupprimeDdePrix(Cledoc, StrToInt(Unique.text),'DETAILDEMPRIX',  'BD0', TtdDetailDemPrix);
        NbError := SupprimeDdePrix(Cledoc, StrToInt(Unique.text),'FOURLIGDEMPRIX', 'BD1', TtdFournDemprix);
      end;
    end
    else
    begin
      if (TobDdePrix <> nil) then
      begin
        //suppression de l'entête de demande de prix
        NbError := SupprimeDdePrix(TobDdePrix, StrToInt(Unique.text));
        if NbError < 0 Then PGIInfo(TexteMessage[3], Ecran.caption);
        //suppression des tables associées à l'entête
        NbError := SupprimeDdePrix(TobArtDemPx, StrToInt(Unique.text));
        if NbError < 0 Then PGIInfo(TexteMessage[4], Ecran.caption);
        NbError := SupprimeDdePrix(TobFrsDemPx, StrToInt(Unique.text));
        if NbError < 0 Then PGIInfo(TexteMessage[5], Ecran.caption);
        NbError := SupprimeDdePrix(TobDetDemPx, StrToInt(Unique.text));
        if NbError < 0 Then PGIInfo(TexteMessage[6], Ecran.caption);
      end;
    end;
  end;

  if NbError < 0 then PGIInfo(TexteMessage[2], Ecran.caption);

  ecran.close;

end ;

procedure TOF_CREATEDEMPRIX.OnUpdate ;
var Ok_Write : Boolean;
begin
  Inherited ;

  Ok_Write := FFromTob;

  repeat
    if TheAction = 'CREATION' then
    begin
      sleep(200);
      TobDdePrix.PutValue('BPP_NATUREPIECEG', NaturePiece.text);
      TobDdePrix.PutValue('BPP_SOUCHE', Souche.text);
      TobDdePrix.PutValue('BPP_NUMERO', NumeroPiece.text);
      TobDdePrix.PutValue('BPP_INDICEG', Indice.text);
      TobDdePrix.PutValue('BPP_UNIQUE', Unique.text);
      TobDdePrix.PutValue('BPP_LIBELLE', Libelle.text);
      TobDdePrix.PutValue('BPP_TRAITE', '-');
      TobDdePrix.PutValue('BPP_ENVOYE', '-');
      TobDdePrix.PutValue('BPP_LASTLIG', '0');
      TOBPAram.PutValue('RETOUR','X');
      if not FFromTob then Ok_Write := TobDdePrix.InsertDB(nil);
    end
    Else if TheAction = 'MODIFICATION' then
    begin
      TobDdePrix.PutValue('BPP_LIBELLE', Libelle.text);
      TOBPAram.PutValue('RETOUR','X');
      if Not Ffromtob then Ok_Write := TobDdePrix.UpdateDB;
    end;
  Until ok_write;


end ;

procedure TOF_CREATEDEMPRIX.CalculNumUnique ;
var StSql   : String;
    QQ      : TQuery;
    Maxnum  : Integer;
    iInd    : Integer;
    TobSort : Tob;
begin

  if Not FfromTOb then
  begin
    StSQL := 'SELECT ISNULL(MAX(BPP_UNIQUE),0) as NUMMAX FROM PIECEDEMPRIX WHERE ' + WherePiece(Cledoc, ttdPieceDemPrix, true);

    QQ := OpenSQL(StSQl, True);

    MaxNum := QQ.findField('NUMMAX').asInteger;
    inc(MaxNum);

    Ferme(QQ);
  end
  else
  begin
    TobSort := TobDemPrix;
    TobSort.Detail.Sort('BPP_UNIQUE');
    MaxNum := TobSort.Detail[TobDemPrix.detail.count -1].GetInteger('BPP_UNIQUE') + 1;
  end;

  Unique.text := IntToStr(MaxNum);

end;

procedure TOF_CREATEDEMPRIX.OnLoad ;
begin
  Inherited ;

  ChargeTob;

  if (unique.text = '0') or (Unique.text = '') then CalculNumUnique;

  LibPiece.caption := 'Pièce N° ' + NaturePiece.Text + '-' + Souche.text + '-' + NumeroPiece.text + '-' + Indice.Text;
  LIBDdePrix.Caption := 'Demande de Prix N° ' +  Unique.text;

end ;

Procedure TOF_CREATEDEMPRIX.ChargeTob;
Var StSQL : String;
    QQ    : TQuery;
begin

  if TheAction = 'MODIFICATION' then
  begin
    if Not FfromTOb then
    begin
      StSQL := 'SELECT * FROM PIECEDEMPRIX WHERE ' + WherePiece(Cledoc, ttdPieceDemPrix, true) + ' AND BPP_UNIQUE='+ Unique.text;
      QQ := OpenSQL(StSQL, True);
      if Not QQ.eof then TobDdePrix.SelectDB('', QQ);
    end
    else
      TobDdePrix := TobDemPrix.FindFirst(['BPP_UNIQUE'], [Unique.text], true);
    TobDdePrix.PutEcran(ecran);
  end
  else
  begin
    TheAction := 'CREATION';
    if FfromTob then TobDdePrix := Tob.create('PIECEDEMPRIX', TobDemPrix, -1);
  end;

end;

procedure TOF_CREATEDEMPRIX.OnArgument (S : String ) ;
Var critere : String;
    X       : Integer;
    Champ   : string;
    Valeur  : string;
begin
  Inherited ;

  if LaTOB <> nil then
  begin
    fFromTOB := true;
    TOBparam    := LaTOB;
    TobPiece    := TOB(TOBParam.data);
    TobDemPrix  := TOB(TOBPiece.data);
    TOBArtDemPx := TOB(TOBDemPrix.data);
    TobFrsDemPx := TOB(TOBArtDemPx.data);
    TobDetDemPx := TOB(TobFrsDemPx.data);
  end
  else
  begin
    fFromTOB := False;
    TobDdePrix  := Tob.Create('PIECEDEMPRIX', nil, -1);
  end;

  NaturePiece := THEdit(ecran.FindComponent('BPP_NATUREPIECEG'));
  Souche      := THEdit(ecran.FindComponent('BPP_SOUCHE'));
  NumeroPiece := THEdit(ecran.FindComponent('BPP_NUMERO'));
  Indice      := THEdit(ecran.FindComponent('BPP_INDICEG'));
  Unique      := THEdit(ecran.FindComponent('BPP_UNIQUE'));
  Libelle     := THEdit(ecran.FindComponent('BPP_LIBELLE'));

  LibPiece    := THLabel(ecran.FindComponent('TBPP_PIECEG'));
  LibDdePrix  := THLabel(ecran.FindComponent('TBPP_UNIQUE'));

  Traite      := TCheckBox(ecran.findComponent('BPP_TRAITE'));

  //Récupération valeur de argument
  Critere:=(Trim(ReadTokenSt(S)));

  while (Critere <> '') do
  begin
    if Critere <> '' then
      begin
      X := pos (';', Critere) ;
      if x = 0 then
        X := pos ('=', Critere) ;
      if x <> 0 then
        begin
        Champ := copy (Critere, 1, X - 1) ;
        Valeur:= Copy (Critere, X + 1, length (Critere) - X) ;
        ControleChamp(champ, valeur);
        end
      end;
    ControleCritere(Critere);
    Critere   := (Trim(ReadTokenSt(S)));
  end;

  ecran.caption := 'Demande de prix';

  ChargeCleDoc;

  AffichageEcran;

end ;

Procedure TOF_CREATEDEMPRIX.ChargeCleDoc;
begin

  cledoc.NaturePiece := NaturePiece.Text;
  cledoc.Souche      := Souche.Text;
  cledoc.NumeroPiece := StrToInt(NumeroPiece.text);
  Cledoc.indice      := StrToInt(Indice.text);

end;

Procedure TOF_CREATEDEMPRIX.AffichageEcran;
begin
  if TheAction='CREATION' then
  begin
    GetControl('Bdelete').Visible := False;
  end
  else if TheAction='MODIFICATION' then
  begin
    GetControl('Bdelete').Visible := True;
  end;

end;


procedure TOF_CREATEDEMPRIX.OnClose ;
begin
  Inherited ;

  if TheAction= 'CREATION' then
  begin
    if TobDdePrix.GetInteger('BPP_UNIQUE') = 0 then FreeAndNil(TobDdePrix);
  end
  else
    if not FfromTob then FreeAndNil(TobDdePrix);

end ;

procedure TOF_CREATEDEMPRIX.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CREATEDEMPRIX.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CREATEDEMPRIX.ControleChamp(Champ, Valeur: String);
begin

  if Champ = 'NATUREPIECEG' then
    NaturePiece.text := Valeur
  else if Champ = 'SOUCHE' then
    Souche.text      := Valeur
  else if Champ = 'NUMERO' then
    NumeroPiece.Text := Valeur
  else if Champ = 'INDICEG' then
    Indice.text      := Valeur
  else if Champ = 'UNIQUE' then
    Unique.Text      := Valeur
  else if champ = 'ACTION' then
    TheAction        := Valeur;

end;

procedure TOF_CREATEDEMPRIX.ControleCritere(Critere: String);
begin
end;


Initialization
  registerclasses ( [ TOF_CREATEDEMPRIX ] ) ; 
end.

