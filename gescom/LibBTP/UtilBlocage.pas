unit UtilBlocage;

interface

uses Windows,
     Classes,
     HEnt1,
     Paramsoc,
     ComCtrls,
     sysutils,
     HCtrls,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
     HDB,
{$ENDIF}
     controls,
     splash,
     HmsgBox,
     EntGC;

Procedure ControleFactureFinie(duplication : boolean; Var Action : TActionFiche; NaturePiece, Souche, NumPiece : String);
Procedure CreationVerrou;
Procedure CreationVerrouValidation;
procedure DeverouilleValidation (Typedoc : string);
Procedure SupprimeBlocageDoc(NumGuid, TypeDoc, NumDoc : string);

function CreationBlocage(TypeDoc, NumDoc : string; Var UserDoc : String) : String;
Function CtrlBlocageDelDoc(TypeDoc, NumDoc : string; Var UserDoc : String) : String;
Function BlocageDoc(TypeDoc, NumDoc : string; Var UserDoc : String) : String;
Function DeverrouillageEnreg : Boolean;
Function ForceDeblocageDoc(TypeDoc, NumDoc : string) : String;
Function VerrouillageEnreg : Boolean;
function VerrouilleValidation (TypeDoc, NumDoc : string): boolean;
procedure DeblocageVerrouValidation;
//
function BloqueRecupBSV : boolean;
function DeBloqueRecupBSV : boolean;

implementation
uses Facture;


Function CalculGuid : String;
Begin

    Result := AGLGetGUID;

end;


//Function permettant de créer une ligne de blocage dans le cas de plusieurs utilisateurs sur une même pièce
Function CreationBlocage(TypeDoc, NumDoc : string; Var UserDoc : String) : String;
Var Req			: String;
    DateDay : TdateTime;
    QBloc   : TQuery;
Begin

  //if not VerrouillageEnreg then Exit;

  //controle et blocage du document par N° GUID
  Req := 'SELECT * FROM BTBLOCAGE WHERE BTB_TYPE="' + TypeDoc;
  Req := Req + '" AND BTB_IDDOC="' + NumDoc + '" AND BTB_GUID<>"VALIDATION"';
  QBloc := OpenSql (Req, true);

  //Si aucun blocage sur le document cela signifie que l'on continue...
  if Not QBloc.Eof then
  begin
    // je retire l'ancien blocage
    ExecuteSQL('DELETE FROM BTBLOCAGE WHERE BTB_TYPE="' + TypeDoc +'" AND BTB_IDDOC="' + NumDoc + '" AND BTB_GUID<>"VALIDATION"');
    // j'affecte le blocage au nouvel utilisateur
    Result := CalculGuid;
    DateDay := Now;
    //si le blocage qui existe n'est pas pour cet utilisateur on créer une autre ligne afin
    //de gérer en cas de sortie du premier sans déconnexion du deuxième.
    if Qbloc.findfield('BTB_USER').asString <> V_PGI.User then
    begin
      Req := 'INSERT INTO BTBLOCAGE (BTB_GUID, BTB_TYPE, BTB_IDDOC, BTB_USER, BTB_DATECREATION, BTB_HEURECREATION) ';
      Req := Req + 'VALUES ("' + Result + '","'  + TypeDoc + '","' + NumDoc + '","'  + V_PGI.User + '","' + UsDateTime(DateDay) + '","' + USDateTime(DateDay) + '")';
      ExecuteSQL(Req);
    end;
  end;

  Ferme(QBloc);

end;

Function CtrlBlocageDelDoc(TypeDoc, NumDoc : string; Var UserDoc : String) : String;
Var Req			: String;
    QBloc   : TQuery;
    Q		    : TQuery;
Begin
  Result := 'ERREUR';

  //controle et blocage du document par N° GUID
  Req := 'SELECT * FROM BTBLOCAGE WHERE BTB_TYPE="' + TypeDoc;
  Req := Req + '" AND BTB_IDDOC="' + NumDoc + '" AND BTB_GUID<>"VALIDATION"';
  Req := Req + '  AND BTB_USER <> "' + UserDoc + '"';
  QBloc := OpenSql (Req, true);

  if not QBloc.Eof then
  Begin
    Result := 'BLOQUE';
    Q := OpenSQL('SELECT US_ABREGE FROM UTILISAT WHERE US_UTILISATEUR="' + QBloc.FindField('BTB_USER').AsString + '"', False);
    if not Q.EOF then UserDoc := Q.FindField('US_ABREGE').AsString;
    Ferme(Q);
  end;

  Ferme(QBloc);

end;

//Function permettant de controler si le document n'est pas bloqué
//et de faire la sauvegarde dans la table blocage des éléments d'un document
//et de renvoyer le Numéro de GUID associé à l'utilisateur.
Function BlocageDoc(TypeDoc, NumDoc : string; Var UserDoc : String) : String;
Var Req			: String;
    DateDay : TdateTime;
    QBloc, Q		: TQuery;
Begin
    Result := 'ERREUR';
    if not VerrouillageEnreg then Exit;

    //controle et blocage du document par N° GUID
    Req := 'SELECT * FROM BTBLOCAGE WHERE BTB_TYPE="' + TypeDoc;
    Req := Req + '" AND BTB_IDDOC="' + NumDoc + '" AND BTB_GUID<>"VALIDATION"';
    QBloc := OpenSql (Req, true);

    if QBloc.Eof then
       Begin
       Result := CalculGuid;
       DateDay := Now;
       Req := 'INSERT INTO BTBLOCAGE (BTB_GUID, BTB_TYPE, BTB_IDDOC, BTB_USER, BTB_DATECREATION, BTB_HEURECREATION) ';
       Req := Req + 'VALUES ("' + Result + '","'  + TypeDoc + '","' + NumDoc + '","'  + V_PGI.User + '","' + UsDateTime(DateDay) + '","' + USDateTime(DateDay) + '")';
       ExecuteSQL(Req);
       end
    Else
    Begin
      Result := 'BLOQUE';
      Q := OpenSQL('SELECT US_ABREGE FROM UTILISAT WHERE US_UTILISATEUR="' + QBloc.FindField('BTB_USER').AsString + '"', False);
      if not Q.EOF then UserDoc := Q.FindField('US_ABREGE').AsString;
      Ferme(Q);
    end;

    Ferme(QBloc);

    if not DeverrouillageEnreg then
       Result := 'ERREUR';

end;


//fonction permettant de contrôler si le document n'est pas terminé.
//si c'est le cas passage du code action en TaConsult
//sinon aucune modif n'est effectuée
Procedure ControleFactureFinie(duplication : boolean; var Action : TActionFiche; NaturePiece, Souche, NumPiece : String);
Var Req				: String;
    AffDevis	: String;
    QBloc, Q	: TQuery;
Begin
    if duplication then exit;
    if action <> taModif then exit;

    Req := 'SELECT GP_AFFAIREDEVIS FROM PIECE WHERE GP_NATUREPIECEG="' + NaturePiece;
    Req := Req + '" AND GP_SOUCHE="' + Souche;
    Req := Req + '" AND GP_NUMERO="' + NumPiece + '"';
    
    QBloc := OpenSql (Req, true);

    if QBloc.Eof then
       Begin
       Ferme(QBloc);
       Exit
       end;

    AffDevis := QBloc.FindField('GP_AFFAIREDEVIS').AsString;

    Ferme(Qbloc);

    if AffDevis = '' then
       Begin
       Ferme(QBloc);
       Exit
       end;

    Req := 'Select AFF_ETATAFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="' + AffDevis + '"';
    Q := OpenSql(Req, True);

    If Q.Eof then
       Begin
       Ferme(Qbloc);
       Ferme(Q);
       Exit;
       end;

    IF Q.FindField('AFF_ETATAFFAIRE').AsString = 'TER' then Action := TaConsult;

    Ferme(Q);

end;

//Procédure de déverrouillage des documents...
Procedure SupprimeBlocageDoc(NumGuid, TypeDoc, NumDoc : string);
Begin

//	  if not VerrouillageEnreg then Exit;

//  ExecuteSql('DELETE FROM BTBLOCAGE WHERE BTB_GUID="' + NumGuid + '" AND BTB_TYPE="' + TypeDoc + '" AND BTB_IDDOC="' + Numdoc + '"');
  ExecuteSql('DELETE FROM BTBLOCAGE WHERE BTB_GUID="' + NumGuid + '"');

//    if not DeverrouillageEnreg then
//       NumGuid := 'ERREUR';

end;

//Procédure de forçage du déblocage de document...
Function ForceDeblocageDoc(TypeDoc, NumDoc : string): String;
Begin
		Result := '';
	  if not VerrouillageEnreg then Exit;

  	ExecuteSql('DELETE FROM BTBLOCAGE WHERE BTB_TYPE="' + TypeDoc + '" AND BTB_IDDOC="' + Numdoc + '"');

    if not DeverrouillageEnreg then Result := 'ERREUR';
       Result := 'ERREUR';

end;

//Function de verrouillage du document en fonction du code blocage
Function VerrouillageEnreg : Boolean;
Var Degage		 : Boolean;
    INbPassage : integer;
Begin

	  CreationVerrou;

		degage := false;
    INbPassage := 1;
    Result := false;

    //Controle et verrouillage avant vérification du blocage du document
		repeat
    	if ExecuteSql ('UPDATE BTBLOCAGE SET BTB_IDDOC="VERROU",BTB_USER="'+ V_PGI.User +'" WHERE BTB_IDDOC="" AND BTB_GUID="XXX"') <= 0 then
      begin
      	Result := true;
        break;
      end;
      sleep(200);
      Inc(INbPassage);
      if INbPassage > 10 then break;
    until degage;

end;

//Function de déverrouillage du document en fonction du code blocage
Function DeverrouillageEnreg : Boolean;
begin

    Result := True;

    //Déverrouillage du droit de blocage du document
		if ExecuteSql ('UPDATE BTBLOCAGE SET BTB_IDDOC="",BTB_USER="" WHERE BTB_IDDOC="VERROU" AND BTB_USER="'+V_PGI.user+'" AND BTB_GUID="XXX"') <= 0 then
       Begin
       Result := false;
       end;

end;

//Procedure qui va permettre de créer l'enregistrement de verrouillage dans la
//Table blocage si celui-ci n'existe pas !!!
Procedure CreationVerrou;
Var QVerrou : TQuery;
    Req     : String;
Begin

    QVerrou := OpenSql ('SELECT BTB_GUID FROM BTBLOCAGE WHERE BTB_GUID="XXX"', true);

    if QVerrou.Eof then
       Begin
       req := 'INSERT INTO BTBLOCAGE (BTB_GUID, BTB_TYPE, BTB_IDDOC, BTB_USER, BTB_DATECREATION, BTB_HEURECREATION) ';
       req := req + 'VALUES ("XXX","","","","' + UsDateTime(iDate1900) + '","' + USDateTime(iDate1900) + '")';
       ExecuteSQL(req);
       end;

    Ferme(QVerrou);

end;


Procedure CreationVerrouValidation;
Var QVerrou : TQuery;
    Req     : String;
Begin
  QVerrou := OpenSql ('SELECT BTB_GUID FROM BTBLOCAGE WHERE BTB_GUID="VALIDATION"', true);
  if QVerrou.Eof then
  Begin
    req := 'INSERT INTO BTBLOCAGE (BTB_GUID, BTB_TYPE, BTB_IDDOC, BTB_USER, BTB_DATECREATION, BTB_HEURECREATION) ';
    req := req + 'VALUES ("VALIDATION","","","","' + UsDateTime(iDate1900) + '","' + USDateTime(iDate1900) + '")';
    ExecuteSQL(req);
  end;
  Ferme(QVerrou);
end;


function VerrouilleValidation (TypeDoc, NumDoc : string): boolean;
Var Degage		 : Boolean;
    NbPassage : integer;
    UserBloc,LibelleBloc : string;
    QQ : Tquery;
    document : string;
begin
  NbPassage := 0;
  result := true;
	CreationVerrouValidation;
  //
  if pos(TypeDoc,'BLF;LFR;FF;LBT;XXX') = 0 then exit;
  //
  result := false;
  //
  repeat
    if ExecuteSql ('UPDATE BTBLOCAGE SET BTB_TYPE="'+TypeDoc+'",BTB_IDDOC="'+NumDoc+'",BTB_USER="'+ V_PGI.User +'" WHERE BTB_IDDOC="" AND BTB_GUID="VALIDATION"') > 0 then
    begin
      Result := true;
      break;
    end else
    begin
      QQ := OpenSql ('SELECT BTB_USER,BTB_IDDOC,BTB_TYPE FROM BTBLOCAGE WHERE BTB_GUID="VALIDATION"',true);
      if not QQ.eof then
      begin
        UserBloc := QQ.findField ('BTB_USER').AsString;
        LibelleBloc := rechdom('TTUTILISATEUR',UserBloc,false);
        if QQ.findField ('BTB_TYPE').asString <> 'XXX' then
        begin
          Document := ' ' +rechdom('GCNATUREPIECEG',QQ.findField ('BTB_TYPE').asString,false)+' : ' + QQ.findField ('BTB_IDDOC').AsString;
        end else
        begin
          Document := ' Saisie consommations ';
        end;
      end;
      ferme (QQ);
    end;
    //
    Inc(NbPassage);
    if PgiAsk ('ATTENTE: Validation en cours '+'('+InttoStr(NbPassage)+') de '+userbloc+' : '+LibelleBloc+''+Document+'.#13#10Désirez-vous réessayer ?')=mrNo then break;
    //
  until Degage;

end;

procedure DeverouilleValidation (Typedoc : string);
begin
  //
  if pos(TypeDoc,'BLF;LFR;FF;LBT;XXX') = 0 then exit;
  //
  ExecuteSql ('UPDATE BTBLOCAGE SET BTB_TYPE="",BTB_IDDOC="",BTB_USER="" WHERE BTB_GUID="VALIDATION"');
end;

procedure DeblocageVerrouValidation;
begin
  if PgiAsk ('Confirmez-vous le lancement du déblocage du verrou des consommations ?')=mrYes then
  	ExecuteSql ('UPDATE BTBLOCAGE SET BTB_TYPE="",BTB_IDDOC="",BTB_USER="" WHERE BTB_GUID="VALIDATION"');
end;

function BloqueRecupBSV : boolean;
Var QVerrou : TQuery;
    Req     : String;
Begin
  req := 'INSERT INTO BTBLOCAGE (BTB_TYPE, BTB_IDDOC, BTB_USER, BTB_DATECREATION, BTB_HEURECREATION) ';
  req := req + 'VALUES ("VERROUBSV","","'+V_PGI.User+'","' + UsDateTime(iDate1900) + '","' + USDateTime(iDate1900) + '")';
  Result := (ExecuteSQL(req)>0);
end;

function DeBloqueRecupBSV : boolean;
Var QVerrou : TQuery;
    Req     : String;
Begin
  ExecuteSQL('DELETE FROM BTBLOCAGE WHERE BTB_TYPE="VERROUBSV"');
end;

end.
