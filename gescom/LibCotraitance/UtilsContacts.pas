{***********UNITE*************************************************
Auteur  ...... : Franck VAUTRAIN
Créé le ...... : 19/07/2011
Modifié le ... :   /  /
Description .. : Classes de gestion des contact
Suite ........ : récupération des contacts
Suite ........ : pour un tiers client ou fournisseur
Suite ........ : si plusieurs contact chargement liste
Suite ........ : si un seul contact controle en fonction
Suite ........ : boolean si défaut ou non doit être pris
Mots clefs ... : CONTACT; CHARGEMENT; LISTE
*****************************************************************}
unit UtilsContacts;
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
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      Fe_Main,
      EdtREtat,
      EdtRDoc,
      {$ELSE}
      HDB,
      eMul,
      Maineagl,
      UtileAGL,
      {$ENDIF}
      uTOB,
      Lookup,
      uTOF;


Type
    TGestionContact = class(Tobject)
    private
      fForm                     : TForm;
      fCleDocNatPiece           : HString;
      fCleDocSouche             : HString;
      fCleDocNumPiece           : HString;
      fCleDocIndPiece           : HString;
      fTypeContact              : HString;
      fcontact                  : HString;
      fTypeTiers                : HString;
      fTiers                    : HString;
      fAuxiliaire               : HString;
      fLibelleTiers             : HString;
      fNumtel                   : HString;
      fPart                     : HString;
      fAffaire                  : hString;
      fMail                     : hString;
      fQualifMail               : hString;
      fTobContact               : TOB;

    public
      constructor Create(Sender : Tobject);
      destructor  Destroy; override;
      //Déclaration des propriété de la classe
      property form             : TForm           Read fForm             write fForm;
      property CleDocNatPiece   : HString         read fCleDocNatPiece   write fCleDocNatPiece;
      property CleDocSouche     : HString         read fCleDocSouche     write fCleDocSouche;
      property CleDocNumPiece   : HString         read fCleDocNumPiece   write fCleDocNumPiece;
      property CledocIndPiece   : HString         read fCleDocIndPiece   write fCleDocIndPiece;
      property TypeContact      : HString         read fTypeContact      write fTypeContact;
      property Contact          : HString         read fContact          write fContact;
      property TypeTiers        : HString         read fTypeTiers        write fTypeTiers;
      property Tiers            : HString         read fTiers            write fTiers;
      property Auxiliaire       : HString         read fAuxiliaire       write fAuxiliaire;
      property LibelleTiers     : HString         read fLibelleTiers     write fLibelleTiers;
      property Part             : HString         read fPart             write fPart;
      property Numtel           : HString         read fNumtel           write fNumtel;
      property Affaire          : HString         read fAffaire          write fAffaire;
      property Mail             : hString         Read fMail             Write fMail;
      property QualifMail       : HString         Read fQualifMail       Write fQualifMail;
      property TobContact       : TOB             read fTobContact       write fTobContact;

      //Procédures et fonctions de traitements de contacts
      procedure AppelContactTiers(Requete: string);
      procedure AppelContact;
      procedure DecoupeContact(var Param1, Param2: HString);  //code contact
      procedure LectureContactAffaireInterv;
      procedure LectureContactPieceInterv;
      procedure LectureMailContact;
      procedure LectureMailTiers;
      Procedure LectureTobContact;
      procedure RechContactTiers;
      procedure RechercheContact;
      procedure VoirContacts;
    end;

implementation

//création de la classe et création
constructor TGestionContact.Create(sender : Tobject);
begin

  fForm := nil;

  fTypeContact := 'T';
  fcontact     := contact;
  fTypeTiers   := TypeTiers;
  fTiers       := Tiers;
  fAuxiliaire  := Auxiliaire;
  fLibelleTiers:= LibelleTiers;
  fNumtel      := numtel;
  fPart        := Part;
  fTobContact  := Tobcontact;

end;

procedure TGestionContact.RechercheContact;
begin

  if QualifMail = 'SST' then
    LectureContactPieceInterv
  else
    LectureContactAffaireInterv;

  if fContact = '0' then
  begin
    RechContactTiers;
    //FV1 - 14/02/2018 : FS#2901 - MOUTHON : <Erreur> ou violation d'accès sur le mail envoyé au ST ou COT
    if TobContact = nil then exit;
    //
    if TobContact.Detail.Count > 1 then
        AppelContact
    else
      LectureTobContact;
  end
  else
    LectureTobContact; 

end;


destructor TGestionContact.Destroy;
begin

  if Assigned(fTobcontact) then fTobContact.free;

inherited;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Franck Vautrain
Créé le ...... : 11/05/2005
Modifié le ... :   /  /
Description .. : Description .. : Fonction d'affichage des données du
Suite ........ : Contact principal dans la fiche Appel.
Mots clefs ... :
*****************************************************************}
Procedure TGestionContact.RechContactTiers;
var Req        : String;
    TobContact : Tob;
begin

  Req := 'SELECT * FROM CONTACT WHERE ';
  Req := Req + 'C_TYPECONTACT = "'  + FTypeContact  + '" AND ';
  Req := Req + 'C_AUXILIAIRE = "'   + fAuxiliaire   + '" AND ';
  Req := Req + 'C_TIERS = "'        + fTiers        + '" AND ';
  Req := Req + 'C_NATUREAUXI = "'   + FTypeTiers    + '" ';

  TobContact := Tob.Create('LesContacts',Nil, -1);
  TobContact.LoadDetailDBFromSQL('CONTACT',req,false);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Franck Vautrain
Créé le ...... : 11/05/2005
Modifié le ... :   /  /
Description .. : Appel de la recherche des contacts dans le cas ou le client
Suite ........ : dispose de + d'un contact
Mots clefs ... :
*****************************************************************}
procedure TGestionContact.AppelContact;
Var Req      : String;
    ZCONTACT : THEdit;
begin

  Req := '';
  Req := 'C_AUXILIAIRE="' + fAuxiliaire + '"';

  ZContact := THEdit.Create(Application);
  ZContact.Parent := fForm;
  ZCONTACT.Visible := False;
  Zcontact.datatype := 'BTCONTACT';
  Zcontact.plus := Req;

  LookupCombo(ZContact);

  FContact := ZContact.text;

  FreeAndNil(ZContact);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Franck Vautrain
Créé le ...... : 08/11/2011
Modifié le ... :   /  /
Description .. : Affichage du lookup du contact
Suite ........ : Contacts
Mots clefs ... :
*****************************************************************}
procedure TGestionContact.AppelContactTiers(Requete : string);
var ZCONTACT : THEdit;
Begin

  ZContact := THEdit.Create(Application);
  ZContact.Parent := fForm;
  ZCONTACT.Visible := False;
  Zcontact.datatype := 'BTCONTACTTIERS';
  Zcontact.plus := Requete;

  LookupCombo(ZContact);

  FContact := ZContact.text;

  FreeAndNil(ZContact);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Franck Vautrain
Créé le ...... : 08/11/2011
Modifié le ... :   /  /
Description .. : Affichage de la tablette de recherche
Suite ........ : Contacts  et récupération du code
Mots clefs ... :
*****************************************************************}
Procedure TGestionContact.VoirContacts;
Var Argument   : String;
		NumContact : String;
Begin

  Argument := '';
  NumContact := '';

  Argument := 'TYPE='  + fTypeContact + ';';
  Argument := Argument + 'TYPE2=' + fTypeTiers + ';';
  Argument := Argument + 'PART= ' + fPart + ';';
  Argument := Argument + 'TITRE=' + fLibelleTiers + ';';
  Argument := Argument + 'TIERS=' + fTiers + ';';
  Argument := Argument + 'ALLCONTACT';

  NumContact := AGLLanceFiche('YY','YYCONTACT','T;'+ fAuxiliaire,'', Argument);

  if NumContact = '' then
     NumContact := fContact
  Else
     fContact   := NumContact;

end;

Procedure TgestionContact.LectureContactAffaireInterv;
Var Req : string;
    QQ  : Tquery;
begin

  req := 'SELECT BAI_NUMEROCONTACT FROM AFFAIREINTERV WHERE ' ;
  if faffaire <> '' then req := req + 'BAI_AFFAIRE="' + fAffaire + '" AND ';
  if Tiers    <> '' then req := req + 'BAI_TIERSFOU="' + Tiers + '"';

  QQ := OpenSql (req,true,1,'',true);
  if not QQ.eof then
  begin
   	fContact := IntToStr(QQ.findField('BAI_NUMEROCONTACT').AsInteger);
  end;
  ferme (QQ);

end;

Procedure TgestionContact.LectureContactPieceInterv;
Var Req : string;
    QQ  : Tquery;
begin

  Req := 'SELECT BPI_NUMEROCONTACT FROM PIECEINTERV WHERE BPI_NATUREPIECEG="' + CleDocNatPiece + '" AND ';
  Req := Req + 'BPI_SOUCHE="' + CleDocSouche + '" AND ';
  Req := Req + 'BPI_NUMERO="' + CleDocNumPiece + '" AND ';
  Req := Req + 'BPI_TIERSFOU="' + Tiers + '"';

  QQ := OpenSql (req,true,1,'',true);
  if not QQ.eof then
  begin
   	fContact := IntToStr(QQ.findField('BPI_NUMEROCONTACT').AsInteger);
  end;
  ferme (QQ);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Franck Vautrain
Créé le ...... : 08/11/2011
Modifié le ... :   /  /
Description .. : Lecture de la zone contact à partir
Suite ........ : passe le compte auxiliaire et le code contact
Mots clefs ... :
*****************************************************************}
Procedure TGestionContact.LectureTobContact;
Var Req         : String;
Begin

  DecoupeContact(fContact, fNumtel);

  if fContact = '' then exit;

  Req := 'SELECT * FROM CONTACT WHERE C_AUXILIAIRE ="' + fAuxiliaire + '" AND C_NUMEROCONTACT=' + fContact;

  TobContact := Tob.Create('LesContacts',Nil, -1);
  TobContact.LoadDetailDBFromSQL('CONTACT',req,false);

end;

Procedure TGestionContact.LectureMailContact;
Var Req : String;
    QQ  : TQuery;
Begin

  Req := 'SELECT C_RVA FROM CONTACT WHERE ' +
         'C_TYPECONTACT="'  + FtypeContact + '" AND '+
         'C_AUXILIAIRE="'   + fAuxiliaire  + '" AND '+
         'C_NUMEROCONTACT=' + fContact;

  QQ := openSql (Req,true,1,'',true);

  if not QQ.eof then	fMail := QQ.findfield('C_RVA').AsString;

  ferme (QQ);

end;


Procedure TGestionContact.LectureMailTiers;
Var QQ  : TQuery;
    Req : string;
begin

    Req := 'SELECT T_EMAIL,T_RVA FROM TIERS WHERE T_NATUREAUXI="' + FtypeTiers + '" AND T_TIERS="' + fTiers +'"';
    QQ := openSql (Req,true,1,'',true);

    if not QQ.eof then
    begin
      fMail := QQ.findfield('T_EMAIL').AsString;
      if fmail = '' then fMail := QQ.findfield('T_RVA').AsString;
    end;

    ferme (QQ);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Franck Vautrain
Créé le ...... : 08/11/2011
Modifié le ... :   /  /
Description .. : Découpage de la zone contact lorsque celle ci
Suite ........ : passe le compte auxiliaire et le code contact
Mots clefs ... :
*****************************************************************}
procedure TGestionContact.DecoupeContact (Var Param1, Param2 : HString);
Var indice  : integer;
    Chaine  : string;
    param   : string;
Begin

	Chaine := fContact;
  Indice := 0;

	repeat
    param := READTOKENST (Chaine);
    if Param <> '' then
       begin
       if indice = 0 then
          Param1 := Param
       Else
          Param2:=param;
       inc(indice);
       end;
  until Param = '';

  fContact := Param1;

end;


end.
