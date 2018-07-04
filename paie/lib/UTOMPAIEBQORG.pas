{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/06/2002
Modifié le ... :   /  /
Description .. : Gestion des banques des organismes
Suite ........ : Table TBORG - Fiche TBORG
Mots clefs ... : PAIE; PGDUCS ; BQORG
*****************************************************************}
unit UTOMPAIEBQORG;

interface
uses
{$IFDEF EAGLCLIENT}
//unused      UtileAGL,
//unused      eFiche,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,
{$ENDIF}
      Classes,UTOM, UTOB, PgOutils,HCtrls,SysUtils,UtilPGI,Hmsgbox,Controls,PgOutils2;

Type
     TOM_PAIEBQORG = Class(TOM)
     procedure OnNewRecord;  override ;
     procedure OnUpdateRecord; override ;
     procedure OnDeleteRecord; override;
    END;
const
	// libellés des messages
	TexteMessage: array[1..15] of string 	= (
          {1}        'Vous devez renseigner la domiciliation.',
          {2}        'La domiciliation comporte des caractères interdits',
          {3}        'Vous devez renseigner le nom de la banque.',
          {4}        '4' ,
          {5}        '5',
          {6}        'Vous devez renseigner le code banque.',
          {7}        'Vous devez renseigner le code guichet.',
          {8}        'Vous devez renseigner le numéro de compte.',
          {9}        'Vous devez renseigner la clé du RIB.',
          {10}       '10',
          {11}       '11',
          {12}       '12',
          {13}       '13',
          {14}       '14',
          {15}       '15'
                     );

implementation
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/06/2002
Modifié le ... :   /  /    
Description .. : Création d'une banque - PREDEFINI toujours STD,
Suite ........ : NODOSSIER toujours 000000
Mots clefs ... : PAIE; PGDUCS ; BQORG
*****************************************************************}
procedure TOM_PAIEBQORG.OnNewRecord;
var
  QQ : TQuery;
begin
  inherited;
  SetField ('PBO_PREDEFINI','STD');
  SetField ('PBO_NODOSSIER','000000');

  QQ := OpenSql ('SELECT MAX(PBO_CODEBQORG) FROM PAIEBQORG', TRUE);
  if (NOT QQ.EOF) and (QQ <> NIL) and (QQ.Fields[0].AsString <> '') then // PORTAGECWAS
    SetField ('PBO_CODEBQORG',ColleZeroDevant(StrToInt(QQ.Fields[0].AsString) + 1,5))
  else
    SetField ('PBO_CODEBQORG','00001');
  Ferme (QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/06/2002
Modifié le ... :   /  /    
Description .. : Mise à jour enregistrement PAIEBQORG avec contrôle
Suite ........ : des champs obligatoires et contrôle de la clé RIB
Mots clefs ... : PAIE; PGDUCS; PAIEBQORG
*****************************************************************}
procedure TOM_PAIEBQORG.OnUpdateRecord;
Var RibCalcul : String ;
    Reponse : Word ;
    StDom   : String ;
begin
Inherited;
if (GetControlText('PBO_LIBELLEBQORG')='') then
  begin
    SetFocusControl('PBO_LIBELLEBQORG');
    LastError:=3;
    LastErrorMsg:=TexteMessage[LastError];
    exit;
  end;
StDom:=GetControlText('PBO_DOMBQORG') ;
if StDom='' then
  begin
  SetFocusControl('PBO_DOMBQORG');
  LastError:=1;
  LastErrorMsg:=TexteMessage[LastError];
  exit;
  end;
if ExisteCarInter(StDom) then
  begin
  SetFocusControl('PBO_DOMBQORG');
  LastError:=2;
  LastErrorMsg:=TexteMessage[LastError];
  exit;
  end;

if GetControlText('PBO_ETABBQORG')='' then
  BEGIN
    SetFocusControl('PBO_ETABBQORG');
    LastError:=6;
    LastErrorMsg:=TexteMessage[LastError];
    exit;
  END ;
if GetControlText('PBO_GUICHBQORG')='' then
  BEGIN
    SetFocusControl('PBO_GUICHBQORG');
    LastError:=7;
    LastErrorMsg:=TexteMessage[LastError];
    exit;
  END ;
if GetControlText('PBO_NUMCPTEBQORG')='' then
  BEGIN
        SetFocusControl('PBO_NUMCPTEBQORG');
        LastError:=8;
        LastErrorMsg:=TexteMessage[LastError];
        exit;
      END ;
   if GetControlText('PBO_CLEBQORG')='' then
      BEGIN
        SetFocusControl('PBO_CLEBQORG');
        LastError:=9;
        LastErrorMsg:=TexteMessage[LastError];
        exit;
      END ;
// Vérification de la clé RIB
RibCalcul:=VerifRib(GetControlText('PBO_ETABBQORG'),
                    GetControlText('PBO_GUICHBQORG'),
                    GetControlText('PBO_NUMCPTEBQORG')) ;
If RibCalcul<>GetControlText('PBO_CLEBQORG') Then
   begin
     Reponse:= HShowMessage('1;Clé erronée;Confirmez-vous la clé du RIB ?;Q;YN;N;N;','','') ;
     if Reponse<>mrYes Then
       BEGIN
         SetFocusControl('PBO_CLEBQORG');
         SetField('PBO_CLEBQORG',RibCalcul) ;
         Exit ;
       END ;
   end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/06/2002
Modifié le ... :   /  /    
Description .. : Suppression d'un enregistrement TBORG
Suite ........ : Contrôle de la non utilisation de la banque dans
Suite ........ : la table ORGANISMEPAIE
Mots clefs ... : PAIE; PGDUCS; TBORG
*****************************************************************}
procedure TOM_PAIEBQORG.OnDeleteRecord;
var ExisteCod : Boolean;
    NomChamp:array[1..1] of Hstring;
    ValChamp:array[1..1] of variant;
begin
  inherited;
NomChamp[1]:='POG_RIBDUCSEDI';
ValChamp[1]:=GetField('PBO_CODEBQORG');
ExisteCod:=RechEnrAssocier('ORGANISMEPAIE',NomChamp,ValChamp);
if ExisteCod=TRUE then
   begin
     LastError:=1;
     LastErrorMsg:='Attention! Cette banque est utilisée par un organisme,'+
                   '#13#10Vous ne pouvez la supprimer!';
     ExisteCod:=False;
   end;
end;


Initialization

registerclasses([TOM_PAIEBQORG]) ;
end.
