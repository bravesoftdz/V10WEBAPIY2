{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 10/04/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : YYIMPORTOBJET ()
Mots clefs ... : TOF;YYIMPORTOBJET
*****************************************************************}
Unit UTOFYYIMPORTOBJET ;

Interface

Uses StdCtrls, Controls, Classes, db, forms, sysutils,
    {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
     ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF , M3FP, Utob, ubob;

Type
  TOF_YYIMPORTOBJET = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Implementation

procedure BOBJournalEvenement(err : integer; stErr: string; FileName : string  ) ;
var TOBJnal : TOB ;
    NumEvt,Version : integer ;
    Nom,Libelle : string ;
    QQ : TQuery ;
    ListRecap : Tstringlist;
begin
NumEvt:=0 ;
ListRecap:=TStringList.create ;
AglReadInfoBob ( FileName,Version,Nom,Libelle ) ;
TOBJnal:=TOB.Create('JNALEVENT', Nil, -1) ;
TOBJnal.PutValue('GEV_TYPEEVENT', 'BOB');
TOBJnal.PutValue('GEV_LIBELLE', 'Insertion BOB '+NOM);
TOBJnal.PutValue('GEV_DATEEVENT', V_PGI.DateEntree);
TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
QQ:=OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True) ;
if Not QQ.EOF then NumEvt:=QQ.Fields[0].AsInteger ;
Ferme(QQ) ;
Inc(NumEvt) ;
TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);
ListRecap.Add ('Insertion BOB ');
ListRecap.Add ('NOM : '+Nom);
ListRecap.Add ('VERSION : '+intToStr(Version));
ListRecap.Add ('Fichier : '+FileName);
ListRecap.Add ('Version base : '+IntToStr(V_PGI.NumVersionSoc)) ;

ListRecap.Add ('');
if err<>0 then
    begin
    TOBJnal.PutValue ('GEV_ETATEVENT', 'ERR');
    ListRecap.Add ('Erreur lors de l''insertion de la BOB :');
    ListRecap.Add ('     '+stErr);
    end else
    begin
    TOBJnal.PutValue('GEV_ETATEVENT', 'OK');
    ListRecap.Add('Traitement correct');
    end;
TOBJnal.PutValue('GEV_BLOCNOTE', ListRecap.Text);
TOBJnal.InsertDB(Nil) ;
TOBJnal.Free ;
ListRecap.free ;
end;


procedure TOF_YYIMPORTOBJET.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_YYIMPORTOBJET.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_YYIMPORTOBJET.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_YYIMPORTOBJET.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_YYIMPORTOBJET.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_YYIMPORTOBJET.OnClose ;
begin
  Inherited ;
end ;

{
Suite ........ : Les valeurs en retour :
Suite ........ : 0  : OK
Suite ........ : 1  : Intégration déjà effectuée
Suite ........ : -1 : Erreur d'écriture dans la table YMYBOBS
Suite ........ : -2 : Erreur d'intégration dans la fonction AglImportBob
Suite ........ : -3 : Erreur de lecture du fichier BOB.
Suite ........ : -4 : Erreur inconnue.
}

function AGLYYImportObjet( parms: array of variant; nb: integer ) : variant;
var FileName : String ;
    stErr : string;
BEGIN
Result:=0;
FileName:=string(parms[0]) ;
if FileExists(FileName) then
   begin
   result:=AGLIntegreBob (FileName,FALSE,TRUE );       //AglImportBob
   case result of
        0    : stErr:='';
        1    : stErr:='Intégration déjà effectuée';
        -1   : stErr:='Erreur d''écriture dans la table YMYBOBS ';
        -2   : stErr:='Erreur d''intégration dans la fonction AglImportBob';
        -3   : stErr:='Erreur de lecture du fichier BOB.';
        -4   : stErr:='Erreur inconnue.';
        -100 : stErr:='BOB sans control.';
   else
     stErr:='Erreur inconnue.';
   end;
   BOBJournalEvenement(result,stErr, FileName);
   if result<>0 then PGIError( stErr ,'Erreur pendant l''import')
             else PGIInfo( 'Import réussi','Import d''objets');
   end
   else exit;
END;

Initialization
  registerclasses ( [ TOF_YYIMPORTOBJET ] ) ;
  RegisterAglFunc( 'YYIMPORTOBJET', False,1,AGLYYImportObjet) ;
end.
