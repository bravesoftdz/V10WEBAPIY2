{***********UNITE*************************************************
Auteur  ...... : Heng Lek CHHOEU
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. : Routine pour les extionsions d'écritures d'engagements.
Suite ........ : La table d'extension est CENGAGEMENT
Mots clefs ... :
*****************************************************************}
unit uLibEngage;

interface
uses
  {$IFNDEF EAGLCLIENT}
    db,
   {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}

  {$IFDEF VER150}
  Variants,
  {$ENDIF}

  UTOB,
  hCtrls,
  SysUtils,
  utilPGI;

procedure CPutTOBEngage ( TheTOB : TOB ;  Valeur : TOB ) ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Heng Lek CHHOEU
Créé le ...... : 27/03/2006
Modifié le ... : 27/03/2006
Description .. : Rend le TOB d'engagement pointé par TheTOB
Suite ........ : TheTOB: Tob contient le lien vers le tob engagement.
Mots clefs ... :
*****************************************************************}
function CGetTOBEngage ( TheTOB : TOB ) : TOB ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Heng Lek CHHOEU
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. : Mettre une valeur dans le tob engagagement:
Suite ........ : TheTOB: Tob contient le lien vers le tob engagement
Suite ........ : Nom: nom du champ d'engagement
Suite ........ : Valeur: valeur d'engagement à maj ou à créer.
Mots clefs ... :
*****************************************************************}
procedure CPutValueTOBEngage ( TheTOB : TOB ;  Nom : string ; Valeur : Variant ) ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Heng Lek CHHOEU
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. : Rend la valeur d'engagement
Suite ........ : TheTOB: Tob contient le lien vers l'engagement
Suite ........ : Nom: champ d'engagement qu'on veut obtenir la valeur
Mots clefs ... :
*****************************************************************}
function CGetValueTOBEngage ( TheTOB : TOB ; Nom : string ) : Variant ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Heng Lek CHHOEU
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. : Rend un TOB d'une ligne extension engagement
Suite ........ : TheTOB:Tob d'écriture
Suite ........ : TheParent:parent de tob engagement lequel sera créé.
Mots clefs ... :
*****************************************************************}
function CSelectDBTOBEngage( TheTOB : TOB ; TheParent : TOB ) : TOB ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Heng Lek CHHOEU
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. : Supprimer une ligne extension engagement
Suite ........ : vQ:Querry contient la clé de l'engagement(écriture ou
Suite ........ : engagement)
Mots clefs ... :
*****************************************************************}
procedure CDeleteDBTOBEngage( vQ : TQuery ) ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Heng Lek CHHOEU
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. : Supprimer une pièce engagement
Suite ........ : TheTOB Tob contient la clé d'une pièce (écriture ou
Suite ........ : engagement) et positionne sur la première ligne.
Mots clefs ... :
*****************************************************************}
procedure CSupprimerEcrEngage( TheTOB : TOB ; vDossier : String = '' ) ;

function CCreateDBTOBEngage( TheTOB : TOB ; TheParent : TOB ; Q : TQuery ) : TOB ;

procedure CMAJTOBEngage( TheTOB : TOB ) ;

function CCreateTOBEngage( TheTOB : TOB ; TheParent : TOB ) : TOB ;

procedure CFreeTOBEngage ( TheTob : TOB ) ;

implementation

procedure CPutTOBEngage ( TheTOB : TOB ;  Valeur : TOB ) ;
begin
 TheTob.AddChampSupValeur('PENGAGE',longInt(Valeur),false) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Heng Lek CHHOEU
Créé le ...... : 27/03/2006
Modifié le ... : 27/03/2006
Description .. : Rend le TOB d'engagement pointé par TheTOB
Suite ........ : TheTOB: Tob contient le lien vers le tob engagement.
Mots clefs ... :
*****************************************************************}
function CGetTOBEngage ( TheTOB : TOB ) : TOB ;
var
 lP : Variant ;
begin
  result := nil ;
  if TheTOB.FieldExists('PENGAGE') then
  begin
    lP     := TheTOB.GetValue('PENGAGE') ;
    if ( VarAsType(lP, VarString) <> #0 ) and ( VarAsType(lP, VarInteger) <> - 1 ) then
    result := TOB(LongInt(lP)) ;
  end ; // if
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Heng Lek CHHOEU
Créé le ...... : 27/03/2006
Modifié le ... :   /  /    
Description .. : Mettre une valeur dans le tob engagagement:
Suite ........ : TheTOB: Tob contient le lien vers le tob engagement
Suite ........ : Nom: nom du champ d'engagement
Suite ........ : Valeur: valeur d'engagement à maj ou à créer.
Mots clefs ... :
*****************************************************************}
procedure CPutValueTOBEngage ( TheTOB : TOB ;  Nom : string ; Valeur : Variant ) ;
var
 lTOB : TOB ;
begin
 lTOB := CGetTOBEngage(TheTOB) ;
 if lTOB <> nil then
  lTOB.PutValue(Nom,Valeur) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Heng Lek CHHOEU
Créé le ...... : 27/03/2006
Modifié le ... :   /  /    
Description .. : Rend la valeur d'engagement
Suite ........ : TheTOB: Tob contient le lien vers l'engagement
Suite ........ : Nom: champ d'engagement qu'on veut obtenir la valeur
Mots clefs ... : 
*****************************************************************}
function CGetValueTOBEngage ( TheTOB : TOB ; Nom : string ) : Variant ;
var
 lTOB : TOB ;
begin
 result := #0 ;
 lTOB := CGetTOBEngage(TheTOB) ;
 if lTOB <> nil then
  result := lTOB.GetValue(Nom) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Heng Lek CHHOEU
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. : Rend un TOB d'une ligne extension engagement
Suite ........ : TheTOB:Tob d'écriture
Suite ........ : TheParent:parent de tob engagement lequel sera créé.
Mots clefs ... : 
*****************************************************************}
function CSelectDBTOBEngage( TheTOB : TOB ; TheParent : TOB ) : TOB ;
var
 lSt   : string ;
 lQ    : TQuery ;
begin                 {ATTENTION NUMECHE}
  lSt := 'SELECT * FROM CENGAGEMENT WHERE CEN_JOURNAL="' + TheTOB.GetValue('E_JOURNAL')      + '" ' +
        'AND CEN_EXERCICE="'            + TheTOB.GetValue('E_EXERCICE')                  + '" ' +
        'AND CEN_DATECOMPTABLE="'       + usDateTime(TheTOB.GetValue('E_DATECOMPTABLE')) + '" ' +
        'AND CEN_NUMEROPIECE='          + intToStr(TheTOB.GetValue('E_NUMEROPIECE'))     + ' '  +
        'AND CEN_NUMLIGNE='             + intToStr(TheTOB.GetValue('E_NUMLIGNE'))        + ' '  +
        'AND CEN_NUMECHE='              + intToStr(TheTOB.GetValue('E_NUMECHE'))         + ' '  +
        'AND CEN_QUALIFPIECE="'         + TheTOB.GetValue('E_QUALIFPIECE')               + '" ' ;
  lQ     := OpenSql( lSt , true ) ;
  result := TOB.Create('CENGAGEMENT',TheParent,-1) ;
  result.SelectDB('',lQ) ;
  CPutTOBEngage(TheTOB,result) ;
  Ferme(lQ) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Heng Lek CHHOEU
Créé le ...... : 27/03/2006
Modifié le ... :   /  /    
Description .. : Supprimer une ligne extension engagement
Suite ........ : vQ:Querry contient la clé de l'engagement(écriture ou 
Suite ........ : engagement)
Mots clefs ... : 
*****************************************************************}
procedure CDeleteDBTOBEngage( vQ : TQuery ) ;
begin {ATTENTION NUMECHE}
  ExecuteSQL( 'DELETE FROM CENGAGEMENT WHERE CEN_JOURNAL="' + vQ.FindField('E_JOURNAL').asString           + '" ' +
              'AND CEN_EXERCICE="'            + vQ.FindField('E_EXERCICE').asString                     + '" ' +
              'AND CEN_DATECOMPTABLE="'       + usDateTime(vQ.FindField('E_DATECOMPTABLE').asDateTime)  + '" ' +
              'AND CEN_NUMEROPIECE='          + vQ.FindField('E_NUMEROPIECE').asString                  + ' '  +
              'AND CEN_NUMLIGNE='             + vQ.FindField('E_NUMLIGNE').asString                     + ' '  +
              'AND CEN_NUMECHE='              + vQ.FindField('E_NUMECHE').asString                     + ' '  +
              'AND CEN_QUALIFPIECE="'         + vQ.FindField('E_QUALIFPIECE').asString                  + '" ' )
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Heng Lek CHHOEU
Créé le ...... : 27/03/2006
Modifié le ... :   /  /    
Description .. : Supprimer une pièce engagement
Suite ........ : TheTOB Tob contient la clé d'une pièce (écriture ou 
Suite ........ : engagement) et positionne sur la première ligne.
Mots clefs ... : 
*****************************************************************}
procedure CSupprimerEcrEngage( TheTOB : TOB ; vDossier : String = '' ) ;
begin
 if TheTOB.GetValue('E_NUMLIGNE') = 1 then
   begin
   ExecuteSQL( 'DELETE FROM ' + GetTableDossier( vDossier, 'CENGAGEMENT' ) +
               ' WHERE CEN_JOURNAL="'        + THeTOB.GetValue('E_JOURNAL')                   + '" ' +
                  'AND CEN_EXERCICE="'       + TheTOB.GetValue('E_EXERCICE')                  + '" ' +
                  'AND CEN_DATECOMPTABLE="'  + usDateTime(TheTOB.GetValue('E_DATECOMPTABLE')) + '" ' +
                  'AND CEN_NUMEROPIECE='     + intToStr(TheTOB.GetValue('E_NUMEROPIECE'))     + ' '  +
                  'AND CEN_QUALIFPIECE="'    + TheTOB.GetValue('E_QUALIFPIECE')               + '" ' )
   end ;
end ;

function CCreateDBTOBEngage( TheTOB : TOB ; TheParent : TOB ; Q : TQuery ) : TOB ;
begin
 result:= TOB.CreateDB('CENGAGEMENT',TheParent,-1,Q) ;
 CPutTOBEngage(TheTOB,result) ;
end ;

procedure CMAJTOBEngage( TheTOB : TOB ) ;
var
 lTOB : TOB ;
begin {ATTENTION NUMECHE}
 lTOB :=  CGetTOBEngage(TheTOB) ;
 if lTOB = nil then exit ;
 lTOB.PutValue('CEN_EXERCICE'      , TheTOB.GetValue('E_EXERCICE') ) ;
 lTOB.PutValue('CEN_JOURNAL'       , TheTOB.GetValue('E_JOURNAL') ) ;
 lTOB.PutValue('CEN_DATECOMPTABLE' , TheTOB.GetValue('E_DATECOMPTABLE') ) ;
 lTOB.PutValue('CEN_NUMEROPIECE'   , TheTOB.GetValue('E_NUMEROPIECE') ) ;
 lTOB.PutValue('CEN_NUMLIGNE'      , TheTOB.GetValue('E_NUMLIGNE') ) ;
 lTOB.PutValue('CEN_NUMECHE'       , TheTOB.GetValue('E_NUMECHE') ) ;
 lTOB.PutValue('CEN_QUALIFPIECE'   , TheTOB.GetValue('E_QUALIFPIECE') ) ;


end ;

function CCreateTOBEngage( TheTOB : TOB ; TheParent : TOB ) : TOB ;
begin
 result:= TOB.Create('CENGAGEMENT',TheParent,-1) ;
 CPutTOBEngage(TheTOB,result) ;
 CMAJTOBEngage(TheTOB) ;
end ;

procedure CFreeTOBEngage ( TheTob : TOB ) ;
var
 lTOB : TOB ;
begin
 lTOB := CGetTOBEngage(TheTob) ;
 if Assigned(lTOB) then FreeAndNil(lTOB) ;
 if TheTOB.FieldExists('PENGAGE') then
   TheTob.PutValue('PENGAGE',-1) ;
end ;


end.
