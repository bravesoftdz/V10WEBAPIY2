{***********UNITE*************************************************
Auteur  ...... : Paul Chapuis
Créé le ...... : 03/04/2001
Modifié le ... : 03/04/2001
Description .. : Fonctions génériques pour liens OLE
Mots clefs ... : OLE;DICO;
*****************************************************************}
unit UtilOle;

interface

uses HCtrls,OleDicoPGI,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     Utob;

Procedure InitTobDico( Nom, Prefixe : string; TheTobDico : Tob);

implementation

{initialise une tob avec la liste des champs utilisables dans
 les documents pour affichage dans le Dico Ole
}
Procedure InitTobDico( Nom, Prefixe : string; TheTobDico : Tob);
var T1 : Tob;
    Q : TQuery;
    i : integer;
    st : string;
begin
Q:=OpenSql('SELECT * FROM DECHAMPS WHERE (DH_PREFIXE="'+Prefixe+'") AND (DH_CONTROLE like "%D%")',True);
T1 := TOB.Create(Nom,theTobDico,-1);
T1.LoadDetailDB('DECHAMPS','','DH_NUMCHAMP',Q,false);
Ferme(Q);
{ prendre le libellé en mémoire pour les champs libres }
For i:=T1.detail.count-1 downto 0 do
    begin
    st:=ChampToLibelle(T1.detail[i].getValue('DH_NOMCHAMP'));
    if (Length(st)> 0) and (copy(st,1,2)='.-') then T1.detail[i].free
                 else T1.detail[i].putValue('DH_LIBELLE',st);
    end;
end;


end.
