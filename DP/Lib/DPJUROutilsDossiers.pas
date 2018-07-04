{***********UNITE*************************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 29/01/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
unit DPJUROutilsDossiers;
////////////////////////////////////////////////////////////////////////////////
interface
////////////////////////////////////////////////////////////////////////////////
uses
   UTOB;
////////////////////////////////////////////////////////////////////////////////
function DPJURCreationDossierJuridique(sGuidPerDos_p : string; sTypeDos_p : string = 'STE'; iNoOrdre_p : integer = 1) : boolean;
////////////////////////////////////////////////////////////////////////////////
implementation
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 29/01/2008
Modifié le ... :   /  /    
Description .. : Création d'un dossier juridique HORS TOM avec les valeurs 
Suite ........ : essentielles par défaut
Mots clefs ... : 
*****************************************************************}
function DPJURCreationDossierJuridique(sGuidPerDos_p : string; sTypeDos_p : string = 'STE'; iNoOrdre_p : integer = 1) : boolean;
var
   OBAnnuaire_l, OBJuridique_l : TOB;
begin
   //--- Création de la tob Annuaire
   OBAnnuaire_l := TOB.Create('ANNUAIRE', Nil, -1) ;
   OBAnnuaire_l.SelectDB('"' + sGuidPerDos_p + '"', Nil, TRUE) ;

   //--- Création de la tob juridique
   OBJuridique_l := TOB.Create('JURIDIQUE', Nil, -1);
   OBJuridique_l.InitValeurs;

   OBJuridique_l.PutValue ('JUR_GUIDPERDOS', OBAnnuaire_l.GetValue ('ANN_GUIDPER'));
   OBJuridique_l.PutValue ('JUR_TYPEDOS', 'STE');
   OBJuridique_l.PutValue ('JUR_NOORDRE', 1);
   OBJuridique_l.PutValue ('JUR_CODEDOS', '&#@');
   OBJuridique_l.PutValue ('JUR_NOMDOS', OBAnnuaire_l.GetValue ('ANN_NOMPER'));
   OBJuridique_l.PutValue ('JUR_DOSLIBELLE',OBAnnuaire_l.GetValue ('ANN_NOM1'));

   OBJuridique_l.PutValue ('JUR_FORME',OBAnnuaire_l.GetValue ('ANN_FORME'));
   OBJuridique_l.PutValue ('JUR_CAPDEV',OBAnnuaire_l.GetValue ('ANN_CAPDEV'));
   OBJuridique_l.PutValue ('JUR_CAPITAL',OBAnnuaire_l.GetValue ('ANN_CAPITAL'));
   OBJuridique_l.PutValue ('JUR_NBDROITSVOTE',OBAnnuaire_l.GetValue ('ANN_CAPNBTITRE'));
   OBJuridique_l.PutValue ('JUR_NBTITRESCLOT',OBAnnuaire_l.GetValue ('ANN_CAPNBTITRE'));
   OBJuridique_l.PutValue ('JUR_MOISCLOTURE',OBAnnuaire_l.GetValue ('ANN_MOISCLOTURE'));
   OBJuridique_l.PutValue ('JUR_VALNOMINCLOT',OBAnnuaire_l.GetValue ('ANN_CAPVN'));

   OBJuridique_l.PutValue ('JUR_DATEEXP',OBAnnuaire_l.GetValue ('ANN_CAPVN'));
   OBJuridique_l.PutValue ('JUR_RESP', '');

   result := OBJuridique_l.InsertDB(nil, FALSE);
   OBJuridique_l.Free;
   OBAnnuaire_l.Free;
end;

end.
