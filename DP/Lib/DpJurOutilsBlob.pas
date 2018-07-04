{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 28/04/2003
Modifié le ... :
Description .. : Unité commune entre le DP et JURI pour la gestion des blobs
Mots clefs ... : DP;JURI
*****************************************************************}
unit DpJurOutilsBlob;

interface

uses
   {$IFDEF EAGLCLIENT}
   {$ELSE}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
   {$ENDIF}
   utob, comctrls, Classes, hCtrls, SysUtils, hRichOLE, rtfcounter, HEnt1;


/////////// ENTETES DE FONCTIONS ////////////

procedure BlobGetCode(sControlName_p : string;
                      var sEmploiBlob_p, sPrefixe_p, sRangBlob_p, sLibelle_p : string);
function  BlobChange(OBBlob_p : TOB; sPrefixe_p, sGuidPer_p, sRangBlob_p : string; reChamp_p : TRichEdit ) : boolean;
procedure BlobEnreg (OBBlob_p : TOB; sGuidPer_p : string; reChamp_p : TRichEdit );
procedure BlobsEnreg(OBBlob_p : TOB );
procedure BlobsSupprime(sPrefixe_p, sGuidPer_p, sRangBlob_p : string);

function  BlobChampCreation(sPrefixe_p, sGuidPer_p, sRangBlob_p, sEmploiBlob_p, sLibelle_p : string;
                            reChamp_p : TRichEdit ) : TOB;
function BlobCreation(sPrefixe_p, sGuidPer_p, sRangBlob_p, sEmploiBlob_p, sLibelle_p, sTexte_p : string) : TOB;

function  ChangeBlobEnTOB(Leblob: string): TOB;


/////////// IMPLEMENTATION ////////////

implementation


/////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure BlobGetCode(sControlName_p : string;
                      var sEmploiBlob_p, sPrefixe_p, sRangBlob_p, sLibelle_p : string);
begin
//Attention : à synchroniser avec TOM_Annuaire.CreationChampBlob
   if sControlName_p = 'OLE_ACTIVITE' then
   begin
      sPrefixe_p := 'ANN';
      sEmploiBlob_p := 'ACT';
      sRangBlob_p := '1';
      sLibelle_p := 'Activité';
   end
   else if sControlName_p = 'OLE_BLOCNOTES' then
   begin
      sPrefixe_p := 'JUR';
      sEmploiBlob_p := 'BLO';
      sRangBlob_p := '2';
      sLibelle_p := 'Bloc Notes';
   end
   else if sControlName_p = 'OLE_REGMATTXT' then
   begin
      sPrefixe_p := 'JUR';
      sEmploiBlob_p := 'REG';
      sRangBlob_p := '3';
      sLibelle_p := 'Régime matrimonial';
   end
   else if sControlName_p = 'OLE_OBJETSOC' then
   begin
      sPrefixe_p := 'JUR';
      sEmploiBlob_p := 'OBJ';
      sRangBlob_p := '4';
      sLibelle_p := 'Objet social';
   end
   else if sControlName_p = 'OLE_ACTIVITE2' then //LM20070315
   begin
      sPrefixe_p := 'ANN';
      sEmploiBlob_p := 'OBJ';
      sRangBlob_p := '5';
      sLibelle_p := 'Activité secondaire';
   end
   else if sControlName_p = 'OLE_CONVENTION' then
   begin
      sPrefixe_p := 'ANL';
      sEmploiBlob_p := 'CNV';
      sRangBlob_p := '1';
      sLibelle_p := 'Conventions';
   end
   else if sControlName_p = 'OLE_CONVLIBRE' then
   begin
      sPrefixe_p := 'ANL';
      sEmploiBlob_p := 'CVL';
      sRangBlob_p := '2';
      sLibelle_p := 'Conventions libres';
   end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/03/2003
Modifié le ... :   /  /    
Description .. : test sur nRangBlob_p='0' pour fiche en création
Mots clefs ... :
*****************************************************************}
function BlobChampCreation(sPrefixe_p, sGuidPer_p, sRangBlob_p, sEmploiBlob_p, sLibelle_p : string;
                           reChamp_p : TRichEdit) : TOB;
var
   OBBlob_l : TOB;
   sTexte_l : string;
begin
   OBBlob_l := TOB.Create('LIENSOLE', nil, -1);
   OBBlob_l.AddChampSupValeur('CHARGE', '-');
   if (sGuidPer_p <> '' ) and
      ( OBBlob_l.SelectDB('"' + sPrefixe_p + '";"' + sGuidPer_p + '";"' + sRangBlob_p + '"', nil )) then
   begin
      // Chargement
      sTexte_l := OBBlob_l.GetValue('LO_OBJET');
      StringToRich(reChamp_p, sTexte_l);
   end
   else
   begin
      // Creation
      OBBlob_l.PutValue('LO_TABLEBLOB', sPrefixe_p);
      OBBlob_l.PutValue('LO_IDENTIFIANT', sGuidPer_p);
      OBBlob_l.PutValue('LO_RANGBLOB', sRangBlob_p);
      OBBlob_l.PutValue('LO_QUALIFIANTBLOB', 'DAT');
      OBBlob_l.PutValue('LO_EMPLOIBLOB', sEmploiBlob_p);
      OBBlob_l.PutValue('LO_LIBELLE', sLibelle_p);
      OBBlob_l.PutValue('LO_PRIVE','-');
      OBBlob_l.PutValue('LO_OBJET', '');
      StringToRich(reChamp_p, '');      
   end;
   OBBlob_l.PutValue('CHARGE', 'X');
   result := OBBlob_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : GHA
Créé le ...... : 11/02/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function BlobCreation(sPrefixe_p, sGuidPer_p, sRangBlob_p, sEmploiBlob_p, sLibelle_p, sTexte_p : string) : TOB;
var
   OBBlob_l : TOB;
begin
   OBBlob_l := TOB.Create('LIENSOLE', nil, -1);
   OBBlob_l.PutValue('LO_TABLEBLOB', sPrefixe_p);
   OBBlob_l.PutValue('LO_IDENTIFIANT', sGuidPer_p);
   OBBlob_l.PutValue('LO_RANGBLOB', sRangBlob_p);
   OBBlob_l.PutValue('LO_QUALIFIANTBLOB', 'DAT');
   OBBlob_l.PutValue('LO_EMPLOIBLOB', sEmploiBlob_p);
   OBBlob_l.PutValue('LO_LIBELLE', sLibelle_p);
   OBBlob_l.PutValue('LO_PRIVE','-');
   OBBlob_l.PutValue('LO_OBJET', sTexte_p);
   result := OBBlob_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 07/11/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
function BlobChange(OBBlob_p : TOB; sPrefixe_p, sGuidPer_p, sRangBlob_p : string; reChamp_p : TRichEdit) : boolean;
var
  OBBlob_l : TOB;
  sOldValeur_l, sNewValeur_l : string;
  sOldValeurRTF_l, sNewValeurRTF_l : string;
begin
   result := false;
   OBBlob_l := OBBlob_p.FindFirst(['LO_TABLEBLOB', 'LO_IDENTIFIANT', 'LO_RANGBLOB'], [sPrefixe_p, sGuidPer_p, sRangBlob_p], true);
   if OBBlob_l = nil then
      exit;
   // Si pas en cours de chargement
   if OBBlob_l.GetValue('CHARGE') = '-' then
      exit;

   sNewValeur_l := RichToString( reChamp_p );
   sOldValeur_l := OBBlob_l.GetValue('LO_OBJET');
   sNewValeurRTF_l := GetRTFStringText(sNewValeur_l);
   sOldValeurRTF_l := GetRTFStringText(sOldValeur_l);

   if (sOldValeurRTF_l = '') and (sNewValeurRTF_l = #$D#$A) then
      exit;

   if (sNewValeurRTF_l <> sOldValeurRTF_l) or (sNewValeur_l <> sOldValeur_l) then
   begin
      OBBlob_l.PutValue('LO_OBJET', sNewValeur_l);
      result := true;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/03/2003
Modifié le ... :   /  /    
Description .. : on affecte le guidper aux enregistrement des blobs
Mots clefs ... :
*****************************************************************}
procedure BlobEnreg(OBBlob_p : TOB; sGuidPer_p : string; reChamp_p : TRichEdit);
var
   sNewValeur_l : string;
begin
   sNewValeur_l := RichToString( reChamp_p );
   OBBlob_p.PutValue('LO_DATEBLOB', Date);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : GHA
Créé le ...... : 11/02/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure BlobsEnreg(OBBlob_p : TOB);
var
   nOBInd_l : integer;
begin
   OBBlob_p.DelChampSup('CHARGE', true);
   for nOBInd_l := 0 to OBBlob_p.Detail.Count - 1 do
   begin
      OBBlob_p.Detail[nOBInd_l].PutValue('LO_DATEBLOB', Date);
   end;
   OBBlob_p.InsertOrUpdateDB(false);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 14/11/2003
Modifié le ... :   /  /    
Description .. : Suprime les blobs de la table LIENSOLE correspondant aux 
Suite ........ : paramètres
Mots clefs ... : 
*****************************************************************}
procedure BlobsSupprime(sPrefixe_p, sGuidPer_p, sRangBlob_p : string);
var
   sWhere_l : string;
begin
   if sGuidPer_p <> '' then
      sWhere_l := 'LO_IDENTIFIANT = "' + sGuidPer_p + '"';

   if sPrefixe_p <> '' then
   begin
      if sWhere_l <> '' then
         sWhere_l := sWhere_l + ' AND ';
      sWhere_l := sWhere_l + 'LO_TABLEBLOB = "' + sPrefixe_p + '"';
   end;

   if sRangBlob_p <> '' then
   begin
      if sWhere_l <> '' then
         sWhere_l := sWhere_l + ' AND ';
      sWhere_l := sWhere_l + 'LO_RANGBLOB = ' + sRangBlob_p;
   end;

   if sWhere_l = '' then exit;

   ExecuteSQL('DELETE FROM LIENSOLE WHERE ' + sWhere_l);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 30/01/2008
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function ChangeBlobEnTOB(Leblob: string): TOB;
var
  aStream: TMemoryStream;
begin
  Result := nil;
  if Leblob <> '' then
  begin
    aStream := TMemoryStream.Create;
    try
      aStream.Write(Leblob[1], Length(Leblob));
      aStream.Seek(0, 0);
      DecompressBinaryObjectStream(aStream);
      if assigned(aStream) then
      begin
        with TTOBUser.create do
        try
          Result := LoadTOBFromBinStream(aStream);
          Result.ReaffecteNomTable('_TOB', True);
        finally
          Free;
        end;
      end;
    finally
      Freeandnil(aStream);
    end;
  end;
end;

/////////////////////////////////////////////////////////////////
end.
