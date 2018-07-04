{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 04/09/2002
Modifié le ... : 28/02/2003 : Compatibilité CWAS
Description .. : Classe pour la fusion de documents
Mots clefs ... :
*****************************************************************}

unit CLASSHistorisation;

interface

uses
   UTOB;

type
   THistorisation = class
      public
      private
         OBElement_c   : TOB;
         OBHistoPre_c  : TOB;
         OBHistoNew_c  : TOB;
         
         sTable_c      : string;
         sTableHisto_c : string;
         sPref_c       : string;
         sPrefHisto_c  : string;
         sArguments_c  : string;
         sCle_c        : string;
         sOrdre_c      : string;
         sChampMax_c   : string;
         sTypeModif_c  : string;

         procedure Init(OBElement_p : TOB; sTable_p, sTableHisto_p,
                        sCle_p, sOrdre_p, sChampMax_p,
                        sTypeModif_p, sArguments_p : string);

         // $$$JP 05/12/05 - function  InitWhere : string;
         function  Precedents : boolean;
         procedure MiseAJour;
         procedure Complement;
         function  AHistoriser : boolean;
         procedure Enregistre;
   end;
/////////// ENTETES DE FONCTION ET PROCEDURES ////////////

function  Historisation(OBElement_p : TOB; sTable_p, sTableHisto_p,
                        sCle_p, sOrdre_p, sChampMax_p : string;
                        sTypeModif_p : string; sArguments_p : string = '') : boolean;

//////////////////////////////////////////////////////////////////
implementation
//////////////////////////////////////////////////////////////////
uses
   hctrls,

   {$IFDEF VER150}
   Variants,
   {$ENDIF}

   SysUtils, HEnt1, UJUROUTILS;

   
//////////////////////////////////////////////////////////////////
// Procédures globales d'accès à la classe                      //
//////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 07/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function  Historisation(OBElement_p : TOB;
                        sTable_p, sTableHisto_p,
                        sCle_p, sOrdre_p, sChampMax_p : string;
                        sTypeModif_p : string; sArguments_p : string = '') : boolean;
var
   FHisto_l : THistorisation;
   bAHistoriser_l, bPrecedent_l : boolean;
begin
   FHisto_l := THistorisation.Create;
   FHisto_l.Init(OBElement_p, sTable_p, sTableHisto_p,
                 sCle_p, sOrdre_p, sChampMax_p,
                 sTypeModif_p, sArguments_p);
   
   bPrecedent_l := FHisto_l.Precedents;
   FHisto_l.MiseAJour;
   FHisto_l.Complement;

   if bPrecedent_l then
      bAHistoriser_l := FHisto_l.AHistoriser
   else
      bAHistoriser_l := true;
   if bAHistoriser_l then
      FHisto_l.Enregistre;

   FHisto_l.OBHistoNew_c.Free;
   FHisto_l.OBHistoPre_c.Free;
   FHisto_l.Free;
   result := bAHistoriser_l;
end;

//////////////////////////////////////////////////////////////////
// Méthodes de la classe                                        //
//////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 07/09/2004
Modifié le ... : 07/09/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure THistorisation.Init(OBElement_p : TOB; sTable_p, sTableHisto_p,
                              sCle_p, sOrdre_p, sChampMax_p : string;
                              sTypeModif_p, sArguments_p : string);
begin
   sTable_c      := sTable_p;
   sTableHisto_c := sTableHisto_p;
   sPref_c       := TableToPrefixe(sTable_c);
   sPrefHisto_c  := TableToPrefixe(sTableHisto_c);

   sCle_c        := sCle_p;
   sOrdre_c      := sOrdre_p;
   sChampMax_c   := sChampMax_p;
   sTypeModif_c  := sTypeModif_p;
   sArguments_c  := sArguments_p;
   OBElement_c   := OBElement_p;

   OBHistoPre_c := TOB.Create(sTableHisto_c, nil, -1);
   OBHistoNew_c := TOB.Create(sTableHisto_c, nil, -1);   
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 16/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
// $$$ JP 05/12/05 - inutilisé (warning Delphi)
{function THistorisation.InitWhere : string;
var
   sWhere_l, sCleHisto_l, sChampHisto_l, sChamp_l : string;
begin
   sCleHisto_l := TableToCle1(sTableHisto_c);
   sChampHisto_l := READTOKENPipe(sCleHisto_l, ',');

   while sCleHisto_l <> '' do
   begin
      sChamp_l := sChampHisto_l;
      READTOKENPipe(sChamp_l, '_');
      sChamp_l := sPref_c + '_' + sChamp_l;
      if sWhere_l <> '' then
         sWhere_l := sWhere_l + ' AND ';

      sWhere_l := sWhere_l +
                  sChampHisto_l + ' = ' +
                  FieldToQuotedStr(sChampHisto_l, OBElement_c.GetValue(sChamp_l));

      sChampHisto_l := READTOKENPipe(sCleHisto_l, ',');
   end;
   sChampMax_c := sChampHisto_l;
   sOrdre_c := sChampHisto_l;
   result := sWhere_l;
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 07/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function THistorisation.Precedents : boolean;
var
   sRequete_l, sWhere_l : string;
//   sCleTable_l : string;
begin
   sRequete_l := 'SELECT * FROM ' + sTableHisto_c;
//   sWhere_l := InitWhere;
   sWhere_l := sCle_c;
   if sWhere_l <> '' then
      sRequete_l := sRequete_l + ' WHERE ' + sWhere_l;

   if sOrdre_c <> '' then
      sRequete_l := sRequete_l + ' ORDER BY ' + sOrdre_c;

   OBHistoPre_c.LoadDetailDBFromSQL(sTableHisto_c, sRequete_l);
   result := (OBHistoPre_c.Detail.Count <> 0);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 07/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure THistorisation.MiseAJour;
var
   iInd_l : integer;
   sChamp_l : string;
   vValeur_l : variant;
begin
   for iInd_l := 1 to OBElement_c.NbChamps do
   begin
      sChamp_l := OBElement_c.GetNomChamp(iInd_l);
      READTOKENPipe(sChamp_l, '_');
      if OBHistoNew_c.FieldExists(sPrefHisto_c + '_' + sChamp_l) then
      begin
         vValeur_l := Unassigned;
         vValeur_l := OBElement_c.GetValue(sPref_c + '_' + sChamp_l);
         OBHistoNew_c.PutValue(sPrefHisto_c + '_' + sChamp_l, vValeur_l);
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 07/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure THistorisation.Complement;
var
   sChamp_l, sValeur_l, sType_l : string;
begin
   while sArguments_c <> '' do
   begin
      sChamp_l  := READTOKENPipe(sArguments_c, '=');
      sValeur_l := READTOKENST(sArguments_c);
      if OBHistoNew_c.FieldExists(sChamp_l) then
      begin
         sType_l := ChampToType(sChamp_l);
         sType_l := UpperCase(sType_l);
         if (Copy(sType_l, 1, 7) = 'VARCHAR') or (sType_l = 'COMBO') then
            OBHistoNew_c.PutValue(sChamp_l, sValeur_l)
         else if (sType_l = 'DATE') then
            OBHistoNew_c.PutValue(sChamp_l, StrToDate(sValeur_l))
         else if (sType_l = 'DOUBLE') then
            OBHistoNew_c.PutValue(sChamp_l, StrToFloat(sValeur_l))
         else if (sType_l = 'INTEGER') then
            OBHistoNew_c.PutValue(sChamp_l, StrToInt(sValeur_l));
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 08/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function THistorisation.AHistoriser : boolean;
var
   iInd_l : integer;
   sChamp_l : string;
   vValeurPre_l, vValeurNew_l : variant;
   bDifferent_l : boolean;
begin
   bDifferent_l := false;
   iInd_l := 1;
   while not bDifferent_l and (iInd_l <= OBHistoNew_c.NbChamps) do
   begin
      sChamp_l := OBHistoNew_c.GetNomChamp(iInd_l);
      if (sChamp_l <> sPrefHisto_c + '_DATECREATION') and
         (sChamp_l <> sPrefHisto_c + '_DATEMODIF') and
         (sChamp_l <> sPrefHisto_c + '_APPMODIF') and
         (sChamp_l <> sPrefHisto_c + '_UTILISATEUR') and
         (sChamp_l <> sPrefHisto_c + '_TYPEMODIF') and
         (sChamp_l <> sChampMax_c) then
      begin
         vValeurPre_l := OBHistoPre_c.Detail[0].GetValue(sChamp_l);
         vValeurNew_l := OBHistoNew_c.GetValue(sChamp_l);
         if vValeurPre_l <> vValeurNew_l then
            bDifferent_l := true;
      end;
      Inc(iInd_l);
   end;
   result := bDifferent_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 07/09/2004
Modifié le ... : 07/09/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure THistorisation.Enregistre;
begin
   if OBHistoPre_c.Detail.Count = 0 then
      OBHistoNew_c.PutValue(sChampMax_c, 1)
   else
      OBHistoNew_c.PutValue(sChampMax_c, OBHistoPre_c.Detail[0].GetValue(sChampMax_c) + 1);

   if OBHistoNew_c.FieldExists(sPrefHisto_c + '_DATECREATION') then
      OBHistoNew_c.PutValue(sPrefHisto_c + '_DATECREATION', Date);
   if OBHistoNew_c.FieldExists(sPrefHisto_c + '_DATEMODIF') then
      OBHistoNew_c.PutValue(sPrefHisto_c + '_DATEMODIF', Date);
   if OBHistoNew_c.FieldExists(sPrefHisto_c + '_APPMODIF') then
      OBHistoNew_c.PutValue(sPrefHisto_c + '_APPMODIF', NomHalley);
   if OBHistoNew_c.FieldExists(sPrefHisto_c + '_TYPEMODIF') then
      OBHistoNew_c.PutValue(sPrefHisto_c + '_TYPEMODIF', sTypeModif_c);
   if OBHistoNew_c.FieldExists(sPrefHisto_c + '_UTILISATEUR') then
      OBHistoNew_c.PutValue(sPrefHisto_c + '_UTILISATEUR', V_PGI.User);

   OBHistoNew_c.InsertDB(nil);
end;

/////////////////////////////////////////////////////////////////

end.
