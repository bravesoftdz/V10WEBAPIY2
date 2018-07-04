{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 04/09/2002
Modifié le ... : 28/02/2003 : Compatibilité CWAS
Description .. : Classe pour la fusion de documents
Mots clefs ... :
*****************************************************************}

unit CLASSAnnuLien;

interface

uses
{$IFNDEF EAGLCLIENT}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ELSE}
{$ENDIF}
   UTOB;

type
   TAnnuLien = class
      public
         sGuidPerDos_c, sGuidPer_c, sPerass1Guid_c,
         sCodeDos_c, sTypeDos_c, sFonction_c, sForme_c, sFonctLibelle_c : string;
         sAffiche_c, sDefaut_c, sRacine_c, sTiers_c, sInfo_c, sFonctLibbelle_c : string;
         sDosLibelle_c, sJurForme_c, sTypePer_c, sNoAdmin_c, sNoIdentif_c, sCompltNoAdmin_c : string;
         sFamPer_c, sTitres_c : string;
         nNoOrdre_c, nTri_c, nOnglet_c : integer;
         nTTNBPP_c, nTTNBNP_c, nTTNBUS_c, nTTNBTOTUS_c, nTTNBTOTPP_c {nTTNBTOT_c,} : integer;
         nNbTitresClot_c, nNbDroitsVote_c, {nVoixTotal_c,} nVoixAge_c, nVoixAgo_c : integer;
         dPctVoix_c, dPctBenef_c, dPctCap_c : double;
         dtExpDate_c : TDateTime;

         function  LienValeurDefaut : boolean;
         procedure LienCalculDefaut;
         function  OngletChoix(sFonction_p : string) : integer;
         procedure Organisme( bVisible_p : boolean);
         procedure TitresEtDroits( var nNbTitresClot_p, nNbDroitsVote_p : integer );

         procedure TotauxPCTCalcule( nNbTitresClot_p, nTtNbTotUS_p, nTtNbTot_p : integer;
                     var dPctBenef_p, dPctCap_p : double );
         function  TotauxVoixCalcule( nNbDroitsVote_p, nVoixTotal_p : integer) : double;
         procedure VoixCalcule( nNbPP_p, nNbUS_p, nNbNP_p : Integer;
                                var nVoixAgo_p, nVoixAge_p, 
                                nTTNBTOTUS_p, nTTNBTOTPP_p : integer);
         procedure GetTypePer;                     
      private

         procedure LienAffecte( sParametres_p : string);
         function  LienAutreQue( sListeFonction_p : string ) : string;
         function  LienAvecJuri( sListeFonction_p : string ) : string;
         function  LienCourant : string;
         function  LienExiste( sWhere_p : string ) : boolean;
         Procedure LienInit(nNoOrdre_p : integer;
                           sGuidPerDos_p, sGuidPer_p, sTypeDos_p, sFonction_p : string;
                           sForme_p : string = ''; sCodeDos_p : string = '' );
         function  LienPrepareINS( OBMere_p : TOB) : TOB;
         procedure LienPrepareMAJ(OBLien_p : TOB);
         procedure LienPrepare(OBLien_p : TOB);                  
         procedure SupprimeUnLien( sFonction_p : string );
         function  SupprimeAutresLiens( sFonction_p : string ) : boolean;
         procedure SupprimeIntervenant;

   end;

/////////// ENTETES DE FONCTION ET PROCEDURES ////////////

procedure CreationUnLien(nNoOrdre_p  : integer;
                        sGuidPerDos_p, sGuidPer_p, sTypeDos_p, sFonction_p, sForme_p, sCodeDos_p : string;
                        sOptions_p : string = '');

function InsOrUpdDepuisTOB(OBDetail_p : TOB;
                        sOptions_p : string = '') : boolean;

procedure SupprimeUnLien(iNoOrdre_p : integer;
                         sGuidPerDos_p, sGuidPer_p, sTypeDos_p, sFonction_p : string );

{procedure SupprimeUnLienJURI(nNoOrdre_p : integer;
                       sGuidPerDos_p, sGuidPer_p, sTypeDos_p, sFonction_p : string );

procedure SupprimeUnLienDP(nNoOrdre_p : integer;
                          sGuidPerDos_p, sGuidPer_p, sTypeDos_p, sFonction_p : string;
                          bIntSeulmt_p : Boolean = FALSE );
}
procedure ANLTotauxPCTCalcule(nNbTitresClot_p, nTtNbTotUS_p, nTtNbTot_p : integer;
                              var dPctBenef_p, dPctCap_p : double );
function  ANLTotauxVoixCalcule(nNbDroitsVote_p, nVoixTotal_p : integer) : double;
procedure ANLVoixCalcule(nNbPP_p, nNbUS_p, nNbNP_p : Integer;
                         var nVoixAgo_p, nVoixAge_p,
                         nTTNBTOTUS_p, nTTNBTOTPP_p : integer);
function  ANLOngletChoix(sFonction_p : string; var sTitres_p : string) : integer;

implementation

uses
   hctrls, SysUtils, DpJurOutils, hmsgbox, AnnOutils, HEnt1, Controls, DPJUROutilsBlob;

{*****************************************************************
* Procédures globales d'accès à la classe                        *
*****************************************************************}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/04/2003
Modifié le ... :   /  /    
Description .. : par exemple sOptions_p :
   'ANL_TTNBPP=10;ANL_TTNBNP=20;ANL_TTNBUS=30;ANL_PERASS1GUID=80;ANL_EXPDATE=20/02/2004'

Mots clefs ... : 
*****************************************************************}
procedure CreationUnLien( nNoOrdre_p : integer;
                        sGuidPerDos_p, sGuidPer_p, sTypeDos_p, sFonction_p, sForme_p, sCodeDos_p : string;
                        sOptions_p : string = '');
var
   FLien_l : TAnnuLien;
   OBLien_l : TOB;
begin
   FLien_l := TAnnuLien.Create;
   FLien_l.LienInit(nNoOrdre_p, sGuidPerDos_p, sGuidPer_p, sTypeDos_p, sFonction_p,
                 sForme_p, sCodeDos_p );
   FLien_l.LienAffecte( sOptions_p );
                 
   if Not FLien_l.LienExiste( FLien_l.LienCourant ) then
   begin
      OBLien_l := FLien_l.LienPrepareINS( nil );
      OBLien_l.InsertDB( Nil );
      OBLien_l.Free;
   end;

   FLien_l.Free;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : JP 16/10/2003
Créé le ...... : 27/11/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function InsOrUpdDepuisTOB(OBDetail_p : TOB;
                        sOptions_p : string = '') : boolean;
var
   FLien_l : TAnnuLien;
begin
   result := false;
   FLien_l := TAnnuLien.Create;
   FLien_l.LienInit( OBDetail_p.GetValue('ANL_NOORDRE'), OBDetail_p.GetValue('ANL_GUIDPERDOS'),
                     OBDetail_p.GetValue('ANL_GUIDPER'), OBDetail_p.GetValue('ANL_TYPEDOS'),
                     OBDetail_p.GetValue('ANL_FONCTION'), OBDetail_p.GetValue('ANL_FORME'),
                     OBDetail_p.GetValue('ANL_CODEDOS'));
                     
   FLien_l.LienAffecte( sOptions_p );
   FLien_l.LienPrepareMAJ(OBDetail_p);
   OBDetail_p.InsertOrUpdateDB;
   FLien_l.Free;
   result := true;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/02/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure SupprimeUnLien(iNoOrdre_p : integer;
                         sGuidPerDos_p, sGuidPer_p, sTypeDos_p, sFonction_p : string );
var
   OBLien_l, OBDetail_l : TOB;
   sMsg_l, sGuidConv_l : string;
   iOBInd_l : integer;
begin
   OBLien_l := TOB.Create('ANNULIEN', nil, -1);
   OBLien_l.LoadDetailDBFromSQL('ANNULIEN',
               'SELECT * FROM ANNULIEN ' +
               'WHERE ANL_GUIDPERDOS = "' + sGuidPerDos_p + '" ' +
               '  AND ANL_GUIDPER = "' + sGuidPer_p + '" ' +
               '  AND ANL_TYPEDOS = "' + sTypeDos_p + '" ' +
               '  AND ANL_NOORDRE = ' + IntToStr(iNoOrdre_p) + '');

   if (sTypeDos_p = 'STE') and (sFonction_p = 'INT') and (OBLien_l.Detail.Count > 1) then
   begin
      sMsg_l := 'Cette personne a d''autres fonctions dans ce dossier.'#13#10+
                'Vous ne pouvez pas détruire cette ligne si vous ne supprimez pas ses autres fonctions dans ce dossier.'#13#10+
                'Voulez-vous détruire tous les autres liens de cette personne dans le dossier ?';

      //On propose la destruction des autres liens
      if (PgiAsk(sMsg_l, TitreHalley) = MrYes) then
      begin
         for iOBInd_l := 0 to OBLien_l.Detail.Count - 1 do
         begin
            // supression blob conventions
            sGuidConv_l := OBLien_l.Detail[iOBInd_l].GetString('ANL_CONVTXT');
            BlobsSupprime('ANL', sGuidConv_l, '');
            OBLien_l.Detail[iOBInd_l].DeleteDB;
         end;
      end;
   end
   else
   begin
      OBDetail_l := OBLien_l.FindFirst(['ANL_TYPEDOS', 'ANL_FONCTION'], [sTypeDos_p, sFonction_p], true);
      if OBDetail_l <> nil then
      begin
         // supression blob conventions
         sGuidConv_l := OBDetail_l.GetString('ANL_CONVTXT');
         BlobsSupprime('ANL', sGuidConv_l, '');
         OBDetail_l.DeleteDB;
         OBDetail_l.free;
      end;
      // Lien intervenant
      OBDetail_l := OBLien_l.FindFirst(['ANL_TYPEDOS', 'ANL_FONCTION'], ['STE', 'INT'], true);
      if (OBLien_l.Detail.Count = 1) and (OBDetail_l <> nil) then
      begin
         // supression blob conventions
         sGuidConv_l := OBDetail_l.GetString('ANL_CONVTXT');
         BlobsSupprime('ANL', sGuidConv_l, '');
         OBDetail_l.DeleteDB;
      end;
   end;

   OBLien_l.Free;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{procedure SupprimeUnLienJURI(nNoOrdre_p : integer;
                       sGuidPerDos_p, sGuidPer_p, sTypeDos_p, sFonction_p : string );
var
   FLien_l : TAnnuLien;
   bContinue_l : boolean;
   sWhere_l : string;
begin
   FLien_l := TAnnuLien.Create;
   FLien_l.LienInit(nNoOrdre_p, sGuidPerDos_p, sGuidPer_p, sTypeDos_p, sFonction_p);

   bContinue_l := true;

   if sFonction_p = 'INT' then
      bContinue_l := FLien_l.SupprimeAutresLiens('INT;TRS');

   if bContinue_l then
   begin
      FLien_l.SupprimeUnLien(sFonction_p);    // Suppression de l'enregistrement sélectionné
      if sFonction_p <> 'INT' then
         FLien_l.SupprimeIntervenant; // suppression du lien INTERVENANT (loi NRE)
   end;

   FLien_l.Free;
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 31/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
{procedure SupprimeUnLienDP(nNoOrdre_p : integer;
                       sGuidPerDos_p, sGuidPer_p, sTypeDos_p, sFonction_p : string;
                       bIntSeulmt_p : Boolean = FALSE );
var
   FLien_l : TAnnuLien;
begin
   FLien_l := TAnnuLien.Create;
   FLien_l.LienInit(nNoOrdre_p, sGuidPerDos_p, sGuidPer_p, sTypeDos_p, sFonction_p );

   if not bIntSeulmt_p then
      FLien_l.SupprimeUnLien(sFonction_p);       // Lien en cours
   FLien_l.SupprimeIntervenant;  // suppression du lien INTERVENANT (loi NRE)
   FLien_l.Free;
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 31/05/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function ANLOngletChoix(sFonction_p : string; var sTitres_p : string) : integer;
var
   nNumOnglet_l : integer;
   OBTypeFonct_l : TOB;
begin
   nNumOnglet_l := -1;
   OBTypeFonct_l := TOB.Create('JUTYPEFONCT', nil, -1);
   OBTypeFonct_l.LoadDetailDBFromSQL('JUTYPEFONCT',
                                     'select JTF_FENETRE, JTF_TITRE ' +
                                     'from JUTYPEFONCT ' +
                                     'where JTF_FONCTION = "' + sFonction_p + '"');
   if OBTypeFonct_l.Detail.Count > 0 then
   begin
      nNumOnglet_l := OBTypeFonct_l.Detail[0].GetInteger('JTF_FENETRE');
      sTitres_p := OBTypeFonct_l.Detail[0].GetString('JTF_TITRE');
   end;
   OBTypeFonct_l.Free;
   Result := nNumOnglet_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 30/05/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure ANLTotauxPCTCalcule(nNbTitresClot_p, nTtNbTotUS_p, nTtNbTot_p : integer;
                              var dPctBenef_p, dPctCap_p : double);
begin
   dPctBenef_p := 0;
   dPctCap_p := 0;
   if nNbTitresClot_p > 0 then
   begin
      dPctBenef_p := Arrondi((nTtNbTotUS_p / nNbTitresClot_p), 9);
      dPctCap_p :=   Arrondi((nTtNbTot_p / nNbTitresClot_p), 9);
//      dPctBenef_p := Arrondi((1.0 *(10000 * nTtNbTotUS_p / nNbTitresClot_p) / 10000), 9);
//      dPctCap_p := Arrondi((1.0 *(10000 * nTtNbTot_p / nNbTitresClot_p) / 10000), 9);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 30/05/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function ANLTotauxVoixCalcule(nNbDroitsVote_p, nVoixTotal_p : integer) : double;
begin
   Result := 0;
   if nNbDroitsVote_p > 0 then
      Result := Arrondi((nVoixTotal_p / nNbDroitsVote_p), 9);
//      Result := Arrondi((1.0 * (10000 * nVoixAge_p / nNbDroitsVote_p)/10000),9);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 30/05/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure ANLVoixCalcule(nNbPP_p, nNbUS_p, nNbNP_p : Integer;
                         var nVoixAgo_p, nVoixAge_p,
                         nTTNBTOTUS_p, nTTNBTOTPP_p : integer);
begin
   nTTNBTOTUS_p := nNbPP_p + nNbUS_p;
   nTTNBTOTPP_p := nNbPP_p + nNbNP_p;
//   nTTNBTOT_c := nNbPP_p + nNbUS_p + nNbNP_p;
   nVoixAgo_p := nTTNBTOTUS_p;
   nVoixAge_p := nTTNBTOTPP_p;
//   nVoixTotal_c := nVoixAgo_p + nVoixAge_p;
end;


{*****************************************************************
* Méthodes de la classe                                          *
*****************************************************************}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 11/04/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
Procedure TAnnuLien.LienInit(nNoOrdre_p : integer;
                sGuidPerDos_p, sGuidPer_p, sTypeDos_p, sFonction_p : string;
                sForme_p : string = ''; sCodeDos_p : string = '' );
begin
   sGuidPerDos_c := sGuidPerDos_p;
   nNoOrdre_c := nNoOrdre_p;
   sGuidPer_c := sGuidPer_p;
   sTypeDos_c := sTypeDos_p;
   sFonction_c := sFonction_p;
   sForme_c := sForme_p;
   sCodeDos_c := sCodeDos_p;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 22/04/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
Procedure TAnnuLien.LienAffecte( sParametres_p : string );
var
   sValChamp_l, sNomChamp_l : string;
begin
   while sParametres_p <> '' do
   begin
      sValChamp_l := ReadTokenSt(sParametres_p);
      sNomChamp_l := ReadTokenPipe(sValChamp_l, '=');
{      if sNomChamp_l = 'ANL_CODEDOS' then
         sCodeDos_c := sValChamp_l
      else if sNomChamp_l = 'ANL_FORME' then
         sForme_c := sValChamp_l
      else if sNomChamp_l = 'ANL_TYPEPER' then
         sTypePer_c := sValChamp_l
      else} if sNomChamp_l = 'ANL_TTNBPP' then
         nTTNBPP_c := StrToInt(sValChamp_l)
      else if sNomChamp_l = 'ANL_TTNBNP' then
         nTTNBPP_c := StrToInt(sValChamp_l)
      else if sNomChamp_l = 'ANL_TTNBUS' then
         nTTNBUS_c := StrToInt(sValChamp_l)
      else if sNomChamp_l = 'ANL_TTNBTOT' then
         nTTNBTOTPP_c := StrToInt(sValChamp_l)
      else if sNomChamp_l = 'ANL_PERASS1GUID' then
         sPerass1Guid_c := sValChamp_l
      else if sNomChamp_l = 'ANL_EXPDATE' then
         dtExpDate_c := StrToDate(sValChamp_l);
//      else
   end;

end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 31/03/2003
Modifié le ... :   /  /    
Description .. : lien autre que intervenant
Mots clefs ... :
*****************************************************************}
function  TAnnuLien.LienAutreQue(sListeFonction_p : string ) : string;
var
   sRequete_l, sCompReq_l, sFonction_l : string;
begin
   sRequete_l := 'where ANL_GUIDPERDOS = "' + sGuidPerDos_c + '" ' +
                 '  AND ANL_GUIDPER = "' + sGuidPer_c + '" ' +
                 '  AND ANL_TYPEDOS = "' + sTypeDos_c + '"';
   sCompReq_l := '';
   while sListeFonction_p <> '' do
   begin
      sFonction_l := ReadTokenST(sListeFonction_p);
      if sCompReq_l <> '' then
         sCompReq_l := sCompReq_l + '  AND ';
      sCompReq_l := sCompReq_l + '  ANL_FONCTION <> "' + sFonction_l + '"';
   end;
   if sCompReq_l <> '' then
      sRequete_l := sRequete_l + '  AND (' + sCompReq_l + ')';
   Result := sRequete_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 31/03/2003
Modifié le ... :   /  /    
Description .. : lien INT n'est pas utilisé par un dossier juridique
Mots clefs ... :
*****************************************************************}
function  TAnnuLien.LienAvecJuri( sListeFonction_p : string ) : string;
var
   sRequete_l, sCompReq_l, sFonction_l : string;
begin
   sRequete_l := 'where ANL_GUIDPERDOS = "' + sGuidPerDos_c + '" ' +
                 '  AND ANL_GUIDPER = "' + sGuidPer_c + '" ' +
                 '  AND ANL_TYPEDOS = "' + sTypeDos_c + '"' +
                 '  AND ANL_CODEDOS <> "&#@"';
   sCompReq_l := '';
   while sListeFonction_p <> '' do
   begin
      sFonction_l := ReadTokenST(sListeFonction_p);
      if sCompReq_l <> '' then
         sCompReq_l := sCompReq_l + '  OR ';
      sCompReq_l := sCompReq_l + '  ANL_FONCTION = "' + sFonction_l + '"';
   end;
   if sCompReq_l <> '' then
      sRequete_l := sRequete_l + '  AND (' + sCompReq_l + ')';
   Result := sRequete_l;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
function  TAnnuLien.LienCourant : string;
var
   sRequete_l : string;
begin
   sRequete_l := 'where ANL_GUIDPERDOS = "' + sGuidPerDos_c + '" ' +
                 '  AND ANL_GUIDPER = "' + sGuidPer_c + '" ' + 
                 '  AND ANL_NOORDRE = ' + IntToStr(nNoOrdre_c) +
                 '  AND ANL_TYPEDOS = "' + sTypeDos_c + '"' +
                 '  AND ANL_FONCTION = "' + sFonction_c + '"';
   result := sRequete_l;
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 18/09/02
Procédure .... : InitAnnuLien
Description .. : Ajoute des éléments an sous-menu courant
Paramètres ... :
*****************************************************************}
function TAnnuLien.LienExiste( sWhere_p : string ) : boolean;
var
   sRequete_l : string;
begin
   sRequete_l := 'select ANL_GUIDPER from ANNULIEN ' + sWhere_p;
   Result := ExisteSQL( sRequete_l );
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 01/04/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TAnnuLien.LienPrepareINS( OBMere_p : TOB) : TOB;
var
   OBLien_l : TOB;
begin
   OBLien_l := TOB.Create('ANNULIEN', OBMere_p, -1);
   OBLien_l.InitValeurs;

   OBLien_l.PutValue('ANL_GUIDPERDOS', sGuidPerDos_c);
   OBLien_l.PutValue('ANL_TYPEDOS', sTypeDos_c);
   OBLien_l.PutValue('ANL_NOORDRE', nNoOrdre_c);
   OBLien_l.PutValue('ANL_GUIDPER', sGuidPer_c);
   OBLien_l.PutValue('ANL_FONCTION', sFonction_c);
   OBLien_l.PutValue('ANL_CODEDOS', sCodeDos_c);

   LienPrepare(OBLien_l);
   result := OBLien_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 30/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TAnnuLien.LienPrepareMAJ(OBLien_p : TOB);
begin
   nTTNBPP_c      := OBLien_p.GetValue('ANL_TTNBPP');
   nTTNBNP_c      := OBLien_p.GetValue('ANL_TTNBNP');
   nTTNBUS_c      := OBLien_p.GetValue('ANL_TTNBUS');
   sPerass1Guid_c := OBLien_p.GetValue('ANL_PERASS1GUID');
   dtExpDate_c    := OBLien_p.GetValue('ANL_EXPDATE');
   nTri_c         := OBLien_p.GetValue('ANL_TRI');
   sTiers_c       := OBLien_p.GetValue('ANL_TIERS');

   LienPrepare(OBLien_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 08/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TAnnuLien.LienPrepare(OBLien_p : TOB);
var
   sLibelle_l, sNom_l : string;
begin
   LienValeurDefaut;
   if (sDefaut_c = 'X') and (sFonction_c <>' STE') then
      OBLien_p.PutValue('ANL_FORME', sForme_c)
   else
      OBLien_p.PutValue('ANL_FORME', '');

   TraiteChoixPersonne(sGuidPer_c, sLibelle_l, sNom_l, sTypePer_c);
   OBLien_p.PutValue('ANL_NOMPER', sNom_l );
   OBLien_p.PutValue('ANL_TYPEPER', sTypePer_c);

   if GetValChamp('JUTYPEFONCT','JTF_TIERS', 'JTF_FONCTION = "' + sFonction_c + '"') = 'X' then
      OBLien_p.PutValue('ANL_TYPEPERDOS', sTypePer_c);

   OBLien_p.PutValue('ANL_TRI', nTri_c);
   OBLien_p.PutValue('ANL_TIERS', sTiers_c);
   OBLien_p.PutValue('ANL_TTNBPP', nTTNBPP_c);
   OBLien_p.PutValue('ANL_TTNBNP', nTTNBNP_c);
   OBLien_p.PutValue('ANL_TTNBUS', nTTNBUS_c);
   OBLien_p.PutValue('ANL_PERASS1GUID', sPerass1Guid_c);
   OBLien_p.PutValue('ANL_EXPDATE', dtExpDate_c);

   nOnglet_c := OngletChoix( sFonction_c );
   TitresEtDroits( nNbTitresClot_c, nNbDroitsVote_c );

   VoixCalcule( nTTNBPP_c, nTTNBUS_c, nTTNBNP_c, nVoixAgo_c, nVoixAge_c, nTTNBTOTUS_c, nTTNBTOTPP_c );
   OBLien_p.PutValue('ANL_VOIXAGO', nVoixAgo_c);
   OBLien_p.PutValue('ANL_VOIXAGE', nVoixAge_c);
   OBLien_p.PutValue('ANL_TTNBTOTUS', nTTNBTOTUS_c);
   OBLien_p.PutValue('ANL_TTNBTOT', nTTNBTOTPP_c);

   dPctVoix_c := TotauxVoixCalcule( nNbDroitsVote_c, nVoixAgo_c + nVoixAge_c);
   OBLien_p.PutValue('ANL_TTPCTVOIX', dPctVoix_c );

   TotauxPCTCalcule( nNbTitresClot_c, nTTNBTOTUS_c, nTTNBTOTPP_c, dPctBenef_c, dPctCap_c );
   OBLien_p.PutValue('ANL_TTPCTBENEF', dPctBenef_c);
   OBLien_p.PutValue('ANL_TTPCTCAP', dPctCap_c);

   LienCalculDefaut;
   OBLien_p.PutValue('ANL_RACINE', sRacine_c);
   OBLien_p.PutValue('ANL_AFFICHE', sAffiche_c );
   OBLien_p.PutValue('ANL_INFO', Copy(sInfo_c, 1, 17));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TAnnuLien.SupprimeUnLien( sFonction_p : string );
var
   OBLien_l : TOB;
   sCle_l : string;
begin
   OBLien_l := TOB.Create('ANNULIEN',nil, -1);
   sCle_l := '"' + sGuidPerDos_c + '";"' + sTypeDos_c +'";' +
             IntToStr(nNoOrdre_c) + ';"' + sFonction_p +'"' +';"' + sGuidPer_c + '"';

   if OBLien_l.SelectDB(sCle_l, nil ) then
      OBLien_l.DeleteDB;

   OBLien_l.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 31/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TAnnuLien.SupprimeAutresLiens( sFonction_p : string ) : boolean;
var
   bContinue_l : boolean;
   OBLien_l : TOB;
   QRY_l : TQuery;
   sMsg_l : string;
   nLienInd_l : integer;
begin
   bContinue_l := true;
   OBLien_l := TOB.Create('ANNULIEN',nil, -1);
   QRY_l := OpenSQL('SELECT * ' +
                    'FROM ANNULIEN ' +
                    LienAutreQue( sFonction_p ), true);

   OBLien_l.LoadDetailDB( 'ANNULIEN', '', '', QRY_l, false);
   Ferme(QRY_l);
   if OBLien_l.Detail.Count > 0 then
   begin
      sMsg_l := 'Cette personne a d''autres fonctions dans ce dossier.'#13#10+
                'Vous ne pouvez pas détruire cette ligne si vous ne supprimez pas ses autres fonctions dans ce dossier.'#13#10+
                'Voulez-vous détruire tous les autres liens de cette personne dans le dossier ?';

      //On propose la destruction des autres liens
      if (PgiAsk(sMsg_l, TitreHalley) = MrYes) then
      begin
         for nLienInd_l := 0 to OBLien_l.Detail.Count - 1 do
            OBLien_l.Detail[nLienInd_l].DeleteDB;
      end
      else
         bContinue_l := false;
   end;
   OBLien_l.Free;
   // Pas d'autres liens
   result := bContinue_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 31/03/2003
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TAnnuLien.SupprimeIntervenant;
begin
   if not LienExiste( LienAutreQue( 'INT' ) ) and
      not LienExiste( LienAvecJuri( 'INT' ) ) then
   begin
      SupprimeUnLien('INT');
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 11/04/2003
Modifié le ... : 19/12/2003
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TAnnuLien.Organisme( bVisible_p : boolean);
var
   QRYReq_l : TQuery;
begin
   if bVisible_p then
   begin
      // rapatrier les nos de l'organisme
      QRYReq_l := OpenSQL('select ANN_NOADMIN, ANN_NOIDENTIF, ANN_COMPLTNOADMIN, ANN_FAMPER ' +
                          'from ANNUAIRE ' +
                          'where ANN_GUIDPER = "' + sGuidPer_c + '"', TRUE);
      if not QRYReq_l.eof then
      begin
         sNoAdmin_c := QRYReq_l.FindField('ANN_NOADMIN').AsString;
         sNoIdentif_c := QRYReq_l.FindField('ANN_NOIDENTIF').AsString;
         sCompltNoAdmin_c := QRYReq_l.FindField('ANN_COMPLTNOADMIN').AsString;
         sFamPer_c := QRYReq_l.FindField('ANN_FAMPER').AsString;
      end;
      Ferme(QRYReq_l);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
function TAnnuLien.OngletChoix(sFonction_p : string) : integer;
//var
//   nNumOnglet_l : integer;
//   QRYReq_l : TQuery;
begin
{   nNumOnglet_l := -1;
   QRYReq_l := OpenSQL('select JTF_FENETRE, JTF_TITRE from JUTYPEFONCT ' +
                       'where JTF_FONCTION = "' + sFonction_p + '"', true);
   if not QRYReq_l.EOF then
   begin
      nNumOnglet_l := QRYReq_l.FindField('JTF_FENETRE').AsInteger;
      sTitres_c := QRYReq_l.FindField('JTF_TITRE').AsString;
   end;
   Ferme(QRYReq_l);
   Result := nNumOnglet_l;}
   Result := ANLOngletChoix(sFonction_p, sTitres_c);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TAnnuLien.LienValeurDefaut : boolean;
var
  QRYReq_l : TQuery;
  sRequete_l : string;
begin
   result := false;
   sAffiche_c := RechDom('JUTYPEFONCTAFF', sFonction_c, FALSE);
   sInfo_c := '';
   sRacine_c := '';

   sRequete_l := 'select JFT_DEFAUT, JFT_RACINE, JFT_TYPEDOS, JFT_TRI, JFT_TIERS, JFT_TITRE ' +
                 'from JUFONCTION ' +
                 'where JFT_FONCTION = "' + sFonction_c + '"' +
                 'and JFT_FORME = "' + sForme_c + '"';
                 
   QRYReq_l := OpenSQL(sRequete_l, true);
   if not QRYReq_l.Eof then
   begin
      sDefaut_c := QRYReq_l.FindField('JFT_DEFAUT').AsString;
      sRacine_c := QRYReq_l.FindField('JFT_RACINE').AsString;
      sTypeDos_c := QRYReq_l.FindField('JFT_TYPEDOS').AsString;
      nTri_c := QRYReq_l.FindField('JFT_TRI').AsInteger;
      sTiers_c := QRYReq_l.FindField('JFT_TIERS').AsString;
      sTitres_c := QRYReq_l.FindField('JFT_TITRE').AsString;
      result := true;
    end;
    Ferme(QRYReq_l);
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/04/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TAnnuLien.LienCalculDefaut;
begin
   case nOnglet_c of
      0 :
      begin
         if sTypeDos_c = 'STE' then
         begin
            if (sTypePer_c <> '') then
            begin
//               sRacine_c := GetValChamp('JUTYPEPER', 'JTP_RACINE', 'JTP_TYPEPER = "' + sTypePer_c + '"');
               sAffiche_c := RechDom('JUTYPEPERAFF', sTypePer_c, FALSE);
            end;
         end
         else
         begin
            if (sTypePer_c <> '') then
               sRacine_c := GetValChamp('JUTYPEPER', 'JTP_RACINE', 'JTP_TYPEPER = "' + sTypePer_c + '"');

            if sRacine_c = '' then
               sRacine_c := GetValChamp('JUTYPEFONCT', 'JTF_RACINE', 'JTF_FONCTION = "' + sFonction_c + '"');
         end;
      end;
      1 : sInfo_c := 'Intervenant';
      2 :
      begin
         sInfo_c := '';
         if nTTNBNP_c <> 0 then
            sInfo_c := Format('%d NP', [nTTNBNP_c]);
         if sInfo_c <> '' then
            sInfo_c := sInfo_c  + ' ';
         if nTTNBUS_c <> 0 then
            sInfo_c := sInfo_c + Format('%d US', [nTTNBUS_c]);
         if sInfo_c <> '' then
            sInfo_c := sInfo_c  + ' ';
         if nTTNBPP_c <> 0 then
           sInfo_c := sInfo_c + Format('%d PP', [nTTNBPP_c]);
      end;
      3 :
      begin
         if sPerass1Guid_c <> '' then
            sInfo_c := RechDom('ANNUAIRE', sPerass1Guid_c, FALSE);
      end;
      4..5 :
      begin
         if dtExpDate_c <> iDate1900 then
            sInfo_c := DateToStr(dtExpDate_c);
      end;
      6 :
      begin
         if (sTypePer_c <> '') then
         begin
            sRacine_c := GetValChamp('JUTYPEPER', 'JTP_RACINE', 'JTP_TYPEPER = "' + sTypePer_c + '"');
            sAffiche_c := RechDom('JUTYPEPERAFF', sTypePer_c, FALSE);
         end;
      end;
      else
         sInfo_c := '';
   end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 13/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
function TAnnuLien.TotauxVoixCalcule(nNbDroitsVote_p, nVoixTotal_p : integer) : double;
begin
   Result := ANLTotauxVoixCalcule(nNbDroitsVote_p, nVoixTotal_p);
{   Result := 0;
   if nNbDroitsVote_p > 0 then
      Result := Arrondi((nVoixTotal_p / nNbDroitsVote_p), 9);
//      Result := Arrondi((1.0 * (10000 * nVoixAge_p / nNbDroitsVote_p)/10000),9);}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 14/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TAnnuLien.VoixCalcule( nNbPP_p, nNbUS_p, nNbNP_p : Integer;
                     var nVoixAgo_p, nVoixAge_p,
                     nTTNBTOTUS_p, nTTNBTOTPP_p : integer);
begin
   ANLVoixCalcule(nNbPP_p, nNbUS_p, nNbNP_p,
                  nVoixAgo_p, nVoixAge_p,
                  nTTNBTOTUS_p, nTTNBTOTPP_p);
{   nTTNBTOTUS_p := nNbPP_p + nNbUS_p;
   nTTNBTOTPP_p := nNbPP_p + nNbNP_p;
//   nTTNBTOT_c := nNbPP_p + nNbUS_p + nNbNP_p;
   nVoixAgo_p := nTTNBTOTUS_p;
   nVoixAge_p := nTTNBTOTPP_p;
//   nVoixTotal_c := nVoixAgo_p + nVoixAge_p;}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 14/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TAnnuLien.TotauxPCTCalcule( nNbTitresClot_p, nTtNbTotUS_p, nTtNbTot_p : integer;
                     var dPctBenef_p, dPctCap_p : double);
begin
   ANLTotauxPCTCalcule(nNbTitresClot_p, nTtNbTotUS_p, nTtNbTot_p,
                       dPctBenef_p, dPctCap_p);

{   dPctBenef_p := 0;
   dPctCap_p := 0;
   if nNbTitresClot_p > 0 then
   begin
      dPctBenef_p := Arrondi((nTtNbTotUS_p / nNbTitresClot_p), 9);
      dPctCap_p :=   Arrondi((nTtNbTot_p / nNbTitresClot_p), 9);
//      dPctBenef_p := Arrondi((1.0 *(10000 * nTtNbTotUS_p / nNbTitresClot_p) / 10000), 9);
//      dPctCap_p := Arrondi((1.0 *(10000 * nTtNbTot_p / nNbTitresClot_p) / 10000), 9);
   end;}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 14/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TAnnuLien.TitresEtDroits( var nNbTitresClot_p, nNbDroitsVote_p : integer);
var
   sRequete_l, sWhere_l : string;
   QRYReq_l : TQuery;
begin
   nNbTitresClot_p := 0;
   nNbDroitsVote_p := 0;

   sRequete_l := 'select C1.ANN_GUIDPER, C1.ANN_NOM1, C1.ANN_NOM2, C1.ANN_NOM3, ' +
                 '       C1.ANN_TYPEPER, C2.JTF_FONCTION, C2.JTF_FONCTLIBELLE, C2.JTF_AFFICHE, ' +
                 '       C3.JUR_CODEDOS, C3.JUR_DOSLIBELLE, C3.JUR_FORME, C3.JUR_NBTITRESCLOT, ' +
                 '       C3.JUR_NBDROITSVOTE ' +
                 'FROM   ANNUAIRE C1, JUTYPEFONCT C2, JURIDIQUE C3 ';

   if (sGuidPer_c <> '') then
   begin
      sWhere_l := 'where (C1.ANN_GUIDPER = "' + sGuidPer_c + '" ' +
                '  and  C2.JTF_FONCTION = "' + sFonction_c + '" ' +
                '  and  C3.JUR_GUIDPERDOS = "' + sGuidPerDos_c + '" ' +
                '  and  C3.JUR_TYPEDOS = "' + sTypeDos_c + '" ' +
                '  and  C3.JUR_NOORDRE = ' + IntToStr(nNoOrdre_c) + ')';
   end
   else
   begin
      sWhere_l := 'where (C3.JUR_GUIDPERDOS = "' + sGuidPerDos_c + '" ' +
                '  and  C3.JUR_TYPEDOS = "' + sTypeDos_c + '" ' +
                '  and  C3.JUR_NOORDRE = ' + IntToStr(nNoOrdre_c) +
                '  and  C2.JTF_FONCTION = "' + sFonction_c + '")';
   end;

   QRYReq_l := OpenSQL(sRequete_l + sWhere_l, true);
   if not QRYReq_l.EOF then
   begin
      sDosLibelle_c := QRYReq_l.FindField('JUR_DOSLIBELLE').AsString;
      sJurForme_c := QRYReq_l.FindField('JUR_FORME').AsString;
      sFonctLibelle_c := QRYReq_l.FindField('JTF_FONCTLIBELLE').AsString;

      nNbTitresClot_p := QRYReq_l.FindField('JUR_NBTITRESCLOT').AsInteger;
      nNbDroitsVote_p := QRYReq_l.FindField('JUR_NBDROITSVOTE').AsInteger;
   end;
   Ferme(QRYReq_l);
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : JP 16/10/2003
Créé le ...... : 26/11/2003
Modifié le ... :   /  /    
Description .. : Pour affecter le type personne du lien à partir du type 
Suite ........ : personne annuaire
Suite ........ : Déplacé ici le 26/11/2003, car peut être utile à JURI
Mots clefs ... : 
*****************************************************************}
procedure TAnnuLien.GetTypePer;
var
   TOBTypePer :TOB;
begin
   TOBTypePer := TOB.Create ('le type lien', nil, -1);
   try
      TOBTypePer.LoadDetailFromSQL ('SELECT ANN_TYPEPER FROM ANNUAIRE WHERE ANN_GUIDPER = "' + sGuidPer_c + '"');
      if TOBTypePer.Detail.Count > 0 then
         sTypePer_c := TOBTypePer.Detail [0].GetValue ('ANN_TYPEPER');
   finally
          TOBTypePer.Free;
   end;
end;

{*****************************************************************
*****************************************************************}

/////////////////////////////////////////////////////////////////

end.
