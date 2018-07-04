{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/09/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MODELESDOC ()
Mots clefs ... : TOF;MODELESDOC
*****************************************************************}
Unit dpTofModelesDoc;
//////////////////////////////////////////////////////////////////
Interface
//////////////////////////////////////////////////////////////////
Uses
   UTOF, FileCtrl, HDB,
{$IFDEF EAGLCLIENT}
   eMul, uTob,
{$ELSE}
   dbtables, mul,
{$ENDIF}
   Classes, sysutils, HCtrls, HTB97, AGLInit,windows,
   HMsgBox, Controls, ParamSoc, hent1, ComCtrls, CBPPath;

//////////////////////////////////////////////////////////////////
Type
   TOF_MODELESDOC = Class (TOF)
         procedure OnArgument (sParametres_p : String ) ; override ;
         procedure OnCancel                 ; override ;
         procedure OnClose                  ; override ;
         procedure Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

      private
         sModeleDir_f : string;
         bModele_f : boolean;
         aosTemp_f : array of string;

         procedure OnClick_BDELETE( Sender : TObject );
         procedure OnClick_BINSERT( Sender : TObject );
         procedure OnClick_BPROPRIETE( Sender : TObject );
         procedure OnDblClick_FLISTE( Sender : TObject );
         function LanceModele( SModeleId_p, SFileId_p : String; sNoDossier_p, sTypeGed_p : string ) : string;

         Procedure ListeTempAjouter( sDocument_p : string);
         Procedure ListeTempSupprimer;
   end ;

//////////////////////////////////////////////////////////////////
Implementation
//////////////////////////////////////////////////////////////////
uses
  dpJurOutils, YNewDocument,

  {$IFDEF VER150}
  Variants,
  {$ENDIF}

  UDossierSelect,
  UGedFiles, UtilGed, UGedDP, GalMenuDisp, UDocumentDepuisModele;

//////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 10/09/2003
Modifié le ... : 10/09/2003
Description .. : Affiche / cache les boutons selon le mode de
Suite ........ : fonctionnement
Mots clefs ... :
*****************************************************************}
procedure TOF_MODELESDOC.OnArgument (sParametres_p : String ) ;
var
   sMode_l : string;
begin
   Inherited ;
   sMode_l := ReadTokenParam('MODE', sParametres_p );
   bModele_f := (sMode_l = 'MODELE');

   SetControlVisible('BINSERT', bModele_f);
   SetControlVisible('BDELETE', bModele_f);
   SetControlVisible('BPROPRIETE', bModele_f);
   SetControlVisible('BOUVRIR', Not bModele_f);

   if bModele_f then
   begin
      TToolBarButton97(GetControl( 'BDELETE' )).OnClick := OnClick_BDELETE;
      TToolBarButton97(GetControl( 'BINSERT' )).OnClick := OnClick_BINSERT;
      TToolBarButton97(GetControl( 'BPROPRIETE' )).OnClick := OnClick_BPROPRIETE;
      Ecran.OnKeyDown := Form_OnKeyDown;
   end;
   THDBGrid(GetControl( 'FLISTE' )).OnDblClick := OnDblClick_FLISTE;

   sModeleDir_f := GetParamSocSecur('SO_MDREPDOCUMENTS', TCbpPath.GetCegidDistriStd+'\BUREAU');
   if Not DirectoryExists(sModeleDir_f) then
   begin
      try
         ForceDirectories(sModeleDir_f);
      except;
      end;
   end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 10/09/2003
Modifié le ... : 10/09/2003
Description .. : Annulation :
Suite ........ : - mode paramétrage : rien;
Suite ........ : - mode sélection : la sélection éventuelle est effacée
Mots clefs ... :
*****************************************************************}
procedure TOF_MODELESDOC.OnCancel () ;
begin
   Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 10/09/2003
Modifié le ... : 10/09/2003
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MODELESDOC.OnClose ;
begin
   Inherited ;
   if not bModele_f then
      ListeTempSupprimer;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 10/09/2003
Modifié le ... : 10/09/2003
Description .. : Création d'un nouveau modèle de document ou ouverture 
Suite ........ : d'un modèle de document existant
Mots clefs ... : 
*****************************************************************}
procedure TOF_MODELESDOC.OnClick_BINSERT( Sender : TObject );
var sResultDoc_l : String;
begin
 sResultDoc_l := LanceModele('','', VH_Doss.NoDossier, 'DOC' );
 if sResultDoc_l <> '##NONE##' then AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 10/09/2003
Modifié le ... : 10/09/2003
Description .. : Propriété d'un modèle de document
Mots clefs ... :
*****************************************************************}
procedure TOF_MODELESDOC.OnDblClick_FLISTE( Sender : TObject );
var
    SModeleId_l, SFileId_l : String;
    sResultDoc_l : String;

    SDocGuidModele,SFileGuidModele : String;
    SCheminFichier,SNomModele      : String;
    SExtModele,SCurGed,SArboGed    : String;
    SDocument,SDocGuid,SFileGuid   : String;
    InfoGed                        : TInfoDocGed;
    ParGed                         : TParamGedDoc;
    TNGed                          : TTreeNode;
    i                              : Integer;
begin
  if VarIsNull(GetField('YDO_DOCGUID')) then exit;

  //--- Affichage des paramètres d'un modèle
  if bModele_f then
  begin
     SModeleId_l := GetField('YDO_DOCGUID');
     SFileId_l := GetFileId(SModeleId_l);
     sResultDoc_l := LanceModele( SModeleId_l, SFileId_l, VH_Doss.NoDossier, 'DOC' );
     if sResultDoc_l <> '##NONE##' then
        AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
  end

  else
  //--- Utilisation d'un modèle
  begin
   //--- On extrait le modèle sur disque dans un répertoire temporaire
   SDocGuidModele  := GetField('YDO_DOCGUID');
   SFileGuidModele := GetFileId(SDocGuidModele);
   SNomModele := GetFileNameGed(SFileGuidModele);
   SExtModele := UpperCase ( ExtractFileExt (sNomModele) );
   SNomModele := Copy(SNomModele, 1, Length(SNomModele)-Length(SExtModele));
   ForceDirectories(V_GedFiles.TempPath+'BUREAU\GED\');
   // Construit un nom de fichier pour le modèle
   SCheminFichier := V_GedFiles.TempPath+'BUREAU\GED\'+SNomModele+SExtModele;
   i := 0;
   While (FileExists(SCheminFichier)) and (i<100) do
   begin
     i := i + 1;
     SCheminFichier := V_GedFiles.TempPath+'BUREAU\GED\'+SNomModele+IntToStr(i)+SExtModele;
   end;
   V_GedFiles.Extract(SCheminFichier, SFileGuidModele);

   //--- Fusion applicable uniquement pour les documents Word ou excel
   if (SExtModele='.DOC') or (SExtModele='.XLS') then
    begin
     // on fusionne avec le PARSER ...
     SDocument:=FusionneModele(VH_Doss.GuidPer,'DP','',SCheminFichier);
     DeleteFile(PChar(SCheminFichier));
     // RP 08/11/07 FQ 11797 : ne pas créer le doc ds la GED si annulation de la fusion
     // si échec ou annulation, on ressort
     if SDocument='' then exit;
{     // si échec de fusion, on utilisera directement le modèle extrait
     if SDocument='' then
       SDocument := SCheminFichier
     else
       DeleteFile(PChar(SCheminFichier)); }
    end
   //--- sinon le document est directement le modèle extrait
   else
     SDocument:=SCheminFichier;

   //--- Détermine L'armoire/Classeur/intercalaire en cours
   if (FMenuDisp.GedDP.TV.Selected<>Nil) then
    begin
     TNGed := FMenuDisp.GedDP.GetCurrentNodeGed (FMenuDisp.GedDP.TV.Selected);
     if (TNGed<>Nil) and (TNGed.Data<>Nil) then sCurGed := TGedDPNode(TNGed.Data).CodeGed;
    end
   else
    sCurGed := 'D'; // Par défaut Dossier Permanent
   SArboGed:=Copy (sCurGed,1,1);

   //--- Intégration du fichier fusionné dans la GED Système
   SFileGUID := V_GedFiles.Import(SDocument);
   DeleteFile(PChar(SDocument)); // suppression du document résiduel dans temp
   if SFileGUID='' then exit;

   //--- Intégration en tant que nouveau document dans la ged métier (donnera un docguid)
   ParGed.SDocGUID  := '';
   ParGed.SFileGUID := SFileGUID;
   ParGed.NoDossier := VH_Doss.NoDossier;
   ParGed.CodeGed   := sCurGed;

   //--- Définition des paramètres pour le document
   sResultDoc_l:=ShowNewDocument(ParGed, False,FALSE,FALSE,SArboGed);

   //--- Si on a annulé, on supprime dans la GED système
   if (sResultDoc_l='##NONE##') then
    begin
     V_GedFiles.Erase(SFileGUID);
     exit;
    end;

   sDocGuid := ReadTokenSt(sResultDoc_l); // élimine le NoDossier
   sDocGuid := ReadTokenSt(sResultDoc_l); // élimine le CodeGed
   sDocGuid := ReadTokenSt(sResultDoc_l); // lit le DocGuid

   //--- On a maintenant un SDocGuid à jour :
   // on va extraire pour modification le document dans C:\Documents and Settings\Le_User\Application Data\CEGID\GED
   InfoGed:=InitialiserInfoDocGed (VH_Doss.NoDossier,SDocGuid);
   ExtraireDocumentGed (InfoGed);
   OuvrirDocumentGed (InfoGed);

   // Fermeture du multi-critère
   Ecran.Close;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 12/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_MODELESDOC.OnClick_BPROPRIETE( Sender : TObject );
var
   SModeleId_l, SFileId_l : String;
   sResultDoc_l : String;
begin
   if VarIsNull(GetField('YDO_DOCGUID')) then exit;
   SModeleId_l := GetField('YDO_DOCGUID');
   SFileId_l := GetFileId( SModeleId_l );
   sResultDoc_l := LanceModele( SModeleId_l, SFileId_l, VH_Doss.NoDossier, 'DOC' );

   if sResultDoc_l <> '##NONE##' then
      AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 10/09/2003
Modifié le ... : 10/09/2003
Description .. : Suppression d'un modèle de document
Mots clefs ... :
*****************************************************************}
procedure TOF_MODELESDOC.OnClick_BDELETE( Sender : TObject );
var
   nModeleId_l : String;
begin
   if VarIsNull(GetField('YDO_DOCGUID')) then exit;
   if PGIAsk('Confirmez-vous la suppression de ce modèle?', Ecran.Caption) = mrNo then
      exit;

   nModeleId_l := GetField('YDO_DOCGUID');
   ExecuteSQL('DELETE FROM YDOCUMENTS WHERE YDO_DOCGUID = "' + nModeleId_l +'"');
   ExecuteSQL('DELETE FROM YDOCFILES WHERE YDF_DOCGUID = "' + nModeleId_l +'"');
   ExecuteSQL('DELETE FROM YMODELES WHERE YMO_DOCGUID = "' + nModeleId_l +'"');

   AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 12/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_MODELESDOC.LanceModele( SModeleId_p, SFileId_p : String; sNoDossier_p, sTypeGed_p : string ) : string;
var
   sResultDoc_l : String;
   ParGed: TParamGedDoc;
begin
   ParGed.SDocGUID := SModeleId_p;
   ParGed.SFileGUID := SFileId_p;
   ParGed.NoDossier := sNoDossier_p;
   ParGed.CodeGed := 'D';
   ParGed.TypeGed := sTypeGed_p;
   sResultDoc_l := ShowNewDocument(ParGed, False, true);
   result := sResultDoc_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 01/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOF_MODELESDOC.ListeTempAjouter( sDocument_p : string);
var
   nInd_l : integer;
begin
   if sDocument_p = '' then
      exit;
   nInd_l := Length(aosTemp_f);
   SetLength(aosTemp_f, nInd_l + 1);
   aosTemp_f[nInd_l] := sDocument_p;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 01/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOF_MODELESDOC.ListeTempSupprimer;
var
   nInd_l : integer;
begin
   for nInd_l := 0 to Length(aosTemp_f) - 1 do
   begin
      if (aosTemp_f[nInd_l] <> '') and FileExists(aosTemp_f[nInd_l]) then
         DeleteFile(Pchar (aosTemp_f[nInd_l]));
      aosTemp_f[nInd_l] := '';
   end;
   SetLength(aosTemp_f, 0);
end;

procedure TOF_MODELESDOC.Form_OnKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
 case Key of
  // Suppression : Ctrl + Suppr
  VK_DELETE :
       if (Shift = [ssCtrl]) and (TFMul(Ecran).FListe.Focused) and GetControl('BDELETE').Visible then
         OnClick_BDELETE (Nil);
  else
       TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
 end;
end;


//////////////////////////////////////////////////////////////////
Initialization
  registerclasses ( [ TOF_MODELESDOC ] ) ; 
end.

