{***********UNITE*************************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 14/01/2002
Modifié le ... : 24/01/2002
Description .. : Source TOF de la FICHE : ARTPHOTO ()
Mots clefs ... : TOF;ARTPHOTO;PHOTO;ARTICLE
*****************************************************************}
Unit ARTPHOTO_TOF ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,UTOB,
     FileCtrl,UtilArticle,UtilGC,ParamSoc,UtilDispGC,AGLInit,
     extctrls,EntGC,ed_tools;

Type
  TOF_ARTPHOTO = Class (TOF)
    public
      TOBInfoArt : TOB;
      procedure OnUpdate                 ; override ;
    private
      RepertoirePhotos : string;
      Extension : Integer;
      TOB_ProfilArt : TOB;
      BGenerique : Boolean;
      procedure SetLastError (Num : integer; ou : string );
      function  NB_Photos(var BigPhoto:boolean) : integer;
      function  AddBkSlash(S : String) : String;
      function  CreationDesArticles(NBTotalPhotos : integer) : boolean;
      function  CreerArticle(NomPhoto : string; EnSerie : boolean): boolean;
  end ;

function AttacheLaPhoto(LaTOBArticle:TOB; GA_ARTICLE,PathPhoto,NomPhoto:String; Ext:Integer):TOB;

const
	// libellés des messages
	TexteMessage: array[1..9] of string 	= (
          {1}        'Vous devez renseigner un répertoire contenant les photos des articles.'
          {2}        ,'Le répertoire de stockage des photos n''existe pas.'
          {3}        ,'Le répertoire ne contient pas de photos d''articles dans le format sélectionné.'
          {4}        ,'Vous devez renseigner un profil.'
          {5}        ,'Le profil n''existe pas.'
          {6}        ,'Le paramètre société "Photo pour les fiches" n''est pas renseigné.'
          {7}        ,'Vous devez renseigner un masque par défaut.'
          {8}        ,'Impossible de lancer le traitement car une ou plusieurs photos dépassent la taille maximum autorisée (> 1 Mo).'
          {9}        ,'Vous devez sélectionner "Génération automatique" du code article dans les paramètres société.'
                     );

Implementation

procedure TOF_ARTPHOTO.SetLastError (Num : integer; ou : string );
begin
  if ou<>'' then SetFocusControl(ou);
  LastError:=Num;
  LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
end ;

function TOF_ARTPHOTO.AddBkSlash(S : String) : String;
begin
  result := S;
  if S = '' then exit;
  if S[Length(S)] <> '\' then result := result + '\';
end;

function TOF_ARTPHOTO.NB_Photos(var BigPhoto:boolean) : integer;
var sr : TSearchRec;
    Ext : string;
begin
  Result:=0;
  if TRadioButton(GetControl('RBJPEG')).Checked then Extension:=1 else Extension:=2;
  if FindFirst(RepertoirePhotos+'*.*', faAnyFile, sr) = 0 then
    begin
    repeat
      Ext:=Uppercase(ExtractFileExt(sr.Name));
      if ((Extension=1) and ((Ext='.JPEG') OR (Ext='.JPG'))) OR
         ((Extension=2) and (Ext='.BMP')) then
         begin
         if (sr.Size > 1048576) then BigPhoto:=True;
         Inc(Result);
         end;
    until FindNext(sr) <> 0;
    FindClose(sr);
    end;
end;

procedure TOF_ARTPHOTO.OnUpdate ;
var CodeProfil : string;
    Q : TQuery;
    NBdePhotos : integer;
    GrossePhoto : boolean;
begin
  Inherited;
  if not GetParamSoc('SO_GCNUMARTAUTO') then begin SetLastError(9,''); exit; end;
  RepertoirePhotos := AddBkSlash(GetControlText('REPERTOIRE'));
  //Test l'existance du répertoire sélectionné
  if RepertoirePhotos='' then begin SetLastError(1,'REPERTOIRE'); exit; end;
  if not DirectoryExists(RepertoirePhotos) then begin SetLastError(2,'REPERTOIRE'); exit; end;
  //et que ce répertoire contient des photos
  GrossePhoto:=False; NBdePhotos:=NB_Photos(GrossePhoto);
  if GrossePhoto then begin SetLastError(8,'REPERTOIRE'); exit; end;
  if NBdePhotos=0 then begin SetLastError(3,'REPERTOIRE'); exit; end;
  //Test l'existance du profil
  CodeProfil := GetControlText('PROFILARTICLE');
  if CodeProfil='' then begin SetLastError(4,'PROFILARTICLE'); exit; end;
  //Vérifie qu'un type de photo est défini par défaut
  if VH_GC.GCPHOTOFICHE='' then begin SetLastError(6,'REPERTOIRE'); exit; end;
  BGenerique := TCheckBox(GetControl('RBGENERIQUE')).Checked;
  if BGenerique and (GetControlText('MASQUEDEFAUT')='') then
    begin SetLastError(7,'MASQUEDEFAUT'); exit; end;
  //Charge le profil dans une TOB
  Q := OpenSQL('SELECT * FROM PROFILART WHERE GPF_PROFILARTICLE="'+CodeProfil+'"',True,-1,'',true);
  try
    if not Q.Eof then
      begin
      TOB_ProfilArt := TOB.Create('PROFILART',nil,-1);
      TOB_ProfilArt.SelectDB('',Q);
      //Lance la création des articles à partir des photos
      if CreationDesArticles(NBdePhotos) then PGIInfo(TraduireMemoire('Traitement terminé.'),Ecran.Caption)
      else PGIInfo(TraduireMemoire('Le traitement a été interrompu.'),Ecran.Caption);
      end
    else SetLastError(5,'PROFILARTICLE');
  Finally
    Ferme(Q);
    TOB_ProfilArt.Free;
    end;
end ;

function TOF_ARTPHOTO.CreationDesArticles(NBTotalPhotos : integer) : boolean;
var CreationSerie : boolean;
    sr : TSearchRec;
    Ext : string;
    NbArt : integer;
begin
  NbArt := 1 ;
  CreationSerie:=TRadioButton(GetControl('RBSERIE')).Checked;

  if CreationSerie then InitMoveProgressForm(Ecran,Ecran.Caption,
     TraduireMemoire('Création en série d''articles à partir d''un répertoire de photos.'),
     NBTotalPhotos,True,True);

  if FindFirst(RepertoirePhotos+'*.*', faAnyFile, sr) = 0 then
    begin
    repeat
      Ext:=Uppercase(ExtractFileExt(sr.Name));
      if ((Extension=1) and ((Ext='.JPEG') OR (Ext='.JPG'))) OR
         ((Extension=2) and (Ext='.BMP')) then
             begin
             if CreationSerie then
               begin
               if Not MoveCurProgressForm(TraduireMemoire('Création de l''article '+IntToStr(NbArt)+
                  ' sur '+IntToStr(NBTotalPhotos))) then Break
               else CreerArticle(Sr.Name,CreationSerie);
               Inc(NbArt);
               end
             else CreerArticle(Sr.Name,CreationSerie);
             end;
    until (FindNext (sr) <> 0) ;
    FindClose (sr) ;
    end;
  if CreationSerie then FiniMoveProgressForm ;
  Result := True;
end;

function TOF_ARTPHOTO.CreerArticle(NomPhoto : string; EnSerie : boolean): boolean;
var TOB_Article : TOB;
    bOkChrono : boolean ;
    i, iRecupChrono : integer;
    CodeArticle, StArticle, NumChrono, stChronoArt, stSQL : string;
begin
  Result:=False;
  TOB_Article := TOB.Create('ARTICLE',nil,-1);
  with TOB_Article do
    begin
    //Libelle de l'article : nom de la photo
    PutValue('GA_LIBELLE',Copy(NomPhoto,1,Length(NomPhoto)-Length(ExtractFileExt(NomPhoto))));
    //Type de l'article
    PutValue('GA_TYPEARTICLE','MAR');
    if BGenerique then
      begin
      PutValue('GA_STATUTART','GEN');
      PutValue('GA_DIMMASQUE',GetControlText('MASQUEDEFAUT'));
      end
    else PutValue('GA_STATUTART','UNI');
    //Tout le reste
    PutValue('GA_COMPTAARTICLE',TOB_ProfilArt.GetValue('GPF_COMPTAARTICLE'));
    PutValue('GA_TENUESTOCK',TOB_ProfilArt.GetValue('GPF_TENUESTOCK'));
    PutValue('GA_LOT',TOB_ProfilArt.GetValue('GPF_LOT'));
    PutValue('GA_NUMEROSERIE',TOB_ProfilArt.GetValue('GPF_NUMEROSERIE'));
    PutValue('GA_CONTREMARQUE',TOB_ProfilArt.GetValue('GPF_CONTREMARQUE'));
    PutValue('GA_REMISEPIED',TOB_ProfilArt.GetValue('GPF_REMISEPIED'));
    PutValue('GA_REMISELIGNE',TOB_ProfilArt.GetValue('GPF_REMISELIGNE'));
    PutValue('GA_ESCOMPTABLE',TOB_ProfilArt.GetValue('GPF_ESCOMPTABLE'));
    PutValue('GA_FAMILLETAXE1',TOB_ProfilArt.GetValue('GPF_CODETAXE'));
    PutValue('GA_COMMISSIONNABLE',TOB_ProfilArt.GetValue('GPF_COMMISSIONNABL'));
    PutValue('GA_CALCPRIXHT',TOB_ProfilArt.GetValue('GPF_CALCPRIXHT'));
    PutValue('GA_CALCPRIXTTC',TOB_ProfilArt.GetValue('GPF_CALCPRIXTTC'));
    PutValue ('GA_COEFCALCHT',TOB_ProfilArt.GetValue('GPF_COEFCALCHT'));
    PutValue('GA_COEFCALCTTC',TOB_ProfilArt.GetValue('GPF_COEFCALCTTC'));
    PutValue('GA_CALCAUTOHT',TOB_ProfilArt.GetValue('GPF_CALCAUTOHT'));
    PutValue('GA_CALCAUTOTTC',TOB_ProfilArt.GetValue('GPF_CALCAUTOTTC'));
    PutValue('GA_TARIFARTICLE',TOB_ProfilArt.GetValue('GPF_TARIFARTICLE'));
    PutValue('GA_PAYSORIGINE',TOB_ProfilArt.GetValue('GPF_PAYSORIGINE'));
    PutValue('GA_COEFFG',TOB_ProfilArt.GetValue('GPF_COEFCALCPR'));
    PutValue('GA_DPRAUTO',TOB_ProfilArt.GetValue('GPF_CALCAUTOPR'));
    PutValue('GA_ARRONDIPRIX',TOB_ProfilArt.GetValue('GPF_ARRONDIPRIX'));
    PutValue('GA_ARRONDIPRIXTTC',TOB_ProfilArt.GetValue('GPF_ARRONDIPRIXTTC'));
    PutValue('GA_PRIXUNIQUE',TOB_ProfilArt.GetValue('GPF_PRIXUNIQUE'));
    PutValue('GA_COLLECTION',TOB_ProfilArt.GetValue('GPF_COLLECTION'));
    PutValue('GA_FOURNPRINC',TOB_ProfilArt.GetValue('GPF_FOURNPRINC'));
    For i:=1 to 3 do PutValue('GA_FAMILLENIV'+IntToStr(i),TOB_ProfilArt.GetValue('GPF_FAMILLENIV'+IntToStr(i)));
    For i:=1 to $A do PutValue('GA_LIBREART'+format('%x',[i]),TOB_ProfilArt.GetValue('GPF_LIBREART'+format('%x',[i])));
    For i:=1 to 3 do PutValue('GA_VALLIBRE'+IntToStr(i),TOB_ProfilArt.GetValue('GPF_VALLIBRE'+IntToStr(i)));
    For i:=1 to 3 do PutValue('GA_DATELIBRE'+IntToStr(i),TOB_ProfilArt.GetValue('GPF_DATELIBRE'+IntToStr(i)));
    For i:=1 to 3 do PutValue('GA_CHARLIBRE'+IntToStr(i),TOB_ProfilArt.GetValue('GPF_CHARLIBRE'+IntToStr(i)));
    For i:=1 to 3 do PutValue('GA_BOOLLIBRE'+IntToStr(i),TOB_ProfilArt.GetValue('GPF_BOOLLIBRE'+IntToStr(i)));
    PutValue('GA_DATESUPPRESSION',iDate2099);
    if EnSerie then
      begin
      // Récupération du chrono article : 20 tentatives avant message d'erreur
      iRecupChrono := 0 ;
      repeat
        inc( iRecupChrono ) ;
        stChronoArt := GetColonneSQL('PARAMSOC', 'SOC_DATA', 'SOC_NOM="SO_GCCOMPTEURART"') ;
        CodeArticle:=AttribNewCode('ARTICLE','GA_CODEARTICLE',GetParamsoc('SO_GCLGNUMART'),trim(GetParamsoc('SO_GCPREFIXEART')),stChronoArt,'',GetParamsoc('SO_GCCABARTICLE'));
        NumChrono:=ExtraitChronoCode(CodeArticle);
        stSql := 'update PARAMSOC set SOC_DATA="' + NumChrono +
                 '" where SOC_NOM="SO_GCCOMPTEURART" and SOC_DATA="' + stChronoArt + '"' ;
        bOkChrono := ( ExecuteSQL( stSQL ) = 1 ) ;
      until (( bOkChrono ) or ( iRecupChrono > 20 )) ;
      if Not bOkChrono then begin SetLastError(10,'GA_CODEARTICLE') ; Exit end ;

      PutValue('GA_CODEARTICLE', CodeArticle);
      StArticle := CodeArticleUnique2(CodeArticle,''); PutValue('GA_ARTICLE', StArticle);
      AttacheLaPhoto(TOB_ARTICLE,StArticle,RepertoirePhotos,NomPhoto,Extension);
      InsertDB(nil);
      Free;
      end
    else
      begin
      //Lance la fiche Article
      AddChampSupValeur('PathPhoto',RepertoirePhotos);
      AddChampSupValeur('NamePhoto',NomPhoto);
      AddChampSupValeur('Extension',Extension);
      TheTOB := TOB_Article;
      DispatchTTArticle(taCreat,'','TYPEARTICLE=MAR;ARTICLEPHOTO=X','');
      end;
    end;
end;

// JD
function AttacheLaPhoto(LaTOBArticle:TOB; GA_ARTICLE,PathPhoto,NomPhoto:String; Ext:Integer):TOB;
var S1 : TmemoryStream ;
    TOB_LIENSOLE : TOB;
    StImage, NomPhotoJpeg : string;
begin
  //Attache la photo
  TOB_LIENSOLE := TOB.Create('LIENSOLE',LaTOBArticle,-1);
  with TOB_LIENSOLE do
    begin
    PutValue('LO_TABLEBLOB','GA');
    //if Ext=1 then PutValue('LO_QUALIFIANTBLOB','PHJ')
    //else PutValue('LO_QUALIFIANTBLOB','PHO');
    PutValue('LO_QUALIFIANTBLOB','PHJ');
    PutValue('LO_EMPLOIBLOB',VH_GC.GCPHOTOFICHE);
    PutValue('LO_IDENTIFIANT',GA_ARTICLE);
    PutValue('LO_RANGBLOB',0);
    PutValue('LO_LIBELLE',NomPhoto);
    S1 := TMemoryStream.Create;
    try
      NomPhotoJpeg := NomPhoto;
      if Ext=2 then
      begin
        NomPhotoJpeg := Copy(NomPhoto,1,Length(NomPhoto)-3) + 'jpg';
        BitmapToJPeg(PathPhoto+NomPhoto,PathPhoto+NomPhotoJpeg);
      end;
      S1.LoadFromFile(PathPhoto+NomPhotoJpeg);
      SetLength(StImage,S1.Size);
      S1.Seek(0,0);
      //S.Read(St[1],S.Size);
      S1.Read(pchar(StImage)^,S1.Size);
      PutValue('LO_OBJET',StImage);
    Finally
      S1.Free;
      end;
    end;
  Result := TOB_LIENSOLE;
end;

Initialization
  registerclasses ( [ TOF_ARTPHOTO ] ) ;
end.
