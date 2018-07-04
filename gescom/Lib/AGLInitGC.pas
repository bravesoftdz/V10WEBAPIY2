unit AGLInitGC;
{==========================
Publication de fonctions et de procédures pour les scripts
===========================}

interface
uses {$IFDEF VER150} variants,{$ENDIF} HCtrls,M3VM ,
{$IFDEF EAGLCLIENT}
     etablette,
{$ELSE}
     tablette, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     Hent1 ,classes,AGLinit, EntGC, Jpeg,lookup, extctrls, ParamSoc ;

Function  DispatchRecherche (G_Code : THCritMaskEdit; Num : Integer;
                             stWhere, stRange, stMul : string ) : string ;
Function  DispatchRechArt (G_Code : THCritMaskEdit; Num : Integer;
                             stWhere, stRange, stMul : string ) : string ;

{ 'Paramétrage' accès fiche article }
function  NatureFicheArticle: String;
function  NatureMulArticle: String;
function  NomFicheArticle: String;
function  NomMulArticle: String;
function  NomMulRechArticle: String;
function  AglNatureFicheArticle(Parms: Array of variant; Nb: Integer ): Variant;
function  AglNomFicheArticle(Parms: Array of variant; Nb: Integer ): Variant;
function  AglNatureMulArticle(Parms: Array of variant; Nb: Integer ): Variant;
function  AglNomMulArticle(Parms: Array of variant; Nb: Integer ): Variant;
{$IFDEF GPAO}
function  AGL_VH_GP(Parms: Array of variant; Nb: Integer ): Variant;
{$ENDIF}
// DBR : pour récupérer dans la recherche de tiers, les natures de tiers liées à une nature de pièce
function AGLLesNaturesTiers ( parms: array of variant; nb: integer ) : Variant;


implementation

uses HMsgBox,M3FP, SysUtils,
{$IFDEF EAGLCLIENT}
     eMul,eFiche,eFichList,MaineAGL,
{$ELSE}
     DB,Fiche, Mul, FichList,Fe_Main,
{$ENDIF}
{$IFDEF GPAO}
	  EntGP,
{$ENDIF}
{$IFDEF BTP}
    UtilVariables,
{$ENDIF}
     Forms,
     utom, utof, utob, UtomArticle ,stdctrls, TiersUtil, HDimension, UtilPGI, UtilArticle, UtomLiensOle,UtilRessource;

procedure AGLLanceJustifTiers ( parms: array of variant; nb: integer ) ;
begin
// LanceJustifTiers(String(parms[0])) ;
end;

procedure AGLAppliqueProfilArticle( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_Article) then TOM_Article(OM).AppliqueProfilArticle(Parms[1],Parms[2]) else exit;
end;

procedure AGLVerifModifDimension( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_Article) then TOM_Article(OM).VerifModifDimension else exit;
end;

procedure AGLEditArtLie( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_Article) then TOM_Article(OM).EditArtLie(Parms[1]) else exit;
end;

function AGLSelectArticle( parms: array of variant; nb: integer ) : variant ;
begin
  result:='' ;
end;

function AGLExisteGRC( parms: array of variant; nb: integer ) : variant ;
begin
result:=False;
if (ctxGRC in V_PGI.PGIContexte) then
  result:=True ;
end;

//Retourne le code article unique à partir du code générique
Function AGLCodeArticleUnique2( parms: array of variant; nb: integer ): variant;
Begin
   result:=CodeArticleUnique2(string (Parms [1]),string (Parms [2]));
End;

procedure AGLInitRechArticle( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     Arguments ,critere: string;
     ChampMul,ValMul : string;
      CC : TComponent;
      x:integer;
begin
F:=TForm(Longint(Parms[0])) ;
if not (F is TFMul) then exit;
Arguments:=string(Parms[1]);
Repeat
 Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
 if Critere<>'' then
    begin
    x:=pos('=',Critere);
    if x<>0 then
       begin
       ChampMul:=copy(Critere,1,x-1);
       ValMul:=copy(Critere,x+1,length(Critere));
       end;
       CC:=TFMul(F).FindComponent(ChampMul);
       If (CC is THEdit) then THEdit(CC).Text:=ValMul;
       If (CC is THValComboBox) then THValComboBox(CC).Value:=ValMul;
       if (CC is TCheckBox) then
           begin
           if (ValMul='X') then TCheckBox(CC).state:=cbChecked;
           if (ValMul='-') then TCheckBox(CC).state:=cbUnChecked;
           if (ValMul=' ') then TCheckBox(CC).state:=cbGrayed;
           end;
       if (CC is THLabel) then THLabel(CC).caption:=ValMul;
    end;
until  Critere='';

end;

procedure AGLInitMulDimArticle( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     Arguments ,critere: string;
     ChampMul,ValMul : string;
     CodeArticle,DimMAsque,LibelleArticle : string;
      CC : TComponent;
      x,i:integer;
      C1:THValComboBox;
      C2:THLabel;
      QMasque : TQuery ;
begin
F:=TForm(Longint(Parms[0])) ;
if not (F is TFMul) then exit;
Arguments:=string(Parms[1]);
Repeat
 Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
 if Critere<>'' then
    begin
    x:=pos('=',Critere);
    if x<>0 then
       begin
       ChampMul:=copy(Critere,1,x-1);
       ValMul:=copy(Critere,x+1,length(Critere));
       end;
       CC:=TFMul(F).FindComponent(ChampMul);
       If (CC is THEdit) then THEdit(CC).Text:=ValMul;
       If (CC is THValComboBox) then THValComboBox(CC).Value:=ValMul;
       if (CC is TCheckBox) then
           begin
           if (ValMul='X') then TCheckBox(CC).state:=cbChecked;
           if (ValMul='-') then TCheckBox(CC).state:=cbUnChecked;
           if (ValMul=' ') then TCheckBox(CC).state:=cbGrayed;
           end;
       if (CC is THLabel) then THLabel(CC).caption:=ValMul;

       if ChampMul='GA_DIMMASQUE' then DimMasque := ValMul;
       if ChampMul='GA_LIBELLE' then LibelleArticle := ValMul;
    end;
until  Critere='';

CodeArticle:= THEdit(TFMul(F).FindComponent('GA_CODEARTICLE')).Text;
TFMul(F).caption:='Dimensions de l''article '+CodeArticle+' '+LibelleArticle;

If (DimMasque<>'') then
    BEGIN
    if ctxMode in V_PGI.PGIContexte
    then QMasque:=OpenSQL('SELECT GDM_TYPE1,GDM_TYPE2,GDM_TYPE3,GDM_TYPE4,GDM_TYPE5 FROM DIMMASQUE WHERE GDM_MASQUE="'+DimMasque+'" and GDM_TYPEMASQUE="'+VH_GC.BOTypeMasque_Defaut+'"',TRUE,-1,'',true)
    else QMasque:=OpenSQL('SELECT GDM_TYPE1,GDM_TYPE2,GDM_TYPE3,GDM_TYPE4,GDM_TYPE5 FROM DIMMASQUE WHERE GDM_MASQUE="'+DimMasque+'"',TRUE,-1,'',true) ;
    if not QMasque.EOF then
        begin
        for i:=1 to MaxDimension do
            begin
            C1:=THValComboBox(TFMul(F).FindComponent('GA_CODEDIM'+InttoStr(i)));
            C2:=THLabel(TFMul(F).FindComponent('LGRILLE'+InttoStr(i)));
            if (QMasque.Findfield('GDM_TYPE'+IntToStr(i)).AsString<>'') then
                begin
                C1.Visible:=TRUE;
                C1.Plus:='GDI_TYPEDIM="DI'+InttoStr(i)+'" AND GDI_GRILLEDIM="'+QMasque.FindField('GDM_TYPE'+IntToStr(i)).AsString+'" ORDER BY GDI_RANG' ;
                C1.DataType:='GCDIMENSION' ;
                C2.Visible:=TRUE;
                C2.Caption:= RechDom('GCGRILLEDIM'+IntToStr(i),QMasque.Findfield('GDM_TYPE'+IntToStr(i)).AsString,FALSE) ;
                //TFMul(F).ChangeTitres('DIM'+IntToStr(i),C2.Caption);
                end
            else begin
                C1.Visible:=FALSE;
                C2.Visible:=FALSE;
                end;

            end;
        end;
    Ferme(QMasque);
    end;
end;

procedure AGLInitMulDepot( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     Arguments ,critere: string;
     ChampMul,ValMul : string;
     CodeArticle,ArticleCaption,LibelleArticle : string;
     sCaption : string;
      CC : TComponent;
      x:integer;
      sJoker : string ;
begin
F:=TForm(Longint(Parms[0])) ;
if not (F is TFMul) then exit;
Arguments:=string(Parms[1]);
sJoker:='';
sCaption:='';
Repeat
 Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
 if Critere<>'' then
    begin
    x:=pos('=',Critere);
    if x<>0 then
       begin
       ChampMul:=copy(Critere,1,x-1);
       ValMul:=copy(Critere,x+1,length(Critere));
       end;
    if  ChampMul='GQ_ARTICLE' then
        begin
        CodeArticle:=ValMul;
        ArticleCaption:=Valmul;
        end
    else if ChampMul='LIBELLE' then
            LibelleArticle:=Valmul
         else if copy(ChampMul,1,6)='LIBDIM' then
                begin
                if ValMul<>'' then sCaption:=sCaption+' '+Valmul
                end
              else
                begin
                if valmul<>'' then
                    begin
                    CodeArticle:=Codearticle+sJoker+ValMul;
                    sJoker:='';
                    end
                else
                    sJoker:=sJoker+'___';
                end;
    end; // critere <> ''
until  Critere='';

CC:=TFMul(F).FindComponent('GQ_ARTICLE');
If (CC is THEdit) then THEdit(CC).Text:=CodeArticle;
TFMul(F).caption:=ArticleCaption+' '+LibelleArticle+' '+sCaption;
end;

procedure AGLDispatchTT (parms: array of variant; nb: integer ) ;
begin
if Assigned(V_PGI.DispatchTT) then V_PGI.DispatchTT(Parms[0],StringToAction(parms[1]),Parms[2],Parms[3],Parms[4]) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Paul Chapuis
Créé le ...... : 30/03/2000
Modifié le ... :   /  /
Description .. : Fonction de recherche générique
Mots clefs ... : RECHERCHE
*****************************************************************}
{---------
Appel depuis Script avec G_Code = nom du controle
----------}
Function  DispatchRecherche (G_Code : THCritMaskEdit; Num : Integer;stWhere, stRange, stMul : string ) : string ;
Begin
    Result := '';
    Case Num of
        1: begin { ARTICLE }
           if stMul = '' then
               begin
{$IFDEF BTP}
               stMul := 'BTARTICLE_RECH';
{$ELSE}
{$IFDEF CHR}
               stMul := 'HRPRESTATIONS_MUL';
{$ELSE}
               if (ctxAffaire in V_PGI.PGIContexte) then stMul := 'AFARTICLE_RECH'
               else if (ctxMode in V_PGI.PGIContexte) then stMul := '*'
               else if (ctxGPAO in V_PGI.PGIContexte) then stMul := NomMulRechArticle
               else stMul := 'GCARTICLE_RECH';
{$ENDIF}
{$ENDIF}
               end;
           Result := GetArticleRecherche_Disp (G_Code, stWhere, stRange, stMul);
           end;
        2: begin { TIERS }
           Result := GetTiersRecherche (G_Code, stWhere, stRange, stMul);
           end;
        3: begin { RESSOURCE }   //GM
           Result := GetRessourceRecherche (G_Code, stWhere, stRange, stMul);
           end;
        4: begin { ARTICLE par code article GA_CODEARTICLE  }
           if stMul = '' then
               begin
               if (ctxAffaire in V_PGI.PGIContexte) then stMul := 'AFARTICLE_RECH'
               else stMul := 'GCARTICLE_RECH';
               end;
           Result := GetArticleRecherche_Disp (G_Code, stWhere, stRange, stMul,True);
           end;
        5: Begin {VARIABLE POUR METRES} //FV
             Result := GetVariableRecherche (G_Code, stWhere, stRange, stMul,True);
           end;
    end ;
End;

Function  DispatchRechArt (G_Code : THCritMaskEdit; Num : Integer;
                             stWhere, stRange, stMul : string ) : string ;
Begin
    Result := '';
    Case Num of
        1: begin { ARTICLE }
           if stMul = '' then
               begin
{$IFDEF BTP}
               stMul := 'BTARTICLE_RECH';
{$ELSE}
               if (ctxAffaire in V_PGI.PGIContexte) then stMul := 'AFARTICLE_RECH'
               else if (ctxMode in V_PGI.PGIContexte) then stMul := '*' //AC 13/09/01
               else stMul := 'GCARTICLE_RECH';
{$ENDIF}
               end;
           Result := GetRechArticle (G_Code, stWhere, stRange, stMul);
           end;
    end ;
End;

Function AGLDispatchRecherche (parms: array of variant; nb: integer ) : variant ;
var G_Code : THCritMaskEdit;
    F : TForm;
begin
F := TForm (Longint (Parms [0]));
G_Code := THCritMaskEdit (F.FindComponent (string (Parms [1])));
Result := DispatchRecherche (G_Code, integer (Parms [2]), string (Parms [3]),
                             string (Parms [4]), string (Parms [5]));
{if Assigned(V_PGI.DispatchTT) then V_PGI. Result:=DispatchRecherche(Parms[0],Parms[1]) ; }
end;


////////// provisoire ///////////////////////////////////
function AGLOuvreFicheMul( parms: array of variant; nb: integer ) : variant ;
{Var F : TForm ;}
begin
  result:='' ;
  {F:=TForm(Longint(Parms[0])) ;
  if F is TFMul then TheMulQ:=TQuery(F.FindComponent('Q')) else }
  TheMulQ:=Nil ;
  result:=AGLLanceFiche(string(parms[1]),string(parms[2]),string(parms[3]),string(parms[4]),string(parms[5]));
end;


////////// OT : Dupliquer un article ///////////////////////////////////
procedure AGLDuplication( parms: array of variant; nb: integer );
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_Article) then TOM_Article(OM).Duplication('') else exit;
end;

/////////////////// THGrid: récupération d'une valeur ////////////
function AGLGridGetCell( parms: array of variant; nb: integer ):variant ;
var  F : TForm ;
     NomGrid :string ;
     Col :Integer ;
     CC :THGrid ;
begin
result:='' ;
F:=TForm(Longint(Parms[0])) ;
NomGrid :=string(Parms[1]) ;
Col:=StrToInt(Parms[2]) ;
CC:=THGrid(F.FindComponent(NomGrid)) ;
Result:=CC.Cells[Col,CC.ROW] ;
end;

///////////////////// Affectation d'une référence à un article de l'article ////////////
procedure AGLAffectationReference(parms:array of variant; nb: integer) ;
var F :TForm ;
    OM :TOM ;
Begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFicheListe) then OM:=TFFicheListe(F).OM else exit ;
if (OM is TOM_Catalogu) then TOM_Catalogu(OM).AffectationReference else exit ;
End;

////////////////// Gestion du code article pour le catalogue AC ////////////////
procedure AGLCodeUnique( parms: array of variant; nb: integer );
var F :TForm ;
    OM :TOM ;
Begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit ;
if (OM is TOM_Catalogu) then TOM_Catalogu(OM).CodeUnique else exit;
End;



{***********A.G.L.***********************************************
Auteur  ...... : AV
Créé le ...... : 26/01/2001
Modifié le ... : 26/01/2001
Description .. : FONCTION A NE PAS UTILISER
Suite ........ : ... en cours de suppression ...
Suite ........ :
Suite ........ : Remplacée par AGLGETPARAMSOC
Mots clefs ... :
*****************************************************************}
function AGLParamSoc ( parms: array of variant; nb: integer ) : variant ;
var ch : string;
begin
ch := string(Parms[1]) ;
if ch='SO_GCMULTIDEPOTS' then
result := VH_GC.GCMultiDepots
else result := '';
end;


////// Récupération du context pour script des fiches ///////
Function AGLIsGestionInterne( parms: array of variant; nb: integer ):Variant;
Begin   // 1=GI SCOT,2=GA TEMPO...
Result := 0;
if ctxAffaire in V_PGI.PGIContexte then Result := 2;
If ctxScot in V_PGI.PGIContexte then Result:=1;
End;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 28/06/2002
Description .. : Retourne la nature de la fiche article
*****************************************************************}
function NatureFicheArticle: String; 
begin
{$IFDEF GPAO}
  if VH_GP.FicheArticleGP then
    Result := 'W'
  else
    Result := 'GC';
{$ELSE}
  {$IFDEF BTP}
  Result := 'BTP';
  {$ELSE}
    {$IFDEF GIGI}
     Result :='AFF'
     {$ELSE}
       {$IFDEF GPAOLIGHT}
        if VH_GC.OASeria then
          Result := 'W'
        else
          Result := 'GC';
       {$ELSE GPAOLIGHT}
        Result := 'GC';
       {$ENDIF GPAOLIGHT}
     {$ENDIF}
  {$ENDIF}
{$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 28/06/2002
Description .. : Retourne la nature du mul article
*****************************************************************}
function NatureMulArticle: String;
begin
{$IFDEF GPAO}
  if VH_GP.FicheArticleGP then
    Result := 'W'
  else
    Result := 'GC';
{$ELSE}
  {$IFDEF BTP}
  Result := 'BTP';
  {$ELSE}
    {$IFDEF GIGI}
     Result :='AFF'
     {$ELSE}
       {$IFDEF GPAOLIGHT}
        if VH_GC.OASeria then
          Result := 'W'
        else
          Result := 'GC';
       {$ELSE GPAOLIGHT}
        Result := 'GC';
       {$ENDIF GPAOLIGHT}
     {$ENDIF}
  {$ENDIF}
{$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 28/06/2002
Description .. : Retourne le nom de la fiche article
*****************************************************************}
function NomFicheArticle: String; 
begin
{$IFDEF GPAO}
  if VH_GP.FicheArticleGP then
    Result := 'WARTICLE_FIC'
  else
    Result := 'GCARTICLE';
{$ELSE}
  {$IFDEF BTP}
  Result := 'BTARTICLE';
  {$ELSE}
    {$IFDEF GIGI}
     Result :='AFARTICLE'
     {$ELSE}
       {$IFDEF GPAOLIGHT}
        if VH_GC.OASeria then
          Result := 'WARTICLE_FIC'
        else
          Result := 'GCARTICLE';
       {$ELSE GPAOLIGHT}
        Result := 'GCARTICLE';
       {$ENDIF GPAOLIGHT}
     {$ENDIF}
  {$ENDIF}
{$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 28/06/2002
Description .. : Retourne le nom du mul article
*****************************************************************}
function NomMulArticle: String; 
begin
  { Mul par défaut }
  Result := 'GCARTICLE_MUL';
{$IFDEF GPAO}
  if VH_GP.FicheArticleGP then
    Result := 'WARTICLE_MUL';
{$ELSE}
  {$IFDEF GPAOLIGHT}
  if VH_GC.OASeria then
    Result := 'WARTICLE_MUL'
  else
    Result := 'GCARTICLE_MUL';
  {$ELSE GPAOLIGHT}
  Result := 'GCARTICLE_MUL';
  {$ENDIF GPAOLIGHT}
{$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 28/06/2002
Description .. : Retourne le nom du mul de recherche d'un article
*****************************************************************}
function NomMulRechArticle: String; 
begin
  { Mul par défaut }
  Result := 'GCARTICLE_RECH';
{$IFDEF GPAO}
  if VH_GP.FicheArticleGP then
    Result := 'WARTICLE_MUL';
{$ELSE}
  {$IFDEF GPAOLIGHT}
  if VH_GC.OASeria then
    Result := 'WARTICLE_MUL'
  else
    Result := 'GCARTICLE_RECH'
  {$ELSE GPAOLIGHT}
  Result := 'GCARTICLE_RECH';
  {$ENDIF GPAOLIGHT}
{$ENDIF}
end;

function AglNatureFicheArticle(Parms: Array of variant; Nb: Integer ): Variant;
begin
 	Result := NatureFicheArticle;
end;

function AglNomFicheArticle(Parms: Array of variant; Nb: Integer ): Variant;  
begin
 	Result := NomFicheArticle;
end;

function AglNatureMulArticle(Parms: Array of variant; Nb: Integer ): Variant;
begin
 	Result := NatureMulArticle;
end;

function AglNomMulArticle(Parms: Array of variant; Nb: Integer ): Variant;  
begin
	Result := NomMulArticle;
end;



{$IFDEF GPAO}

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 28/06/2002
Modifié le ... :   /  /
Description .. : 'Accès' à VH_GP depuis le script
Mots clefs ... :
*****************************************************************}
function AGL_VH_GP(Parms: Array of variant; Nb: Integer ): Variant;
begin
	if Parms[0] = 'FicheArticleGP' then Result := VH_GP.FicheArticleGP
   else
      Result := null;
end;
{$ENDIF}

procedure AGLSetImage( parms: array of variant; nb: integer ) ;
Var I : TImage ;
    Vide : Boolean ;
    FileName : String ;
    Q : TDataSet ;
    F : TForm ;
{$IFDEF EAGLCLIENT}
    s : TStringStream ;
{$ELSE}
    s : TmemoryStream ;
{$ENDIF}
begin
  F:=TForm(Longint(Parms[0])) ;
  I:=TImage(F.FindComponent(String(Parms[1]))) ;
  Vide:=Boolean(Parms[2]) ;
  if i=Nil then exit ;
  Q:=TFFicheListe(F).Ta ;
  if vide then
     BEGIN
     if Q.State in [dsInsert,dsEdit] then else Q.Edit ;
     I.picture.graphic:=Nil ;
     ENd else
     BEGIN
     if GetFileName(tfOpenBMP,'*.BMP',FileName) then
        BEGIN
        if Q.State in [dsInsert,dsEdit] then else Q.Edit ;
        I.Picture.LoadFromFile(Filename) ;
{$IFDEF EAGLCLIENT}
        s:=TStringStream.Create('');
{$ELSE}
        s:=TMemoryStream.Create;
{$ENDIF}
        I.picture.graphic.savetostream(s) ;
        TOM_LiensOle(TFFicheListe(F).OM).ChargeStream (s) ;
        S.Free;
        ENd ;
     END ;
End;

procedure AGLSetJpegImage( parms: array of variant; nb: integer ) ;
Var I : TImage ;
    Vide : Boolean ;
    FileName : String ;
    Q : TDataSet ;
    F : TForm ;
{$IFDEF EAGLCLIENT}
    s : TStringStream ;
{$ELSE}
    s : TmemoryStream ;
{$ENDIF}
    Scale : integer ;
begin
  F:=TForm(Longint(Parms[0])) ;
  I:=TImage(F.FindComponent(String(Parms[1]))) ;
  Vide:=Boolean(Parms[2]) ;
  Scale:=strToInt(Parms[3]) ;
  if i=Nil then exit ;
  Q:=TFFicheListe(F).Ta ;
  if vide then
     BEGIN
     if Q.State in [dsInsert,dsEdit] then else Q.Edit ;
     I.picture.graphic:=Nil ;
     ENd else
     BEGIN
     if GetFileName(tfOpenBMP,'*.JPG;*.JPEG',FileName) then
        BEGIN
        if Q.State in [dsInsert,dsEdit] then else Q.Edit ;
        I.Picture.LoadFromFile(Filename) ;
        SetJPEGOptions (I,Scale) ;
{$IFDEF EAGLCLIENT}
        s:=TStringStream.Create('');
{$ELSE}
        s:=TMemoryStream.Create;
{$ENDIF}
        I.picture.graphic.savetostream(s) ;
        TOM_LiensOle(TFFicheListe(F).OM).ChargeStream (s) ;
        S.Free;
        ENd ;
     END ;
End;

procedure AGLSetJpegOptions( parms: array of variant; nb: integer ) ;
Var I : TImage ;
    F : TForm ;
    Scale : integer ;
begin
  F:=TForm(Longint(Parms[0])) ;
  I:=TImage(F.FindComponent(String(Parms[1]))) ;
  Scale:=strToInt(Parms[2]) ;
  if i=Nil then exit ;
  SetJPEGOptions (I,Scale) ;
end;

function NatureTiersPourNaturePiece (stNaturePiece : string) : string;
var stNatureTiers : string;
begin
  stNatureTiers := GetInfoParPiece (stNaturePiece, 'GPP_NATURETIERS');
  Result := '';
  while stNatureTiers <> '' do
  begin
    if (Result <> '') then Result := Result + ',';
    Result := Result + '"' + ReadTokenst (stNatureTiers) + '"';
  end;
end;

function AGLLesNaturesTiers ( parms: array of variant; nb: integer ) : Variant;
var stNaturePiece : string;
begin
  stNaturePiece := Parms[1];

  if (stNaturePiece = 'ZZ1') then // Facture + avoir
  begin
  Result := NatureTiersPourNaturePiece ('FAC') + ',' +
              NatureTiersPourNaturePiece ('AVS') + ',' +
              NatureTiersPourNaturePiece ('AVC');
  end else
  begin
  Result := NatureTiersPourNaturePiece (stNaturePiece);
  end;
end;

//////////////////////////////////////////////////////////////////////////////
procedure initM3Gescom();
begin
(*
{$IFNDEF SANSCOMPTA}
RegisterAglProc( 'ParamVentil', FALSE , 5, AGLParamVentil);
{$ENDIF}
*)
RegisterAglProc( 'LanceJustifTiers', FALSE , 1, AGLLanceJustifTiers);
RegisterAglProc( 'AppliqueProfilArticle', TRUE , 2, AGLAppliqueProfilArticle);
RegisterAglProc( 'VerifModifDimension', TRUE , 0, AGLVerifModifDimension);
RegisterAglProc( 'EditArtLie', TRUE , 1, AGLEditArtLie);
RegisterAglFunc( 'SelectArticle', FALSE , 1, AGLSelectArticle);
RegisterAglFunc( 'ExisteGRC', FALSE , 0, AGLExisteGRC );
RegisterAglFunc( 'CodeArticleUnique2',True,1,AGLCodeArticleUnique2);
RegisterAglProc( 'InitRechArticle', TRUE , 1, AGLInitRechArticle);
RegisterAglProc( 'InitMulDimArticle', TRUE , 1, AGLInitMulDimArticle);
RegisterAglProc( 'InitMulDepot', TRUE , 1, AGLInitMulDepot);
{RegisterAglProc( 'InitMulDispo_Dim', TRUE , 1, AGLInitMulDispo_Dim);}
{RegisterAglProc( 'InitGCArticle_Mul', TRUE , 0, AGLInitGCArticle_Mul); }
RegisterAglProc( 'DispatchTT', FALSE , 4, AGLDispatchTT);
RegisterAglFunc( 'DispatchRecherche', TRUE, 5, AGLDispatchRecherche);
RegisterAglProc('AffectationReference',True,0,AGLAffectationReference) ; // Catalogue AC
RegisterAglProc('CodeUnique',True,0,AGLCodeUnique) ; // Catalogue AC

RegisterAglProc ( 'Duplication', TRUE , 0, AGLDuplication);
RegisterAglFunc( 'GridGetCell', True,2,AglGridGetCell) ;

RegisterAglFunc( 'OuvreFicheMul', TRUE , 5, AGLOuvreFicheMul );
RegisterAglFunc( 'AGLParamSoc', TRUE, 1, AGLParamSoc );


RegisterAglFunc( 'AFIsGestionInterne',False,0,AGLIsGestionInterne);

{ GPAO V1.0 Thierry }
RegisterAglFunc('AGLNatureFicheArticle', False, 1, AGLNatureFicheArticle);
RegisterAglFunc('AGLNomFicheArticle', False, 1, AGLNomFicheArticle);
RegisterAglFunc('AGLNatureMulArticle', False, 1, AGLNatureMulArticle);
RegisterAglFunc('AGLNomMulArticle', False, 1, AGLNomMulArticle);

{$IFDEF GPAO }
RegisterAglFunc('AGL_VH_GP', False, 1, Agl_VH_GP);
{$ENDIF}
RegisterAglProc( 'SetImage', TRUE , 2, AGLSetImage); //JD
RegisterAglProc( 'SetJpegImage', TRUE , 2, AGLSetJpegImage);
RegisterAglProc( 'SetJpegOptions', TRUE , 2, AGLSetJpegOptions);
RegisterAglFunc ('LesNaturesTiers', TRUE , 2, AGLLesNaturesTiers); // DBR
end;

Initialization
initM3Gescom();
end.
