{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 02/03/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
unit UComOutils;
//////////////////////////////////////////////////////////////////
interface
//////////////////////////////////////////////////////////////////
uses
   DB,
   Forms, UBob, HEnt1, sysutils, Hctrls, hMsgBox, UTOB, CBPPath,
   windows, Graphics, Controls, Classes, TypInfo, ParamSoc;


//////////////////////////////////////////////////////////////////

Type TArrayOfString = array of string;
type TTCloseAction = procedure (Sender : TObject; var FAction_p : TCloseAction) of object;
//////////////////////////////////////////////////////////////////
function  IntToStrCar(iInt_p, iLong_p : integer) : string;
function  StrToXCar(sChaine_p : string; iLong_p : integer; sChar_p : string = '0') : string;
function  MaxVersion(sRep_p, sFile_p, sExt_p : string; iCurVers_p : integer) : integer;

procedure VideFicBase(sCodeProduit_p : string);
function  ReadTokenInv(var sChaine_p : string; sSep_p : string) : string;
function  ExtractRacine(sRacine_p, sDirComp_p : string) : string;

procedure LargeurColonnes(gGrille_p : THGrid);
procedure LargeurUneColonne(gGrille_p : THGrid; iCol_p : integer);
function  NbColonnes(GListe_p : THGrid; sListeChamps_p : string) : integer;

procedure Rebranche(Controle_p : TControl; sEvent_p : string; neProc_p : TNotifyEvent); overload;
procedure Rebranche(Logo_p : THImage; sEvent_p : string; neProc_p : TNotifyEvent); overload;
procedure Rebranche(Fiche_p : TForm; sEvent_p : string; neProc_p : TTCloseAction); overload;

// *** Affichage icônes dans grilles
procedure WriteIconeFromImageList(GraphList_p : TImageList; ACanvas_p : TCanvas;
                                  ARect_p : TRect; iColWidth_p, iIcone_p : integer);

procedure WriteIconeFromBitMap(Bitmap_p : TBitmap; ACanvas_p : TCanvas;
                               ARect_p : TRect; iColWidth_p : integer);

procedure WriteIcone(AIco_p : TBitmap; ACanvas_p : TCanvas;
                     ARect_p : TRect; iWidth_p : integer);

function MyProcCalcMulGlobal(sFonction_p, sParams_p, sWhere_l : Hstring;
                       DS_p : TDataset; bTotal_p : Boolean) : Hstring ;


function InitRep(sRepTrav_p : string) : boolean;
function BoolToStrS(bBool_p : boolean) : string;

function GetWindowsDrive : String;
function BOBGetGoodRep(sCodeProd_p : string; bDDE_p : boolean) : string;
function JURGetEnvVar(sVariable_p : string; sValeurDefaut_p : string = '') : boolean;

// *** Gestion des groupes de Travail *** GHA 02/2008, FQ 149
// >> Cegid Expert Juridique
function  GererCritereGroupeConfTousJUR (FiltreDonnees:string='') : String;
function  GererCritereGroupeConfJUR (ComboGroupeConf: TControl; SansGroupeConf: Boolean; bFullGroups: Boolean=False; FiltreDonnees:string='') : String;
// >> Cegid Expert CFE
function  GererCritereGroupeConfTousCFE (FiltreDonnees:string='') : String;
function  GererCritereGroupeConfCFE (ComboGroupeConf: TControl; SansGroupeConf: Boolean; bFullGroups: Boolean=False; FiltreDonnees:string='') : String;



{$IFDEF EAGLCLIENT}
function TOBParamInit(sAction_p, sParam_p : string; bWait_p : boolean) : TOB;
{$ENDIF EAGLCLIENT}

//////////////////////////////////////////////////////////////////
implementation

//////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 25/01/2008
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
function JURGetEnvVar(sVariable_p : string; sValeurDefaut_p : string = '') : boolean;
var
   sValeur_l : string;
begin
   sValeur_l := Uppercase(GetEnvVar(sVariable_p));
   result := (sValeur_l = sValeurDefaut_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 15/04/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function  IntToStrCar(iInt_p, iLong_p : integer) : string;
var
   sChaine_l : string;
begin
   sChaine_l := IntToStr(iInt_p);

   while Length(sChaine_l) < iLong_p do
      sChaine_l := '0' + sChaine_l;
   result := sChaine_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 14/11/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function  StrToXCar(sChaine_p : string; iLong_p : integer; sChar_p : string = '0') : string;
begin
   while Length(sChaine_p) < iLong_p do
      sChaine_p := sChar_p + sChaine_p;
   result := sChaine_p;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 15/04/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function MaxVersion(sRep_p, sFile_p, sExt_p : string; iCurVers_p : integer) : integer;
var
   iFind_l, iFileVers_l : integer;
   sFile_l : string;
   srFile_l : TSearchRec;
begin
   SetCurrentDir(sRep_p);
   iFileVers_l := 0;
   iFind_l := FindFirst(sRep_p + '\' + sFile_p + '_*'  + sExt_p, faAnyFile, srFile_l);
   while iFind_l = 0 do
   begin
      sFile_l := srFile_l.Name;
      try
         iFileVers_l := StrToInt(Copy(srFile_l.Name, Length(sFile_p) + 2, 3));
      except
         result := iCurVers_p;
         sysutils.FindClose(srFile_l);
         exit;
      end;

      if (iFileVers_l > iCurVers_p ) then
         iCurVers_p := iFileVers_l;

      iFind_l := FindNext(srFile_l);
   end;

   sysutils.FindClose(srFile_l);
   result := iCurVers_p;
end;


{***********A.G.L.***********************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 10/11/2005
Modifié le ... :   /  /
Description .. : Nettoyage des fichiers en base
Mots clefs ... : 
*****************************************************************}
procedure VideFicBase(sCodeProduit_p : string);
var
   sWhere_l : string;
begin
   if sCodeProduit_p = 'CFE5' then
   begin
   sWhere_l := 'WHERE YFS_CODEPRODUIT IN ("EPSF", "CFE", "M0", "M2", "M3A", "M3B", "M4", "P0", "P2", "TNS")' +
               '  AND YFS_NODOSSIER = ""';
   end
   else
   begin
      sWhere_l := 'WHERE YFS_CODEPRODUIT = "' + sCodeProduit_p + '"' +
                  '  AND YFS_NODOSSIER = ""';

   end;
   if not ExisteSQL('SELECT * FROM YFILESTD ' + sWhere_l) then exit;

   ExecuteSQL('DELETE FROM NFILES WHERE EXISTS ' +
              ' (SELECT * FROM YFILESTD ' + sWhere_l +
              '    AND YFS_FILEGUID = NFI_FILEGUID)');

   ExecuteSQL('DELETE FROM NFILEPARTS WHERE EXISTS ' +
              ' (SELECT * FROM YFILESTD ' + sWhere_l +
              '    AND YFS_FILEGUID = NFS_FILEGUID)');
   ExecuteSQL('DELETE FROM YFILESTD ' + sWhere_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 10/11/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
{procedure MajDonnees(sRepBob_p : string);
var
   sSQL_l     : string;
   sRequete_l : string;
   sChemin_l  : string;
   sFicLigne_l : string;
   srSQL_l    : TSearchRec;
   iVersion_l : integer;
   iRet_l     : integer;
   tfSQL_l    : TextFile;
BEGIN
   sChemin_l := sRepBob_p + '\';
   iRet_l := FindFirst(sChemin_l + '*.SQL', faAnyFile, srSQL_l);
   while iRet_l = 0 do
   begin
      sSQL_l := srSQL_l.Name;
      iVersion_l := ValeurI(Copy(sSQL_l, 5, 4));

      if (iVersion_l > V_PGI.NumVersionSoc) and
      not ExisteSQL('SELECT YB_BOBNAME FROM YMYBOBS WHERE YB_BOBNAME = "' + sSQL_l + '"') then
      begin

         AssignFile(tfSQL_l, sChemin_l + srSQL_l.Name);
         Reset(tfSQL_l);
         Readln(tfSQL_l, sFicLigne_l);

         while not EOF(tfSQL_l) do
         begin
            sRequete_l := '';
            while not EOF(tfSQL_l) and (sFicLigne_l <> '###') do
            begin
               sRequete_l := sRequete_l + ' ' + sFicLigne_l;
               Readln(tfSQL_l, sFicLigne_l);
            end;
            if sRequete_l <> '' then
               ExecuteSQL(sRequete_l);
            Readln(tfSQL_l, sFicLigne_l);
         end;

         CloseFile(tfSQL_l);

         ExecuteSQL('INSERT INTO YMYBOBS ' +
                    '(YB_BOBNAME, YB_BOBLIBELLE, YB_BOBVERSION, YB_BOBDATECREAT, YB_BOBDATEMODIF) ' +
                    'VALUES ("' + sSQL_l + '", "' + sSQL_l + '", 1, "' +
                    USDateTime(Date) + '", "' + USDateTime(Date) + '")');
         end;
      iRet_l := FindNext(srSQL_l);
   end;
   sysutils.FindClose(srSQL_l);
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/06/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure LargeurColonnes(gGrille_p : THGrid);
var
   nRowInd_l, nColInd_l : integer;
   nColLargeurTmp_l : integer;
   tnColLargeur_l : array of integer;
//   tnColLargeur_l : array of integer;
begin
//   SetLength(tnColLargeurTmp_l, gGrille_p.ColCount);
   SetLength(tnColLargeur_l, gGrille_p.ColCount);

   for nRowInd_l := 0 to gGrille_p.RowCount - 1 do
   begin
      for nColInd_l := 0 to gGrille_p.ColCount - 1 do
      begin
         if (gGrille_p.ColWidths[nColInd_l] = -1) then
         begin
            tnColLargeur_l[nColInd_l] := -1;
         end
         else
         begin
            nColLargeurTmp_l := gGrille_p.Canvas.TextWidth(gGrille_p.Cells[ nColInd_l, nRowInd_l ]) + 15;
            if tnColLargeur_l[nColInd_l] < nColLargeurTmp_l then
               tnColLargeur_l[nColInd_l] := nColLargeurTmp_l;
{
            tnColLargeurTmp_l[nColInd_l] := gGrille_p.Canvas.TextWidth(gGrille_p.Cells[ nColInd_l, nRowInd_l ]) + 15;
            if tnColLargeur_l[nColInd_l] < tnColLargeurTmp_l[nColInd_l] then
               tnColLargeur_l[nColInd_l] := tnColLargeurTmp_l[nColInd_l];
}
         end;
      end;
   end;

   for nColInd_l := 0 to gGrille_p.ColCount - 1 do
   begin
      if tnColLargeur_l[nColInd_l] > 15 then
         gGrille_p.ColWidths[nColInd_l] := tnColLargeur_l[nColInd_l];
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 28/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure LargeurUneColonne(gGrille_p : THGrid; iCol_p : integer);
var
   nRowInd_l : integer;
   nColLargeurTmp_l : integer;
   iColLargeur_l : integer;
begin
   iColLargeur_l := gGrille_p.ColWidths[iCol_p];
   for nRowInd_l := 0 to gGrille_p.RowCount - 1 do
   begin
      if (gGrille_p.ColWidths[iCol_p] = -1) then
      begin
         iColLargeur_l := -1;
      end
      else
      begin
         nColLargeurTmp_l := gGrille_p.Canvas.TextWidth(gGrille_p.Cells[iCol_p, nRowInd_l ]) + 15;
         if iColLargeur_l < nColLargeurTmp_l then
            iColLargeur_l := nColLargeurTmp_l;
      end;
   end;

   gGrille_p.ColWidths[iCol_p] := iColLargeur_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/06/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function NbColonnes(GListe_p : THGrid; sListeChamps_p : string) : integer;
var
   iNbCol_l : integer;
   sChamp_l : string;
begin
   iNbCol_l := GListe_p.FixedCols;
   while sListeChamps_p <> '' do
   begin
      sChamp_l := READTOKENST(sListeChamps_p);
      GListe_p.Cells[iNbCol_l, 0] := sChamp_l;
      Inc(iNbCol_l);
   end;
   result := iNbCol_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 04/01/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure Rebranche(Controle_p : TControl; sEvent_p : string; neProc_p : TNotifyEvent);// overload;
begin
   SetMethodProp(Controle_p, sEvent_p, TMethod(neProc_p));
end ;

procedure Rebranche(Logo_p : THImage; sEvent_p : string; neProc_p : TNotifyEvent); //overload;
begin
   SetMethodProp(Logo_p, sEvent_p, TMethod(neProc_p));
end ;

procedure Rebranche(Fiche_p : TForm; sEvent_p : string; neProc_p : TTCloseAction); //overload;
begin
   SetMethodProp(Fiche_p, sEvent_p, TMethod(neProc_p));
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Meriaux
Créé le ...... : 24/10/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure WriteIconeFromImageList(GraphList_p : TImageList; ACanvas_p : TCanvas;
                                  ARect_p : TRect; iColWidth_p, iIcone_p : integer);
var
   AIco_l : TBitmap;
begin
   AIco_l := TBitmap.Create;
   GraphList_p.GetBitmap(iIcone_p, AIco_l);

   WriteIcone(AIco_l, ACanvas_p, ARect_p, iColWidth_p);

   AIco_l.Free;
end;

procedure WriteIconeFromBitMap(Bitmap_p : TBitmap; ACanvas_p : TCanvas;
                               ARect_p : TRect; iColWidth_p : integer);
begin
   WriteIcone(Bitmap_p, ACanvas_p, ARect_p, iColWidth_p);
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Meriaux
Créé le ...... : 24/10/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure WriteIcone(AIco_p : TBitmap; ACanvas_p : TCanvas;
                     ARect_p : TRect; iWidth_p : integer);
var

   iX_l, iY_l : integer;
begin
   AIco_p.Transparent := TRUE;
   if AIco_p.Width > iWidth_p then
      iX_l := ARect_p.left
   else
      iX_l := ARect_p.left + (iWidth_p - AIco_p.Width) div 2;

   if AIco_p.Height > (ARect_p.Bottom - ARect_p.Top) then
      iY_l := ARect_p.Top
   else
      iY_l := ARect_p.Top + ((ARect_p.Bottom - ARect_p.Top) - AIco_p.Height) div 2;

   ACanvas_p.Draw(iX_l, iY_l, AIco_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 02/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function ReadTokenInv(var sChaine_p : string; sSep_p : string) : string;
var
   iInd_l, iLg_l : integer;
   bOK_l : boolean;
   sResult_l : string;
begin
   sResult_l := sChaine_p;

   iLg_l := length(sChaine_p);
   iInd_l := iLg_l;
   bOK_l := false;
   while (iInd_l > 0) and not BOK_l do
   begin
      bOK_l := (sChaine_p[iInd_l] = sSep_p);
      if not bOK_l then
         Dec(iInd_l);
   end;
   if bOk_l then
   begin
      iLg_l := iLg_l - iInd_l;
      sResult_l := Copy(sChaine_p, iInd_l + 1 , iLg_l);
      sChaine_p := Copy(sChaine_p, 1, iInd_l - 1);
   end;

   result := sResult_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 08/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function  ExtractRacine(sRacine_p, sDirComp_p : string) : string;
var
   sPartage_l, sPartTmp_l : string;
   bDirOK_l : boolean;
begin
   // Racine c:\quadra ou \\serveur\quadra$
   if Copy(sDirComp_p, 1, 2) = '\\' then
   begin
      sPartage_l := '\\';
      sDirComp_p := Copy(sDirComp_p, 3, length(sDirComp_p) - 2);
   end;

   bDirOK_l := false;
   while not bDirOK_l and (sDirComp_p <> '') do
   begin
      sPartTmp_l := READTOKENPipe(sDirComp_p, '\');
      bDirOK_l := Pos(sRacine_p, sPartTmp_l) <> 0;
      sPartage_l := sPartage_l + sPartTmp_l + '\';
   end;

   result := sPartage_l;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 13/07/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function InitRep(sRepTrav_p : string) : boolean;
var
   bDroitsOK_l : boolean;
begin
   bDroitsOK_l := true;
   if not DirectoryExists(sRepTrav_p)   then
      bDroitsOK_l := ForceDirectories(sRepTrav_p);

   if not bDroitsOK_l then
      PGIError('Création du répertoire ' + sRepTrav_p + ' impossible.#13#10' +
               'Vérifier vos droits d''accès à ce répertoire.');

   Result := bDroitsOK_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/10/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function BoolToStrS(bBool_p : boolean) : string;
begin
   if bBool_p then
      result := 'X'
   else
      result := '-';
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 08/10/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetWindowsDrive: String;
const
   ciMaxWinPathLen_l = MAX_PATH + 1;
var
   lwLong_l : LongWord;
begin
   SetLength(Result, ciMaxWinPathLen_l);
   lwLong_l := GetWindowsDirectory(PChar(Result), ciMaxWinPathLen_l);
   if lwLong_l > 0 then
      SetLength(Result, lwLong_l)
   else
      Result := '';
   Result := IncludeTrailingPathDelimiter(ExtractFileDrive(Result)) ;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 16/11/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function MyProcCalcMulGlobal(sFonction_p, sParams_p, sWhere_l : Hstring;
                       DS_p :TDataset; bTotal_p : Boolean) : Hstring ;
var
   fPct_l : Extended;
   ////
   dtDateNais_l : TDatetime;
   wYear1_l, wMonth1_l, wDay1_l :Word;
   wYear2_l, wMonth2_l, wDay2_l :Word;
   wAge_l :Word;
   ////
   sDevise_l, sEtat_l : string;
   iCapLib_l : integer;
   bAppelPub_l : boolean;
begin
   result := '???';
   try
      // Affichage état selon capital libéré
      if sFonction_p = 'CAPLIB' then
      begin
         if sParams_p = 'ANN_CAPLIB' then
         begin
            sDevise_l := DS_p.FindField('ANN_CAPDEV').AsString;
            iCapLib_l := DS_p.FindField('ANN_CAPLIB').AsInteger;
            bAppelPub_l := (DS_p.FindField('JUR_APPELPUB').AsString = 'X');

            if (iCapLib_l >= 37000) and (sDevise_l = 'EUR') and not bAppelPub_l then
               sEtat_l := ''
            else if (iCapLib_l >= 225000) and (sDevise_l = 'EUR') and bAppelPub_l then
               sEtat_l := ''
            else
               sEtat_l := 'A vérifier';

            result := sEtat_l;
         end;
      end
      // Affichage icones
      else if sFonction_p = 'ICONE' then
      begin
         if DS_p.FindField(sParams_p) = nil then exit;
         sEtat_l := DS_p.FindField(sParams_p).AsString;
         if sEtat_l <> '' then
            result :='#ICO#' + IntToStr(98);
      end
      // Calcul de l'âge depuis une date
      else if sFonction_p = 'AGE' then
      begin
         if DS_p.FindField(sParams_p) = nil then exit;
         dtDateNais_l := DS_p.FindField(sParams_p).AsDateTime;

         if (dtDateNais_l = iDate1900) or (dtDateNais_l > Date) then
            exit;

         DecodeDate(dtDateNais_l, wYear1_l, wMonth1_l, wDay1_l);
         DecodeDate(Date, wYear2_l, wMonth2_l, wDay2_l);

         wAge_l := 12 * (wYear2_l - wYear1_l) + (wMonth2_l - wMonth1_l) ;
         if wDay2_l < wDay1_l then
            wAge_l := wAge_l - 1;

         Result := Format('%d', [wAge_l div 12]);
      end
      // Formatage du pourcentage depuis un float
      else if sFonction_p = 'PRCT' then
      begin
         if DS_p.FindField(sParams_p) = nil then exit;
         fPct_l := (DS_p.FindField(sParams_p).AsFloat) * 100;
         Result := FloatToStrF(fPct_l, ffFixed, 100, 2) + '%';
      end
   except
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/09/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function BOBGetGoodRep(sCodeProd_p : string; bDDE_p : boolean) : string;
var
   sRepBob_l : string;
begin
   {$IFNDEF EAGLSERVER}
   sRepBob_l := UpperCase(TcbpPath.GetCegidDistriBob);
   if Pos('\CEGID EXPERT\', sRepBob_l) = 0 then
   	insert('\CEGID EXPERT', sRepBob_l, Pos('\BOB', sRepBob_l));

   if sCodeProd_p <> '' then
      sRepBob_l := sRepBob_l + '\' + sCodeProd_p;
   ForceDirectories(sRepBob_l);
   {$ELSE}
   if bDDE_p then
      sRepBob_l := GetWindowsDrive + 'CWS\BOB\DDE'
   else
      sRepBob_l := GetWindowsDrive + 'CWS\BOB\' + sCodeProd_p;
   DDWriteln('');
   DDWriteln('Rep Bob : '+ sRepBob_l);
   DDWriteln('');
   {$ENDIF EAGLSERVER}

   result := sRepBob_l;
end;

{$IFDEF EAGLCLIENT}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 08/10/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOBParamInit(sAction_p, sParam_p : string; bWait_p : boolean) : TOB;
var
   OBParam_l : TOB;
begin
   OBParam_l := TOB.Create('le param', nil, -1);

   OBParam_l.AddChampSupValeur('ACTION', sAction_p);
   OBParam_l.AddChampSupValeur('PARAM', sParam_p);
   OBParam_l.AddChampSupValeur('WAIT', BoolToStrS(bWait_p));

   result := OBParam_l;
end;
{$ENDIF EAGLCLIENT}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G. Harlez
Créé le ...... : 14/02/2008
Modifié le ... : 14/02/2008
Description .. : Filtre sur les groupes de travail
Suite ........ : Cegid Expert Juridique
Mots clefs ... : FQ 149
*****************************************************************}
function GererCritereGroupeConfJUR(ComboGroupeConf: TControl;
  SansGroupeConf, bFullGroups: Boolean; FiltreDonnees: string): String;
var
    XxWhere, SValeur, SCritGroupes : String;
begin
  if FiltreDonnees = '' then
    XXWhere := 'NOT EXISTS (SELECT 1 FROM GRPDONNEES, LIENDOSGRP '
                         +'left join dossier on jur_guidperdos = dos_guidper '
				                 +'left join annuaire on jur_guidperdos = ann_guidper '
                         +'WHERE GRP_NOM = LDO_NOM AND GRP_NOM = "GROUPECONF" AND GRP_ID = LDO_GRPID '
                         + 'AND LDO_NODOSSIER =DOS_NODOSSIER)'
  else
    XXWhere := 'not exists (select 1 from DOSSIERGRP_FILTREDONNEES ' +
                           'where DOG_NODOSSIER=DOS_NODOSSIER and DOG_FILTREDONNEES="'+FiltreDonnees+'") ' ; //-LM20071008

  SCritGroupes := '';
  if (ComboGroupeConf<>nil) then
  begin
    //--- Uniquement les dossiers sans groupe : rien à modifier
    if (SansGroupeConf) then
    //--- Tous les groupes du user
    else if (THMultiValComboBox (ComboGroupeConf).Tous) then
    begin
       //--- Groupes auxquels appartient l'utilisateur + dossiers sans groupes
       if Not bFullGroups then
           XxWhere := GererCritereGroupeConfTousJUR (FiltreDonnees) //LM20071008
       //--- Non limité (vision de tous les dossiers même si l'utilisateur n'est dans aucun groupe !)
       else
           XxWhere := '';
    end
    //--- Uniquement les groupes sélectionnés par l'utilisateur
    else
    begin
       SValeur:=THMultiValComboBox (ComboGroupeConf).Text;
       while (SValeur<>'') do
       begin
           if SCritGroupes<>'' then SCritGroupes := SCritGroupes + ' OR ';
           if FiltreDonnees ='' then
             SCritGroupes:=SCritGroupes+'GRP_CODE="'+ReadToKenPipe (SValeur,';')+'"'
           else
             SCritGroupes:=SCritGroupes+'DOG_GROUPECONF="'+ReadToKenPipe (SValeur,';')+'"';
       end;
       if SCritGroupes<>'' then
       begin
           // XxWhere := '('+XxWhere+') OR ('+SCritGroupes+')'; => Non, nouveau comportement :
           // lorsqu'on a sélectionné des groupes, on ne voit plus les dossiers sans groupes
           if FiltreDonnees ='' then
             //XxWhere := 'EXISTS (SELECT 1 FROM DOSSIERGRP WHERE DOG_NODOSSIER=DOS_NODOSSIER AND ('+SCritGroupes+'))'
             // MD 14/11/2007 - Lisibilité => JOIN au lieu des WHERE
             // XxWhere := 'EXISTS (SELECT 1 FROM GRPDONNEES, LIENDOSGRP WHERE GRP_NOM = LDO_NOM AND GRP_NOM = "GROUPECONF" AND GRP_ID = LDO_GRPID '
             //                 + 'AND LDO_NODOSSIER=DOS_NODOSSIER AND ('+SCritGroupes+'))'
             XxWhere := 'EXISTS (SELECT 1 FROM GRPDONNEES '
                               +'LEFT JOIN LIENDOSGRP ON GRP_NOM = LDO_NOM AND GRP_ID = LDO_GRPID '
                               +'left join dossier on jur_guidperdos = dos_guidper '
				                       +'left join annuaire on jur_guidperdos = ann_guidper '
                               +'WHERE GRP_NOM = "GROUPECONF" AND LDO_NODOSSIER=DOS_NODOSSIER AND ('+SCritGroupes+'))'
           else
            XxWhere := 'exists (select 1 from DOSSIERGRP_FILTREDONNEES '+
                               'where DOG_NODOSSIER=DOS_NODOSSIER and ('+SCritGroupes+') '+
                               'and DOG_FILTREDONNEES="'+FiltreDonnees+'") '; //-LM20071008
       end
     end;
  end;

 if XxWhere<>'' then Result := '(' + XxWhere + ')'
 else Result := '';
end;
////////////////////////////////////////////////////////////////////////////////
function GererCritereGroupeConfTousJUR(FiltreDonnees: string): String;
begin
  Result := '';

  if FiltreDonnees='' then  //LM20071008
  begin
    // Dossier sans groupes
    if GetParamsocDpSecur('SO_MDDOSSANSGRP', True) then
      Result := 'NOT EXISTS (SELECT 1 FROM GRPDONNEES '
                           +'LEFT JOIN LIENDOSGRP ON GRP_NOM=LDO_NOM AND GRP_ID=LDO_GRPID '
                           +'left join dossier on jur_guidperdos = dos_guidper '
				                   +'left join annuaire on jur_guidperdos = ann_guidper '
                           +'WHERE GRP_NOM="GROUPECONF" AND LDO_NODOSSIER=DOS_NODOSSIER) OR ';
    // Dossiers dans des groupes auxquels l'utilisateur en cours a droit
    Result := Result
                  + 'EXISTS (SELECT 1 FROM GRPDONNEES '
                           +'LEFT JOIN LIENDOSGRP ON GRP_NOM=LDO_NOM AND GRP_ID=LDO_GRPID '
                           +'LEFT JOIN LIENDONNEES ON GRP_NOM=LND_NOM AND GRP_ID=LND_GRPID '
                           +'left join dossier on jur_guidperdos = dos_guidper '
				                   +'left join annuaire on jur_guidperdos = ann_guidper '
                           +'WHERE GRP_NOM="GROUPECONF" AND LDO_NODOSSIER=DOS_NODOSSIER AND LND_USERID="'+V_PGI.User+'")'
    end
  else
    Result := 'not exists (select 1 FROM DOSSIERGRP_FILTREDONNEES ' +
                          'where DOG_NODOSSIER=DOS_NODOSSIER  and DOG_FILTREDONNEES="'+ FiltreDonnees+'") ' +
              'or exists (select 1 from DOSSIERGRP_FILTREDONNEES, USERCONF_FILTREDONNEES '+
                         'where UCO_GROUPECONF = DOG_GROUPECONF '+
                         'and UCO_USER="'+ V_PGI.User+'" ' +
                         'and DOG_NODOSSIER=DOS_NODOSSIER '+
                         'and DOG_FILTREDONNEES="'+ FiltreDonnees+'") ' ; //LM20071008
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G. Harlez
Créé le ...... : 14/02/2008
Modifié le ... : 14/02/2008
Description .. : Filtre sur les groupes de travail
Suite ........ : Cegid Expert CFE
Mots clefs ... : FQ 149
*****************************************************************}
function GererCritereGroupeConfCFE(ComboGroupeConf: TControl;
  SansGroupeConf, bFullGroups: Boolean; FiltreDonnees: string): String;
var
    XxWhere, SValeur, SCritGroupes : String;
begin
  if FiltreDonnees = '' then
    XXWhere := 'NOT EXISTS (SELECT 1 FROM GRPDONNEES, LIENDOSGRP '
                         +'left join dossier on l02_guidperpere = dos_guidper '
				                 +'left join annuaire on l02_guidperpere = ann_guidper '
                         +'WHERE GRP_NOM = LDO_NOM AND GRP_NOM = "GROUPECONF" AND GRP_ID = LDO_GRPID '
                         + 'AND LDO_NODOSSIER =DOS_NODOSSIER)'
  else
    XXWhere := 'not exists (select 1 from DOSSIERGRP_FILTREDONNEES ' +
                           'where DOG_NODOSSIER=DOS_NODOSSIER and DOG_FILTREDONNEES="'+FiltreDonnees+'") ' ; //-LM20071008

  SCritGroupes := '';
  if (ComboGroupeConf<>nil) then
  begin
    //--- Uniquement les dossiers sans groupe : rien à modifier
    if (SansGroupeConf) then
    //--- Tous les groupes du user
    else if (THMultiValComboBox (ComboGroupeConf).Tous) then
    begin
       //--- Groupes auxquels appartient l'utilisateur + dossiers sans groupes
       if Not bFullGroups then
           XxWhere := GererCritereGroupeConfTousCFE (FiltreDonnees) //LM20071008
       //--- Non limité (vision de tous les dossiers même si l'utilisateur n'est dans aucun groupe !)
       else
           XxWhere := '';
    end
    //--- Uniquement les groupes sélectionnés par l'utilisateur
    else
    begin
       SValeur:=THMultiValComboBox (ComboGroupeConf).Text;
       while (SValeur<>'') do
       begin
           if SCritGroupes<>'' then SCritGroupes := SCritGroupes + ' OR ';
           if FiltreDonnees ='' then
             SCritGroupes:=SCritGroupes+'GRP_CODE="'+ReadToKenPipe (SValeur,';')+'"'
           else
             SCritGroupes:=SCritGroupes+'DOG_GROUPECONF="'+ReadToKenPipe (SValeur,';')+'"';
       end;
       if SCritGroupes<>'' then
       begin
           // XxWhere := '('+XxWhere+') OR ('+SCritGroupes+')'; => Non, nouveau comportement :
           // lorsqu'on a sélectionné des groupes, on ne voit plus les dossiers sans groupes
           if FiltreDonnees ='' then
             //XxWhere := 'EXISTS (SELECT 1 FROM DOSSIERGRP WHERE DOG_NODOSSIER=DOS_NODOSSIER AND ('+SCritGroupes+'))'
             // MD 14/11/2007 - Lisibilité => JOIN au lieu des WHERE
             // XxWhere := 'EXISTS (SELECT 1 FROM GRPDONNEES, LIENDOSGRP WHERE GRP_NOM = LDO_NOM AND GRP_NOM = "GROUPECONF" AND GRP_ID = LDO_GRPID '
             //                 + 'AND LDO_NODOSSIER=DOS_NODOSSIER AND ('+SCritGroupes+'))'
             XxWhere := 'EXISTS (SELECT 1 FROM GRPDONNEES '
                               +'LEFT JOIN LIENDOSGRP ON GRP_NOM = LDO_NOM AND GRP_ID = LDO_GRPID '
                               +'left join dossier on l02_guidperpere = dos_guidper '
				                       +'left join annuaire on l02_guidperpere = ann_guidper '
                               +'WHERE GRP_NOM = "GROUPECONF" AND LDO_NODOSSIER=DOS_NODOSSIER AND ('+SCritGroupes+'))'
           else
            XxWhere := 'exists (select 1 from DOSSIERGRP_FILTREDONNEES '+
                               'where DOG_NODOSSIER=DOS_NODOSSIER and ('+SCritGroupes+') '+
                               'and DOG_FILTREDONNEES="'+FiltreDonnees+'") '; //-LM20071008
       end
     end;
  end;

 if XxWhere<>'' then Result := '(' + XxWhere + ')'
 else Result := '';
end;
////////////////////////////////////////////////////////////////////////////////
function GererCritereGroupeConfTousCFE(FiltreDonnees: string): String;
begin
  Result := '';

  if FiltreDonnees='' then  //LM20071008
  begin
    // Dossier sans groupes
    if GetParamsocDpSecur('SO_MDDOSSANSGRP', True) then
      Result := 'NOT EXISTS (SELECT 1 FROM GRPDONNEES '
                           +'LEFT JOIN LIENDOSGRP ON GRP_NOM=LDO_NOM AND GRP_ID=LDO_GRPID '
                           +'left join dossier on l02_guidperpere = dos_guidper '
				                   +'left join annuaire on l02_guidperpere = ann_guidper '
                           +'WHERE GRP_NOM="GROUPECONF" AND LDO_NODOSSIER=DOS_NODOSSIER) OR ';
    // Dossiers dans des groupes auxquels l'utilisateur en cours a droit
    Result := Result
                  + 'EXISTS (SELECT 1 FROM GRPDONNEES '
                           +'LEFT JOIN LIENDOSGRP ON GRP_NOM=LDO_NOM AND GRP_ID=LDO_GRPID '
                           +'LEFT JOIN LIENDONNEES ON GRP_NOM=LND_NOM AND GRP_ID=LND_GRPID '
                           +'left join dossier on l02_guidperpere = dos_guidper '
				                   +'left join annuaire on l02_guidperpere = ann_guidper '
                           +'WHERE GRP_NOM="GROUPECONF" AND LDO_NODOSSIER=DOS_NODOSSIER AND LND_USERID="'+V_PGI.User+'")'
    end
  else
    Result := 'not exists (select 1 FROM DOSSIERGRP_FILTREDONNEES ' +
                          'where DOG_NODOSSIER=DOS_NODOSSIER  and DOG_FILTREDONNEES="'+ FiltreDonnees+'") ' +
              'or exists (select 1 from DOSSIERGRP_FILTREDONNEES, USERCONF_FILTREDONNEES '+
                         'where UCO_GROUPECONF = DOG_GROUPECONF '+
                         'and UCO_USER="'+ V_PGI.User+'" ' +
                         'and DOG_NODOSSIER=DOS_NODOSSIER '+
                         'and DOG_FILTREDONNEES="'+ FiltreDonnees+'") ' ; //LM20071008
end;
////////////////////////////////////////////////////////////////////////////////
end.
