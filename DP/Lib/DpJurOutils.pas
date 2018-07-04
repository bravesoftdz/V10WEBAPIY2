{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 10/12/2002
Modifié le ... : 28/02/2003 : Compatibilité CWAS
Description .. : Unité commune entre le DP et JURI
Mots clefs ... : DP;JURI
*****************************************************************}
unit DpJurOutils;

interface

uses
   hCtrls,
   {$IFDEF EAGLCLIENT}
   eMul, MaineAgl, MailOL, AGLUtilOle,
   {$ELSE}
   {$IFDEF EAGLSERVER}

   {$ELSE}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
   db, Mul, Fe_main, MailOL, AGLUtilOle,
   {$ENDIF}
   {$ENDIF}
{$IFDEF VER150}
   Variants,
{$ENDIF}
   HDB, Hqry, forms, classes, UTOB, UCoherence,
   SysUtils, hmsgbox, controls, HStatus, HEnt1,
   comctrls, RtfCounter, ParamSoc,
   uTom, uTof, UGedFiles;


/////////// POINTEURS DE FONCTIONS ET PROCEDURES ////////////
type PTOB = ^TOB;
type TPP_TVI_I  = function (tvParams_p : array of variant; nParamsNb_p : integer) : integer;
//type TPP_OBTVI_I  = function (OBTheTOB_p : PTOB; tvParams_p : array of variant; nParamsNb_p : integer) : integer;

type TPF_QSTV_I = function ( qQuery_p : THQuery; sData_p : string; tvVal_p : Array Of Variant) : integer;
type TPDataset = ^TDataset;
/////////// ENTETES DE FONCTION ET PROCEDURES ////////////
{$IFNDEF EAGLSERVER}
function  AjoutOutlook( nature, sGuidEvt_p : String) : integer;

function  GrilleParcours( // dbgGrille_p : THMyGrid;
                          // qQuery_p : THQuery;
                          frmMul : TFMul;
                          sDataVal_p : string;
                          nDataMax_p : integer;
                          funcInit_p : TPF_QSTV_I;
                          funcTraitement_p : TPP_TVI_I;
                          bFinSelect : boolean = true) : integer;

procedure ModeEdition(DS: TDataset);

//function SupprimeMvt(parms : array of variant; nb : integer) : integer;

function  InitAjoutOutlook( qQuery_p : THQuery;
                          sData_p : string ;
                          var tvVal_p : Array Of Variant) : integer;

function  InitSetValChamp( qQuery_p : THQuery;
                          sData_p : string ;
                          var tvVal_p : Array Of Variant) : integer;

function  InitGetValChamp( qQuery_p : THQuery;
                          sData_p : string ;
                          var tvVal_p : Array Of Variant) : integer;

function  InitPurgeDoc( qQuery_p : THQuery;
                          sData_p : string ;
                          var tvVal_p : Array Of Variant) : integer;
function  PurgeDoc( sGuidEvt_p, sJevCodeOp_p, sJurNomDos_p : string; TheTOB_p : TOB; var bErreur_p : boolean) : string;

function  MontreOnglet( sEcran_p : TForm; sOnglet_p : string ) : boolean;

function  OnElipsisClick( Domaine, ListeSel, QuoiSel, SQL, ChampRet1, ChampRet2 : string;
                          var ValChamp: variant ) : string; overload;
function  OnElipsisClick( Domaine, ListeSel, QuoiSel, SQL, ChampRet1, ChampRet2, ChampRet3 : string;
                          var ValChamp1, ValChamp2 : variant ) : string; overload;

function  CalculNumCle (NomChamp,NomTable:string; strWhere:string='') : integer; // $$$ JP 05/05/04: ajout condition where
function  DupliquerElement(sTable_p, sCle1_p, sCle2_p : string) : string;
function  GetNomCompPer(sGuidPer_p : string) : string;
function  GetCapital(sGuidPer_p : string; var sDevise_p : string;
                     var dCapital_p : double; var iCapNbTitre_p : integer) : boolean;

function  GetNomPer(sGuidPer_p : string) : string;
function  GetNomAvecCivOuForme(sGuidPer_p : string) : string;
function  GetLienPer(sGuidPer_p : string) : string;
function  GetNomDos(ValChamp : string) : string;
function  GetLibOp(ValChampOp,ValChampDos : string) : string;
function  GetEmail(sGuidPer_p : string) : string;
function  RendFormeGen(Forme: String) : String;
function  GetValChamp(Table,Champ,Cle : string) : variant;
function  GetValsChamps(sTable_p, sLesChamps_p, sCle_p : string) : string;


procedure SetValChamp( sTable_p : string;
                       tsCle_p, tsCleVal_p : array of string;
                       sChamp_p, sChampVal_p : string);

function  PlusHeure( dtHeure_p : TDateTime; nInterval_p : integer; sTypeInterval_p : string;
                    var nReliqJour_p : integer ) : TDateTime; overload;

function  PlusHeure( dtDateEtHeure_p : TDateTime; nInterval_p : integer;
                    sTypeInterval_p : string ) : TDateTime; overLoad;

function  JURBoolToStr(Val : boolean) : string;
function  JURBoolToInt(Val : boolean) : integer;
function  JURStrToBool(Val : string) : boolean;
function  JURVarToBool(Val : variant) : boolean;

function  AffecteNomDocMaquette(NomDoc : string) : string;
function  ReadTokenParam( sParametre_p, sListeParametres_p : string ): string;

function  ControleCarInterdits( sCode_p, sCarInterdits_p, sExceptions_p : string ) : string;
function  ControleCode( sCode_p, sExceptions_p : string ) : string;

procedure LanceWord(NomDoc,DocPerso : string);

function  CodeSiren( sSiren_p : string ) : string;
function  DecodeSiren( sSiren_p : string ) : string;
function  TraduitBoolean(Val : string) : string;
function  TraduitDate(Val : string) : TDateTime;
function  TypeNombre(typeCh : string) : boolean;
function  TypeDate(typeCh : string) : boolean;
function  TTToChampLib(sTablette_p : string) : string;

function  FiltreTOB( OBSource_p : TOB; aosNomChamp_p, aosValeur_p : array of string) : TOB;
function  CommencePar(sSubStr_p, sSource_p : string) : boolean;
{$endif}
{$ifndef EAGL}
procedure  MaxChamp( sTable_p, sChamp_p : string; var vValeur_p : variant; sWhere_p : string = '' ); overload;
procedure  MaxChamp( sTable_p, sChamp_p : string; var nValeur_p : integer; sWhere_p : string = '' ); overload;
procedure  MaxChamp( sTable_p, sChamp_p : string; var bValeur_p : boolean; sWhere_p : string = '' ); overload;
procedure  MaxChamp( sTable_p, sChamp_p : string; var dValeur_p : TDateTime; sWhere_p : string = '' ); overload;
procedure  MaxChamp( sTable_p, sChamp_p : string; var sValeur_p : string; sWhere_p : string = '' ); overload;
{$else}
{$endif}
function  MaxChampX(sTable_p, sChamp_p : string; sWhere_p : string = '') : variant;
function  MaxCodeDossier : integer;
function  FindFirstUpper( OBSource_p : TOB; aosNomChamp_p, aosValeur_p: array of string) : TOB;
// BM 31/03/2006 : déplacé dans UCoherence.pas
//procedure Verifcoherence(sGuidPerDos_p, sCodedos_p : String);

/////////////////////////////////////////////////////////////////////////////
function  ChampType(sChamp_p : string) : string;
function  ChampTypeVarchar(sChamp_p : string) : boolean;


function  ChangeFormeJuridique(LaTom_p : Tom; sMajForce_p : string; PDataset_p : TPDataset) : boolean ; overload;
function  ChangeFormeJuridique(var sCodeInsee_p, sForme_p, sFormePrive_p, sFormeSte_p, sFormeSci_p,
                               sDORFormeAsso_p, sCoop_p, sDFIRegleFisc_p, sDORSectionBnc_p,
                               sPPPM_p, sTypeSCI_p : string) : boolean ; overload;

function  ChangeCodeNaf(LaTom:Tom; Ctrl:TControl; codeNafIni:string) : boolean ; //LM20070516
procedure AfficheInfoLienPersonne(LaTom:TObject; GuidPer, ZoneGuidPer, ZoneLibelle: string); //LM20070703
procedure cElipsisPerAssClick(LaTom:Tom; Sender: TObject); //LM20070516
function RecupNomGerant(GuidPerDos : String; var RefGerant:string): String;   //de TomOrga... LM20070516

function VerrouilleSiBaseComptable(Ecran : TOM; guid : string ) : boolean ;  //LM20070620

/////////////////////////////////////////////////////////////////////////////
const
  cstCarInterdits_g : string = '&"''(-_)=~#{[|`\^@]}°+$ù*,;:!¨£%µ?./§¤<>';


/////////// IMPLEMENTATION ////////////
implementation

uses
{$IFNDEF EAGLSERVER}
//LM20070516 uses
   UtilMulTraitmt , //LM20070516 UtilMulTraitmt ;
{$endif}
  uSatUtil,
  utoolsrtf,
  AnnOutils,
  dpTOFAnnusel,
  eventDecla,
  stdctrls
;



{$IFNDEF EAGLSERVER}
{*****************************************************************
Auteur ....... : MD
Date ......... : ??/??/??
Procédure .... : AjoutOutlook
Description .. :
Paramètres ... : Nature de l'évènement
                 Numéro de l'évènement
*****************************************************************}
function AjoutOutlook(nature, sGuidEvt_p : String) : integer;
// #### Msg ne serait-il pas mieux en Var dans les paramètres ?
var
  bAlerte : boolean;
  Duree : TDateTime;
  dDate_l : TDate;
  tHeure_l : TTime;
  NbMinutes : TTimeStamp;
  Msg, NomPer, Sujet, Lieu, BlocNote, Email, Txt : string;
{$IFDEF VER150}
  Contenu : HTStringList;
{$ELSE}
  Contenu : TStringList;
{$ENDIF}
  TEvt : TOB;
begin
   result := 0;
   TEvt := TOB.Create('JUEVENEMENT',nil,-1);
   if TEvt.SelectDB('"' + sGuidEvt_p + '"', nil) then
   begin
      NomPer := GetNomCompPer(TEvt.GetValue('JEV_GUIDPER'));
      Sujet := NomPer + ' - ' + TEvt.GetValue('JEV_EVTLIBELLE');
      BlocNote := TEvt.GetValue('JEV_NOTEEVT');
      if (Nature='REU') or (Nature='ACT') then // $$$ JP 23/09/03: prendre en compte également la famille "ACT"ivité
      begin
         bAlerte := TEvt.GetValue('JEV_ALERTE')='X';
         Duree := TDateTime(TEvt.GetValue('JEV_DATEFIN'))-TDateTime(TEvt.GetValue('JEV_DATE'));
         NbMinutes := DateTimeToTimeStamp(Duree);
         Lieu := TEvt.GetValue('JEV_LIEU');
         dDate_l := TEvt.GetValue('JEV_DATE');
         tHeure_l := TEvt.GetValue('JEV_DATE');

         AddRDV( Sujet,
                 Lieu,
                 BlocNote,
                 StrToDateTime( FormatDateTime('dd/mm/yyyy 00:00:00', dDate_l)),
                 StrToDateTime( FormatDateTime('30/12/1899 hh:nn:ss', tHeure_l)),
                 Trunc(NbMinutes.Time/1000/60),
                 bAlerte);
         Msg := 'dans le calendrier.';
         result := 1;
      end
      else if Nature='MSG' then // gestion d'un mail
      begin
         // String to TStrings
{$IFDEF VER150}
         Contenu := HTStringList.Create;
{$ELSE}
         Contenu := TStringList.Create;
{$ENDIF}
         // #### en attendant que GetRTFStringText renvoie St au lieu de vide
         Txt := GetRTFStringText(BlocNote);
         if Txt<>'' then
           Contenu.Text := Txt
         else
           Contenu.Text := BlocNote;
         Email := TEvt.GetValue('JEV_EMAIL');
         SendMail(Sujet,Email,'',Contenu,'');
         Contenu.Free;
         Msg := 'dans la boite d''envoi.';
      end
      else  // gestion d'une tache
      begin
         AddTache( Sujet,
                   BlocNote,
                   TDateTime(TEvt.GetValue('JEV_DATE')),
                   TDateTime(TEvt.GetValue('JEV_DATEFIN')));
         Msg := 'dans le gestionnaire des tâches.';
         result := 1;
      end;
    end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 10/12/02
Procédure .... : GrilleParcours
Description .. : Parcours la grille et effectue le traitement défini à l'appel
Paramètres ... : La grille
                 Le query de la fiche
                 La liste des données
                 Le nombre de données
                 Pointeur vers la fonction d'initialisation
                 Pointeur vers la procédure de traitement
*****************************************************************}

function GrilleParcours(  // dbgGrille_p : THMyGrid;
                          // qQuery_p : THQuery;
                          frmMul : TFMul;
                          sDataVal_p : string;
                          nDataMax_p : integer;
                          funcInit_p : TPF_QSTV_I;
                          funcTraitement_p : TPP_TVI_I;
                          bFinSelect : boolean = true) : integer; //overload;
// MD 01/06/04 : rajout frmMul qui est le mul appelant, pour fetchlestous en eagl
var
   nRow_l, nValMax_l, nElemNb_l, nValRetour_l : integer;
   tvVal_l : array of variant;
   dbgGrille_l : THDBGrid; // en eagl, casté en thgrid par Hdb
   qQuery_l : THQuery;
begin
   Result := 0;
   if frmMul=Nil then exit;
   dbgGrille_l := frmMul.FListe;
   qQuery_l := frmMul.Q;

   nElemNb_l := 0;
   SetLength( tvVal_l, nDataMax_p);

   if dbgGrille_l.AllSelected then
   begin
{$IFDEF EAGLCLIENT}
      if not frmMul.FetchLesTous then
         PGIInfo('Impossible de récupérer tous les enregistrements')
      else
{$ENDIF}
      begin
        // Toutes les lignes sélectionnées
        InitMove( qQuery_l.RecordCount, '' );
        qQuery_l.First;
        for nRow_l := 0 to qQuery_l.RecordCount - 1 do
        begin
           MoveCur(False);
           nValMax_l := funcInit_p( qQuery_l, sDataVal_p, tvVal_l );
           nValRetour_l := funcTraitement_p( tvVal_l, nValMax_l );
           nElemNb_l := nElemNb_l + nValRetour_l;
           qQuery_l.Next;
        end;
      end;
      FiniMove;
   end
   else
   begin
      if dbgGrille_l.NbSelected <> 0 then
      begin
         // Quelques lignes sélectionnées
         InitMove( dbgGrille_l.NbSelected, '' );
         for nRow_l := 0 to dbgGrille_l.NbSelected - 1 do
         begin
            MoveCur(False);
            dbgGrille_l.GotoLeBookmark( nRow_l );
            {$IFDEF EAGLCLIENT}
            qQuery_l.TQ.Seek(dbgGrille_l.Row - 1) ;
            {$ENDIF}
            nValMax_l := funcInit_p( qQuery_l, sDataVal_p, tvVal_l );
            nValRetour_l := funcTraitement_p( tvVal_l, nValMax_l );
            nElemNb_l := nElemNb_l + nValRetour_l;
         end;
         FiniMove;
      end
      else
         PGIInfo( 'Aucuns élément sélectionné');
   end;
   // déselectionne
   if bFinSelect then
      FinTraitmtMul(frmMul);

   result := nElemNb_l;

end;

{*****************************************************************
Auteur ....... : MD
Date ......... : ??/??/??
Procédure .... : ModeEdition
Description .. : Préparation pour l'appel de la procédure AGLSetValChamp
Paramètres ... : Le query de la fiche
*****************************************************************}
procedure ModeEdition(DS: TDataset);
begin
   if DS = Nil then
      exit;
   if not (DS.State in [dsInsert, dsEdit]) then
      DS.Edit;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 19/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function FiltreTOB( OBSource_p : TOB; aosNomChamp_p, aosValeur_p : array of string) : TOB;
var
   bSourceOK_l : boolean;
   nSourceInd_l, nSubInd_l : integer;
   sSourceVal_l, sSubVal_l : string;
   OBFin_l, OBTemp_l : TOB;
begin
   OBFin_l := TOB.Create('', nil, -1);
   for nSourceInd_l := 0 to OBSource_p.Detail.Count - 1 do
   begin
      bSourceOK_l := true;
      nSubInd_l := 0;
      while bSourceOK_l and (nSubInd_l < Length(aosNomChamp_p)) do
      begin
          sSubVal_l := UpperCase(aosValeur_p[nSubInd_l]);
          sSourceVal_l := UpperCase(OBSource_p.Detail[nSourceInd_l].GetValue(aosNomChamp_p[nSubInd_l]));
          if not CommencePar(sSubVal_l, sSourceVal_l) then
             bSourceOK_l := false;
          Inc(nSubInd_l);
      end;
      //
      if bSourceOK_l then
      begin
         OBTemp_l := TOB.Create('', OBFin_l, OBFin_l.Detail.Count);
         OBTemp_l.Dupliquer(OBSource_p.Detail[nSourceInd_l], true, true);
      end;
   end;
   result := OBFin_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 19/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function CommencePar(sSubStr_p, sSource_p : string) : boolean;
var
   nSubInd_l : integer;
   bCommencePar_l : boolean;
begin
    nSubInd_l := 1;
    bCommencePar_l := true;
    if (sSource_p = #0) then
      sSource_p := '';
      
    if (sSubStr_p = '') and (sSource_p <> '') then
      bCommencePar_l := false
    else if (sSubStr_p <> '') and (sSource_p = '') then
      bCommencePar_l := false;

    while bCommencePar_l and (nSubInd_l <= Length(sSubStr_p)) do
    begin
       bCommencePar_l := sSubStr_p[nSubInd_l] = sSource_p[nSubInd_l];
       Inc(nSubInd_l);
    end;
    result := bCommencePar_l;
end;


{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 01/01/2000
Modifié le ... : 27/06/2006
Description .. : DupliquerElement
Suite ........ : NCX 20/10/00 Remplacement '' par " pour duplication click 
Suite ........ : droit
Suite ........ : BM 27/06/2006 : QUERY -> TOB pour compatibilité eAGL
Mots clefs ... : TABLE, CLE1, CLE2
*****************************************************************}
function DupliquerElement(sTable_p, sCle1_p, sCle2_p : string) : string;
var
  OBOld_l : TOB;   //    , OBNew_l, sWhereNew_l
  sChamp_l, sWhereOld_l, sGuidNew_l : string;
begin
   result := '';

   // pour Table=='JUBIBACTION' ou 'JUBIBRUBRIQUE' , voir JuriOutils.pas
   if sTable_p = 'JUEVENEMENT' then
   begin
      sChamp_l := 'JEV_GUIDEVT';
   end
   else if sTable_p = 'ANNUAIRE' then
   begin
      sChamp_l := 'ANN_GUIDPER';
   end
   else
      exit;

   sGuidNew_l := AglGetGuid();
   sWhereOld_l := sChamp_l + ' = "' + sCle1_p + '"';
//   sWhereNew_l := sChamp_l + ' = "' + sGuidNew_l + '"';

   OBOld_l := TOB.Create(sTable_p, nil, -1);
   OBOld_l.LoadDetailDBFromSQL(sTable_p,
                               'select * from ' + sTable_p +
                               ' where ' + sWhereOld_l);

   if OBOld_l.Detail.Count = 0 then
   begin
      OBOld_l.Free;
      exit;
   end;

   OBOld_l.Detail[0].PutValue(sChamp_l, sGuidNew_l);
   OBOld_l.Detail[0].SetAllModifie(true);
   OBOld_l.Detail[0].SetAllCleWithDateModif(true);
   OBOld_l.Detail[0].InsertDB(nil);
{   OBNew_l := TOB.Create(sTable_p, nil, -1);
//   OBNew_l := OpenSQL('select * from ' + sTable_p + ' where ' + sWhereNew_l);
   OBNew_l.InitValeurs(true);
   OBNew_l.Dupliquer(OBOld_l, true, true, true);
   // nouvelle clé
   OBNew_l.Detail[0].PutValue(sChamp_l, sGuidNew_l);
   OBNew_l.InsertDB(nil);
   OBNew_l.Free;}

   OBOld_l.Free;

   Result := sGuidNew_l;
end;


{*****************************************************************
Auteur ....... : BM
Date ......... : 10/12/02
sFonction_c ..... : InitAjoutOutlook
Description .. : Préparation pour l'appel de la fonction AGLAjoutOutlook
Paramètres ... : Le query de la fiche
                 La liste des données
                 Le tableau de variant à renseigner avec les données :
                 - 0 à n  : clés de l'enregistrement
Renvoie ...... : Le nombre de données
*****************************************************************}
function InitAjoutOutlook( qQuery_p : THQuery; sData_p : string ; var tvVal_p : Array Of Variant) : integer;
var
   nValInd_l : integer;
   sData_l, sValeur_l : string;
begin
   // Clé de l'enregistrement
   sData_l := ReadTokenSt( sData_p );
   nValInd_l := 0;
   while sData_l <> '' do
   begin
       sValeur_l := VarToStr(qQuery_p.FindField( sData_l ).AsVariant);
       tvVal_p[nValInd_l] := sValeur_l;
      inc(nValInd_l);
      sData_l := ReadTokenSt( sData_p );
   end;
   result := nValInd_l;
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 10/12/02
Procédure .... : InitSetValChamp
Description .. : Préparation pour l'appel de la procédure AGLSetValChamp
Paramètres ... : Le query de la fiche
                 La liste des données
                 Le tableau de variant à renseigner avec les données :
                 - 0 sTable_c,
                 - 1 nom = valeur du champ,
                 - 2 à n clé de l'enregistrement
Renvoie ...... : Le nombre de données
*****************************************************************}
function InitSetValChamp( qQuery_p : THQuery; sData_p : string ; var tvVal_p : Array Of Variant) : integer;
var
   nValInd_l : integer;
   sData_l : string;
begin
   // Table
   tvVal_p[0] := ReadTokenPipe( sData_p , '|' );
   // Champ à mettre à jour et sa valeur
   sData_l := ReadTokenPipe( sData_p , '|' );
   tvVal_p[1] := ReadTokenPipe( sData_l , '|' );
   // Clé de l'enregistrement
   sData_l := ReadTokenSt( sData_p );
   nValInd_l := 2;
   while sData_l <> '' do
   begin
       tvVal_p[nValInd_l] := sData_l + '=' + VarToStr(qQuery_p.FindField( sData_l ).AsVariant);
      inc(nValInd_l);
      sData_l := ReadTokenSt( sData_p );
   end;
   result := nValInd_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 26/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function InitGetValChamp(qQuery_p : THQuery; sData_p : string ; var tvVal_p : Array Of Variant) : integer;
var
   nValInd_l : integer;
   sData_l : string;
begin
   sData_l := ReadTokenSt(sData_p);
   nValInd_l := 0;
   while sData_l <> '' do
   begin
      if Copy(sData_l, 1, 1) = '@' then
          tvVal_p[nValInd_l] := Copy(sData_l, 2, Length(sData_l) - 1)
      else
          tvVal_p[nValInd_l] := qQuery_p.FindField(sData_l).AsVariant;
      inc(nValInd_l);
      sData_l := ReadTokenSt(sData_p);
   end;
   result := nValInd_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 09/07/2003
Modifié le ... : 09/07/2003
Description .. : Préparation pour l'appel de la procédure AGLSetValChamp
Suite ........ : Paramètres :
Suite ........ :    Le query de la fiche
Suite ........ :    La liste des données
Suite ........ :    Le tableau de variant à renseigner avec les données :
Suite ........ :    - 0 sTable_c,
Suite ........ :    - 1 nom = valeur du champ,
Suite ........ :    - 2 à n clé de l'enregistrement
Suite ........ : Renvoie : Le nombre de données
Mots clefs ... : 
*****************************************************************}
function InitPurgeDoc( qQuery_p : THQuery; sData_p : string ; var tvVal_p : Array Of Variant) : integer;
var
   nValInd_l : integer;
   sData_l, sValeur_l : string;
begin
   sData_l := ReadTokenSt( sData_p );
     tvVal_p[0] := StrToInt( sData_l );
   sData_l := ReadTokenSt( sData_p );
   nValInd_l := 1;
   while sData_l <> '' do
   begin
       sValeur_l := VarToStr(qQuery_p.FindField( sData_l ).AsVariant);
       tvVal_p[nValInd_l] := sValeur_l;
      inc(nValInd_l);
      sData_l := ReadTokenSt( sData_p );
   end;
   result := nValInd_l;
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 10/12/02
Procédure .... : SetValChamp
Description .. : Met à jour l'enregistrement suite au changement
Paramètres ... : La sTable_c
                 Les clés
                 Les valeurs des clés
                 Le champ à traiter
                 La valeur du champ à traiter
*****************************************************************}
procedure SetValChamp( sTable_p : string;
                       tsCle_p, tsCleVal_p : array of string;
                       sChamp_p, sChampVal_p : string);
var
   sRequete_l, sSet_l, sWhere_l : string;
   nInd_l, nMax_l : integer;
begin
   nMax_l := Length(tsCle_p);
   sWhere_l := '';
   for nInd_l := 0 to nMax_l - 1 do
   begin
      if sWhere_l <> '' then
         sWhere_l := sWhere_l + ' AND ';
      sWhere_l := sWhere_l + tsCle_p[nInd_l] + ' = "' + tsCleVal_p[nInd_l] + '"';
   end;
   sSet_l := sChamp_p + ' = "' + sChampVal_p + '"';

   sRequete_l := 'update ' + sTable_p  + ' ' +
                 'set ' + sSet_l + ' ' +
                 'where ' + sWhere_l;
   ExecuteSql( sRequete_l );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 09/07/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function PurgeDoc(sGuidEvt_p, sJevCodeOp_p, sJurNomDos_p : string;
                  TheTOB_p : TOB; var bErreur_p : boolean) : string;
var
   QRYEvenement_l :  TQuery ;
   sJevNomDoc_l, sRequete_l : string;
   OBFille_l : TOB;
begin
   sRequete_l := 'SELECT JEV_DOCNOM FROM JUEVENEMENT WHERE JEV_GUIDEVT = "' + sGuidEvt_p + '"';
   QRYEvenement_l := OpenSQL(sRequete_l, true,-1,'',true);
   If not QRYEvenement_l.Eof then
      sJevNomDoc_l := QRYEvenement_l.findField('JEV_DOCNOM').asString
   else
      sJevNomDoc_l := '';

   Ferme(QRYEvenement_l);

   bErreur_p := false;
   if sJevNomDoc_l = '' then
   begin
      sJevNomDoc_l := 'inconnu';
      bErreur_p := true;
   end
   else if not FileExists(sJevNomDoc_l) then
   begin
      if not V_GedFiles.Erase(sJevNomDoc_l) then
		begin
	      sJevNomDoc_l := 'inexistant : ' + sJevNomDoc_l;
   	   bErreur_p := true;
      end;
   end
   else if not DeleteFile( sJevNomDoc_l ) then
   begin
      sJevNomDoc_l := 'suppression impossible : ' + sJevNomDoc_l;   
      bErreur_p := true;
   end;

   if bErreur_p then
   begin
      OBFille_l := TOB.create('', TheTOB_p, -1);
      OBFille_l.AddChampSup('GUIDEVT', False);
      OBFille_l.putValue('GUIDEVT', sGuidEvt_p);
      OBFille_l.AddChampSup('CODEOP', False);
      OBFille_l.putValue('CODEOP', sJevCodeOp_p);
      OBFille_l.AddChampSup('NOMDOS', False);
      OBFille_l.putValue('NOMDOS', sJurNomDos_p);
      OBFille_l.AddChampSup('DOCNOM', False);
      OBFille_l.putValue('DOCNOM', sJevNomDoc_l);
   end;
   result := sGuidEvt_p;
end;

// Les fonctions ANNUAIRE pour JURI
{*****************************************************************
Auteur ....... : BM
Date ......... : 14/01/2003
Fonction ..... : GetNomCompPer
Description .. : Nom complet
Paramètres ... : Code personne
Renvoie ...... : Le nom
*****************************************************************}
function GetNomCompPer(sGuidPer_p : string) : string;
var
  QQ : TQuery;
begin
  QQ:=OpenSQL('SELECT ANN_NOM1, ANN_NOM2 from ANNUAIRE where ANN_GUIDPER = "' + sGuidPer_p + '"',TRUE,-1,'',true);
  result:='';
  if Not QQ.EOF then
    result:=QQ.FindField('ANN_NOM1').AsString + ' '+ QQ.FindField('ANN_NOM2').AsString;
  Ferme(QQ) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 17/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetCapital(sGuidPer_p : string; var sDevise_p : string;
                       var dCapital_p : double; var iCapNbTitre_p : integer) : boolean;
var
   QRY_l : TQuery;
begin
   QRY_l := OpenSQL('SELECT ANN_CAPITAL, ANN_CAPNBTITRE, ANN_CAPDEV ' +
                    'from ANNUAIRE ' +
                    'where ANN_GUIDPER = "' + sGuidPer_p + '"', TRUE,-1,'',true);
   sDevise_p := '';
   dCapital_p := 0;
   iCapNbTitre_p := 0;
   result := (Not QRY_l.EOF);
   if Not QRY_l.EOF then
   begin
      sDevise_p := QRY_l.FindField('ANN_CAPDEV').AsString;
      dCapital_p := QRY_l.FindField('ANN_CAPITAL').AsFloat;
      iCapNbTitre_p := QRY_l.FindField('ANN_CAPNBTITRE').AsInteger;
   end;
   Ferme(QRY_l) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 14/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetNomAvecCivOuForme(sGuidPer_p : string) : string;
var
   QRY_l : TQuery;
   sNomAvecCivOuForme_l : string;
begin
   QRY_l := OpenSQL('SELECT ANN_PPPM, ANN_FORME, ANN_CVA, ANN_NOM1, ANN_NOM2 ' +
                    'from ANNUAIRE ' +
                    'where ANN_GUIDPER = "' + sGuidPer_p + '"', TRUE,-1,'',true);
   sNomAvecCivOuForme_l := '';
   if Not QRY_l.EOF then
   begin
     if QRY_l.FindField('ANN_PPPM').AsString = 'PM' then
        sNomAvecCivOuForme_l := QRY_l.FindField('ANN_FORME').AsString + ' '
     else
        sNomAvecCivOuForme_l := QRY_l.FindField('ANN_CVA').AsString + ' ';
     sNomAvecCivOuForme_l := sNomAvecCivOuForme_l +
                             QRY_l.FindField('ANN_NOM1').AsString + ' ' +
                             QRY_l.FindField('ANN_NOM2').AsString;
   end;
   Ferme(QRY_l) ;
   result := sNomAvecCivOuForme_l;
end;


{*****************************************************************
Auteur ....... : BM
Date ......... : 14/01/2003
Fonction ..... : GetNomPer
Description .. : Nom complet
Paramètres ... : Code personne
Renvoie ...... : La valeur trouvée
*****************************************************************}
function GetNomPer(sGuidPer_p : string) : string;
var
  QQ : TQuery;
begin
  QQ:=OpenSQL('SELECT ANN_NOMPER from ANNUAIRE where ANN_GUIDPER = "' + sGuidPer_p + '"',TRUE,-1,'',true);
  result:='';
  if Not QQ.EOF then
    result:=QQ.FindField('ANN_NOMPER').AsString;
  Ferme(QQ) ;
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 14/01/2003
Fonction ..... : GetLienPer
Description .. : Nom de recherche du lien
Paramètres ... : Code personne
Renvoie ...... : La valeur trouvée
*****************************************************************}
function GetLienPer(sGuidPer_p : string) : string;
var
  QQ : TQuery;
begin
  QQ:=OpenSQL('SELECT ANL_NOMPER from ANNULIEN where ANL_GUIDPER = "'+ sGuidPer_p + '"', TRUE,-1,'',true);
  result:='';
  if Not QQ.EOF then
    result:=QQ.FindField('ANL_NOMPER').AsString;
  Ferme(QQ) ;
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 14/01/2003
Fonction ..... : GetNomDos
Description .. : Nom dossier
Paramètres ... : Code dossier
Renvoie ...... : La valeur trouvée
*****************************************************************}
function GetNomDos(ValChamp : string) : string;
var
  QQ : TQuery;
begin
  QQ:=OpenSQL('SELECT JUR_NOMDOS from JURIDIQUE where JUR_CODEDOS="'+ValChamp+'"',TRUE,-1,'',true);
  result:='';
  if Not QQ.EOF then
    result:=QQ.FindField('JUR_NOMDOS').AsString;
  Ferme(QQ) ;
end;


function GetLibOp(ValChampOp,ValChampDos : string) : string;
var
  QQ : TQuery;
begin
  QQ:=OpenSQL('SELECT JOP_LIBELLE from JUDOSOPER where JOP_CODEOP="'+ValChampOp+'" and JOP_CODEDOS="'+ValChampDos+'"',TRUE,-1,'',true);
  result:='';
  if Not QQ.EOF then
    result:=QQ.FindField('JOP_LIBELLE').AsString;
  Ferme(QQ) ;
end;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 02/01/03
Fonction ..... : GetEmail
Description .. : Récupère l'email de la personne de l'annuaire
Paramètres ... : Le code de la personne
Renvoie ...... : L'email
*****************************************************************}
function GetEmail(sGuidPer_p : string) : string;
var
  QQ : TQuery;
begin
  QQ:=OpenSQL('SELECT ANN_EMAIL from ANNUAIRE where ANN_GUIDPER = "'+ sGuidPer_p + '"', TRUE,-1,'',true);
  result:='';
  if Not QQ.EOF then
    result:=QQ.FindField('ANN_EMAIL').AsString;
  Ferme(QQ) ;
end;

function RendFormeGen(Forme: String) : String;
// Rend ANN_FORMEGEN à partir de ANN_FORME
var QQ: TQuery;
begin
  QQ:=OpenSQL('select JFJ_FORME,JFJ_FORMEGEN from juformejur where jfj_forme="'+Forme+'"',TRUE,-1,'',true);
  if Not QQ.EOF then
   begin
    if QQ.FindField('JFJ_FORMEGEN').AsString <> '' then
      Result := QQ.FindField('JFJ_FORMEGEN').AsString
    else
      Result := QQ.FindField('JFJ_FORME').AsString;
   end
  else
    Result := '';
  Ferme (QQ);
end;

function GetValChamp(Table,Champ,Cle : string) : variant;
var
  QQ : TQuery;
begin
  QQ:=OpenSQL('SELECT '+Champ+' from '+Table+' where '+Cle,TRUE,-1,'',true);
  result:='';
  if Not QQ.EOF then
    result:=QQ.FindField(Champ).AsVariant;
  Ferme(QQ) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 19/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetValsChamps(sTable_p, sLesChamps_p, sCle_p : string) : string;
var
   QRY_l : TQuery;
   sChamp_l, sResult_l : string;
begin
   QRY_l := OpenSQL('SELECT ' + sLesChamps_p + ' from ' + sTable_p + ' where ' + sCle_p, TRUE,-1,'',true);
   sResult_l := '';
   if Not QRY_l.EOF then
   begin
      sChamp_l := Trim(READTOKENPipe(sLesChamps_p, ','));
      while sChamp_l <> '' do
      begin
         sResult_l := sResult_l + VarToStr(QRY_l.FindField(sChamp_l).AsVariant) + ';';
         sChamp_l := Trim(READTOKENPipe(sLesChamps_p, ','));
      end;
   end;
   Ferme(QRY_l) ;
   result := sResult_l;
end;

function CalculNumCle (NomChamp,NomTable:string; strWhere:string='') : integer; // $$$ JP 05/05/04: ajout condition where
//function CalculNumCle (NomChamp,NomTable : string) : integer;
var
  QQ : TQuery;
begin
  Result:=1;
  if strWhere = '' then
      QQ:=OpenSQL ('SELECT MAX(' + NomChamp + ') AS ' + NomChamp + ' FROM ' + NomTable, TRUE,-1,'',true)
  else
      // $$$ JP 05/05/04 - condition where
      QQ:=OpenSQL ('SELECT MAX(' + NomChamp + ') AS LEMAX FROM ' + NomTable + ' WHERE ' + strWhere, TRUE,-1,'',true);

  // MD 27/07/01
  // if not QQ.EOF then QQ.First else begin Ferme(QQ); exit; end;
  if (Not QQ.Eof) and (QQ.Fields[0].AsString<>'') then
    Result := QQ.Fields[0].AsInteger + 1;
  Ferme(QQ) ;
end;

function OnElipsisClick(Domaine,ListeSel,QuoiSel,SQL,ChampRet1,ChampRet2 : string;var ValChamp: variant) : string;
var
  Q : TQuery;
begin
   result:='';
   ValChamp := AGLLanceFiche(Domaine, ListeSel, QuoiSel, '', '');
   if ValChamp='' then
      exit;

   SQL := SQL + '"' + ValChamp + '"';

   Q := OpenSQL(SQL,true,-1,'',true);
   if not Q.EOF then
   begin
      result := Q.FindField(ChampRet1).AsString;
      if ChampRet2<>'' then
         result := result+' '+Q.FindField(ChampRet2).AsString;
   end;
   Ferme(Q);
end;

function  OnElipsisClick( Domaine, ListeSel, QuoiSel, SQL, ChampRet1, ChampRet2, ChampRet3 : string;
                          var ValChamp1, ValChamp2 : variant ) : string; overload;
var
  Q : TQuery;
begin
   ValChamp1 := AGLLanceFiche(Domaine, ListeSel, QuoiSel, '', '');

   result := ValChamp1;
   if ValChamp1 = '' then
      exit;

   if (ListeSel='ANNUAIRE_SEL') or (ListeSel='ANNUAIRE_SELLITE') then
      SQL := SQL+ValChamp1
   else
      SQL := SQL+'"'+ValChamp1+'"';

   if ExisteSQL( SQL ) then
   begin
      Q := OpenSQL(SQL,true,-1,'',true);

      result := Q.FindField(ChampRet1).AsString;
      if ChampRet2<>'' then
         result := result+' '+Q.FindField(ChampRet2).AsString;
      if ChampRet3<>'' then
         ValChamp2 := Q.FindField(ChampRet3).AsString;
      Ferme(Q);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 17/12/2002
Modifié le ... :   /  /
Description .. : Détermine s'il existe des controles valides sur l'onglet
Suite ........ : Paramètres ... : La fiche, l'onglet
Suite ........ : Renvoie ...... : True s'il y a des controles, false sinon
Mots clefs ... :
*****************************************************************}
function MontreOnglet( sEcran_p : TForm; sOnglet_p : string ) : boolean;
var
   nFICompNb_l : integer;
   bCompVisible_l : boolean;
begin
   bCompVisible_l := false;
   For nFICompNb_l := 0 to sEcran_p.ComponentCount -1 do
   begin
      if ( sEcran_p.Components[nFICompNb_l] is TControl ) and
         ( TControl(sEcran_p.Components[nFICompNb_l]).Parent is TTabSheet ) and
         ( TControl(sEcran_p.Components[nFICompNb_l]).Visible ) and
         ( TControl(sEcran_p.Components[nFICompNb_l]).ClassName <> 'TBevel' ) then
      begin
         if UpperCase( TTabSheet( TControl(sEcran_p.Components[nFICompNb_l]).Parent ).Name )
                     = UpperCase( sOnglet_p ) then
            bCompVisible_l := true;
      end;
   end;
   result := bCompVisible_l;
end;

{*****************************************************************
Auteur ....... : B. MERIAUX
Date ......... : 15/01/03
Fonction ..... : PlusHeure
Description .. : Ajoute un interval à une heure selon un type d'interval
                 et met à jour le reliquat sous forme de nombre de jours
Paramètres ... : L'heure, l'interval, le type d'interval, le reliquat de jours
Renvoie ...... : L'heure incrémentée
*****************************************************************}
function PlusHeure( dtHeure_p : TDateTime; nInterval_p : integer; sTypeInterval_p : string;
                     var nReliqJour_p : integer ) : TDateTime;
var
   wHeure_l, wMinute_l, wSeconde_l, wMSeconde_l : Word;
   nIntervalSeconde_l, nReliqJour_l, nSeconde_l : integer;
begin
   result := dtHeure_p;
   if nInterval_p = 0 then exit;
   if not ( sTypeInterval_p[1] in ['S','M','H'] ) then exit;

   if sTypeInterval_p = 'S' then
      nIntervalSeconde_l := nInterval_p
   else if sTypeInterval_p = 'M' then
      nIntervalSeconde_l := nInterval_p * 60
   else if sTypeInterval_p = 'H' then
      nIntervalSeconde_l := nInterval_p * 3600
   else
      nIntervalSeconde_l := 0;

   DecodeTime( dtHeure_p, wHeure_l, wMinute_l, wSeconde_l, wMSeconde_l );
   nSeconde_l := wSeconde_l + (60 * wMinute_l) + (3600 * wHeure_l);
   nSeconde_l := nSeconde_l + nIntervalSeconde_l;
   nReliqJour_p := 0;
   if nSeconde_l < 0 then
   begin
      // Complément à 24h00
      nReliqJour_p := -1 * ( nSeconde_l div 86400 ) + 1;
      nSeconde_l := nSeconde_l + ( nReliqJour_p * 86400 );
   end;

   wMinute_l := nSeconde_l div 60;
   wSeconde_l := nSeconde_l - ( wMinute_l * 60 );
   wHeure_l := wMinute_l div 60;
   wMinute_l := wMinute_l - ( wHeure_l * 60 );
   nReliqJour_l := (wHeure_l div 24);
   wHeure_l := wHeure_l - ( nReliqJour_l * 24 );
   dtHeure_p := EncodeTime( wHeure_l, wMinute_l, wSeconde_l, wMSeconde_l );

   nReliqJour_p := nReliqJour_p + nReliqJour_l;
   if (nInterval_p < 0) then
      nReliqJour_p := -1 * nReliqJour_p;

   result := dtHeure_p;
end;
function PlusHeure( dtDateEtHeure_p : TDateTime; nInterval_p : integer;
                    sTypeInterval_p : string ) : TDateTime;
var
   nReliqJour_l : integer;
begin
   dtDateEtHeure_p := PlusHeure( dtDateEtHeure_p, nInterval_p, sTypeInterval_p, nReliqJour_l );
   dtDateEtHeure_p := PlusDate( dtDateEtHeure_p, nReliqJour_l, 'J' );
   Result := dtDateEtHeure_p;
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 22/11/02
Fonction ..... : BoolToStr
Description .. : Convertit un booléen en chaîne
Paramètres ... : Le booléen (true/false)
Renvoie ...... : La chaîne correspondante ( "X"/"-" ou "-" si indéterminé)
*****************************************************************}
function JURBoolToStr(Val : boolean) : string;
begin
   if Val = True then
      result := 'X'
   else if Val = False then
      result :='-'
   else
      result := '-';
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 26/02/03
Fonction ..... : BoolToInt
Description .. : Convertit un booléen en entier
Paramètres ... : Le booléen (true/false)
Renvoie ...... : L'entier correspondant ( 0 ou 1)
*****************************************************************}
function JURBoolToInt(Val : boolean) : integer;
begin
   if Val = True then
      result := 1
   else if Val = False then
      result := 0
   else
      result := 0;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 22/11/02
Fonction ..... : StrToBool
Description .. : Convertit une chaîne en booléen
Paramètres ... : La chaîne  ( "X"/"-")
Renvoie ...... : Le booléen correspondant (true/false ou false si indéterminé)
*****************************************************************}
function JURStrToBool(Val : string) : boolean;
begin
   if ( Val = 'X' ) or ( Val = '1' ) then
      result := True
   else if ( Val = '-' ) or ( Val = '0' ) then
      result := False
   else
      result := False;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 26/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{function SupprimeMvt(parms : array of variant; nb : integer) : integer;
var
   iResult_l : integer;
begin
   result := 0;
   if string(parms[0]) = 'X' then
   begin
      // Restauration des titres
   end;

   iResult_l := ExecuteSQL('DELETE FROM JUMVTTITRES ' +
                           'WHERE JMT_GUIDPERDOS = ' + string(parms[1]) +
                           '  AND JMT_TYPEDOS    = "' + string(parms[2]) + '"' +
                           '  AND JMT_NATURECPT  = "' + string(parms[3]) + '"' +
                           '  AND JMT_NOORDRE    = ' + string(parms[4]));
   result := iResult_l + 1;
end;}
{*****************************************************************
Auteur ....... : BM
Date ......... : 22/11/02
Fonction ..... : VarToBool
Description .. : Convertit un variant en booléen
Paramètres ... : La chaîne  ( "X"/"-")
Renvoie ...... : Le booléen correspondant (true/false ou false si indéterminé)
*****************************************************************}
function JURVarToBool(Val : variant) : boolean;
begin
   result := StrToBool( VarToStr( Val ));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function AffecteNomDocMaquette(NomDoc : string) : string;
var
  i : integer;
  Racine,Ext,FileName: string;
begin
  i:=0;
  Racine := Copy(NomDoc,0,Length(NomDoc)-4);
  Ext := Copy(NomDoc,Length(NomDoc)-2,3);
  repeat
    i:=i+1;
    FileName:=Racine+'_'+IntToStr(i)+'.'+Ext;
  until not FileExists(FileName);
  result := FileName;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 10/09/2003
Modifié le ... :   /  /    
Description .. : Extrait une valeur depuis une liste d'argumens présentés 
Suite ........ : sous la forme :
Suite ........ : PARAMETRE1=valeur1;PARAMETRE2=valeur2;PARAME
Suite ........ : TRE3=valeur3...
Mots clefs ... : 
*****************************************************************}
function ReadTokenParam( sParametre_p, sListeParametres_p : string ): string;
var
   sParametre_l, sValeur_l : string;
begin
   repeat
      sValeur_l := READTOKENST( sListeParametres_p );
      sParametre_l := READTOKENPipe( sValeur_l, '=' );
      if sParametre_l <> sParametre_p then
         sValeur_l := '';
   until (sParametre_l = sParametre_p ) or ( sListeParametres_p = '');
   result := sValeur_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 08/09/2003
Modifié le ... :   /  /    
Description .. : Recherche de caractères interdits dans une chaine
Mots clefs ... : 
*****************************************************************}
function ControleCarInterdits( sCode_p, sCarInterdits_p, sExceptions_p : string ) : string;
var
   sCarInterdit_l : string;
   nPos_l : integer;
   bPresent_l : boolean;
begin
   for nPos_l := 1 to Length( sExceptions_p ) do
   begin
      StringReplace( sCarInterdits_p, sExceptions_p[nPos_l], '', [rfReplaceAll, rfIgnoreCase] );
   end;

   nPos_l := 1;
   bPresent_l := false;
   sCarInterdit_l := '';
   while not bPresent_l and ( nPos_l <= Length( sCarInterdits_p ) ) do
   begin
      if Pos( sCarInterdits_p[nPos_l], sCode_p ) <> 0 then
      begin
         bPresent_l := true;
         sCarInterdit_l := sCarInterdits_p[nPos_l];
      end;
      nPos_l := nPos_l + 1;
   end;
   result := sCarInterdit_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 08/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function ControleCode( sCode_p, sExceptions_p : string ) : string;
begin
   result := ControleCarInterdits( sCode_p, cstCarInterdits_g, sExceptions_p );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure LanceWord(NomDoc,DocPerso : string);
var
  RepBib,RepPerso,Doc : string;
begin
   RepBib := GetParamSocSecur('SO_JUREPBIBLE', '');
   RepPerso := GetParamSocSecur('SO_JUREPPERSO', '');
   Doc := NomDoc;
   if (DocPerso='X') and (RepPerso<>'') then
      Doc:=RepPerso   //doc perso
   else if DocPerso='-' then
      Doc:=RepBib;                    //doc bible (sans chemin origine)
   if DocPerso<>'' then
      Doc := Doc + '\' + NomDoc;           //doc bible (avec chemin origine)
   if FileExists(Doc) then
      OpenDoc( Doc, wdWindowStateMaximize, wdPrintView)
   else
      PGIInfo('Le document n''existe pas :'#10#13 + Doc,'Lancement de Word');
end;


{***********A.G.L.***********************************************
Auteur  ...... : E. PLIEZ
Créé le ...... : 27/01/2000
Modifié le ... :   /  /
Description .. : Traduit la valeur X ou - en 1 ou 0
Mots clefs ... : BOOLEAN X - TRUE FALSE
*****************************************************************}
function TraduitBoolean(Val : string) : string;
begin
  if Val='X' then result := '1'
  else if Val='-' then result := '0'
  else if Val='1' then result := 'X'
  else if Val='0' then result :='-'
  else result := '';
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TraduitDate(Val : string) : TDateTime;
begin
  if Val='' then result := iDate1900 else result := StrToDateTime(Val);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TypeNombre(typeCh : string) : boolean;
begin
  result := (TypeCh='E') or (TypeCh='N');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TypeDate(typeCh : string) : boolean;
begin
  result := (TypeCh='D');
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 10/10/02
Fonction ..... : CodeSiren
Description .. : Ajoute les blancs
Paramètres ... : Le numéro SIREN sans blancs
Renvoie ...... : Le numéro SIREN avec blancs
*****************************************************************}

function CodeSiren( sSiren_p : string ) : string;
var
   sNewSiren_l : string;
   nInd_l : integer;
begin
   result := sSiren_p;
   if (sSiren_p <> '') and (Pos(' ', sSiren_p) = 0) then
   begin
      nInd_l := 1;
      while ( nInd_l < Length( sSiren_p) + 1 ) do
      begin
         sNewSiren_l := sNewSiren_l + sSiren_p[nInd_l];
         if ( nInd_l <> Length( sSiren_p) ) and ( nInd_l mod 3 = 0 ) then
            sNewSiren_l := sNewSiren_l + ' ';
         inc(nInd_l);
      end;
      result := sNewSiren_l;
   end;
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 10/10/02
Fonction ..... : DecodeSiren
Description .. : Enlève les blancs
Paramètres ... : Le numéro SIREN avec blancs
Renvoie ...... : Le numéro SIREN sans blancs
*****************************************************************}

function DecodeSiren( sSiren_p : string ) : string;
var Indice, Longueur : Integer;
    ChResultat : String;
    St         : string;
begin
   if sSiren_p = '' then exit;
   longueur:=Length (sSiren_p);
   indice:=1;
   repeat
      if (sSiren_p [Indice]<>' ') then
         St := St + sSiren_p [Indice];
      inc (Indice);
   until (Indice=Longueur+1);
   ChResultat:= St;
   result := ChResultat;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 26/01/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TTToChampLib(sTablette_p : string) : string;
var
   sRequete_l : string;
   QRYRequete_l : TQuery;
begin
   sRequete_l := 'SELECT DO_CHAMPLIB ' +
                 'FROM DECOMBOS ' +
                 'WHERE DO_COMBO = "' + sTablette_p + '"';
   QRYRequete_l := OpenSQL(sRequete_l, True,-1,'',true);
   if not QRYRequete_l.EOF then
      result := QRYRequete_l.FindField('DO_CHAMPLIB').AsString
   else
      result := '';
   Ferme(QRYRequete_l);
end;

{$endif}

{$ifndef EAGL}
{*****************************************************************
Auteur ....... : BM
Date ......... : 06/10/02
Procédure .... : MaxChamp
Description .. : Recher le MAX d'un champ
Paramètres ... : La table, le champ
                 OUT : la valeur de type variant
*****************************************************************}

// Redéfinition pour une valeur de type variant
procedure MaxChamp( sTable_p, sChamp_p : string; var vValeur_p : variant; sWhere_p : string = '' ); overload;
var
   QRY_l : TQuery;
   sRequete_l : string;
begin
   vValeur_p := '';
   sRequete_l := 'SELECT MAX(' + sChamp_p + ') FROM AS ' + sChamp_p + ' ' + sTable_p;
   if sWhere_p <> '' then
      sRequete_l := sRequete_l + ' WHERE ' + sWhere_p;
   QRY_l := OpenSQL( sRequete_l, true ,-1,'',true);
   if (Not QRY_l.eof) and (QRY_l.FindField(sChamp_p).AsVariant <> null) then
      vValeur_p := QRY_l.FindField(sChamp_p).AsVariant;
   Ferme(QRY_l);
end;

// Redéfinition pour une valeur de type integer
procedure MaxChamp( sTable_p, sChamp_p : string; var nValeur_p : integer; sWhere_p : string = '' ); overload;
var
   QRY_l : TQuery;
   sRequete_l : string;
begin
   nValeur_p := 0;
   sRequete_l := 'SELECT MAX(' + sChamp_p + ') AS ' + sChamp_p + ' FROM ' + sTable_p;
   if sWhere_p <> '' then
      sRequete_l := sRequete_l + ' WHERE ' + sWhere_p;
   QRY_l := OpenSQL( sRequete_l, true ,-1,'',true);

   if (Not QRY_l.eof)  and (QRY_l.FindField(sChamp_p).AsVariant <> null) then
   begin
      nValeur_p := QRY_l.FindField(sChamp_p).AsInteger;
   end;
   Ferme(QRY_l);
end;

// Redéfinition pour une valeur de type boolean
procedure MaxChamp( sTable_p, sChamp_p : string; var bValeur_p : boolean; sWhere_p : string = '' ); overload;
var
   QRY_l : TQuery;
   sRequete_l : string;
begin
   bValeur_p := false;
   sRequete_l := 'SELECT MAX(' + sChamp_p + ') AS ' + sChamp_p + ' FROM ' + sTable_p;
   if sWhere_p <> '' then
      sRequete_l := sRequete_l + ' WHERE ' + sWhere_p;
   QRY_l := OpenSQL( sRequete_l, true ,-1,'',true);

   if (Not QRY_l.eof) and (QRY_l.FindField(sChamp_p).AsVariant <> null) then
   begin
      {$ifdef eAGLClient}
      if QRY_l.FindField(sChamp_p).AsString = 'X' then
         bValeur_p := true;
      {$else}
      bValeur_p := QRY_l.FindField(sChamp_p).AsBoolean;
      {$endif}
   end;
   Ferme(QRY_l);
end;

// Redéfinition pour une valeur de type datetime
procedure MaxChamp( sTable_p, sChamp_p : string; var dValeur_p : TDateTime; sWhere_p : string = '' ); overload;
var
   QRY_l : TQuery;
   sRequete_l : string;
begin
   dValeur_p := 0;
   sRequete_l := 'SELECT MAX(' + sChamp_p + ') AS ' + sChamp_p + ' FROM ' + sTable_p;
   if sWhere_p <> '' then
      sRequete_l := sRequete_l + ' WHERE ' + sWhere_p;
   QRY_l := OpenSQL( sRequete_l, true ,-1,'',true);

   if (Not QRY_l.eof) and (QRY_l.FindField(sChamp_p).AsVariant <> null) then
      dValeur_p := QRY_l.FindField(sChamp_p).AsDateTime;
   Ferme(QRY_l);
end;

// Redéfinition pour une valeur de type string
procedure MaxChamp( sTable_p, sChamp_p : string; var sValeur_p : string; sWhere_p : string = '' ); overload;
var
   QRY_l : TQuery;
   sRequete_l : string;
begin
   sValeur_p := '';
   sRequete_l := 'SELECT MAX(' + sChamp_p + ') AS ' + sChamp_p + ' FROM ' + sTable_p;
   if sWhere_p <> '' then
      sRequete_l := sRequete_l + ' WHERE ' + sWhere_p;
   QRY_l := OpenSQL( sRequete_l, true ,-1,'',true);

   if (Not QRY_l.eof) and (QRY_l.FindField(sChamp_p).AsVariant <> null) then
      sValeur_p := QRY_l.FindField(sChamp_p).AsString;
   Ferme(QRY_l);
end;
{$else}
{$endif}


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 28/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function MaxChampX(sTable_p, sChamp_p : string; sWhere_p : string = '') : variant;
var
   OBTable_l : TOB;
   sRequete_l : string;
begin
   sRequete_l := 'SELECT MAX(' + sChamp_p + ') AS ' + sChamp_p + ' FROM ' + sTable_p;
   if sWhere_p <> '' then
      sRequete_l := sRequete_l + ' WHERE ' + sWhere_p;

   OBTable_l := TOB.Create(sTable_p, nil, -1);
   OBTable_l.LoadDetailDBFromSQL(sTable_p, sRequete_l);

   if (OBTable_l.Detail.Count > 0) and OBTable_l.FieldExists(sChamp_p) then
      result := OBTable_l.Detail[0].GetValue(sChamp_p)
   else
      result := Unassigned;
   OBTable_l.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 04/01/2005
Modifié le ... : 09/09/2005
Description .. : Prise en compte des codes dossier de type alphabétique
Suite ........ : Recherche du max dans le cas 00000045 et 00040
Mots clefs ... : 
*****************************************************************}
function MaxCodeDossier : integer;
var
   OBTable_l : TOB;
   sRequete_l, sCur_l : string;
   iInd_l, iCur_l, iMax_l : integer;
begin
   sRequete_l := 'SELECT JUR_CODEDOS FROM JURIDIQUE ' +
                 'WHERE (JUR_CODEDOS <> "&#@" AND JUR_CODEDOS <> "")';
   OBTable_l := TOB.Create('JURIDIQUE', nil, -1);
   OBTable_l.LoadDetailDBFromSQL('JURIDIQUE', sRequete_l);

   //$$$JP 05/12/05 - warning Delphi: inutile -> iCur_l := 0;
   iMax_l := 0;
   for iInd_l := 0 to OBTable_l.Detail.Count - 1 do
   begin
      sCur_l := OBTable_l.Detail[iInd_l].GetString('JUR_CODEDOS');
      if IsNumeric(sCur_l) then
      begin
         iCur_l := StrToInt(sCur_l);
         if iCur_l > iMax_l then
            iMax_l := iCur_l;
      end;
   end;
   OBTable_l.Free;
   result := iMax_l;
end;

{function MaxCodeDossier : integer;
var
   OBTable_l : TOB;
   sRequete_l : string;
begin
   sRequete_l := 'SELECT MAX(JUR_CODEDOS) AS JUR_CODEDOS ' +
                 'FROM JURIDIQUE ' +
                 'GROUP BY JUR_CODEDOS HAVING ISNUMERIC(JUR_CODEDOS) > 0 ' +
                 'ORDER BY JUR_CODEDOS DESC';

   OBTable_l := TOB.Create('JURIDIQUE', nil, -1);
   OBTable_l.LoadDetailDBFromSQL('JURIDIQUE', sRequete_l);

   if (OBTable_l.Detail.Count > 0) then
      result := OBTable_l.Detail[0].GetValue('JUR_CODEDOS')
   else
      result := 0;
   OBTable_l.Free;
end;}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 09/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function FindFirstUpper( OBSource_p : TOB; aosNomChamp_p, aosValeur_p : array of string) : TOB;
var
   bOBFin_l, bARFin_l : boolean;
   nOBInd_l, nARInd_l : integer;
   sOBVal_l, sARVal_l : string;
   OBFin_l : TOB;
begin
   OBFin_l := nil;
   bOBFin_l := false;
   nOBInd_l := 0;
   while not bOBFin_l and (nOBInd_l < OBSource_p.Detail.Count) do
   begin
      nARInd_l := 0;
      bARFin_l := true;
      while bARFin_l and (nARInd_l < Length(aosNomChamp_p)) do
      begin
          sOBVal_l := UpperCase(OBSource_p.Detail[nOBInd_l].GetValue(aosNomChamp_p[nARInd_l]));
          sARVal_l := UpperCase(aosValeur_p[nARInd_l]);
          bARFin_l := (sOBVal_l = sARVal_l);
          Inc(nARInd_l);
      end;
      if bARFin_l then
      begin
         OBFin_l := OBSource_p.Detail[nOBInd_l];
         bOBFin_l := true;
      end;
      Inc(nOBInd_l);
   end;
   result := OBFin_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 19/02/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function ChampTypeVarchar(sChamp_p : string) : boolean;
var
   sTypeChamp_l : string;
begin
   sTypeChamp_l := ChampType(sChamp_p);
   result := (Copy(sTypeChamp_l, 1, 7) = 'VARCHAR');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 28/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function ChampType(sChamp_p : string) : string;
var
   sRequete_l : string;
   OBRequete_l : TOB;
begin
   sRequete_l := 'SELECT DH_TYPECHAMP ' +
                 'FROM DECHAMPS ' +
                 'WHERE DH_NOMCHAMP = "' + sChamp_p + '"';
   OBRequete_l := TOB.Create('DECHAMPS', nil, -1);
   OBRequete_l.LoadDetailDBFromSQL('DECHAMPS', sRequete_l);
   if OBRequete_l.Detail.Count > 0 then
      result := OBRequete_l.GetValue('DH_TYPECHAMP')
   else
      result := '';
   OBRequete_l.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 13/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{procedure Verifcoherence(sGuidPerDos_p, sCodedos_p : String);
var
   FCoherence_l : TCoherence;
begin
  FCoherence_l := TCoherence.Create;
  FCoherence_l.ControlNbTitre(sGuidPerDos_p);
  FCoherence_l.ControlNbVoix(sGuidPerDos_p);
  FCoherence_l.ControlCapital(sGuidPerDos_p);
  FCoherence_l.ControlAssAct(sGuidPerDos_p, sCodeDos_p);
  FCoherence_l.ControlComCo(sGuidPerDos_p);
  FCoherence_l.free;
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 19/11/2007
Modifié le ... : 19/11/2007
Description .. : Passage d'un pointeur sur le Dataset de la TOM en 
Suite ........ : paramètre car LaTom.ForceUpdate ne se comporte pas de 
Suite ........ : la même façon en 2 tiers et en WebAccess
Mots clefs ... : 
*****************************************************************}
function  ChangeFormeJuridique(LaTom_p :Tom; sMajForce_p : string; PDataset_p : TPDataset) : boolean ; overload;
//vrai si validée par utilisateur
var
   sCodeInsee_l, sForme_l, sFormePrive_l, sFormeSte_l, sPPPM_l, sTypeSCI_l : string;
   sFormeSci_l, sCoop_l, sDORFormeAsso_l, sDFIRegleFisc_l, sDORSectionBnc_l : string;
begin
   result := false;
   sCodeInsee_l := Latom_p.GetField('ANN_CODEINSEE');

   if not ChangeFormeJuridique(sCodeInsee_l, sForme_l, sFormePrive_l, sFormeSte_l, sFormeSci_l,
                              sDORFormeAsso_l, sCoop_l, sDFIRegleFisc_l, sDORSectionBnc_l,
                              sPPPM_l, sTypeSCI_l) then exit;

   ModeEdition(PDataset_p^);
   LaTom_p.ForceUpdate;

   if sMajForce_p <> '' then
      LaTom_p.setField(sMajForce_p, LaTom_p.getField(sMajForce_p));

    LaTom_p.SetField('ANN_PPPM', sPPPM_l);

   LaTom_p.SetField('ANN_COOP', sCoop_l);
  // FQ 11396 - pas moyen que le checked soit propagé si champ lié !
   if LaTom_p.TableName = 'ANNUAIRE' then
      LaTom_p.SetField('ANN_COOP', sCoop_l);

   LaTom_p.SetField('ANN_FORME', sForme_l);
   LaTom_p.SetField('ANN_FORMEGRPPRIVE', sFormePrive_l);
   LaTom_p.SetField('ANN_FORMESTE', sFormeSte_l);
   LaTom_p.SetField('ANN_FORMESCI', sFormeSci_l);
   LaTom_p.SetField('ANN_CODEINSEE', sCodeInsee_l);

   LaTom_p.SetControlText('DFI_REGLEFISC', sDFIRegleFisc_l);
   LaTom_p.SetControlText('DOR_SECTIONBNC', sDORSectionBnc_l);

   LaTom_p.SetField('ANN_TYPESCI', sTypeSCI_l);
   result := true;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 27/11/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function ChangeFormeJuridique(var sCodeInsee_p, sForme_p, sFormePrive_p, sFormeSte_p, sFormeSci_p,
                              sDORFormeAsso_p, sCoop_p, sDFIRegleFisc_p, sDORSectionBnc_p,
                              sPPPM_p, sTypeSCI_p : string) : boolean ;
//vrai si validée par utilisateur
var
   sOldCodeInsee_l, sResult_l : string;
begin
  result := false  ;
  sOldCodeInsee_l := sCodeInsee_p;
  sResult_l := AGLLanceFiche('YY', 'YFORMESINSEE_SEL', '', '', sCodeInsee_p);

  if sResult_l = 'FALSE' then exit;

  sCodeInsee_p  := READTOKENST(sResult_l);     {1}
  sForme_p      := READTOKENST(sResult_l);     {2}
  sFormePrive_p := READTOKENST(sResult_l);     {3}
  sFormeSte_p   := READTOKENST(sResult_l);     {4}
  sFormeSci_p   := READTOKENST(sResult_l);     {5}
  sDORFormeAsso_p  := READTOKENST(sResult_l);  {6}
  sCoop_p          := READTOKENST(sResult_l);  {7}
  sDFIRegleFisc_p  := READTOKENST(sResult_l);  {8}
  sDORSectionBnc_p := READTOKENST(sResult_l);  {9}

  if sOldCodeInsee_l = sCodeInsee_p then exit;

  if (Copy(sCodeInsee_p, 1, 1) = '1') then
      sPPPM_p := 'PP'
  else
      sPPPM_p := 'PM';

  if sCodeInsee_p = '6540' then
      sTypeSCI_p := 'LOC'
  else if sCodeInsee_p= '6541' then
      sTypeSCI_p := 'CVT'
  else if sCodeInsee_p= '6542' then
      sTypeSCI_p := 'ATT';
  //nb : le raz est fait si nécessaire sur le OnUpdateRecord

  result := true ;
end ;

function  ChangeCodeNaf(LaTom:Tom; Ctrl:TControl; codeNafIni:string) : boolean ; //LM20070516
var s, act, lib  : string ;
    sl : HTStrings ;
begin

  s := LaTom.GetControlText(Ctrl.name) ; //code naf
  if pos('2', Ctrl.name)>0 then act:='OLE_ACTIVITE2' else act:='OLE_ACTIVITE' ;

  sl:=HTStringList.Create;
  RichToStrings(TRichEdit(LaTom.getControl(act)), sl) ;

  if ((s<>'')
   and ((Trim(CbpRTFtoText(sl.Text, false, false, 0))='')) or (codeNafIni<>s)) then
  begin //si code naf change  => activité = CO_LIBRE provenant de YYCODENAF
    lib := rechDom('YYCODENAF', s, false, '', true) ;
    if lib<>'' then LaTom.setControlText(act, lib) ;
  end
  else if trim(s)='' then
    LaTom.setControlText(act, '') ;

  sl.free ;

end ;

procedure AfficheInfoLienPersonne(LaTom:TObject; GuidPer, ZoneGuidPer, ZoneLibelle: string);//LM20070703
var NomPer, Libelle, TypePer, Classe: string;
    cp : tcomponent ;
begin
  TraiteChoixPersonne(GuidPer, Libelle, NomPer, TypePer);
  Classe:=LaTom.ClassName ;
  if upperCase(copy(Classe,1,3)) = 'TOM' then
  begin
    if Tom(LaTom).GetControl(ZoneLibelle + 'SOV') <> nil then
      Tom(LaTom).SetControlText(ZoneLibelle + 'SOV', Libelle);
    Tom(LaTom).SetControlText(ZoneLibelle, Libelle);
  end
  else if inString(Classe, ['TFVierge']) then
  begin
    cp:=TForm(LaTom).FindComponent(ZoneLibelle + 'SOV') ;
    if cp<>nil then TEdit(cp).text:= Libelle ;
    cp:=TForm(LaTom).FindComponent(ZoneLibelle) ;
    if cp<>nil then TEdit(cp).text:= Libelle ;
  end ;
end;


procedure cElipsisPerAssClick(LaTom:Tom; Sender: TObject);  //LM20070516
var sChampCode_l, sChampNom_l, sValeurCode_l, sValeurNom_l, sSansDossier_l: string;
begin
  if sender=nil then exit ;
  if pos('cjs', Application.ExeName) >0 then sSansDossier_l := 'X' else sSansDossier_l := '-';

  sChampNom_l := Tcontrol(Sender).Name;
  sChampCode_l := Copy(sChampNom_l, 2, Length(sChampNom_l));

  sValeurNom_l := LaTom.GetControlText(sChampNom_l) ;

  sValeurCode_l := LancerAnnuSel ('ANN_NOMPER=' + sValeurNom_l, '', ';;' + sSansDossier_l);
  if (sValeurCode_l<>'') and  (sValeurCode_l<> laTom.GetControlText(sChampCode_l)) then
  begin
    LaTom.forceUpdate ;
    LaTom.SetField(sChampCode_l, sValeurCode_l);
  end ;
end;

function RecupNomGerant(GuidPerDos : String; var RefGerant:string): String;   //de TomOrga... LM20070516
var Q1 : TQuery;
begin
  Result := '';
  RefGerant := '';
  Q1 :=  OpenSQL('select ANL_NOMPER, ANL_GUIDPER from ANNULIEN ' +
                 'where ANL_GUIDPERDOS="' + GuidPerDos + '" and ANL_FONCTION="PFC"', True,-1,'',true);
  if not((Q1=nil) or (Q1.EOF)) then
  begin
    Result := Q1.Fields[0].AsString;
    RefGerant := Q1.Fields[1].AsString;
  end;
  Ferme(Q1);
end;

function VerrouilleSiBaseComptable(Ecran : TOM; guid : string ) : boolean ;  //LM20070620
var q:TQuery ;
begin
  result:=false ;
  with Ecran do
  begin
    q := openSQL( 'select DOS_ABSENT from DOSSIER ' +
                  'inner join DOSSAPPLI on DAP_NODOSSIER = DOS_NODOSSIER ' +
                  'where DOS_GUIDPER="' + guid + '" and DAP_NOMEXEC in ("CCS5.EXE", "eCCS5.EXE")', true,-1,'',true);
    if q.Eof or ((not q.Eof) and (q.FindField('DOS_ABSENT').asString='X')) then else
    begin
      setControlEnabled('JUR_DUREEEXPREC', false) ;
      setControlEnabled('JUR_DATEFINEX', false) ;
      setControlEnabled('JUR_DUREEEX', false) ;
      setControlEnabled('JUR_DUREEEXPREC', false) ;
      setControlEnabled('JUR_MOISCLOTURE', false) ;
      result:=true ;
    end ;
    ferme(q);
  end ;
end ;

end.

