unit CPVersion;

{--------------------------- GP Code Spécif : --------------------------------

51199 : Escompte fournisseur   Passe en standard
51198 : Modid entete pièce     Passe en standard
51197 : Nouveau module relance Passe en standard
51196 : Affectation automatique TL0 tiers sur TL0 ecr  Passe en standard ???  (mais pas sur la saisie)
51195 : Lettre virements pour CCMP  Passe en standard
51194 : Consultation des écritures en PGE : ancienne version non optimisée
51193 : Accès à la saisie en mode libre    Passe en standard
51192 : Lettre prélèvements
51191 : Accès à l'ancienne version Pointage
51190 : Accès aux anciennes éditions de pointage
51189 : Accès à l'ancienne version extourne
51188 : Stop écart de conversion
51187 : GetCumuls optimisé
51186 : Onglet avancès sur éditions standards
51185 : Ancienne version de getcumuls non compatible EAGL
51500 : CCS7 en mode conso gestion
51501 : creation d'un fichier de sauvegarde en saisie de bordereau
51502 : Gestion des standard niveau Cegid
51503 : Accès à l' ancienne version de la consultation des écritures
51504 : Specif KPMG - FQ 23294
51204 : Sauvegarde du folio en cours de saisie
51205 : Autorise tous les caractères pour G_GENERAL ou T_AUXILIAIRE
51206 : Encaissement / Décaissement en masse (DIRECT NRJ)
51207 : Pas de sauvegarde dossier format TRA au lancement de la PURGE
51208 : Saisie paramétrable en PCL
51209 : ancien reclassement
51210 : validation d'écritures simulées avec une REFGESCOM
51211 : activation GED en entreprise
51212 : Paramétrage des masques en saisie
51213 : Activation du nouveau pointage
51214 : Choix de la devise dans les états de synthèse
51215 : Enregistrement spécifique des TOb en saisie (sans les Blob) pour pb CWAS
51310 : Specif CERIC
51320 : Specif TESSI TVA : possibilité de réimputation sur pièce GESCOM + onglet avancé (CPREIMPUT_TOF.PAS)
51330 : maj des solde des comptes sans TRANSACTION : permet de revenir sur l'ancien code utilisant la procedure TRANSACTION
51340 : CCMP --> OPTIMISATION des chargements des infos des comptes dans les modules génératin des ENCA / DECA 
51801--51899 pour spécif compta Lek & Gérard 220107
51801 : Netting pour Europe assistance.
{---------------------------------------------------------------------------------------}

interface

uses
  CPTypeCons;

{Fonctions à supprimer}
function EstComptaSansAna : Boolean;
function EstComptaPackAvance : Boolean;
function EstbaseConso : Boolean;
{Fin Fonctions à supprimer}

function IsDossierCabinet(NumDossier : string) : Boolean;
function _IsDossierNetExpert : Boolean;
function EstSpecif(St : string) : Boolean;
Function EstSpecifCompta(St : String) : Boolean ; //Lek
function GetFromPCL : Boolean;
function EstImmoPGI : Boolean;
function EstComptaIFRS : Boolean;
function CPEstComptaAgricole : Boolean;
function EstComptaTreso : Boolean;
function CPPresenceEtafi : Boolean;
function CPEstTvaActivee : Boolean;

//function RenseignelaSerie(QuelProduit : tQuelProduit) : Byte ;
function IsSerialise(Product : LongWord) : Integer;

function GetRecupPCL   : Boolean;
function GetRecupLTL   : Boolean;
function GetRecupComSx : Boolean;
function GetRecupSISCOPGI : Boolean;
function GetRecupCegid : Boolean;

procedure SetParamUseEtats;

implementation

uses
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$IFNDEF EAGLSERVER}
  Web,
  {$ENDIF EAGLSERVER}
  {$ENDIF EAGLCLIENT}
  
  {$IFDEF COMPTA}
  galOutil,
  {$ENDIF}
  uTob, HCtrls, HEnt1, Ent1, SysUtils, Forms, ParamSoc, cbpPath;


{Retourne True si la tréso est utilisée
{---------------------------------------------------------------------------------------}
function EstComptaTreso : Boolean;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRESO}
  Result := True; {FQ 10184}
  {$ELSE IF}
    Result := not (ctxPCL in V_PGI.PGIContexte) and
              ( (ctxTreso in V_PGI.PGIContexte) or GetParamSocSecur('SO_MODETRESO',False) );
  {$ENDIF TRESO}
end;

{Spécificités agricoles
{---------------------------------------------------------------------------------------}
function CPEstComptaAgricole : Boolean;
{---------------------------------------------------------------------------------------}
begin
  {MVG 11/10/2006 FQ 18948 RSI => RS
   MVG 28/03/2007 FQ 1987}
  {$IFDEF COMPTA}
  Result := ExisteSql('SELECT DFI_REGIMFISCDIR FROM DPFISCAL Where DFI_REGLEFISC="BA" and (DFI_REGIMFISCDIR="RN" or DFI_REGIMFISCDIR="RS") and '+
                      'DFI_GUIDPER="'+ GetGuidPer(V_PGI.NoDossier) + '"' );
  {$ELSE}
  Result := False;
  {$ENDIF COMPTA}
end;

{Détermine si c'est un dossier cabinet
{---------------------------------------------------------------------------------------}
function IsDossierCabinet(NumDossier : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  ChSql      : string;
  RSql       : Tquery;
  ChResultat : Boolean;
begin
 ChResultat := False;

 ChSql := 'SELECT DOS_CABINET FROM DOSSIER WHERE DOS_NODOSSIER = "' + NumDossier + '"';
 RSql  := OpenSql(ChSql, True);
 if (not RSql.Eof) then
  ChResultat := RSql.FindField('DOS_CABINET').AsString = 'X';
 Ferme(RSql);

 IsDossierCabinet := ChResultat;
end;

{---------------------------------------------------------------------------------------}
function EstSpecif(St : string) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := (Pos(St, V_PGI.Specif) > 0);
end;

Function EstSpecifCompta(St : String) : Boolean ; //Lek
var specifList : String   ;
BEGIN
 If AglExistParamSoc('SO_I_CODESPECIF') then
 begin
   specifList := GetParamSocSecur('SO_I_CODESPECIF', '') ;
   Result     := (Pos(St,specifList) > 0) ;
   if Not result then Result := (Pos(St,V_PGI.Specif) > 0) ;
 end
 else Result:=(Pos(St,V_PGI.Specif) > 0) ;
END ;

{---------------------------------------------------------------------------------------}
function GetFromPCL : Boolean;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF NOVH}
  Result := True; 
  {$ELSE}
  Result := VH^.FromPCL;
  {$ENDIF}
end;

{Cette fonction est devenue inutile = > à Supprimer
{---------------------------------------------------------------------------------------}
function EstComptaPackAvance : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
end;

{Cette fonction est devenue inutile = > à Supprimer
{---------------------------------------------------------------------------------------}
function EstComptaSansAna : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := False;
end;

{Cette fonction est-elle devenue inutile ? = > à Supprimer ?
{---------------------------------------------------------------------------------------}
function EstbaseConso : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := False;
  {$IFDEF CONSOCERIC}
  {$IFDEF NOVH}
  Result := False;
  {$ELSE}
  if VH^.IsBaseConso then Result := True;
  {$ENDIF NOVH}
  {$ENDIF CONSOCERIC}
end;

{---------------------------------------------------------------------------------------}
function EstImmoPGI : Boolean ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF PGIIMMO}
  Result := True;
  {$ELSE}
  Result := False;
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
function EstComptaIFRS : Boolean;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF NOVH}
  Result := False;
  {$ELSE}
  Result := ( ( (ctxCompta in V_PGI.PGIContexte) and ( VH^.OkModCPPackIFRS ) )
              or ( V_PGI.VersionDemo )
             ) and (VH^.PaysLocalisation <> CodeISOES) ; {XVI 24/02/2005 XVI Ce modiule n'existe pas en Espagne.}
  {$ENDIF NOVH}
end;


{Type tQuelProduit = (exeCCS7,exeCCS5,exeCCS3,exeCCIMPEXP,exeCCMP,exeCCAUTO) ;

{---------------------------------------------------------------------------------------}
//function RenseignelaSerie(QuelProduit : tQuelProduit) : Byte ;
{---------------------------------------------------------------------------------------}
(*
begin
  Result := 5;
{$IFNDEF EAGLSERVER}
  HalSocIni:='CEGIDPGI.INI' ;

  {$IFDEF IMMOSIC}
  HalSocIni:='CEGIDPGISIC.INI' ;
  {$ENDIF}

  // STD
  V_PGI.NumVersionBase := 850;
{$IFNDEF ANAL2}
  V_PGI.NumVersion:='8.1.0' ;
  V_PGI.NumBuild:='002.001';
  V_PGI.DateVersion:=EncodeDate(2007,10,25);
{$ELSE}
  V_PGI.NumVersion:='9.0.0' ;
  V_PGI.NumBuild:='000.010';
  V_PGI.DateVersion:=EncodeDate(2008,03,03);
{$ENDIF}
  HalSocIni:='CEGIDPGI.INI' ;
  {$IFDEF IMMOSIC}
  HalSocIni:='CEGIDPGISIC.INI' ;
  {$ENDIF}

  V_PGI.ScriptApp           := 'SCRIPTAPP_CPTA' ; // FQ 22316 : positionner la variable V_PGI.ScriptApp

  case QuelProduit Of
     ExeCCS3 : BEGIN
               If EstImmoPGI Then
                 BEGIN
                 {$IFDEF EAGLCLIENT}
                   {$IFDEF INVSAL}
                   NomHalley:='eINVENTAIRE S3' ;
                   TitreHalley:='INVENTAIRE S3' ;
                   {$ELSE}
                   NomHalley:='eSERVANTISSIMMO S3' ;
                   TitreHalley:='SERVANTISSIMMO S3' ;
                   {$ENDIF}
                 {$ELSE}
                   {$IFDEF INVSAL}
                   NomHalley:='INVENTAIRE S3' ;
                   TitreHalley:='INVENTAIRE S3' ;
                   {$ELSE}
                   NomHalley:='SERVANTISSIMMO S3' ;
                   TitreHalley:='SERVANTISSIMMO S3' ;
                   {$ENDIF}
                 {$ENDIF}
                 END;
                {$IFNDEF BUSINESSPLACE}
                  NomHalley:='Comptabilité S3' ;
                  TitreHalley:='Comptabilité' ;
                {$ENDIF}
               END ;
     ExeCCS5 : BEGIN
               If EstImmoPGI Then
                 BEGIN
                 {$IFDEF EAGLCLIENT}
                   {$IFDEF INVSAL}
                   NomHalley:='eINVENTAIRE S5' ;
                   TitreHalley:='INVENTAIRE S5' ;
                   {$ELSE}
                   NomHalley:='eSERVANTISSIMMO S5' ;
                   TitreHalley:='SERVANTISSIMMO S5' ;
                   {$ENDIF}
                 {$ELSE}
                   {$IFDEF INVSAL}
                   NomHalley:='INVENTAIRE S5' ;
                   TitreHalley:='INVENTAIRE S5' ;
                   {$ELSE}
                   NomHalley:='SERVANTISSIMMO S5' ;
                   TitreHalley:='SERVANTISSIMMO S5' ;
                   {$ENDIF}
                 {$ENDIF}
                 END Else
                 BEGIN
                 {$IFDEF EAGLCLIENT}
                 {$IFDEF TESTSIC}
                   TitreHalley:='Comptabilité - TESTSIC';
                   NomHalley:='eComptabilité S5 - TESTSIC';
                 {$ELSE}
                   NomHalley:='eComptabilité S5' ;
                   TitreHalley := 'Comptabilité' ;
                 {$ENDIF}
                 {$ELSE}
                 NomHalley:='Comptabilité S5' ;
                 {$IFNDEF HVCL}
                  TitreHalley := 'Comptabilité';

                  {$ENDIF HVCL}
                  {$ENDIF}
                 END ;
               END ;
     ExeCCS7 : BEGIN
               If EstImmoPGI Then
                 BEGIN
                 {$IFDEF EAGLCLIENT}
                   {$IFDEF INVSAL}
                   NomHalley:='eINVENTAIRE S7' ;
                   TitreHalley:='INVENTAIRE S7' ;
                   {$ELSE}
                   NomHalley:='eSERVANTISSIMMO S7' ;
                   TitreHalley:='SERVANTISSIMMO S7' ;
                   {$ENDIF}
                 {$ELSE}
                   {$IFDEF INVSAL}
                   NomHalley:='INVENTAIRE S7' ;
                   TitreHalley:='INVENTAIRE S7' ;
                   {$ELSE}
                   NomHalley:='SERVANTISSIMMO S7' ;
                   TitreHalley:='SERVANTISSIMMO S7' ;
                   {$ENDIF}
                 {$ENDIF}
                 END ;
               END ;
     ExeCIIMPEXP : BEGIN
                   NomHalley:='SERVANTISSIMO S5' ;
                   TitreHalley:='SERVANTISSIMO IMPORT EXPORT S5' ;
                   V_PGI.LaSerie:=S5 ;
                   V_PGI.NumVersion:='1.0.0' ;
                   V_PGI.NumBuild:='000.001' ;
                   END ;
     ExeCCIMPEXP : BEGIN
                   // Modif Fiche 12760 SBO
                   NomHalley := 'Comptabilité' ;
                   {$IFNDEF HVCL}
                   TitreHalley := 'Imports / Exports' ;
                   {$ENDIF HVCL}
                   V_PGI.LaSerie:=S5 ;
                   Application.HelpFile:=TcbpPath.GetCegidDistriDoc+'\CCS5.chm' ;
                   V_PGI.NumMenuPop:=27 ;
                   V_PGI.OfficeMsg:=True ;
                   V_PGI.OutLook:=TRUE ;
                   V_PGI.NoModuleButtons:=True ;
                   V_PGI.NbColModuleButtons:=1 ;
                   END ;
     ExeCCMP : BEGIN
                   {$IFDEF CCMPGC}
                   NomHalley:= 'Règlements' ;
                   TitreHalley := 'Règlements' ;
                   {$ELSE}
                   {$IFDEF TESTSIC}
                      TitreHalley:='Suivi des règlements - TESTSIC';
                      NomHalley:='Suivi des règlements S5 - TESTSIC';
                   {$ENDIF}
                   NomHalley:= 'Suivi des règlements S5' ;
                   TitreHalley := 'Suivi des règlements' ;
                   {$ENDIF}
                   Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CCS5.chm' ;
                   V_PGI.NumMenuPop:=27 ;
                   V_PGI.OfficeMsg:=True ;
                   V_PGI.OutLook:=TRUE ;
                   V_PGI.NoModuleButtons:=True ;
                   V_PGI.NbColModuleButtons:=1 ;
                 end;
     ExeCCAUTO : BEGIN
                     NomHalley:= 'Comptabilité S5' ;
                     {$IFNDEF HVCL}
                     TitreHalley := 'Imports / Exports' ;
                     {$ENDIF HVCL}
                     V_PGI.LaSerie:=S5 ;
                     Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CCS5.chm' ;
                   END ;
     ExePGIMAJVER: BEGIN
                     TitreHalley:='Administration' ;
                     V_PGI.LaSerie:=S5 ;
                     Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CCS5.chm' ;
                   END ;
     ExeCTS5     : begin
                    {$IFDEF TESTSIC}
                    NomHalley           := 'Trésorerie S5 - TESTSIC';
                    TitreHalley         := 'Trésorerie - TESTSIC';
                    {$ELSE}
                    NomHalley           := 'Trésorerie S5';
                    TitreHalley         := 'Trésorerie';
                    {$ENDIF TESTSIC}
                     Application.HelpFile:= TcbpPath.GetCegidDistriDoc+ '\CTS5.CHM';
                   end;
     ExeCTS3     : begin
                     NomHalley           := 'Trésorerie S3';
                     TitreHalley         := 'Trésorerie';
                     Application.HelpFile:= TcbpPath.GetCegidDistriDoc+ '\CTS5.CHM';
                   end;

     ExeCCSTD     : begin
                     NomHalley           := 'Standard Comptabilité S5';
                     TitreHalley         := 'Standards Comptabilité';
                     V_PGI.LaSerie       := S5;
                     Application.HelpFile:= '';
                   end;
     ExeCOMSX     : begin
                      {$IFDEF TESTSIC}
                      TitreHalley:='Communications Sx - TESTSIC';
                      NomHalley:='COMSX - TESTSIC';
                      {$ELSE}
                      NomHalley:= 'COMSX';
                      TitreHalley := 'Communications Sx' ;
                      {$ENDIF}
                      V_PGI.LaSerie       := S5;
                      Application.HelpFile:=TcbpPath.GetCegidDistriDoc+'\ComSx.chm';
                    end;
     ExeAutre    : ;
     END ;

   Application.Title := TitreHalley ;
  {$IFDEF EAGLCLIENT}
  {$ELSE}
    SetDidacticielPath (TCbpPath.getCegidDistriDoc);
  {$ENDIF}

{$ENDIF EAGLSERVER}
end;
*)
{---------------------------------------------------------------------------------------}
function IsSerialise(Product : LongWord) : Integer;
{---------------------------------------------------------------------------------------}
begin
  Result := c_CheckProductSeriaPGI(Product);
end;

{---------------------------------------------------------------------------------------}
function CPPresenceEtafi : Boolean;
{---------------------------------------------------------------------------------------}
var
  stFileName : string;
begin
  {$IFDEF EAGLCLIENT}
  stFileName := 'eEtafiS5.exe';
  {$ELSE}
  stFileName := 'EtafiS5.exe';
  {$ENDIF}

  Result := FileExists(ExtractFilePath(Application.ExeName) + '\' + stFileName);
end;

{---------------------------------------------------------------------------------------}
function CPEstTvaActivee : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := GetParamSocSecur('SO_CPPCLSAISIETVA', False);
end;

{---------------------------------------------------------------------------------------}
function GetRecupPCL : Boolean;
{---------------------------------------------------------------------------------------}
begin
{$IFDEF NOVH}
 Result := True; // Pour La Comsx Server
{$ELSE}
 Result := VH^.RecupPCL;
{$ENDIF}
end;

{---------------------------------------------------------------------------------------}
function GetRecupComSx : Boolean;
{---------------------------------------------------------------------------------------}
begin
{$IFDEF NOVH}
 Result := True; // Pour La Comsx Server
{$ELSE}
 Result := VH^.RecupComSx;
{$ENDIF}
end;

{---------------------------------------------------------------------------------------}
function GetRecupSISCOPGI : Boolean;
{---------------------------------------------------------------------------------------}
begin
{$IFDEF NOVH}
 Result := False; // Pour La Comsx Server
{$ELSE}
 Result := VH^.RecupSISCOPGI;
{$ENDIF}
end;

{---------------------------------------------------------------------------------------}
function GetRecupLTL : Boolean;
{---------------------------------------------------------------------------------------}
begin
{$IFDEF NOVH}
 Result := False; // Pour La Comsx Server
{$ELSE}
 Result := VH^.RecupLTL;
{$ENDIF}
end;

{---------------------------------------------------------------------------------------}
function GetRecupCegid : Boolean;
{---------------------------------------------------------------------------------------}
begin
{$IFDEF NOVH}
 Result := False; // Pour La Comsx Server
{$ELSE}
 Result := VH^.RecupCegid;
{$ENDIF}
end;

{---------------------------------------------------------------------------------------}
function _IsDossierNetExpert : Boolean;
{---------------------------------------------------------------------------------------}
{$IFNDEF EAGLSERVER}
var Q : TQuery;
{$ENDIF EAGLSERVER}
begin
  Result := False;
{$IFNDEF EAGLSERVER}
  Q := OpenSql ('SELECT DOS_NETEXPERT,DOS_NECPSEQ FROM DOSSIER WHERE DOS_NODOSSIER = "' + V_PGI.NoDossier + '"', True);
  if not Q.Eof then
    Result := Q.FindField ('DOS_NETEXPERT').AsString = 'X';
  Ferme (Q);
{$ENDIF EAGLSERVER}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 15/04/2008
Modifié le ... :   /  /
Description .. : Activation de la suppression des documents dans la compta
Mots clefs ... : YMO0408
*****************************************************************}
procedure SetParamUseEtats;
var
  i   : integer;
begin
  {Suivi des règlements :}
  {On vide BanqueCP de son paramétrage associé aux anciens documents

  ListNatDocBQ:='LC1,TRA,VRT,LPR';
  ListNatEtaBQ:='CLC,CLT,CLV,CLP';
  ListNatDocPSce:='LCH,LPR,LVI,LTR';
  ListNatEtaPSce:='CCH,CPR,CVI,CTR';                                }

  {Maj param comptes bancaires si <>""}
  ExecuteSql('UPDATE BANQUECP SET BQ_LETTRECHQ="CLC" WHERE BQ_LETTRECHQ="LC1"');
  ExecuteSql('UPDATE BANQUECP SET BQ_LETTRELCR="CLT" WHERE BQ_LETTRECHQ="TRA"');
  ExecuteSql('UPDATE BANQUECP SET BQ_LETTREVIR="CLV" WHERE BQ_LETTRECHQ="VRT"');
  ExecuteSql('UPDATE BANQUECP SET BQ_LETTREPRELV="CLP" WHERE BQ_LETTRECHQ="LPR"');

  {On ne laisse plus le choix documents/etats dans le paramétrage des scénarii}
  ExecuteSql('UPDATE CPARAMGENER SET CPG_ETATNATURE="CCH", CPG_MODELEENCADECA="LCH", CPG_AVECDOC="-" '
            +'WHERE CPG_AVECDOC="X" AND CPG_ETATNATURE="LCH"');
  ExecuteSql('UPDATE CPARAMGENER SET CPG_ETATNATURE="CPR", CPG_MODELEENCADECA="LPR", CPG_AVECDOC="-" '
            +'WHERE CPG_AVECDOC="X" AND CPG_ETATNATURE="LPR"');
  ExecuteSql('UPDATE CPARAMGENER SET CPG_ETATNATURE="CVI", CPG_MODELEENCADECA="LVI", CPG_AVECDOC="-" '
            +'WHERE CPG_AVECDOC="X" AND CPG_ETATNATURE="LVI"');
  ExecuteSql('UPDATE CPARAMGENER SET CPG_ETATNATURE="CTR", CPG_MODELEENCADECA="LTR", CPG_AVECDOC="-" '
            +'WHERE CPG_AVECDOC="X" AND CPG_ETATNATURE="LTR"');

  //Maj RELANCES
  for i:=1 to 7 do
    ExecuteSQL('UPDATE RELANCE SET RR_MODELE'+inttostr(i)+'="RL'+inttostr(i)+'" WHERE RR_MODELE'+inttostr(i)+'<>""');

  SetParamSoc('SO_CPSUPPRESSIONDOC', True);
  SetParamSoc('SO_CPDOCAVECETAT', True);
end;


end.
