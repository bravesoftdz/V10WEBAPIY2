unit impFicU;

interface

uses
  SysUtils, Classes, Forms, Hctrls, Ent1,
  hmsgbox, HStatus, HEnt1, RappType,
{$IFNDEF EAGLSERVER}
  RapSuppr,
{$ENDIF}

{$IFNDEF EAGLCLIENT}
  Db,         // YMO 24/03/2006 Db et Dbtables dans la directive car impficu rajouté à eCCS5
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  TImpFic,FileCtrl, SaisComm, Utob ;


Type TInfoGeneSISCO= Record
                     ExoSISCO      : TExoDate ;
                     QuelDev       : TSorteMontant ;
                     StPlusExo     : String ; // + si suivant, - si N-1, -- si N-2, --- si n-3 etc...
                     CodeSoc,Nom,Adr1,Adr2,Adr3,Siret,APE,Tel : String ;
                     AnaSurSISCO : Boolean ;
                     ShuntAnaCharge,ShuntAnaProduit : Boolean ;
                     FourchetteImport : TFourchetteImport ;
                     FourchetteSISCO : TTabCptCollSISCO ;
                     LJalLu : HTStringList ;
                     QFiche : TQfiche ;
                     CharRG : TCharRemplace ;
                     CharRT : TCharRemplace ;
                     CharRS : TCharRemplace ;
                     CharRJ : TCharRemplace ;
                     EnEuro : Boolean ;
                     MRT,RGT : String ;
                     InitPCL : Boolean ; // Pour récup PCL
                     OkEnr : tOkEnrSISCO ; // Pour récup PCL
                     AuMoinsUnClient,AuMoinsUnFournisseur : Boolean ; // Pour récup PCL
                     PbEnr : tOkEnrSISCO ; // Pour récup PCL
                     ListePbEcrPCL : HTStringList ; // Pour récup PCL
                     // ajout me
                     lgcpt : string;
                     Numplan : string;
                     Troncaux : Boolean; // ajout me 12-03-2002
                     End ;
Type TJalSISCO     = Record
                     Code,Libelle,Ctr : String ;
                     End ;
{$IFNDEF EAGLSERVER}
{$IFNDEF EAGLCLIENT}
procedure InitMvtImport(Var MvtImport : TFMvtImport) ;
Procedure CreateListeImp(Var InfoImp :TInfoImport) ;
procedure VerifExisteCptLu(Var InfoImp : TInfoImport ; Bourre : Boolean);
Procedure EcrireRapportDoublons(Var FicRap : TextFile ; Var InfoImp : TInfoImport) ;
Procedure EcrireRapportRejets(Var FicRap : TextFile ; Var InfoImp : TInfoImport) ;
Procedure AlimLTabDiv(What : Integer ; L : HTStringList) ;
Procedure AjouteLTabDiv(What : Integer ; L : HTStringList ; St1,St2 : String) ;
Function LeModePaie(L : HTStringList ; Var CptLu : TCptLu ; ChercheSurCat : Boolean = TRUE) : Boolean ;
//function BourreOuTronque(St : String ; fb : TFichierBase) : String ;
Procedure AlimTotauxCptLu(Var InfoImp : TInfoImport ; Var IdentPiece : TIdentPiece) ;
Procedure MajSoldeCompteImport(Var InfoImp : TInfoImport) ;
Procedure RecToClass(Var C1 : TCptLu ; C2 : TCptLuP) ;
Function TraiteImportCompte(St : String ; Var InfoImp : TInfoImport ; QFiche : TQFiche) : tResultImportCpte ;
Function EstUneLigneCptOk(St,Code : String) : Boolean ;
FUNCTION FORMAT1_STRING (st : string; l : Integer) : string ;
Procedure DecomposeNomFic(NomFic : String ; Var LePath,LeNom,Lextension : String) ;
Procedure ChargeScenario(Import,Format,Lequel,CodeFormat : String ; Var Scenario : TScenario ; Auto : Boolean) ;
Procedure CreateFicRapport(InfoImp : TInfoImport) ;
Procedure CreateFicRapportGlobal(InfoImp : TInfoImport) ;
Procedure EcrireRapportDivers(What : Integer ; InfoImp : TInfoImport) ;
Procedure EcrireRapportGlobal(Var InfoImp : TInfoImport ; PbFic : Boolean ; StErr : String = '' ; OkSep : Boolean = TRUE) ;
Procedure EcrireSeparateur(Var FicRap : TextFile) ;
Procedure EcrireRapportDebutIntegration(PasDeMvtIntegres : Boolean ; InfoImp : TInfoImport ; Var FicRap : TextFile ; Var OkRapport : Boolean) ;
Procedure DecodeLeRIB ( Var Etab,Guichet,Numero,Cle,Dom : String ; RIB : String ) ;
Procedure TraiteRibImport(CptAux,Rib,Ville,Pays,RibPrincipal : String ; What : Integer = 0 ) ;
Function EstFichierSISCO(StFichier : String ; OkMess : Boolean) : Boolean ;
Function NomFichierRecupSISCO(StFichier : String ; SuffSup : String = '') : String ;
Function NomFichierImportSU(StFichier : String) : String ;
Function ChercheRibAux(Aux : String ; Var Dom,Etab,Guichet,NumCpt,Cle : String) : Boolean ;
Procedure ChercheContactAux(Aux : String ; Var Ct : TContact) ;
Function FaitStCPTCEGID(Qui : Char ; Q : TQuery) : String ;
Procedure ExporteSousSection(Var F : TextFile ; NomFic : String ; OkAppend,OkMove : Boolean) ;
Function LibSousSection(Sect,Axe : String) : String ;
FUNCTION TravailleFichier ( Force : Boolean ; Var StErr : String ;
                            Var InfoImp : TInfoImport ; QFiche : TQFiche ; ignorel1 : boolean=False) : Boolean ;
procedure ChargeDev(Var InfoImp : TInfoImport) ;
procedure LettrageSurRegroupement(Var InfoImp : TInfoImport ; ChampR : String) ;
procedure ChargeDossierAuto ( Var FDossier : HTStrings ; VireRef : Boolean) ;
procedure FreeDossierAuto ( Var FDossier : HTStrings ) ;
Function ChargeUnDossier(StNom : String ; NomChamp : String) : String ;
Function TrouveNomDossier(CodeSoc,NomChamp : String) : String ;
Procedure ChargeLCodeSoc(NomChamp : String ; LCodeSoc : HtStringList) ;
Function LitLCodeSoc (LCodeSoc : HtStringList ; CodeSoc : String) : String ;
Function ModeJal(Jal : String ; Var InfoImp : TInfoImport) : tModeJal ;
Procedure AlimSoucheBor(LS : HtStringList) ;
Function SetIncNumSB(LS : HtStringList ; Jal : String ; DateP : TDateTime) : Integer;
Procedure TraiteODA(Var InfoImp :TInfoImport)  ;
Procedure TraiteEEXBQ(Var InfoImp :TInfoImport)  ;
Function CreateMPSISCO(CodeMP,LibMP,CatMP,AccMP : String) : Boolean ;
Function InitSQLCpt1(What : Integer ; LeCpt,S1,S2 : String ; Existe : Boolean) : String ;
Function ImportSect(St : String ; Var InfoImp : TInfoImport ; QFiche : TQFiche ; TL17 : Boolean = FALSE) : tResultImportCpte ;
Function ImportTL(St : String) : tResultImportCpte ;

{$ENDIF EAGLCLIENT}
{$ENDIF EAGLSERVER}

Function NewNomFic(NomFic,Suffixe : String) : String ;
Function NewNomFicEtDir(NomFic,Suffixe,SousDir : String) : String ;
Function  FichierOnDisk(NomFic : String ; OkMoveCur : Boolean) : Boolean ;
Function EstJalARecuperer(St,St1 : String) : Boolean ;
procedure VideStringList ( L : HTStringList ) ;
procedure VideListeInfoImp(Var InfoImp : TInfoImport ; OnFree : Boolean) ;
Function VerifEntier(St : String) : Boolean ;
Function VerifMontant(St : String) : Boolean ;

{$IFNDEF EAGLCLIENT}
Function  NatureCptImport(Cpt : String ; Var FI : TFourchetteImport) : String ;
Function LitTRFParam(LeCode,LeNom : String) : String ;
Procedure UpdateTRFParam(St,LeCode,LeNom : String) ;
Procedure ChargeFourchetteCompte(Qui : String ; Var FI : TFourchetteImport) ;
Procedure DechiffreFour(Var Four : TFCha ; Qui,St : String ; MaxE : Integer) ;
Procedure InitRequete(Var Q : TQuery ; What : Integer) ;
Function HalteLa : Boolean ;
function TraiteCodeStatSISCO(Quoi : String ; Var InfoImp :TInfoImport; bCreationSection : boolean)  : boolean;
Function TraiteCorrespCpt(Quoi : Integer ; QFiche : TQfiche ; Var CptI : TCptImport ; Var InfoImp : TInfoImport) : Boolean ;
Procedure SetCreerpar(Q : tQuery ; StWhat : String) ;
Function ChercheRibPrincipal(CptAux : String) : String ;
function  EncodeLeRIB ( Etab,Guichet,Numero,Cle,Dom : String ) : String ;
Procedure VerifUneSouche(CodeSouche : String ; Ana,Simu : Boolean) ;
{$ENDIF}


implementation

Uses
{$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  CPVersion,
{$ENDIF MODENT1}

{$IFNDEF EAGLCLIENT}
{$IFNDEF EAGLSERVER}
LettUtil, LetBatch,
MajTable,
{$ENDIF EAGLSERVER}
Saisutil,
{$ENDIF}
uLibAnalytique, // AlloueAxe
ParamSoc,
ed_Tools;

(****************************************************)
Function GetTenuEuro : Boolean;
begin
{$IFDEF NOVH}
        Result := GetParamsoc ('SO_TENUEEURO');
{$ELSE}
        Result := VH^.TenueEuro;
{$ENDIF}
end;

Function GetRecupPCL : Boolean;
begin
{$IFDEF NOVH}
        Result := TRUE; // Pour La Comsx Server
{$ELSE}
        Result := VH^.RecupPCL;
{$ENDIF}
end;

Function GetRecupComSx : Boolean;
begin
{$IFDEF NOVH}
        Result := TRUE; // Pour La Comsx Server
{$ELSE}
        Result := VH^.RecupComSx;
{$ENDIF}
end;

Function GetRecupSISCOPGI : Boolean;
begin
{$IFDEF NOVH}
        Result := FALSE; // Pour La Comsx Server
{$ELSE}
        Result := VH^.RecupSISCOPGI;
{$ENDIF}
end;

Function GetRecupLTL : Boolean;
begin
{$IFDEF NOVH}
        Result := FALSE; // Pour La Comsx Server
{$ELSE}
        Result := VH^.RecupLTL;
{$ENDIF}
end;

Function GetRecupCegid : Boolean;
begin
{$IFDEF NOVH}
        Result := FALSE; // Pour La Comsx Server
{$ELSE}
        Result := VH^.RecupCegid;
{$ENDIF}
end;

Function GetCPIFDEFCEGID : Boolean;
begin
{$IFDEF NOVH}
        Result := GetParamSocSecur('SO_IFDEFCEGID',False); // Pour La Comsx Server
{$ELSE}
        Result := VH^.CPIFDEFCEGID;
{$ENDIF}
end;

(****************************************************)

Function NewNomFic(NomFic,Suffixe : String) : String ;
Var LePath,LeNom,Lextension : String ;
    i : Integer ;
BEGIN
Result:=NomFic ; If Suffixe='' Then Exit ;
LePath:=ExtractFilePath(NomFic) ;
LExtension:=ExtractFileExt(NomFic) ;
LeNom:=ExtractFileName(NomFic) ;
i:=Pos('.',LeNom) ; If i>0 Then LeNom:=Copy(LeNom,1,i-1) ;
NomFic:=LePath+LeNom+Suffixe+LExtension ;
Result:=NomFic ;
END ;

Function NewNomFicEtDir(NomFic,Suffixe,SousDir : String) : String ;
Var LePath,LeNom,Lextension,LaSousDir : String ;
    i : Integer ;
    OkDir : Boolean ;
BEGIN
LePath:=ExtractFilePath(NomFic) ;
LExtension:=ExtractFileExt(NomFic) ;
LeNom:=ExtractFileName(NomFic) ;
i:=Pos('.',LeNom) ; If i>0 Then LeNom:=Copy(LeNom,1,i-1) ;
LaSousDir:=LePath+SousDir ; OkDir:=TRUE ;
If Not DirectoryExists(LaSousDir) Then OkDir:=CreateDir(LaSousDir) ;
If OkDir Then Result:=LaSousDir+'\'+LeNom+Suffixe+LExtension Else Result:='' ;
END ;

Function FichierOnDisk(NomFic : String ; OkMoveCur : Boolean) : Boolean ;
Var Fichier : TextFile ;
    FichierOk : Boolean ;
    i : Integer ;
BEGIN
i:=0 ; //FichierOk:=FALSE ;
Repeat
  Inc(i) ;
  If OkMoveCur Then MoveCur(FALSE) ;
  AssignFile(Fichier,NomFic) ;
  {$I-} Reset (Fichier) ; {$I+}
  FichierOk:=ioresult=0 ;
  If FichierOk Then Close(Fichier) ;
Until FichierOk Or (i>10000) ;
Result:=FichierOk ;
Delay(2000) ;
END ;
Function EstJalARecuperer(St,St1 : String) : Boolean ;
BEGIN
Result:=TRUE ; If St1='' Then Exit ;
If GetRecupSISCOPGI Then
  BEGIN
  St:=AnsiUpperCase(St) ; St1:=AnsiUpperCase(St1) ;
  If Pos(St,St1)>0 Then Result:=FALSE ;
  END  Else
  BEGIN
  {$IFDEF RECUPPCL}
  if GetCPIFDEFCEGID then
  Begin
       St:=AnsiUpperCase(St) ; St1:=AnsiUpperCase(St1) ;
       If Pos(St,St1)>0 Then Result:=FALSE ;
  end ;
  {$ENDIF}
  END ;
END ;

procedure VideStringList ( L : HTStringList ) ;
Var i : integer ;
BEGIN
if L=Nil then Exit ; if L.Count<=0 then Exit ;
for i:=0 to L.Count-1 do If L.Objects[i]<>NIL Then L.Objects[i].Free ;
L.Clear ;
END ;

procedure VideListeInfoImp(Var InfoImp : TInfoImport ; OnFree : Boolean) ;
BEGIN
VideStringList(InfoImp.LGenLu) ; If OnFree then InfoImp.LGenLu.Free ;
VideStringList(InfoImp.LAuxLu) ; If OnFree then InfoImp.LAuxLu.Free ;
VideStringList(InfoImp.LAnaLu) ; If OnFree then InfoImp.LAnaLu.Free ;
VideStringList(InfoImp.LJalLu) ; If OnFree then InfoImp.LJalLu.Free ;
VideStringList(InfoImp.LMP) ; If OnFree then InfoImp.LMP.Free ;
VideStringList(InfoImp.LMR) ; If OnFree then InfoImp.LMR.Free ;
VideListe(InfoImp.ListeCptFaux) ; If OnFree then InfoImp.ListeCptFaux.Free ;
VideListe(InfoImp.ListePieceFausse) ; If OnFree then InfoImp.ListePieceFausse.Free ;
VideListe(InfoImp.CRListeEnteteDoublon) ; If OnFree then InfoImp.CRListeEnteteDoublon.Free ;
VideStringList(InfoImp.ListeEnteteDoublon) ; If OnFree then InfoImp.ListeEnteteDoublon.Free ;
VideStringList(InfoImp.ListePieceIntegre) ; If OnFree then InfoImp.ListePieceIntegre.Free ;
If OnFree Then InfoImp.ListeEntetePieceFausse.Free ;
VideStringList(InfoImp.ListePbAna) ; If OnFree then InfoImp.ListePbAna.Free ;
VideStringList(InfoImp.LSoucheBOR) ; If OnFree then InfoImp.LSoucheBOR.Free ;
VideStringList(InfoImp.LCS) ; If OnFree then InfoImp.LCS.Free ;
if InfoImp.TobStat<> nil then
begin
  if InfoImp.TOBStat.Detail.Count < 0 then InfoImp.TOBStat.ClearDetail;
If OnFree then InfoImp.TOBStat.Free;
end;
END ;

Function VerifEntier(St : String) : Boolean ;
Var i : Integer ;
BEGIN
Result:=TRUE ;
For i:=1 To length(St) Do
  BEGIN
  If (St[i] in ['0'..'9'])=FALSE Then BEGIN Result:=FALSE ; Exit ; END ;
  END ;
END ;

Function VerifMontant(St : String) : Boolean ;
Var i : Integer ;
BEGIN
Result:=TRUE ;
For i:=1 To length(St) Do
  BEGIN
  If (St[i] in ['0'..'9','.',',','-'])=FALSE Then BEGIN Result:=FALSE ; Exit ; END ;
  END ;
END ;



{$IFNDEF EAGLSERVER}
{$IFNDEF EAGLCLIENT}


procedure InitMvtImport(Var MvtImport : TFMvtImport) ;
BEGIN
Fillchar(MvtImport,SizeOf(MvtImport),#0) ;
With MvtImport do
  BEGIN
  IE_ETATLETTRAGE:='RI'; IE_ECRANOUVEAU:='N';
  IE_ETABLISSEMENT := GetParamsocSecur('SO_ETABLISDEFAUT','001', false); //RR FQ17098 GetParamsoc('SO_ETABLISDEFAUT', False);
  IE_QUALIFPIECE:='N';
  IE_SOCIETE:=V_PGI.CodeSociete; IE_DEVISE:=V_PGI.DevisePivot;
  IE_NUMPIECE:=1 ; IE_DATEECHEANCE:=iDate1900 ;IE_DATEPAQUETMAX:=iDate1900 ;
  IE_DATEPAQUETMIN:=iDate1900 ; IE_DATEPOINTAGE:=iDate1900 ; IE_DATEREFEXTERNE:=iDate1900 ;
  IE_DATERELANCE:=iDate1900 ; IE_DATETAUXDEV:=iDate1900 ; IE_DATEVALEUR:=iDate1900 ;
  IE_DATECREATION:=Date ;
  IE_ORIGINEPAIEMENT:=iDate1900 ; IE_DATECOMPTABLE:=iDate1900 ; IE_LIBREDATE:=iDate1900 ;
  IE_ECHE:='-' ;IE_ENCAISSEMENT:='RIE';IE_CONTROLE:='-' ;IE_LETTRAGEDEV:='-' ;
  IE_OKCONTROLE:='X';IE_SELECTED:='X' ; IE_TVAENCAISSEMENT:='-' ; IE_TYPEANALYTIQUE:='-' ;
  IE_VALIDE:='-' ; IE_ANA:='-' ;IE_INTEGRE:='-' ; IE_QUOTITE:=1 ;
  IE_LIBREBOOL0:='-' ; IE_LIBREBOOL1:='-' ;
  IE_ELEMENTARECUPERER:=FALSE ; IE_COTATION:=1 ; IE_CODEACCEPT:='NON' ;
  END ;
END ;



Procedure CreateListeImp(Var InfoImp :TInfoImport) ;
BEGIN
InfoImp.LGenLu:=HTStringList.Create ;
InfoImp.LAuxLu:=HTStringList.Create ;
InfoImp.LAnaLu:=HTStringList.Create ;
InfoImp.LJalLu:=HTStringList.Create ;
InfoImp.LMP:=HTStringList.Create ;
InfoImp.LMR:=HTStringList.Create ;
InfoImp.LRGT:=HTStringList.Create ;
InfoImp.ListeCptFaux:=TList.Create ;
InfoImp.ListePieceFausse:=TList.Create ;
InfoImp.ListeEntetePieceFausse:=HTStringList.Create ;
InfoImp.ListeEnteteDoublon:=HTStringList.Create ;
InfoImp.ListePieceIntegre:=HTStringList.Create ;
InfoImp.CRListeEnteteDoublon:=TList.Create ;
InfoImp.ListePbAna:=HTStringList.Create ;
InfoImp.LSoucheBOR:=HTStringList.Create ;
InfoImp.LCS:=HTStringList.Create ;
InfoImp.LCS.Sorted:=TRUE ; InfoImp.LCS.duplicates:=DupIgnore ;
InfoImp.TOBStat := TOB.Create ('', nil, -1);
END ;


Procedure DecomposeNomFic(NomFic : String ; Var LePath,LeNom,Lextension : String) ;
//Var i : Integer ;
BEGIN
LePath:=ExtractFilePath(NomFic) ;
LExtension:=ExtractFileExt(NomFic) ;
LeNom:=ExtractFileName(NomFic) ;
END ;


Function NomFichierRecupSISCO(StFichier : String ; SuffSup : String = '') : String ;
Var NomFichier : String ;
BEGIN
NomFichier:=NewNomFicEtDir(StFichier,'Mvt','SISCO') ;
NomFichier:=NewNomFic(NomFichier,'CGE'+SuffSup) ;
Result:=NomFichier ;
END ;

Function NomFichierImportSU(StFichier : String) : String ;
Var NomFichier : String ;
BEGIN
NomFichier:=NewNomFic(StFichier,'CGE') ;
Result:=NomFichier ;
END ;



Procedure StFicRapport(Var St : String ; B : Boolean ; St1 : String) ;
BEGIN
If St1='#' Then
  BEGIN
  If B Then St:=St+'Actif' Else St:=St+'Non Actif' ;
  END Else
  BEGIN
  If Trim(St1)<>'' Then St:=St+St1 Else St:=St+'Non Actif'
  END ;
END ;

Procedure CreateFicRapportGlobal(InfoImp : TInfoImport) ;
Var FicRap : TextFile ;
    St : String ;
BEGIN
If InfoImp.NomFicRapportGlobal='' Then Exit ;
//If InfoImp.NomFic='' Then Exit ;
AssignFile(FicRap,InfoImp.NomFicRapportGlobal) ;
If GetRecupPCL Then
  BEGIN
  {$i-} Append(FicRap) ; {$i+}
  If IoResult<>0 Then {$i-} Rewrite(FicRap) ; {$i+}
  END Else {$i-} Rewrite(FicRap) ; {$i+}
If IoResult<>0 Then Exit ;
If GetRecupPCL Then BEGIN St:='' ; Writeln(FicRap,St) ; END Else
  BEGIN
  St:='--------------------------------------------------------------------' ; Writeln(FicRap,St) ;
  St:=TraduireMemoire('   RAPPORT D''IMPORTATION DU '+DateToStr(Date)) ; Writeln(FicRap,St) ;
  If InfoImp.SC.ImpAuto Then
    BEGIN
    St:=TraduireMemoire('   -> IMPORTATION AUTOMATIQUE ') ; Writeln(FicRap,St) ;
    END ;
  St:='--------------------------------------------------------------------' ; Writeln(FicRap,St) ;
  St:=' ' ; Writeln(FicRap,St) ;
  END ;
CloseFile(FicRap) ;
END ;

Procedure CreateFicRapport(InfoImp : TInfoImport) ;
Var FicRap : TextFile ;
    St : String ;
BEGIN
If InfoImp.NomFicRapport='' Then Exit ;
If InfoImp.NomFic='' Then Exit ;
AssignFile(FicRap,InfoImp.NomFicRapport) ;
{$i-} Rewrite(FicRap) ; {$i+}  If IoResult<>0 Then Exit ;
St:=TraduireMemoire(' ') ; Writeln(FicRap,St) ;
St:='--------------------------------------------------------------------' ; Writeln(FicRap,St) ;
St:=TraduireMemoire('   IMPORTATION DU FICHIER ')+InfoImp.NomFic ; Writeln(FicRap,St) ;
If InfoImp.SC.ImpAuto Then
  BEGIN
  St:=TraduireMemoire('   -> IMPORTATION AUTOMATIQUE ') ; Writeln(FicRap,St) ;
  END ;
St:='--------------------------------------------------------------------' ; Writeln(FicRap,St) ;
St:=' ' ; Writeln(FicRap,St) ;
If InfoImp.Sc.EstCharge Then
  BEGIN
  St:=' Options de scénarios :  ' ; Writeln(FicRap,St) ;
  St:=' Substitution du collectif pour les tiers clients      : ' ; StFicRapport(St,InfoImp.Sc.RCollCliT<>'','#') ; Writeln(FicRap,St) ;
  St:=' Substitution du collectif pour les tiers Fournisseurs : ' ; StFicRapport(St,InfoImp.Sc.RCollFouT<>'','#') ; Writeln(FicRap,St) ;
  St:=' Recalcul de la TVA                                    : ' ; StFicRapport(St,InfoImp.Sc.TraiteTva,'#') ; Writeln(FicRap,St) ;
  St:=' Recalcul des contreparties                            : ' ; StFicRapport(St,InfoImp.Sc.TraiteCtr,'#') ; Writeln(FicRap,St) ;
  St:=' Affectation du RIB principal des clients              : ' ; StFicRapport(St,InfoImp.Sc.ForceRIB,'#')  ; Writeln(FicRap,St) ;
  St:=' Affectation du RIB principal des fournisseurs         : ' ; StFicRapport(St,InfoImp.Sc.ForceRIBFou,'#') ; Writeln(FicRap,St) ;
  St:=' Récupération des A-Nouveaux en détail                 : ' ; StFicRapport(St,InfoImp.Sc.ANDetail,'#')  ; Writeln(FicRap,St) ;
  St:=' Validation des pièces importées                       : ' ; StFicRapport(St,InfoImp.Sc.ValideEcr,'#') ; Writeln(FicRap,St) ;
  St:=' Recalcul du taux devise out                           : ' ; StFicRapport(St,InfoImp.Sc.CalcTauxDevOut,'#') ; Writeln(FicRap,St) ;
  St:=' Mode de paiement de remplacement                      : ' ; StFicRapport(St,FALSE,InfoImp.Sc.RMP)     ; Writeln(FicRap,St) ;
  St:=' Compte général de remplacement                        : ' ; StFicRapport(St,FALSE,InfoImp.Sc.Rgen)    ; Writeln(FicRap,St) ;
  St:=' Tiers Client de remplacement                          : ' ; StFicRapport(St,FALSE,InfoImp.Sc.RCli)    ; Writeln(FicRap,St) ;
  St:=' Tiers Fournisseur de remplacement                     : ' ; StFicRapport(St,FALSE,InfoImp.Sc.RFou)    ; Writeln(FicRap,St) ;
  St:=' Tiers Salarié de remplacement                         : ' ; StFicRapport(St,FALSE,InfoImp.Sc.RSal)    ; Writeln(FicRap,St) ;
  St:=' Tiers Divers de remplacement                          : ' ; StFicRapport(St,FALSE,InfoImp.Sc.RDiv)    ; Writeln(FicRap,St) ;
  St:=' Section Axe 1 de remplacement                         : ' ; StFicRapport(St,FALSE,InfoImp.Sc.RSect1)  ; Writeln(FicRap,St) ;
  St:=' Section Axe 2 de remplacement                         : ' ; StFicRapport(St,FALSE,InfoImp.Sc.RSect2)  ; Writeln(FicRap,St) ;
  St:=' Section Axe 3 de remplacement                         : ' ; StFicRapport(St,FALSE,InfoImp.Sc.RSect3)  ; Writeln(FicRap,St) ;
  St:=' Section Axe 4 de remplacement                         : ' ; StFicRapport(St,FALSE,InfoImp.Sc.RSect4)  ; Writeln(FicRap,St) ;
  St:=' Section Axe 5 de remplacement                         : ' ; StFicRapport(St,FALSE,InfoImp.Sc.RSect5)  ; Writeln(FicRap,St) ;
  St:=' Plan de correspondance sur comptes généraux           : ' ; StFicRapport(St,InfoImp.Sc.CorrespGen,'#') ; Writeln(FicRap,St) ;
  St:=' Plan de correspondance sur tiers                      : ' ; StFicRapport(St,InfoImp.Sc.CorrespAux,'#') ; Writeln(FicRap,St) ;
  St:=' Plan de correspondance sur section axe 1              : ' ; StFicRapport(St,InfoImp.Sc.CorrespSect1,'#') ; Writeln(FicRap,St) ;
  St:=' Plan de correspondance sur section axe 2              : ' ; StFicRapport(St,InfoImp.Sc.CorrespSect2,'#') ; Writeln(FicRap,St) ;
  St:=' Plan de correspondance sur section axe 3              : ' ; StFicRapport(St,InfoImp.Sc.CorrespSect3,'#') ; Writeln(FicRap,St) ;
  St:=' Plan de correspondance sur section axe 4              : ' ; StFicRapport(St,InfoImp.Sc.CorrespSect4,'#') ; Writeln(FicRap,St) ;
  St:=' Plan de correspondance sur section axe 5              : ' ; StFicRapport(St,InfoImp.Sc.CorrespSect5,'#') ; Writeln(FicRap,St) ;
  St:=' Plan de correspondance sur mode de paiement           : ' ; StFicRapport(St,InfoImp.Sc.CorrespMP,'#')  ; Writeln(FicRap,St) ;
  St:=' Plan de correspondance sur journaux                   : ' ; StFicRapport(St,InfoImp.Sc.CorrespJal,'#') ; Writeln(FicRap,St) ;
  St:=' Mode de règlement de remplacement (Création tiers)    : ' ; StFicRapport(St,FALSE,InfoImp.Sc.RMR)      ; Writeln(FicRap,St) ;
  St:=' Régime de TVA de remplacement (Création tiers)        : ' ; StFicRapport(St,FALSE,InfoImp.Sc.RGT)      ; Writeln(FicRap,St) ;
  If InfoImp.SC.ImpAuto Then
    BEGIN
    St:=' Chemin de recherche des fichiers                      : '+InfoImp.SC.Chemin ; Writeln(FicRap,St) ;
    St:=' Préfixe des fichiers à importer                       : '+InfoImp.SC.Prefixe ; Writeln(FicRap,St) ;
    St:=' Suffixe des fichiers à importer                       : '+InfoImp.SC.Suffixe ; Writeln(FicRap,St) ;
    St:=' Contrôle des doublons                                 : ' ; StFicRapport(St,InfoImp.Sc.Doublon,'#') ; Writeln(FicRap,St) ;
    St:=' Forcer le numéro de pièce                             : ' ; StFicRapport(St,InfoImp.Sc.ForcePiece,'#') ; Writeln(FicRap,St) ;
    St:=' Détruire le fichier origine                           : ' ; StFicRapport(St,InfoImp.Sc.DetruitFic,'#') ; Writeln(FicRap,St) ;
    END ;
  END Else
  BEGIN
  St:=' Pas de scénario actif. ' ; Writeln(FicRap,St) ;
  END ;
CloseFile(FicRap) ;
END ;

FUNCTION FORMAT1_STRING (st : string; l : Integer) : string ;
var St1 : String ;
BEGIN
St1:=St ;
if ((l>0) and (l<Length(St1))) then St1:=Copy(St1,1,l) ;
While Length(St1)<l Do St1:=St1+' ' ;
Result:=st1 ;
END ;

Procedure RecToClass(Var C1 : TCptLu ; C2 : TCptLuP) ;
BEGIN
With C1,C2 Do
  BEGIN
  T.Ok                   :=Ok;
  T.Cpt                  :=Cpt;
  T.Collectif            :=Collectif;
  T.CollTiers            :=CollTiers;
  T.Ventilable           :=Ventilable ;
  T.VentilableDansFichier:=VentilableDansFichier ;
  T.Nature               :=Nature;
  T.RegimeTva            :=RegimeTva;
  T.Pointable            :=Pointable ;
  T.Lettrable            :=Lettrable;
  T.SoucheN              :=SoucheN ;
  T.SoucheS              :=SoucheS ;
  T.Axe                  :=Axe ;
  T.LastDate             :=LastDate ;
  T.LastNum              :=LastNum ;
  T.LastNumL             :=LastNumL;
  T.DebitMvt             :=DebitMvt ;
  T.CreditMvt            :=CreditMvt ;
  T.TotDN                :=TotDN ;
  T.TotCN                :=TotCN ;
  T.TotDN1               :=TotDN1 ;
  T.TotCN1               :=TotCN1 ;
  T.TotDAno              :=TotDAno ;
  T.TotCAno              :=TotCAno ;
  T.Tva                  :=Tva ;
  T.TPF                  :=TPF ;
  T.OkMajSolde           :=OkMajSolde ;
  T.TvaEncHT             :=TvaEncHT ;
  T.TvaEnc               :=TvaEnc ;
  T.TP                   :=TP ;
  T.IsTP                 :=IsTP ;
  T.modeSaisie           :=ModeSaisie ;
  T.qualifPiece          :=QualifPiece ;
  T.EstColl              :=EstColl ;
  END ;
END ;

(*function BourreOuTronque(St : String ; fb : TFichierBase) : String ;
var Diff : Integer ;
BEGIN
St:=Trim(St) ; Result:=St ; if (St='') then Exit ;
Diff:=Length(St)-LongMax(fb) ;
if Diff>0 then St:=BourreLess(St,fb) ;
Result:=BourreLaDonc(St,fb) ;
END ;
*)



Function VerifExiste(What : Integer ; ListeCptFaux : TList ; L : HTStringList ; Bourre : Boolean) : Integer ;
Var Q : Tquery ;
    i,LeMax : Integer ;
    St,Cpt,MessSup : String ;
    X : DelInfo ;
    fb : TFichierBase ;
    LeCptLuP : TCptLuP ;
    LeCptLu : TCptLu ;

BEGIN
Result:=0 ; fb:=fbgene ;
Case What Of
  0 : BEGIN
      St:='Select G_GENERAL FROM GENERAUX WHERE G_GENERAL=:CPT' ; fb:=fbgene ;
      END ;
  1 : BEGIN
      St:='Select T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE=:CPT' ; fb:=fbAux ;
      END ;
  2 : BEGIN
      St:='Select S_SECTION FROM SECTION WHERE S_SECTION=:CPT AND S_AXE="A1"' ; fb:=fbAxe1 ;
      END ;
  END ;
Q:=prepareSQL(St,TRUE) ; LeMax:=L.Count-1 ;
InitMove(LeMax,'') ;
for i:=0 to MaxCptLu do
    BEGIN
    MoveCur(FALSE) ;
    LeCptLuP:=TCptLuP(L.Objects[i]) ; Cpt:=LeCptLuP.T.Cpt ; MessSup:='' ;
    If Bourre Then Cpt:=BourreOuTronque(Cpt,fb) ;
    Q.Close ;
    Q.Params[0].AsString:=Cpt ;
    Q.Open ;
    if Q.Eof then
       BEGIN
       Inc(Result) ;
       X:=DelInfo.Create ; X.Gen:='' ; X.aux:='' ; X.Sect:='' ;
       X.GenACreer:=FALSE ; X.AuxACreer:=FALSE ; X.SectACreer:=FALSE ;
       Case What Of
          0 : BEGIN
              X.Gen:=Cpt ; X.GenACreer:=TRUE ;
              END ;
          1 : BEGIN
              X.Aux:=Cpt ; X.AuxACreer:=TRUE ;
              END ;
          2 : BEGIN
              X.Sect:=Cpt ; X.SectACreer:=TRUE ;
              END ;
          END ;
       ClassToRec(LeCptlu,LeCptLuP) ;
       MessSup:=MessErreurCompte(What,LeCptLu) ;
       X.LeCod:=Cpt ; X.LeLib:=MessSup ;
       ListeCptFaux.Add(X);
       END ;
    END ;
Ferme(Q) ;
FiniMove ;
END ;

Procedure AlimLTabDiv(What : Integer ; L : HTStringList) ;
Var Q : TQuery ;
    LeCptLuP : TCptLuP ;
    St : String ;
BEGIN
L.Clear ;
Case What Of
  0 : St:='Select MP_MODEPAIE,MP_CATEGORIE FROM MODEPAIE' ;
  1 : St:='Select MR_MODEREGLE FROM MODEREGL' ;
  2 : St:='Select CC_CODE FROM CHOIXCOD WHERE CC_TYPE="RTV"' ;
  END ;
Q:=OpenSQL(St,TRUE) ;
While Not Q.Eof Do
  BEGIN
  Case What Of
    0 : BEGIN
        LeCptLuP:=TCptLuP.Create ;
        LeCptLuP.T.Cpt:=Trim(Q.Fields[0].AsString) ;
        LeCptLuP.T.Nature:=Trim(Q.Fields[1].AsString) ;
        L.AddObject(Q.Fields[0].AsString,LeCptLuP) ;
        END ;
    Else L.Add(Q.Fields[0].AsString) ;
    END ;
  Q.Next ;
  END ;
Ferme(Q) ;
END ;

Procedure AjouteLTabDiv(What : Integer ; L : HTStringList ; St1,St2 : String) ;
Var LeCptLuP : TCptLuP ;
BEGIN
Case What Of
  0 : BEGIN
      LeCptLuP:=TCptLuP.Create ;
      LeCptLuP.T.Cpt:=Trim(St1) ;
      LeCptLuP.T.Nature:=Trim(St2) ;
      L.AddObject(St1,LeCptLuP) ;
      END ;
  Else L.Add(St1) ;
  END ;
END ;


Function LeModePaie(L : HTStringList ; Var CptLu : TCptLu ; ChercheSurCat : Boolean = TRUE) : Boolean ;
Var LeCptLuP : TCptLuP ;
    i : Integer ;
BEGIN
Result:=True;
If CptLu.Cpt<>'' Then Result:=ChercheCptLu(L,CptLu) Else If ChercheSurCat Then
   BEGIN
   Result:=FALSE ;
   For i:=0 To L.Count-1 Do
     BEGIN
     LeCptLuP:=TCptLuP(L.Objects[i]) ;
     If LeCptLuP.T.Nature=CptLu.Nature Then BEGIN Result:=TRUE ; ClassToRec(CptLu,LeCptLuP) ; Exit ; END ;
     END ;
   END ;
END ;


{=============================================================================}
procedure VerifExisteCptLu(Var InfoImp : TInfoImport ; Bourre : Boolean);
Var GGRien : Boolean ;
    ListeCptFaux : TList ;
begin
ListeCptFaux:=TList.Create ;
ListeCptFaux.Clear ; InfoImp.NbGenFaux:=0 ; InfoImp.NbAuxFaux:=0 ; InfoImp.NbAnaFaux:=0 ;
InfoImp.NbGenFaux:=VerifExiste(0,ListeCptFaux,InfoImp.LGenLu,Bourre) ;
InfoImp.NbAuxFaux:=VerifExiste(1,ListeCptFaux,InfoImp.LAuxLu,Bourre) ;
InfoImp.NbAnaFaux:=VerifExiste(2,ListeCptFaux,InfoImp.LAnaLu,Bourre) ;
GGRien:=false ;
If (InfoImp.NbGenFaux>0) Or (InfoImp.NbAuxFaux>0) Or (InfoImp.NbAnaFaux>0) Then
   BEGIN
   RapportdErreurMvt(ListeCptFaux,8,GGRien,False) ;
   END ;
ListeCptFaux.Clear ; ListeCptFaux.Free ;
end;

{=============================================================================}

Procedure AlimTotalCptLu(L : HTStringList ; Cpt : String ; Var IdentPiece : TIdentPiece) ;
Var LeCptLuP : TCptLuP ;
    i : Integer ;
    SurN : Boolean ;
BEGIN
i:=L.IndexOf(Cpt) ; If Trim(Cpt)='' Then Exit ;
SurN:=TRUE ; If (IdentPiece.DateP>=VH^.Suivant.Deb) and (IdentPiece.DateP<=VH^.Suivant.Fin) then SurN:=FALSE ;
if i>-1 then
  BEGIN
  LeCptLuP:=TCptLuP(L.Objects[i]) ;
  LeCPtLuP.T.OkMajSolde:=TRUE ;
  LeCptLuP.T.LastDate:=IdentPiece.DateP ;
  LeCptLuP.T.LastNum:=IdentPiece.NumPDef ;
  LeCptLuP.T.LastNumL:=IdentPiece.LigneEnCours.NumLig ;
  LeCptLuP.T.DebitMvt:=IdentPiece.LigneEnCours.Debit ;
  LeCptLuP.T.CreditMvt:=IdentPiece.LigneEnCours.Credit ;
  If SurN Then
     BEGIN
     LeCptLuP.T.TotDN:=LeCptLuP.T.TotDN+IdentPiece.LigneEnCours.Debit ;
     LeCptLuP.T.TotCN:=LeCptLuP.T.TotCN+IdentPiece.LigneEnCours.Credit ;
     END Else
     BEGIN
     LeCptLuP.T.TotDN1:=LeCptLuP.T.TotDN1+IdentPiece.LigneEnCours.Debit ;
     LeCptLuP.T.TotCN1:=LeCptLuP.T.TotCN1+IdentPiece.LigneEnCours.Credit ;
     END ;
  If IdentPiece.LigneEnCours.ANouveau Then
    BEGIN
     LeCptLuP.T.TotDAno:=LeCptLuP.T.TotDAno+IdentPiece.LigneEnCours.Debit ;
     LeCptLuP.T.TotCAno:=LeCptLuP.T.TotCAno+IdentPiece.LigneEnCours.Credit ;
    END ;
  END ;
END ;

{=============================================================================}
Procedure AlimTotauxCptLu(Var InfoImp : TInfoImport ; Var IdentPiece : TIdentPiece) ;
BEGIN
If IdentPiece.QualP<>'N' Then Exit ;
If (InfoImp.LgenLu<>Nil) And (IdentPiece.LigneEnCours.Gen<>'') Then
  If (Not IdentPiece.LigneEnCours.ODAnal) And (Not IdentPiece.LigneEnCours.Ana) Then AlimTotalCptLu(InfoImp.LGenLu,IdentPiece.LigneEnCours.Gen,IdentPiece) ;
If (InfoImp.LAuxLu<>Nil) And (IdentPiece.LigneEnCours.Aux<>'') Then
  If Not IdentPiece.LigneEnCours.ODAnal Then AlimTotalCptLu(InfoImp.LAuxLu,IdentPiece.LigneEnCours.Aux,IdentPiece) ;
If (InfoImp.LAnaLu<>Nil) And (IdentPiece.LigneEnCours.Sect<>'') Then AlimTotalCptLu(InfoImp.LAnaLu,IdentPiece.LigneEnCours.Sect,IdentPiece) ;
If (InfoImp.LJalLu<>Nil) And (IdentPiece.JalP<>'') Then
  If (Not (IdentPiece.LigneEnCours.Sect<>'')) Or (IdentPiece.LigneEnCours.ODAnal) Then AlimTotalCptLu(InfoImp.LJalLu,IdentPiece.JalP,IdentPiece) ;
END ;

{=============================================================================}
Function FabricReqCpt(fb : TFichierBase) : TQuery ;
Var Q   : TQuery ;
    SQL : String ;
BEGIN
Q:=TQuery.Create(Application) ;
Q.DatabaseName:=DBSOC.DataBaseName ;
Q.SessionName:=DBSOC.SessionName ;
Q.SQL.Clear ;
Case fb of
   fbGene : BEGIN
            SQL:='UPDATE GENERAUX SET G_TOTALDEBIT=G_TOTALDEBIT+:DD, G_TOTALCREDIT=G_TOTALCREDIT+:CD, ' ;
            Q.SQL.Add(SQL) ;
            SQL:='G_TOTDEBE=G_TOTDEBE+:DE, G_TOTCREE=G_TOTCREE+:CE, G_TOTDEBS=G_TOTDEBS+:DS, G_TOTCRES=G_TOTCRES+:CS, ' ;
            Q.SQL.Add(SQL) ;
            SQL:='G_DATEDERNMVT=:DATEDM, G_DEBITDERNMVT=:DEBITDM, G_CREDITDERNMVT=:CREDITDM, G_NUMDERNMVT=:NUMPDM, G_LIGNEDERNMVT=:NUMLDM, ' ;
            Q.SQL.Add(SQL) ;
            SQL:='G_TOTDEBANO=G_TOTDEBANO+:DA, G_TOTCREANO=G_TOTCREANO+:CA WHERE G_GENERAL=:CPTE ' ;
            Q.SQL.Add(SQL) ;
            END ;
    fbAux : BEGIN
            SQL:='UPDATE TIERS SET T_TOTALDEBIT=T_TOTALDEBIT+:DD, T_TOTALCREDIT=T_TOTALCREDIT+:CD, ' ;
            Q.SQL.Add(SQL) ;
            SQL:='T_TOTDEBE=T_TOTDEBE+:DE, T_TOTCREE=T_TOTCREE+:CE, T_TOTDEBS=T_TOTDEBS+:DS, T_TOTCRES=T_TOTCRES+:CS, ' ;
            Q.SQL.Add(SQL) ;
            SQL:='T_DATEDERNMVT=:DATEDM, T_DEBITDERNMVT=:DEBITDM, T_CREDITDERNMVT=:CREDITDM, T_NUMDERNMVT=:NUMPDM, T_LIGNEDERNMVT=:NUMLDM, ' ;
            Q.SQL.Add(SQL) ;
            SQL:='T_TOTDEBANO=T_TOTDEBANO+:DA, T_TOTCREANO=T_TOTCREANO+:CA WHERE T_AUXILIAIRE=:CPTE ' ;
            Q.SQL.Add(SQL) ;
            END ;
   fbAxe1..fbAXe5 :
            BEGIN
            SQL:='UPDATE SECTION SET S_TOTALDEBIT=S_TOTALDEBIT+:DD, S_TOTALCREDIT=S_TOTALCREDIT+:CD, ' ;
            Q.SQL.Add(SQL) ;
            SQL:='S_TOTDEBE=S_TOTDEBE+:DE, S_TOTCREE=S_TOTCREE+:CE, S_TOTDEBS=S_TOTDEBS+:DS, S_TOTCRES=S_TOTCRES+:CS, ' ;
            Q.SQL.Add(SQL) ;
            SQL:='S_DATEDERNMVT=:DATEDM, S_DEBITDERNMVT=:DEBITDM, S_CREDITDERNMVT=:CREDITDM, S_NUMDERNMVT=:NUMPDM, S_LIGNEDERNMVT=:NUMLDM, ' ;
            Q.SQL.Add(SQL) ;
            SQL:='S_TOTDEBANO=S_TOTDEBANO+:DA, S_TOTCREANO=S_TOTCREANO+:CA WHERE S_AXE=:AXE AND S_SECTION=:CPTE ' ;
            Q.SQL.Add(SQL) ;
            END ;
   fbJal : BEGIN
            SQL:='UPDATE JOURNAL SET J_TOTALDEBIT=J_TOTALDEBIT+:DD, J_TOTALCREDIT=J_TOTALCREDIT+:CD, ' ;
            Q.SQL.Add(SQL) ;
            SQL:='J_DATEDERNMVT=:DATEDM, J_DEBITDERNMVT=:DEBITDM, J_CREDITDERNMVT=:CREDITDM, J_NUMDERNMVT=:NUMPDM, ' ;
            Q.SQL.Add(SQL) ;
            SQL:='J_TOTDEBE=J_TOTDEBE+:DE, J_TOTCREE=J_TOTCREE+:CE, J_TOTDEBS=J_TOTDEBS+:DS, J_TOTCRES=J_TOTCRES+:CS  WHERE J_JOURNAL=:CPTE ' ;
            Q.SQL.Add(SQL) ;
            END ;
   End ;
ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Result:=Q ;
END ;


{=============================================================================}
procedure AttribParamsCpt ( QX : TQuery ; Var LeCptLu : TCptLu ; fb : TFichierBase ) ;
BEGIN
QX.ParamByName('CPTE').AsString:=LeCptLu.Cpt ;
QX.ParamByName('DD').AsFloat:=LeCptLu.TotDN+LeCptLu.TotDN1 ;
QX.ParamByName('CD').AsFloat:=LeCptLu.TotCN+LeCptLu.TotCN1 ;
QX.ParamByName('DE').AsFloat:=LeCptLu.TotDN ; QX.ParamByName('CE').AsFloat:=LeCptLu.TotCN ;
QX.ParamByName('DS').AsFloat:=LeCptLu.TotDN1 ; QX.ParamByName('CS').AsFloat:=LeCptLu.TotCN1 ;
QX.ParamByName('DATEDM').AsDateTime:=LeCptLu.LastDate ;
QX.ParamByName('DEBITDM').AsFloat:=LeCptLu.DebitMvt ;
QX.ParamByName('CREDITDM').AsFloat:=LeCptLu.CreditMvt ;
QX.ParamByName('NUMPDM').AsInteger:=LeCptLu.LastNum ;
If Fb<>fbJal Then
   BEGIN
   QX.ParamByName('DA').AsFloat:=LeCptLu.TotDAno ;
   QX.ParamByName('CA').AsFloat:=LeCptLu.TotCAno ;
   QX.ParamByName('NUMLDM').AsInteger:=LeCptLu.LastNumL ;
   END ;
 if fb=fbAxe1 Then QX.ParamByName('AXE').AsString:=LeCptLu.Axe ;
END ;

{=============================================================================}
procedure MajTotCompteImport(fb : TFichierBase ; L : HTStringList);
Var QCpt : TQuery ;
    LeCptLuP : TCptLuP ;
    LeCptLu : TCptLu ;
    i : Integer ;
begin
QCpt:=FabricReqCpt(Fb) ;
For i:=0 To L.Count-1 Do
  BEGIN
  If MoveCur(FALSE) Then ;
  LeCptLuP:=TCptLuP(L.Objects[i]) ;
  ClassToRec(LeCptLu,LeCptLuP) ;
  If LeCptLu.OkMajSolde Then
    BEGIN
    AttribParamsCpt(QCpt,LeCptLu,fb) ;
    QCpt.ExecSQL ; QCpt.Close ;
    END ;
  END ;
QCpt.Free ;
end ;

{=============================================================================}
Procedure MajSoldeCompteImport(Var InfoImp : TInfoImport) ;
BEGIN
If (Not RecupSISCO) And (Not RecupSERVANT) And (Not RecupSISCOExt)Then
  BEGIN
  InitMove(1000,'') ;
  BeginTrans ;
  If (InfoImp.LGenLu<>NIL) And (InfoImp.LGenLu.Count>0) Then MajTotCompteImport(fbGene,InfoImp.LGenLu) ;
  If (InfoImp.LAuxLu<>NIL) And (InfoImp.LAuxLu.Count>0) Then MajTotCompteImport(fbAux,InfoImp.LAuxLu) ;
  If (InfoImp.LAnaLu<>NIL) And (InfoImp.LAnaLu.Count>0) Then MajTotCompteImport(fbAxe1,InfoImp.LAnaLu) ;
  If (InfoImp.LJalLu<>NIL) And (InfoImp.LJalLu.Count>0) Then MajTotCompteImport(fbJal,InfoImp.LJalLu) ;
  CommitTrans ;
  FiniMove ;
  END ;
END ;

Procedure DecodeLeRIB ( Var Etab,Guichet,Numero,Cle,Dom : String ; RIB : String ) ;
BEGIN
Etab:=Copy(RIB,1,5) ; Guichet:=Copy(RIB,7,5) ; Numero:=Copy(RIB,13,11) ;
Cle:=Copy(RIB,25,2) ; Dom:=Copy(RIB,28,24)   ;
END ;


Procedure TraiteRibImport(CptAux,Rib,Ville,Pays,RibPrincipal : String ; What : Integer = 0 ) ;
Var Etab,Guichet,Numero,Cle,Dom,Principal,St : String ;
    LastNum : Integer ;
    QRib : TQuery ;
    MajAFaire{,ForcePrincipal} : Boolean ;
    RibTrouve : Boolean ;
BEGIN
If Trim(Rib)='' Then Exit ; If Trim(CptAux)='' Then Exit ;
DecodeLeRIB(Etab,Guichet,Numero,Cle,Dom,RIB) ; RibTrouve:=FALSE ;
Etab:=Trim(Etab) ; Guichet:=Trim(Guichet) ; Numero:=Trim(Numero) ; Cle:=Trim(Cle) ;
Dom:=Trim(Dom) ;
If (Etab='') and (Guichet='') and (Numero='') and (Cle='') and (Dom='') Then Exit ;
St:='Select * from RIB Where R_AUXILIAIRE="'+CptAux+'"' ;
QRib:=OpenSQL(St,FALSE) ;
MajAFaire:=TRUE ; LastNum:=0 ; Principal:='X' ;
If Not QRib.EOF Then
   BEGIN
   While Not QRib.Eof Do
     BEGIN
//     RibTrouve:=FALSE ;
     If (QRib.Findfield('R_DOMICILIATION').ASString=Dom) And
        (QRib.Findfield('R_ETABBQ').ASString=Etab) And
        (QRib.Findfield('R_GUICHET').ASString=Guichet) And
        (QRib.Findfield('R_NUMEROCOMPTE').ASString=Numero) And
        (QRib.Findfield('R_CLERIB').ASString=Cle) Then BEGIN MajAFaire:=FALSE ; RibTrouve:=TRUE ; END ;
     If QRib.Findfield('R_PRINCIPAL').ASString='X' Then Principal:='-' ;
     If QRib.Findfield('R_NUMERORIB').ASInteger>LastNum Then LastNum:=QRib.Findfield('R_NUMERORIB').ASInteger ;
     If RibPrincipal='X' Then
       BEGIN
       If RibTrouve Then
         BEGIN
         If QRib.Findfield('R_PRINCIPAL').ASString='-' Then
           BEGIN
           QRib.Edit ;
           QRib.Findfield('R_PRINCIPAL').ASString:='X' ;
           QRib.Post ;
           END ;
         END Else
         BEGIN
         If QRib.Findfield('R_PRINCIPAL').ASString='X' Then
           BEGIN
           QRib.Edit ;
           QRib.Findfield('R_PRINCIPAL').ASString:='-' ;
           QRib.Post ;
           END ;
         END ;
       END ;
     QRib.Next ;
     END ;
   END ;
If MajAFaire Then
   BEGIN
   QRib.insert ;
   InitNew(QRib) ;
   If RibPrincipal='X' Then Principal:='X' ;
   QRib.Findfield('R_AUXILIAIRE').ASString:=CptAux ;
   QRib.Findfield('R_NUMERORIB').ASInteger:=LastNum+1 ;
   QRib.Findfield('R_DOMICILIATION').ASString:=Dom ;
   QRib.Findfield('R_ETABBQ').ASString:=Etab ;
   QRib.Findfield('R_GUICHET').ASString:=Guichet ;
   QRib.Findfield('R_NUMEROCOMPTE').ASString:=Numero ;
   QRib.Findfield('R_CLERIB').ASString:=Cle ;
   QRib.Findfield('R_PRINCIPAL').ASString:=Principal ;
   QRib.Findfield('R_VILLE').ASString:=Ville ;
   QRib.Findfield('R_DEVISE').ASString:=V_PGI.DevisePivot ;
   QRib.Findfield('R_PAYS').ASString:=Pays ;
   If What<>0 Then QRib.Findfield('R_SOCIETE').ASString:=IntToStr(What) ; // Pour recup SISCO S3
   QRib.Post ;
   END ;
Ferme(QRib) ;
END ;

Procedure TraiteContactImport(CptAux : String ; Var Ct : TContact) ;
Var Principal,St : String ;
    LastNum : Integer ;
    QCt : TQuery ;
    MajAFaire,InsertAFAire : Boolean ;
BEGIN
If Trim(CptAux)='' Then Exit ; If (Ct.Nom='') Then Exit ;
St:='Select * from CONTACT Where C_AUXILIAIRE="'+CptAux+'"' ;
QCt:=OpenSQL(St,FALSE) ;
MajAFaire:=FALSE ; InsertAFaire:=TRUE ; LastNum:=0 ; Principal:='X' ;
If Not QCt.EOF Then
   BEGIN
   While Not QCt.Eof Do
     BEGIN
     If (QCt.Findfield('C_NOM').ASString=Ct.Nom) Then BEGIN InsertAFaire:=FALSE ; MajAFaire:=TRUE ; Break ; END ;
     If QCt.Findfield('C_PRINCIPAL').ASString='X' Then Principal:='-' ;
     If QCt.Findfield('C_NUMEROCONTACT').ASInteger>LastNum Then LastNum:=Qct.Findfield('C_NUMEROCONTACT').ASInteger ;
     QCt.Next ;
     END ;
   END ;
If InsertAFaire Then
   BEGIN
   QCt.insert ;
   InitNew(QCt) ;
   QCt.Findfield('C_TYPECONTACT').ASString:='T' ;
   QCt.Findfield('C_AUXILIAIRE').ASString:=CptAux ;
   QCt.Findfield('C_NUMEROCONTACT').ASInteger:=LastNum+1 ;
   QCt.Findfield('C_NOM').ASString:=Ct.Nom ;
   QCt.Findfield('C_SERVICE').ASString:=Ct.Service ;
   QCt.Findfield('C_FONCTION').ASString:=Ct.Fonction ;
   QCt.Findfield('C_TELEPHONE').ASString:=Ct.Tel ;
   QCt.Findfield('C_FAX').ASString:=Ct.Fax ;
   QCt.Findfield('C_PRINCIPAL').ASString:=Principal ;
   QCt.Findfield('C_TELEX').ASString:=Ct.Telex ;
   QCt.Findfield('C_RVA').ASString:=Ct.RVA ;
   QCt.Findfield('C_CIVILITE').ASString:=Ct.Civilite ;
   QCt.Post ;
   END Else
If MajAFaire Then
   BEGIN
   QCt.Edit ;
   If Ct.Service<>'' Then QCt.Findfield('C_SERVICE').ASString:=Ct.Service ;
   If Ct.Fonction<>'' Then QCt.Findfield('C_FONCTION').ASString:=Ct.Fonction ;
   If Ct.Tel<>'' Then QCt.Findfield('C_TELEPHONE').ASString:=Ct.Tel ;
   If Ct.Fax<>'' Then QCt.Findfield('C_FAX').ASString:=Ct.Fax ;
   If Ct.Telex<>'' Then QCt.Findfield('C_TELEX').ASString:=Ct.Telex ;
   If Ct.RVA<>'' Then QCt.Findfield('C_RVA').ASString:=Ct.RVA ;
   If Ct.Civilite<>'' Then QCt.Findfield('C_CIVILITE').ASString:=Ct.Civilite ;
   QCt.Post ;
   END ;
Ferme(QCt) ;
END ;


Function InitSQLCpt(What : Integer) : TQuery ;
Var Q : TQuery ;
    Table,Champ,StWhereSupp : String ;
BEGIN
StWhereSupp:='' ;
Case What Of
  0 : BEGIN
      Table:='GENERAUX' ; Champ:='G_GENERAL' ;
      END ;
  1 : BEGIN
      Table:='TIERS' ; Champ:='T_AUXILIAIRE' ;
      END ;
  2 : BEGIN
      Table:='SECTION' ; Champ:='S_SECTION' ; StWhereSupp:='AND S_AXE=:AXE ' ;
      END ;
  3 : BEGIN
      Table:='JOURNAL' ; Champ:='J_JOURNAL' ;
      END ;
  10 : BEGIN
       Table:='NATCPTE' ; Champ:='NT_TYPECPTE' ; StWhereSupp:='AND NT_NATURE=:NAT ' ;
       END ;
  11 : BEGIN
      Table:='SSSTRUCR' ; Champ:='PS_CODE' ; StWhereSupp:='AND PS_AXE=:AXE AND PS_SOUSSECTION=:SS' ;
      END ;
  12 : BEGIN
      Table:='SOUCHE' ; Champ:='SH_SOUCHE' ; StWhereSupp:='AND SH_TYPE="CPT"' ;
      END ;
  END ;
Q:=PrepareSQL('SELECT * FROM '+Table+' WHERE '+Champ+'=:CPT '+StWhereSupp,FALSE) ;
Result:=Q ;
END ;

Function RecupCptGen(St : String ; Var Gen : TCptImport ; Var InfoImp : TInfoImport) : Boolean ;
Var NatFour : String ;
BEGIN
Fillchar(Gen,SizeOf(Gen),#0) ;
With Gen Do
  BEGIN
  Cpt:=Trim(Copy(St,7,17))        ; Libelle:=Trim(Copy(St,24,35)) ;
  Nature:=Trim(Copy(St,59,3))     ; Lettrable:=Copy(St,62,1) ;
  Pointable:=Copy(St,63,1)        ; Vent1:=Copy(St,64,1) ;
  Vent2:=Copy(St,65,1)            ; Vent3:=Copy(St,66,1) ;
  Vent4:=Copy(St,67,1)            ; Vent5:=Copy(St,68,1) ;
  T0:=Trim(Copy(St,69,3))         ; T1:=Trim(Copy(St,72,3)) ;
  T2:=Trim(Copy(St,75,3))         ; T3:=Trim(Copy(St,78,3)) ;
  T4:=Trim(Copy(St,81,3))         ; T5:=Trim(Copy(St,84,3)) ;
  T6:=Trim(Copy(St,87,3))         ; T7:=Trim(Copy(St,90,3)) ;
  T8:=Trim(Copy(St,93,3))         ; T9:=Trim(Copy(St,96,3)) ;
  Abrege:=Trim(Copy(St,99,17))    ; Sens:=Trim(Copy(St,116,3)) ;
  If (Sens='') Or ((Sens<>'M') And (Sens<>'D') And (Sens<>'C')) Then Sens:='M' ;
  If InfoImp.SC.Majuscule Then Cpt:=AnsiUpperCase(Cpt) ;
  END ;
NatFour:=NatureCptImport(Gen.Cpt,InfoImp.SC.FourchetteImport) ;
If (NatFour<>'---') Then Gen.Nature:=NatFour ;
//If Gen.Nature='' Then Gen.Nature:='DIV' ;
Result:=(Gen.Cpt<>'') And (Gen.Nature<>'') ;
END ;

Function RecupCptAux(St : String ; Var Aux : TCptImport ; Var Ct : TContact ;
                     What : Integer ; Var InfoImp : TInfoImport ; QFiche : TQFiche) : Boolean ;
Var Etab,Guichet,Numero,Cle,Dom : String ;
    CptLu : TCptLu ;
    Decal : Integer ;
    St1,St2 : String ;
BEGIN
Fillchar(Aux,SizeOf(Aux),#0) ; Fillchar(Ct,SizeOf(Ct),#0) ; Decal:=0 ;
With Aux Do
  BEGIN
  Cpt:=Trim(Copy(St,7,17))        ; Libelle:=Trim(Copy(St,24,35)) ;
  If GetRecupSISCOPGI And (Pos('µ',Libelle)>0) Then
    BEGIN
    St1:=Libelle+'µ' ;
    St2:=ReadTokenPipe(St1,'µ') ; Libelle:=St2 ;
    St2:=ReadTokenPipe(St1,'µ') ; CptOrigine:=St2 ;
    END ;
  Nature:=Trim(Copy(St,59,3))     ; Lettrable:=Copy(St,62,1) ;
  Collectif:=Trim(Copy(St,63,17)) ; EAN:=Trim(Copy(St,80,17)) ;
  If InfoImp.SC.Majuscule Then
    BEGIN
    Cpt:=AnsiUpperCase(Cpt)         ; Collectif:=AnsiUpperCase(Collectif) ;
    END ;
  Case What Of
    0 : BEGIN
        T0:=Trim(Copy(St,97,3))         ; T1:=Trim(Copy(St,100,3)) ;
        T2:=Trim(Copy(St,103,3))        ; T3:=Trim(Copy(St,106,3)) ;
        T4:=Trim(Copy(St,109,3))        ; T5:=Trim(Copy(St,112,3)) ;
        T6:=Trim(Copy(St,115,3))        ; T7:=Trim(Copy(St,118,3)) ;
        T8:=Trim(Copy(St,121,3))        ; T9:=Trim(Copy(St,124,3)) ;
        END ;
    1 : BEGIN
        T0:=Trim(Copy(St,97,4))         ; T1:=Trim(Copy(St,101,4)) ;
        T2:=Trim(Copy(St,105,4))        ; T3:=Trim(Copy(St,109,4)) ;
        END ;
    2 : BEGIN
        T0:=Trim(Copy(St,97,17))         ; T1:=Trim(Copy(St,114,17)) ;
        T2:=Trim(Copy(St,131,17))        ; T3:=Trim(Copy(St,148,17)) ;
        T4:=Trim(Copy(St,165,17))        ; T5:=Trim(Copy(St,182,17)) ;
        T6:=Trim(Copy(St,199,17))        ; T7:=Trim(Copy(St,216,17)) ;
        T8:=Trim(Copy(St,233,17))        ; T9:=Trim(Copy(St,250,17)) ;
        Decal:=140 ;
        END ;
    END ;
  Adr1:=Trim(Copy(St,127+Decal,35))     ; Adr2:=Trim(Copy(St,162+Decal,35)) ;
  Adr3:=Trim(Copy(St,197+Decal,35))     ; CodePost:=Trim(Copy(St,232+Decal,9)) ;
  Ville:=Trim(Copy(St,241+Decal,35))    ; Etab:=Trim(Copy(St,300+Decal,5)) ;
  Guichet:=Trim(Copy(St,305+Decal,5))   ; Numero:=Trim(Copy(St,310+Decal,11)) ;
  Cle:=Trim(Copy(St,321+Decal,2))       ; Dom:=Trim(Copy(St,276+Decal,24)) ;
  If (Etab<>'') And (Guichet<>'') And (Numero<>'') And (Cle<>'') And (Dom<>'') Then
    BEGIN
    RIB:=Format_String(Etab,5)+'/'+Format_String(Guichet,5)+'/'+Format_String(Numero,11)+'/'+
         Format_String(Cle,2)+'/'+Dom ;
    End ;
  Pays:=Trim(Copy(St,323+Decal,3))      ; Abrege:=Trim(Copy(St,326+Decal,17)) ;
  Langue:=Trim(Copy(St,343+Decal,3))    ; MultiDevise:=Trim(Copy(St,346+Decal,1)) ;
  Devise:=Trim(Copy(St,347+Decal,3))    ; Tel:=Trim(Copy(St,350+Decal,25)) ;
  Fax:=Trim(Copy(St,375+Decal,25))      ; RegimeTVA:=Trim(Copy(St,400+Decal,3)) ;
  ModeRegl:=Trim(Copy(St,403+Decal,3))  ; Commentaire:=Trim(Copy(St,406+Decal,35)) ;
  NIF:=Trim(Copy(St,441+Decal,17))      ; Siret:=Trim(Copy(St,458+Decal,17)) ;
  APE:=Trim(Copy(St,475+Decal,5))       ; FormeJur:=Trim(Copy(St,714+Decal,3)) ;
  RibP:=Trim(Copy(St,717+Decal,1))      ; TiersTP:=Trim(Copy(St,718+Decal,17)) ;
  IsTP:=Trim(Copy(St,735+Decal,1))      ; AvoirRBT:=Trim(Copy(St,736+Decal,1)) ;
  TiersMP:=Trim(Copy(St,737+Decal,1))   ; TiersEche:=Trim(Copy(St,754+Decal,1)) ;
//  Rep:=Trim(Copy(St,441,17))      ;
  END ;
With Ct Do
  BEGIN
  Nom:=Trim(Copy(St,480+Decal,35))      ; Service:=Trim(Copy(St,515+Decal,35)) ;
  Fonction:=Trim(Copy(St,550+Decal,35)) ; Tel:=Trim(Copy(St,585+Decal,25)) ;
  Fax:=Trim(Copy(St,610+Decal,25))      ; Telex:=Trim(Copy(St,635+Decal,25)) ;
  RVA:=Trim(Copy(St,660+Decal,50))      ; Civilite:=Trim(Copy(St,710+Decal,3)) ;
  Principal:=Trim(Copy(St,713+Decal,1))
  END ;
If (Aux.Nature='') And (Aux.Collectif<>'') Then
  BEGIN
  Fillchar(CptLu,SizeOf(CptLu),#0) ;
  CptLu.Cpt:=BourreOuTronque(Aux.Collectif,fbGene) ;
  If AlimLTabCptLu(0,QFiche[0],InfoImp.LGenLu,Nil,CptLu) Then
    BEGIN
    If CptLu.Nature='COC' Then Aux.Nature:='CLI' Else
     If CptLu.Nature='COF' Then Aux.Nature:='FOU' Else
      If CptLu.Nature='COS' Then Aux.Nature:='SAL' Else
       If CptLu.Nature='COD' Then Aux.Nature:='DIV' ;
    END ;
  END ;
Result:=(Aux.Cpt<>'') And (Aux.Collectif<>'') And (Aux.Nature<>'') ;
If (Aux.Nature<>'CLI') And (Aux.Nature<>'FOU') And (Aux.Nature<>'SAL') And
   (Aux.Nature<>'AUD') And (Aux.Nature<>'AUC') And (Aux.Nature<>'DIV') Then Result:=FALSE ;
END ;


Function RecupSousSection(St : String ; Var SousSection : TCptImport ; Var InfoImp : TInfoImport) : Boolean ;
BEGIN
Fillchar(SousSection,SizeOf(SousSection),#0) ;
With SousSection Do
  BEGIN
  Cpt:=Trim(Copy(St,7,17)) ; Libelle:=Trim(Copy(St,24,35)) ;
  Axe:=Trim(Copy(St,59,3)) ; If Axe='' Then Axe:='A1' ;
  T0:=Trim(Copy(St,62,3))  ; T1:=Trim(Copy(St,69,35)) ;
  Deb:=StrToInt(Copy(St,65,2)) ; Lg:=StrToInt(Copy(St,67,2)) ;
  If InfoImp.SC.Majuscule Then Cpt:=AnsiUpperCase(Cpt) ;
  END ;
Result:=(SousSection.Cpt<>'') And (SousSection.Axe<>'') And ((SousSection.T0<>'') Or (SousSection.Deb>0)) ;
END ;

Function RecupJal(St : String ; Var Jal : TCptImport ; Var InfoImp : TInfoImport) : Boolean ;
Var OkOk : Boolean ;
BEGIN
Fillchar(Jal,SizeOf(Jal),#0) ; OkOk:=TRUE ;
With Jal Do
  BEGIN
  Cpt:=Trim(Copy(St,7,3)) ; Libelle:=Trim(Copy(St,10,35)) ;
  Nature:=Trim(Copy(St,45,3)) ;
  SoucheN:=Trim(Copy(St,48,3)) ; SoucheS:=Trim(Copy(St,51,3)) ;
  Collectif:=Trim(Copy(St,54,17)) ; Axe:=Trim(Copy(St,71,3)) ;
  ModeSaisie:=Trim(Copy(St,74,3)) ;
  CptAuto:=Trim(Copy(St,77,200)) ;
  CptInt:=Trim(Copy(St,277,200)) ;
  Abrege:=Trim(Copy(Cpt,1,17)) ;
  If InfoImp.SC.Majuscule Then Collectif:=AnsiUpperCase(Collectif) ;
  If (Axe='') And ((Nature='ODA') Or (Nature='ANA')) Then Axe:='A1' ;
  If Nature='' Then Nature:='OD' ;
  If SoucheN='' Then SoucheN:=Cpt ; If SoucheS='' Then SoucheS:='SIM' ;
  // Pour un journal à-nouveau pas de souche simu.
  if GetRecupPCL and (Nature = 'ANO') then SoucheS := '';
  If ModeSaisie='' Then ModeSaisie:='-' ;
  If (Nature='BQE') Or (Nature='CAI') Then OkOk:=(Collectif<>'') ;
  OkOk:=OkOk And
       ((Nature='ACH') Or (Nature='ANA') Or (Nature='ANO') Or (Nature='BQE') Or
        (Nature='CAI') Or (Nature='CLO') Or (Nature='ECC') Or (Nature='EXT') Or
        (Nature='OD') Or (Nature='ODA') Or (Nature='REG') Or (Nature='VTE')) ;
  If Collectif<>'' Then Collectif:=BourreOuTronque(Collectif,fbGene) ;
  END ;
Result:=OkOk And (Jal.Cpt<>'') ;
END ;



Procedure CorrigeGen(Var Gen : TCptImport) ;
BEGIN
If (Gen.Nature='COC') Or (Gen.Nature='COF') Or (Gen.Nature='COS') Or (Gen.Nature='COD') Or
   (Gen.Nature='CHA') Or (Gen.Nature='PRO') Then
  BEGIN
  Gen.Pointable:='-' ; Gen.Lettrable:='-' ;
  END ;
If (Gen.Pointable='X') And ((Gen.Nature='BQE') Or (Gen.Nature='CAI')) Then Gen.Lettrable:='-' ;
END ;

Function ImportGene(St : String ; Var InfoImp : TInfoImport ; QFiche : TQFiche) : tResultImportCpte ;
Var Q : TQuery ;
    CptLu : TCptLu ;
    Gen : TCptImport ;
    OkCreat,Existe : Boolean ;
    GenOrigine : String ;
    NatureGene : string;
BEGIN
Result:=ResRien ;
OkCreat:=RecupCptGen(St,Gen,InfoImp) ; GenOrigine:=Gen.Cpt ; TraiteCorrespCpt(0,QFiche,Gen,InfoImp) ;
Gen.Cpt:=BourreOuTronque(Gen.Cpt,fbGene) ;
CptLu.Cpt:=Gen.Cpt ;
Existe:=AlimLTabCptLu(0,QFiche[0],InfoImp.LGenLu,Nil,CptLu) ;
If Existe Or ((Not Existe) And OkCreat) Then
  BEGIN
  (*
  Q:=InitSQLCpt(0) ;
  If Existe Then Q.Params[0].AsString:=Gen.Cpt Else Q.Params[0].AsString:=w_w ;
  Q.Open ;
  *)
  Q:=OpenSQL(InitSQLCpt1(0,Gen.Cpt,'','',Existe),FALSE) ;
  Result:=ResModifier ;
  If Existe Then Q.Edit Else BEGIN Q.Insert ; InitNew(Q) ; Result:=ResCreer ; END ;
  If Not Existe Then
    BEGIN
    Q.FindField('G_GENERAL').AsString:=Gen.Cpt ;
    Q.FindField('G_NATUREGENE').AsString:=Gen.Nature ;
    if GetRecupPCL Or GetRecupSISCOPGI then Q.FindField('G_CONFIDENTIEL').AsString:='0';
    CorrigeGen(Gen) ;
    If (Gen.Nature='TID') Or (Gen.Nature='TIC') Then
      BEGIN
      If InfoImp.SC.RGT<>'' Then Q.FindField('G_REGIMETVA').AsString:=InfoImp.SC.RGT ;
      If InfoImp.SC.RMR<>'' Then Q.FindField('G_MODEREGLE').AsString:=InfoImp.SC.RMR ;
      END ;
    If (Gen.Nature='COC') Or (Gen.Nature='COF') Or (Gen.Nature='COS') Or (Gen.Nature='COD') Then Q.FindField('G_COLLECTIF').AsString:='X' ;
    If Gen.Nature='COC' Then Q.FindField('G_SUIVITRESO').AsString:='ENC' Else
     If Gen.Nature='COF' Then Q.FindField('G_SUIVITRESO').AsString:='DEC' ;
    If Gen.Vent1='X' Then Q.FindField('G_VENTILABLE1').AsString:=Gen.Vent1 ;
{$IFNDEF CCS3}
    If Gen.Vent2='X' Then Q.FindField('G_VENTILABLE2').AsString:=Gen.Vent2 ;
    If Gen.Vent3='X' Then Q.FindField('G_VENTILABLE3').AsString:=Gen.Vent3 ;
    If Gen.Vent4='X' Then Q.FindField('G_VENTILABLE4').AsString:=Gen.Vent4 ;
    If Gen.Vent5='X' Then Q.FindField('G_VENTILABLE5').AsString:=Gen.Vent5 ;
{$ENDIF}
    If Gen.Pointable='X' Then Q.FindField('G_POINTABLE').AsString:=Gen.Pointable ;
    If Gen.Lettrable='X' Then Q.FindField('G_LETTRABLE').AsString:=Gen.Lettrable ;
    If (Gen.Vent1='X') Or( Gen.Vent2='X') Or (Gen.Vent3='X') Or
       (Gen.Vent4='X') Or (Gen.Vent5='X') Then
      Q.FindField('G_VENTILABLE').AsString:='X' ;
    If Gen.Sens<>'' Then Q.FindField('G_SENS').AsString:=Gen.Sens ;
    END
    ELSE
    BEGIN  // fiche 10351 ajout me 22-03-2006
      NatureGene := Q.FindField('G_NATUREGENE').AsString;
      If GetRecupPcl and (not GetRecupComSx) and (Gen.Nature <> NatureGene) Then
      begin
           if (Gen.Nature = 'COC') or (Gen.Nature = 'COF') or
              (NatureGene = 'COC') or (NatureGene = 'COF') then
              begin
                   Q.FindField('G_NATUREGENE').AsString:=Gen.Nature ;
                   Q.FindField('G_LETTRABLE').AsString := '-';    // Fiche 10449
                   Q.FindField('G_COLLECTIF').AsString := 'X';
              end;
      end;
    END;

  Q.FindField('G_LIBELLE').AsString:=Gen.Libelle ;
  If Gen.Abrege<>'' Then Q.FindField('G_ABREGE').AsString:=Gen.Abrege
                    Else Q.FindField('G_ABREGE').AsString:=Copy(Gen.Libelle,1,17) ;
  If Gen.T0<>'' Then Q.FindField('G_TABLE0').AsString:=Gen.T0 ;
  If Gen.T1<>'' Then Q.FindField('G_TABLE1').AsString:=Gen.T1 ;
  If Gen.T2<>'' Then Q.FindField('G_TABLE2').AsString:=Gen.T2 ;
  If Gen.T3<>'' Then Q.FindField('G_TABLE3').AsString:=Gen.T3 ;
  If Gen.T4<>'' Then Q.FindField('G_TABLE4').AsString:=Gen.T4 ;
  If Gen.T5<>'' Then Q.FindField('G_TABLE5').AsString:=Gen.T5 ;
  If Gen.T6<>'' Then Q.FindField('G_TABLE6').AsString:=Gen.T6 ;
  If Gen.T7<>'' Then Q.FindField('G_TABLE7').AsString:=Gen.T7 ;
  If Gen.T8<>'' Then Q.FindField('G_TABLE8').AsString:=Gen.T8 ;
  If Gen.T9<>'' Then Q.FindField('G_TABLE9').AsString:=Gen.T9 ;
  If Q.FindField('G_CODEIMPORT')<>NIL Then Q.FindField('G_CODEIMPORT').AsString:=GenOrigine ;
//  If ((Not Existe) And OkCreat) Then Q.FindField('G_CREERPAR').AsString:='IMP' ;
  If ((Not Existe) And OkCreat) Then SetCreerPar(Q,'G_CREERPAR') ;
  Q.Post ;
  Q.Close ; Ferme(Q) ;
  END ;
END ;

Procedure CorrigeAux(Var Aux : TCptImport) ;
BEGIN
If Aux.TiersTP<>'' Then Aux.IsTP:='-' ;
END ;

Function ImportTiers(St : String ; Var InfoImp : TInfoImport ; QFiche : TQFiche ; What : Integer) : tResultImportCpte ;
// What = 1 : SCOT
//      = 2 : TL à17
Var Q : TQuery ;
    CptLu : TCptLu ;
    Aux : TCptImport ;
{$IFDEF SYSTEMU}
    MRTiers : TCptImport ;
{$ENDIF}
    OkCreat,Existe : Boolean ;
    Ct : TContact ;
    TiersOrigine : String ;
//    LeRegime : String ;
    MRLu : tCptLu ;
BEGIN
Result:=ResRien ;
OkCreat:=RecupCptAux(St,Aux,Ct,What,InfoImp,QFiche) ; TiersOrigine:=Aux.Cpt ; TraiteCorrespCpt(1,QFiche,Aux,InfoImp) ;
Aux.Cpt:=BourreOuTronque(Aux.Cpt,fbAux) ;
If PasDeBlanc And (Aux.Cpt<>'') Then Aux.Cpt:=FindEtReplace(Aux.Cpt,' ','.',TRUE) ;
CptLu.Cpt:=Aux.Cpt ;
Existe:=AlimLTabCptLu(1,QFiche[1],InfoImp.LAuxLu,Nil,CptLu) ;
If Existe Or ((Not Existe) And OkCreat) Then
  BEGIN
  (*
  Q:=InitSQLCpt(1) ;
  If Existe Then Q.Params[0].AsString:=Aux.Cpt Else Q.Params[0].AsString:=w_w ;
  Q.Open ;
  *)
  Q:=OpenSQL(InitSQLCpt1(1,Aux.Cpt,'','',Existe),FALSE) ;
  Result:=ResModifier ;
  If Existe Then Q.Edit Else BEGIN Q.Insert ; InitNew(Q) ; Result:=ResCreer ; END ;
  If Not existe then
    BEGIN
    Q.FindField('T_AUXILIAIRE').AsString:=Aux.Cpt ;
    If GetRecupSISCOPGI And (Aux.CptOrigine<>'') Then Q.FindField('T_TIERS').AsString:=Aux.CptOrigine
                                                  Else Q.FindField('T_TIERS').AsString:=Aux.Cpt ;
    if GetRecupPCL Or GetRecupSISCOPGI then Q.FindField('T_CONFIDENTIEL').AsString := '0';
    Q.FindField('T_LETTRABLE').AsString:=Aux.Lettrable ;
    Q.FindField('T_NATUREAUXI').AsString:=Aux.Nature ;
    Q.FindField('T_COLLECTIF').AsString:=BourreOuTronque(Aux.Collectif,fbGene) ;
    CorrigeAux(Aux) ;
    If InfoImp.SC.RGT<>'' Then
      BEGIN
      MRLu.Cpt:=Aux.RegimeTva ;
      If InfoImp.LRGT.IndexOf(Aux.ModeRegl)<=-1 Then Q.FindField('T_REGIMETVA').AsString:=InfoImp.SC.RGT
                                                Else Q.FindField('T_REGIMETVA').AsString:=Aux.RegimeTVA ;
      END Else
      BEGIN
      If Aux.RegimeTVA<>'' Then Q.FindField('T_REGIMETVA').AsString:=Aux.RegimeTVA Else Q.FindField('T_REGIMETVA').AsString:='FRA' ;
      END ;
    (*
    If Aux.RegimeTVA<>'' Then Q.FindField('T_REGIMETVA').AsString:=Aux.RegimeTVA Else
      If InfoImp.SC.RGT<>'' Then Q.FindField('T_REGIMETVA').AsString:=InfoImp.SC.RGT
        Else Q.FindField('T_REGIMETVA').AsString:='FRA' ;
    *)
    Q.FindField('T_TVAENCAISSEMENT').AsString:='TD' ;
    If Aux.MultiDevise='X' Then Q.FindField('T_MULTIDEVISE').AsString:=Aux.MultiDevise Else
      If Aux.Devise<>'' Then Q.FindField('T_DEVISE').AsString:=Aux.Devise ;
    END ;
  Q.FindField('T_LIBELLE').AsString:=Aux.Libelle ;
  If Aux.Abrege<>'' Then Q.FindField('T_ABREGE').AsString:=Aux.Abrege
                    Else Q.FindField('T_ABREGE').AsString:=Copy(Aux.Libelle,1,17) ;
  If Aux.EAN<>'' Then Q.FindField('T_EAN').AsString:=Aux.EAN ;
  If Aux.T0<>'' Then Q.FindField('T_TABLE0').AsString:=Aux.T0 ;
  If Aux.T1<>'' Then Q.FindField('T_TABLE1').AsString:=Aux.T1 ;
  If Aux.T2<>'' Then Q.FindField('T_TABLE2').AsString:=Aux.T2 ;
  If Aux.T3<>'' Then Q.FindField('T_TABLE3').AsString:=Aux.T3 ;
  If Aux.T4<>'' Then Q.FindField('T_TABLE4').AsString:=Aux.T4 ;
  If Aux.T5<>'' Then Q.FindField('T_TABLE5').AsString:=Aux.T5 ;
  If Aux.T6<>'' Then Q.FindField('T_TABLE6').AsString:=Aux.T6 ;
  If Aux.T7<>'' Then Q.FindField('T_TABLE7').AsString:=Aux.T7 ;
  If Aux.T8<>'' Then Q.FindField('T_TABLE8').AsString:=Aux.T8 ;
  If Aux.T9<>'' Then Q.FindField('T_TABLE9').AsString:=Aux.T9 ;
  If Aux.Adr1<>'' Then Q.FindField('T_ADRESSE1').AsString:=Aux.Adr1 ;
  If Aux.Adr2<>'' Then Q.FindField('T_ADRESSE2').AsString:=Aux.Adr2 ;
  If Aux.Adr3<>'' Then Q.FindField('T_ADRESSE3').AsString:=Aux.Adr3 ;
  If Aux.CodePost<>'' Then Q.FindField('T_CODEPOSTAL').AsString:=Aux.CodePost ;
  If Aux.Ville<>'' Then Q.FindField('T_VILLE').AsString:=Aux.Ville ;
  If Aux.Pays<>'' Then Q.FindField('T_PAYS').AsString:=Aux.Pays ;
  If Aux.Abrege<>'' Then Q.FindField('T_ABREGE').AsString:=Aux.Abrege ;
  If Aux.Langue<>'' Then Q.FindField('T_LANGUE').AsString:=Aux.Langue ;
  If Aux.Tel<>'' Then Q.FindField('T_TELEPHONE').AsString:=Aux.Tel ;
  If Aux.Fax<>'' Then Q.FindField('T_FAX').AsString:=Aux.Fax ;
  If InfoImp.SC.RMR<>'' Then
    BEGIN
    If Aux.ModeRegl<>'' Then
      BEGIN
//      MRLu.Cpt:=Aux.ModeRegl ;
//      If Not ChercheCptLu(InfoImp.LMR,MRLu) Then
      {$IFDEF SYSTEMU}
      If (Aux.TiersMP<>'') And (Aux.TiersEche<>'') Then
        BEGIN
        Fillchar(MRTiers,SizeOf(MRTiers),#0) ;
        MRTiers.Cpt:=Aux.TiersMP+';'+Aux.TiersEche+';' ;
        If TraiteCorrespCpt(9,QFiche,MrTiers,InfoImp) Then Aux.ModeRegl:=MRTiers.Cpt ;
        END ;
      {$ENDIF}
      If InfoImp.LMR.IndexOf(Aux.ModeRegl)<=-1 Then
        BEGIN
        If (Not Existe) And (InfoImp.SC.RMR<>'') Then Q.FindField('T_MODEREGLE').AsString:=InfoImp.SC.RMR ;
        END Else Q.FindField('T_MODEREGLE').AsString:=Aux.ModeRegl ;
      END Else If Not Existe Then If InfoImp.SC.RMR<>'' Then Q.FindField('T_MODEREGLE').AsString:=InfoImp.SC.RMR ;
    END else If Aux.ModeRegl<>'' Then Q.FindField('T_MODEREGLE').AsString:=Aux.ModeRegl ;
  If Aux.Commentaire<>'' Then Q.FindField('T_COMMENTAIRE').AsString:=Aux.Commentaire ;
  If Aux.NIF<>'' Then Q.FindField('T_NIF').AsString:=Aux.Nif ;
  If Aux.Siret<>'' Then Q.FindField('T_SIRET').AsString:=Aux.Siret ;
  If Aux.APE<>'' Then Q.FindField('T_APE').AsString:=Aux.APE ;
  If Aux.FormeJur<>'' Then Q.FindField('T_JURIDIQUE').AsString:=Aux.FormeJur ;
  If Aux.TiersTP<>'' Then Q.FindField('T_PAYEUR').AsString:=Aux.TiersTP ;
  If Aux.IsTP<>'' Then Q.FindField('T_ISPAYEUR').AsString:=Aux.IsTP ;
  If Aux.AvoirRBT<>'' Then Q.FindField('T_AVOIRRBT').AsString:=Aux.AvoirRBT ;
  If Q.FindField('T_CODEIMPORT')<>NIL Then Q.FindField('T_CODEIMPORT').AsString:=TiersOrigine ;
//  If ((Not Existe) And OkCreat) Then Q.FindField('T_CREERPAR').AsString:='IMP' ;
  If ((Not Existe) And OkCreat) Then SetCreerPar(Q,'T_CREERPAR') ;
  Q.Post ;
  If Aux.Rib<>'' Then TraiteRibImport(Aux.Cpt,Aux.Rib,Q.FindField('T_VILLE').AsString,Q.FindField('T_PAYS').AsString,Aux.RibP) ;
  If Ct.Nom<>'' Then TraiteContactImport(Aux.Cpt,Ct) ;
  Q.Close ; Ferme(Q) ;
  END ;
END ;


Function VerifSousPlan(Var SousSection : TCptImport ; Var IndSP : Integer) : boolean ;
Var fb : tFichierBase ;
    i{,j},Lg1 : Integer ;
    SousPlanOk,SousPlanExiste : Boolean ;
    Q : TQuery ;
BEGIN
Result:=FALSE ; SousPlanOk:=FALSE ; SousPlanExiste:=FALSE ; IndSP:=-1;
fb:=AxeTofb(SousSection.Axe) ;
If SousSection.T0<>'' Then
  BEGIN
  For i:=1 To MaxSousPlan Do
    BEGIN
    If VH^.SousPlanAxe[fb,i].Code=SousSection.T0 Then
      BEGIN
      SousPlanExiste:=TRUE ; IndSP:=i ;
      If SousSection.Deb>0 Then
        BEGIN
        SousPlanOk:=SousSection.Deb=VH^.SousPlanAxe[fb,i].Debut ;
        If SousPlanOk Then
          BEGIN
          If SousSection.Lg>0 Then SousPlanOk:=SousSection.Lg=VH^.SousPlanAxe[fb,i].Longueur
                              Else SousSection.Lg:=VH^.SousPlanAxe[fb,i].Longueur ;
          END ;
        END Else
        BEGIN
        SousSection.Deb:=VH^.SousPlanAxe[fb,i].Debut ;
        SousSection.Lg:=VH^.SousPlanAxe[fb,i].Longueur ;
        END ;
      Break ;
      END ;
    END ;
  END Else If SousSection.Deb>0 Then
  BEGIN
  For i:=1 To MaxSousPlan Do
    If VH^.SousPlanAxe[fb,i].Debut=SousSection.Deb Then
      BEGIN
      SousPlanExiste:=TRUE ; IndSP:=i ;
      SousSection.T0:=VH^.SousPlanAxe[fb,i].Code ;
      SousSection.T1:=VH^.SousPlanAxe[fb,i].Lib ;
      SousSection.Lg:=VH^.SousPlanAxe[fb,i].Longueur ;
      SousPlanOk:=TRUE ;
      END ;

  END Else Exit ;
If (Not SousPlanExiste) And (SousSection.T0<>'') And (SousSection.Deb>0) Then
  BEGIN
  If SousSection.Lg=0 Then SousSection.Lg:=Length(SousSection.Cpt) ;
  Q:=OpenSQL('SELECT * FROM STRUCRSE WHERE SS_AXE="'+SousSection.Axe+'" AND SS_SOUSSECTION="'+w_w+'" ',FALSE) ;
  Q.Insert ; InitNew(Q) ;
  Q.FindField('SS_AXE').AsString:=SousSection.Axe ;
  Q.FindField('SS_SOUSSECTION').AsString:=SousSection.T0 ;
  If SousSection.T1='' Then SousSection.T1:='CREE PAR IMPORTATION' ;
  Q.FindField('SS_LIBELLE').AsString:=SousSection.T1 ;
  Q.FindField('SS_CONTROLE').AsString:='-' ;
  Q.FindField('SS_DEBUT').AsInteger:=SousSection.Deb ;
  Q.FindField('SS_LONGUEUR').AsInteger:=SousSection.Lg ;
  Q.Post ;
  Ferme(Q) ;
  SousPlanOk:=TRUE ; IndSP:=0 ;
  CHARGESOUSPLANAXE ;
  For i:=1 To MaxSousPlan Do If VH^.SousPlanAxe[fb,i].Code=SousSection.T0 Then BEGIN IndSP:=i ; Break ;END ;
  END ;
If SousPlanOk Then
  BEGIN
  Lg1:=Length(SousSection.Cpt) ;
  If Lg1<SousSection.Lg Then
    BEGIN
    for i:=Lg1+1 to SousSection.Lg do SousSection.Cpt:=SousSection.Cpt+VH^.Cpta[fb].Cb ;
    END Else If Lg1>SousSection.Lg Then SousSection.Cpt:=Copy(SousSection.Cpt,1,SousSection.Lg) ;
  END ;
Result:=SousPlanOk ;
END ;


Function PlanSousSectionVide(Var OkCreat : Boolean ; Var SousSection : TCptImport ; Var IndSp : Integer) : Boolean ;
Var fb : tFichierBase ;
    OkSousPlan : Boolean ;
    j : Integer ;
    St : String ;
    Code : String ;
BEGIN
Result:=FALSE ;
OkSousPlan:=VerifSousPlan(SousSection,IndSP) ;
If (Not OkSousPlan) Or (IndSP<0) Then BEGIN OkCreat:=FALSE ; Result:=FALSE ; Exit ; END ;
If IndSP=0 Then BEGIN OkCreat:=TRUE ; Result:=FALSE ; Exit ; END ;
fb:=AxeTofb(SousSection.Axe) ;
If VH^.SousPlanAxe[fb,indSP].ListeSP<>NIL Then
  For j:=0 To VH^.SousPlanAxe[fb,IndSP].ListeSP.Count-1 Do
    BEGIN
    St:=VH^.SousPlanAxe[fb,IndSP].ListeSP[j] ;
    Code:='' ; If St<>'' Then Code:=readtokenSt(St) ;
    If Code=SousSection.Cpt Then BEGIN OkCreat:=FALSE ; Result:=TRUE ; Break ; END ;
    END ;
END ;

Function ImportSousSection(St : String ; Var InfoImp : TInfoImport ; QFiche : TQFiche) : tResultImportCpte ;
Var Q : TQuery ;
    SousSection : TCptImport ;
    OkCreat,Existe : Boolean ;
    IndSP : Integer ;
    fb : tFichierBase ;
BEGIN
Result:=ResRien ;
OkCreat:=RecupSousSection(St,SousSection,InfoImp) ;
Existe:=PlanSousSectionVide(OkCreat,SousSection,IndSP) ;
If Existe Or ((Not Existe) And OkCreat) Then
  BEGIN
  (*
  Q:=InitSQLCpt(11) ;
  If Existe Then Q.Params[0].AsString:=SousSection.Cpt Else Q.Params[0].AsString:=w_w ;
  Q.Params[1].AsString:=SousSection.Axe ;
  Q.Params[2].AsString:=SousSection.T0 ;
  Q.Open ;
  *)
  Q:=OpenSQL(InitSQLCpt1(11,SousSection.Cpt,SousSection.Axe,SousSection.T0,Existe),FALSE) ;
  Result:=ResModifier ;
  If Existe Then Q.Edit Else BEGIN Q.Insert ; InitNew(Q) ; Result:=ResCreer ; END ;
  If Not Existe Then
    BEGIN
    Q.FindField('PS_AXE').AsString:=SousSection.Axe ;
    Q.FindField('PS_SOUSSECTION').AsString:=SousSection.T0 ;
    Q.FindField('PS_CODE').AsString:=SousSection.Cpt ;
    END ;
  Q.FindField('PS_LIBELLE').AsString:=SousSection.Libelle ;
  If (Not Existe) And (IndSP>0) Then
    BEGIN
    fb:=AxeTofb(SousSection.Axe) ;
    If VH^.SousPlanAxe[fb,IndSP].ListeSP=NIL Then VH^.SousPlanAxe[fb,IndSP].ListeSP:=HTStringList.Create ;
    VH^.SousPlanAxe[fb,IndSP].ListeSP.Add(SousSection.Cpt+';'+SousSection.Libelle+';') ;
    END ;
  Q.Post ;
  Q.Close ; Ferme(Q) ;
  END ;
END ;


Procedure VerifSouche(Jal : TCptImport) ;
BEGIN
VerifUneSouche(Jal.SoucheN,(Jal.Axe<>''),FALSE) ;
VerifUneSouche(Jal.SoucheS,FALSE,TRUE) ;
END ;

Function ImportJournal(St : String ; Var InfoImp : TInfoImport ; QFiche : TQFiche) : tResultImportCpte ;
Var Q : TQuery ;
    CptLu : TCptLu ;
    Jal : TCptImport ;
    OkCreat,Existe : Boolean ;
//    leFb : TFichierBase ;
    SectOrigine : String ;
BEGIN
Result:=ResRien ;
OkCreat:=RecupJal(St,Jal,InfoImp) ; SectOrigine:=Jal.Cpt ; TraiteCorrespCpt(8,QFiche,Jal,InfoImp) ;
CptLu.Cpt:=Jal.Cpt ;
Existe:=AlimLTabCptLu(3,QFiche[3],InfoImp.LJalLu,Nil,CptLu) ;
If Existe Or ((Not Existe) And OkCreat) Then
  BEGIN
  (*
  Q:=InitSQLCpt(3) ;
  If Existe Then Q.Params[0].AsString:=Jal.Cpt Else Q.Params[0].AsString:=w_w ;
  Q.Open ;
  *)
  Q:=OpenSQL(InitSQLCpt1(3,Jal.Cpt,'','',Existe),FALSE) ;
  Result:=ResModifier ;
  If Existe Then Q.Edit Else BEGIN Q.Insert ; InitNew(Q) ; Result:=ResCreer ; END ;
  If Not Existe Then
    BEGIN
    Q.FindField('J_JOURNAL').AsString:=Jal.Cpt ;
    Q.FindField('J_NATUREJAL').AsString:=Jal.Nature ;
    Q.FindField('J_COMPTEURNORMAL').AsString:=Jal.SoucheN ;
    Q.FindField('J_COMPTEURSIMUL').AsString:=Jal.SoucheS ;
    Q.FindField('J_AXE').AsString:=Jal.Axe ;
    Q.FindField('J_MODESAISIE').AsString:=Jal.ModeSaisie ;
    Q.FindField('J_MULTIDEVISE').AsString:='X' ;
    Q.FindField('J_COMPTEINTERDIT').AsString:=Jal.CptInt ;
    Q.FindField('J_COMPTEAUTOMAT').AsString:=Jal.CptAuto ;
    If Jal.Nature='BQE' Then Q.FindField('J_TYPECONTREPARTIE').AsString:='LIG' ;
    If Jal.Nature='CAI' Then Q.FindField('J_TYPECONTREPARTIE').AsString:='MAN' ;
    END ;
  // CA - 15/01/2003 - Si libellé pas renseigné, on met le code journal
  if Jal.Libelle<>'' then Q.FindField('J_LIBELLE').AsString:=Jal.Libelle
  else Q.FindField('J_LIBELLE').AsString:=Jal.Cpt;
  Q.FindField('J_CONTREPARTIE').AsString:=Jal.Collectif ;
  if GetRecupPCL then Q.FindField('J_COMPTEAUTOMAT').AsString := Jal.Collectif;
  Q.FindField('J_ABREGE').AsString:=Jal.Abrege ;
//  If ((Not Existe) And OkCreat) Then Q.FindField('J_CREERPAR').AsString:='IMP' ;
  If ((Not Existe) And OkCreat) Then SetCreerPar(Q,'J_CREERPAR') ;
  Q.Post ;
  Q.Close ; Ferme(Q) ;
  If OkCreat Then VerifSouche(Jal) ;
  END ;
END ;

Function TraiteImportCompte(St : String ; Var InfoImp : TInfoImport ; QFiche : TQFiche) : tResultImportCpte ;
Var Qui : String ;
BEGIN
Result:=ResRien ;
Qui:=Copy(St,4,3) ;
If Qui='CGE' Then Result:=ImportGene(St,InfoImp,QFiche) Else
  If Qui='CAU' Then Result:=ImportTiers(St,InfoImp,QFiche,0) Else
    If Qui='CAS' Then Result:=ImportTiers(St,InfoImp,QFiche,1) Else
      If Qui='CAE' Then Result:=ImportTiers(St,InfoImp,QFiche,2) Else
        If Qui='SAN' Then Result:=ImportSect(St,InfoImp,QFiche) Else
         If Qui='SAT' Then Result:=ImportSect(St,InfoImp,QFiche,TRUE) Else
           If Copy(Qui,1,2)='TL' Then Result:=ImportTL(St(*,InfoImp,QFiche*)) Else
             If Copy(Qui,1,3)='SSA' Then Result:=ImportSousSection(St,InfoImp,QFiche) Else
                If Copy(Qui,1,3)='JAL' Then Result:=ImportJournal(St,InfoImp,QFiche) ;
END ;

Function EstUneLigneCptOk(St,Code : String) : Boolean ;
Var Cod : String ;
BEGIN
Cod:=Trim(Copy(St,1,3)) ;
Result:=(Cod='¥¥¥') Or (Cod='***') Or (Cod='###') ;
If Result Then Result:=(Trim(Copy(St,4,3))=Code) ;
END ;

Procedure ChargeScenario(Import,Format,Lequel,CodeFormat : String ; Var Scenario : TScenario ; Auto : Boolean) ;
Var Q : TQuery ;
    St : String ;
BEGIN
St:='Select * FROM FMTSUP WHERE FS_IMPORT="'+Import+'" AND FS_NATURE="'+Lequel+'" AND FS_FORMAT="'+Format+'" ' ;
If (*Auto And*) (CodeFormat<>'') Then St:=St+' AND FS_CODE="'+CodeFormat+'" ' ;
Q:=OpenSQL(St,TRUE) ;
Fillchar(Scenario,SizeOf(Scenario),#0) ;
ChargeFourchetteCompte(COTYPEIMP,Scenario.FourchetteImport) ;
If (Not Q.Eof) And (Not GetRecupSISCOPGI) Then With Scenario Do
  BEGIN
  Scenario.EstCharge:=TRUE ;
  If Q.Findfield('FS_NATMVT')<>NIL Then Nature:=Trim(Q.Findfield('FS_NATMVT').AsString) ;
  If Import='-' Then
    BEGIN
    If Q.Findfield('FS_FILTREGENE')<>NIL Then FiltreGen:=Q.Findfield('FS_FILTREGENE').AsString='X' ;
    If Q.Findfield('FS_FILTREAUX')<>NIL Then FiltreAux:=Q.Findfield('FS_FILTREAUX').AsString='X' ;
    If Q.Findfield('FS_NATUREGENE')<>NIL Then NatGen:=Trim(Q.Findfield('FS_NATUREGENE').AsString) ;
    If Q.Findfield('FS_NATUREAUX')<>NIL Then NatAux:=Trim(Q.Findfield('FS_NATUREAUX').AsString) ;
    END Else
    BEGIN
    If Q.Findfield('FS_GENATTEND')<>NIL Then RGen:=Trim(Q.Findfield('FS_GENATTEND').AsString) ;
    If Q.Findfield('FS_CLIATTEND')<>NIL Then RCli:=Trim(Q.Findfield('FS_CLIATTEND').AsString) ;
    If Q.Findfield('FS_FOUATTEND')<>NIL Then RFou:=Trim(Q.Findfield('FS_FOUATTEND').AsString) ;
    If Q.Findfield('FS_SALATTEND')<>NIL Then RSal:=Trim(Q.Findfield('FS_SALATTEND').AsString) ;
    If Q.Findfield('FS_DIVATTEND')<>NIL Then RDiv:=Trim(Q.Findfield('FS_DIVATTEND').AsString) ;
    If Q.Findfield('FS_SECTION1') <>NIL Then RSect1:=Trim(Q.Findfield('FS_SECTION1').AsString) ;
    If Q.Findfield('FS_SECTION2') <>NIL Then RSect2:=Trim(Q.Findfield('FS_SECTION2').AsString) ;
    If Q.Findfield('FS_SECTION3') <>NIL Then RSect3:=Trim(Q.Findfield('FS_SECTION3').AsString) ;
    If Q.Findfield('FS_SECTION4') <>NIL Then RSect4:=Trim(Q.Findfield('FS_SECTION4').AsString) ;
    If Q.Findfield('FS_SECTION5') <>NIL Then RSect5:=Trim(Q.Findfield('FS_SECTION5').AsString) ;
    If Q.Findfield('FS_MPDEFAUT') <>NIL Then RMP:=Trim(Q.Findfield('FS_MPDEFAUT').AsString) ;
    If Q.Findfield('FS_MRDEFAUT') <>NIL Then RMR:=Trim(Q.Findfield('FS_MRDEFAUT').AsString) ;
    If Q.Findfield('FS_TREGDEFAUT') <>NIL Then RGT:=Trim(Q.Findfield('FS_TREGDEFAUT').AsString) ;
    If Q.Findfield('FS_COLLCLI')<>NIL Then If Q.Findfield('FS_COLLCLI').AsString='X' Then RCollCliT:=Trim(Q.Findfield('FS_CPTCOLLCLI').AsString) ;
    If Q.Findfield('FS_COLFOU')<>NIL Then If Q.Findfield('FS_COLFOU').AsString='X' Then RCollFouT:=Trim(Q.Findfield('FS_CPTCOLLFOU').AsString) ;
    If Q.Findfield('FS_CORRIMPG')<>NIL Then BEGIN CorrespGen:=Q.Findfield('FS_CORRIMPG').AsString='X' ; If CorrespGen Then UseCorresp:=TRUE ; END ;
    If Q.Findfield('FS_CORRIMPT')<>NIL Then BEGIN CorrespAux:=Q.Findfield('FS_CORRIMPT').AsString='X' ; If CorrespAux Then UseCorresp:=TRUE ; END ;
    If Q.Findfield('FS_CORRIMP1')<>NIL Then BEGIN CorrespSect1:=Q.Findfield('FS_CORRIMP1').AsString='X' ; If CorrespSect1 Then UseCorresp:=TRUE ; END ;
    If Q.Findfield('FS_CORRIMP2')<>NIL Then BEGIN CorrespSect2:=Q.Findfield('FS_CORRIMP2').AsString='X' ; If CorrespSect2 Then UseCorresp:=TRUE ; END ;
    If Q.Findfield('FS_CORRIMP3')<>NIL Then BEGIN CorrespSect3:=Q.Findfield('FS_CORRIMP3').AsString='X' ; If CorrespSect3 Then UseCorresp:=TRUE ; END ;
    If Q.Findfield('FS_CORRIMP4')<>NIL Then BEGIN CorrespSect4:=Q.Findfield('FS_CORRIMP4').AsString='X' ; If CorrespSect4 Then UseCorresp:=TRUE ; END ;
    If Q.Findfield('FS_CORRIMP5')<>NIL Then BEGIN CorrespSect5:=Q.Findfield('FS_CORRIMP5').AsString='X' ; If CorrespSect5 Then UseCorresp:=TRUE ; END ;
    If Q.Findfield('FS_TRAITETVA')<>NIL Then TraiteTVA:=Q.Findfield('FS_TRAITETVA').AsString='X' ;
    If Q.Findfield('FS_TRAITECTR')<>NIL Then TraiteCTR:=Q.Findfield('FS_TRAITECTR').AsString='X' ;
    If Q.Findfield('FS_RIBCLIENT')<>NIL Then ForceRIB:=Q.Findfield('FS_RIBCLIENT').AsString='X' ;
    If Q.Findfield('FS_ANOUVEAUD')<>NIL Then ANDetail:=Q.Findfield('FS_ANOUVEAUD').AsString='X' ;
    If Q.Findfield('FS_CHEMIN')<>NIL Then Chemin:=Q.Findfield('FS_CHEMIN').AsString ;
    If Q.Findfield('FS_PREFIXE')<>NIL Then Prefixe:=Q.Findfield('FS_PREFIXE').AsString ;
    If Q.Findfield('FS_SUFFIXE')<>NIL Then Suffixe:=Q.Findfield('FS_SUFFIXE').AsString ;
    If Q.Findfield('FS_DOUBLON')<>NIL Then Doublon:=Q.Findfield('FS_DOUBLON').AsString='X' ;
    If Q.Findfield('FS_FORCEPIECE')<>NIL Then ForcePiece:=Q.Findfield('FS_FORCEPIECE').AsString='X' ;
    If Q.Findfield('FS_CORRIMPMP')<>NIL Then BEGIN CorrespMP:=Q.Findfield('FS_CORRIMPMP').AsString='X' ; If CorrespMP Then UseCorresp:=TRUE ; END ;
    If Q.Findfield('FS_RIBFOUR')<>NIL Then ForceRibFou:=Q.Findfield('FS_RIBFOUR').AsString='X' ;
    If Q.Findfield('FS_VALIDEECR')<>NIL Then ValideEcr:=Q.Findfield('FS_VALIDEECR').AsString='X' ;
    If Q.Findfield('FS_CORRIMPJAL')<>NIL Then BEGIN CorrespJal:=Q.Findfield('FS_CORRIMPJAL').AsString='X' ; If CorrespJal Then UseCorresp:=TRUE ; END ;
    If Q.Findfield('FS_TAUXDEVOUT')<>NIL Then CalcTauxDevOut:=Q.Findfield('FS_TAUXDEVOUT').AsString='X' ;
    If Q.Findfield('FS_FILTREAUX')<>NIL Then ShuntPbAna:=Q.Findfield('FS_FILTREAUX').AsString='X' ;
    If Q.Findfield('FS_ECCDIV')<>NIL Then Majuscule:=Q.Findfield('FS_ECCDIV').AsString='X' ;
    If Auto Then
      BEGIN
      ImpAuto:=TRUE ;
      If Q.Findfield('FS_FILTREGENE')<>NIL Then DetruitFic:=Q.Findfield('FS_FILTREGENE').AsString='X' ;
      END ;
    END ;
  END ;
Ferme(Q) ;
{$IFDEF RECUPPCL}
If (GetRecupPCL) And (Not GetRecupCegid) Then
  BEGIN
  Scenario.RMR:='002' ;
  Scenario.RMP:='DIV' ;
  Scenario.RGT:='FRA' ;
  Scenario.ForcePiece:=TRUE ;
  END ;
{$ENDIF}
If (GetRecupSISCOPGI) Then
  BEGIN
  Scenario.RMR:='002' ;
  Scenario.RMP:='DIV' ;
  Scenario.RGT:='FRA' ;
  Scenario.ForcePiece:=TRUE ;
  Scenario.ShuntPbAna:=TRUE ;
  END ;
END ;

Procedure EcrireSeparateur(Var FicRap : TextFile) ;
Var StR : String ;
BEGIN
Str:='***************************************************************************************** ' ; Writeln(FicRap,StR) ;
END ;

Procedure EcrireRapportDoublons(Var FicRap : TextFile ; Var InfoImp : TInfoImport) ;
Var StR : String ;
    j : Integer ;
BEGIN
If InfoImp.CtrlDB Then
  BEGIN
  Str:=' ' ; Writeln(FicRap,StR) ;
  EcrireSeparateur(FicRap) ;
  StR:=TraduireMemoire(' Contrôle de doublons activé') ; Writeln(FicRap,StR) ;
  If (InfoImp.ListeEnteteDoublon<>Nil) And (InfoImp.ListeEnteteDoublon.Count>0) Then
    BEGIN
    StR:=TraduireMemoire(' Des doublons ont été détectés :') ; Writeln(FicRap,StR) ;
    StR:=' ' ; Writeln(FicRap,StR) ;
    StR:=TraduireMemoire(' Jal / N° pièce dans fichier   *      Identification dans la base comptable') ; Writeln(FicRap,StR) ;
    For j:=0 To InfoImp.CRListeEnteteDoublon.Count-1 Do
      BEGIN
      StR:=' '+Format_String(DelInfo(InfoImp.CRListeEnteteDoublon.Items[j]).LeCod,3)+' / ' ;
      StR:=StR+Format_String(DelInfo(InfoImp.CRListeEnteteDoublon.Items[j]).LeLib,23)+' * ' ;
      StR:=StR+DelInfo(InfoImp.CRListeEnteteDoublon.Items[j]).LeMess ; Writeln(FicRap,StR) ;
      END ;
    StR:=' ' ; Writeln(FicRap,StR) ;
    StR:=TraduireMemoire(' Les Doublons du fichier d''origine ont été stockés dans le fichier ')+InfoImp.NomFicDoublon ; Writeln(FicRap,StR) ;
    END Else
    BEGIN
    StR:=TraduireMemoire(' Pas de doublon(s) détecté(s)') ; Writeln(FicRap,StR) ;
    END ;
  EcrireSeparateur(FicRap) ;
  END Else
  BEGIN
  Str:=' ' ; Writeln(FicRap,StR) ;
  EcrireSeparateur(FicRap) ;
  StR:=TraduireMemoire(' Contrôle de doublons non activé') ;
  Writeln(FicRap,StR) ;
  EcrireSeparateur(FicRap) ;
  END ;
END ;

Procedure EcrireRapportRejets(Var FicRap : TextFile ; Var InfoImp : TInfoImport) ;
Var StR : String ;
    j : Integer ;
BEGIN
If (InfoImp.ListeEntetePieceFausse=Nil) Or (InfoImp.ListeEntetePieceFausse.Count=0) Then
  BEGIN
  Str:=' ' ; Writeln(FicRap,StR) ;
  EcrireSeparateur(FicRap) ;
  StR:=TraduireMemoire(' Aucune erreur détectée : Pas de fichier de Rejet') ;
  Writeln(FicRap,StR) ;
  EcrireSeparateur(FicRap) ;
  END Else
  BEGIN
  Str:=' ' ; Writeln(FicRap,StR) ;
  EcrireSeparateur(FicRap) ;
  StR:=TraduireMemoire(' Des erreurs ont été détectées : ') ; Writeln(FicRap,StR) ;
  StR:=' ' ; Writeln(FicRap,StR) ;
  StR:=TraduireMemoire(' Jal *          Référence               * Pièce/Ligne *    Date    * Commentaires') ; Writeln(FicRap,StR) ;
  For j:=0 To InfoImp.ListePieceFausse.Count-1 Do
    BEGIN
    StR:=' '+Format_String(DelInfo(InfoImp.ListePieceFausse.Items[j]).LeCod,3)+' * ' ;
    StR:=StR+Format_String(DelInfo(InfoImp.ListePieceFausse.Items[j]).LeLib,32)+' * ' ;
    StR:=StR+Format_String(DelInfo(InfoImp.ListePieceFausse.Items[j]).LeMess,11)+' * ' ;
    StR:=StR+Format_String(DelInfo(InfoImp.ListePieceFausse.Items[j]).LeMess2,10)+' * ' ;
    StR:=StR+DelInfo(InfoImp.ListePieceFausse.Items[j]).LeMess3 ;
    Writeln(FicRap,StR) ;
    END ;
  StR:=' ' ; Writeln(FicRap,StR) ;
  StR:=TraduireMemoire(' Les pièces fausses du fichier d''origine ont été stockées dans le fichier ')+InfoImp.NomFicRejet ; Writeln(FicRap,StR) ;
  EcrireSeparateur(FicRap) ;
  END ;
END ;

Procedure EcrireRapportDivers(What : Integer ; InfoImp : TInfoImport) ;
Var FicRap : TextFile ;
    StR : String ;
BEGIN
If InfoImp.NomFicRapport='' Then Exit ;
If InfoImp.NomFic='' Then Exit ;
AssignFile(FicRap,InfoImp.NomFicRapport) ;
{$i-} Append(FicRap) ; {$i+}
If IoResult<>0 Then Exit ;
StR:=' ' ; Writeln(FicRap,StR) ;
EcrireSeparateur(FicRap) ;
Case What Of
  0 : BEGIN
      EcrireSeparateur(FicRap) ;
      StR:=' Aucun mouvement n''a pu être intégré' ; Writeln(FicRap,StR) ;
      EcrireSeparateur(FicRap) ;
      END ;
  1 : BEGIN
      EcrireSeparateur(FicRap) ;
      StR:=' Des mouvements ont été intégrés : ' ; Writeln(FicRap,StR) ;
      EcrireSeparateur(FicRap) ;
      END ;
  END ;
CloseFile(FicRap) ;
END ;

Procedure EcrireRapportGlobal(Var InfoImp : TInfoImport ; PbFic : Boolean ; StErr : String = '' ; OkSep : Boolean = TRUE) ;
Var FicRap : TextFile ;
    StR : String ;
BEGIN
If InfoImp.NomFicRapportGlobal='' Then Exit ;
If InfoImp.NomFic='' Then Exit ;
If InfoImp.NomFicOrigine='' Then Exit ;
If Not GetRecupPCL Then If InfoImp.TypeFic<>'1' Then Exit ;
AssignFile(FicRap,InfoImp.NomFicRapportGlobal) ;
{$i-} Append(FicRap) ; {$i+}
If IoResult<>0 Then Exit ;
StR:=' ' ; Writeln(FicRap,StR) ;
If OkSep Then EcrireSeparateur(FicRap) ;
If StErr<>'' Then
  BEGIN
  Writeln(FicRap,StErr) ;
  END Else
  BEGIN
  If GetRecupPCL Then
    BEGIN
    StR:=TraduireMemoire('   IMPORTATION DU FICHIER ')+InfoImp.NomFicOrigine ; Writeln(FicRap,StR) ;
    StR:=TraduireMemoire('   DANS LA SOCIETE ')+V_PGI.NomSociete ; Writeln(FicRap,StR) ;
    If VH^.FromPCL Then BEGIN StR:=TraduireMemoire('   DOSSIER ')+V_PGI.NoDossier ; Writeln(FicRap,StR) ; END ;
    END Else
    BEGIN
    StR:=TraduireMemoire('   IMPORTATION DU FICHIER ')+InfoImp.NomFicOrigine ; Writeln(FicRap,StR) ;
    StR:=TraduireMemoire('   DANS LA SOCIETE ')+V_PGI.NomSociete ; Writeln(FicRap,StR) ;
    END ;
  Case PbFic Of
    TRUE : BEGIN
        StR:=' Des erreurs ont été détectées.' ; Writeln(FicRap,StR) ;
        StR:=' Aucun mouvement n''a pu être intégré.' ; Writeln(FicRap,StR) ;
        If OkSep Then EcrireSeparateur(FicRap) ;
        END ;
    FALSE : BEGIN
        StR:=' Les mouvements ont été intégrés. ' ; Writeln(FicRap,StR) ;
        If OkSep Then EcrireSeparateur(FicRap) ;
        END ;
    END ;
  END ;
CloseFile(FicRap) ;
END ;

Procedure EcrireRapportDebutIntegration(PasDeMvtIntegres : Boolean ; InfoImp : TInfoImport ; Var FicRap : TextFile ; Var OkRapport : Boolean) ;
Var StR : String ;
BEGIN
OkRapport:=FALSE ;
If InfoImp.NomFicRapport<>'' Then
  BEGIN
  AssignFile(FicRap,InfoImp.NomFicRapport) ;
  {$i-} Append(FicRap) ; {$i+}
  If IoResult=0 Then OkRapport:=TRUE ;
  END ;
If PasDeMvtIntegres Then
  BEGIN
  If OkRapport Then
    BEGIN
    EcrireRapportRejets(FicRap,InfoImp) ;
    EcrireRapportDoublons(FicRap,InfoImp) ;
    CloseFile(FicRap) ;
    END ;
  END Else
  BEGIN
  If OkRapport Then
    BEGIN
    StR:=' ' ; Writeln(FicRap,StR) ;
    EcrireSeparateur(FicRap) ;
    StR:=TraduireMemoire(' Des mouvements ont été intégrés : ') ; Writeln(FicRap,StR) ;
    END ;
  END ;
END ;


Function EstFichierSISCO(StFichier : String ; OkMess : Boolean) : Boolean ;
Var Fichier : TextFile ;
    St : String ;
BEGIN
Result:=FALSE ; St:='' ;
AssignFile(Fichier,StFichier) ; {$i-} Reset(Fichier) ; {$i+}
If IoResult<>0 Then St:='0;Fichier '+StFichier+';Fichier Introuvable;E;O;O;O;' Else
  BEGIN
  If Not Eof(Fichier) THen
    BEGIN
    Readln(Fichier,St) ;
    If Copy(St,1,11)='***DEBUT***' Then Result:=TRUE Else St:='0;Fichier '+StFichier+';Ce n''est pas un fichier SISCO !;E;O;O;O;' ;
    END Else St:='0;Fichier '+StFichier+';Ce fichier est vide !;E;O;O;O;' ;
  System.Close(Fichier) ;
  END ;
If (Not Result) And (St<>'') And OkMess Then HShowMessage(St,'','') ;
END ;

Function ChercheRibAux(Aux : String ; Var Dom,Etab,Guichet,NumCpt,Cle : String) : Boolean ;
Var Q : TQuery ;
BEGIN
Dom:='' ; Etab:='' ; Guichet:='' ; NumCpt:='' ; Cle:='' ; Result:=FALSE ; 
Q:=OpenSQL('SELECT * FROM RIB WHERE R_AUXILIAIRE="'+Aux+'" AND R_PRINCIPAL="X"',True) ;
if Not Q.Eof then
  BEGIN
  Dom:=Q.FindField('R_DOMICILIATION').AsString ;
  Etab:=Q.FindField('R_ETABBQ').AsString ;
  Guichet:=Q.FindField('R_GUICHET').AsString ;
  NumCpt:=Q.FindField('R_NUMEROCOMPTE').AsString ;
  Cle:=Q.FindField('R_CLERIB').AsString ;
  Result:=TRUE ;
  END ;
Ferme(Q) ;
END ;

Procedure ChercheContactAux(Aux : String ; Var Ct : TContact) ;
Var Q : TQuery ;
BEGIN
Fillchar(Ct,SizeOf(Ct),#0) ; 
Q:=OpenSQL('SELECT * FROM CONTACT WHERE C_AUXILIAIRE="'+Aux+'" AND C_PRINCIPAL="X"',True) ;
if Not Q.Eof then
  BEGIN
  Ct.Nom:=Q.Findfield('C_NOM').ASString ;
  Ct.Service:=Q.Findfield('C_SERVICE').ASString ;
  Ct.Fonction:=Q.Findfield('C_FONCTION').ASString ;
  Ct.Tel:=Q.Findfield('C_TELEPHONE').ASString ;
  Ct.Fax:=Q.Findfield('C_FAX').ASString ;
  Ct.Principal:=Q.Findfield('C_PRINCIPAL').ASString ;
  Ct.Telex:=Q.Findfield('C_TELEX').ASString ;
  Ct.RVA:=Q.Findfield('C_RVA').ASString ;
  Ct.Civilite:=Q.Findfield('C_CIVILITE').ASString ;
  END ;
Ferme(Q) ;
END ;

Procedure ExporteSousSection(Var F : TextFile ; NomFic : String ; OkAppend,OkMove : Boolean) ;
Var fb : tFichierBase ;
    i,j : Integer ;
    DebutSP,LgSP : Integer ;
    CodeSP,LibSP,Code,Lib : String ;
    StF,St : String ;
    OkOk : Boolean ;
BEGIN
If NomFic<>'' Then
  BEGIN
  AssignFile(F,Nomfic) ; OkOk:=TRUE ;
  If OkAppend Then
    BEGIN
    {$I-} Append(F) ; {$I+}
    If IoResult=0 Then OkOk:=FALSE ;
    END ;
  If OkOk Then ReWrite (F) ;
  END ;
If OkMove Then InitMove(5*MaxSousPlan,'') ;
For fb:=fbAxe1 to fbAxe5 Do For i:=1 To MaxSousPlan Do
  BEGIN
  If OkMove Then MoveCur(FALSE) ;
  DebutSP:=VH^.SousPlanAxe[fb,i].Debut ; LgSP:=VH^.SousPlanAxe[fb,i].Longueur ;
  CodeSP:=VH^.SousPlanAxe[fb,i].Code ;   LibSP:=VH^.SousPlanAxe[fb,i].Lib ;
  If (CodeSP<>'') And  (VH^.SousPlanAxe[fb,i].ListeSP<>NIL) Then
    for j:=0 to VH^.SousPlanAxe[fb,i].ListeSP.Count-1 Do
      BEGIN
      St:=VH^.SousPlanAxe[fb,i].ListeSP[j] ; StF:='' ;
      If St<>'' Then Code:=readtokenSt(St) ; If St<>'' Then Lib:=readtokenSt(St) ;
      If (Code<>'') Then
        BEGIN
        StF:='***SSA'+Format_String(Code,17)
                     +Format_String(Lib,35)
                     +Format_String(fbToAxe(fb),3)
                     +Format_String(CodeSP,3)
                     +formatfloat('00',DebutSP)
                     +formatfloat('00',LgSP)
                     +Format_String(LibSP,35) ;
        END ;
      If StF<>'' Then Writeln(F,StF) ;
      END ;
  END ;
If OkMove Then FiniMove ;
If NomFic<>'' Then CloseFile(F) ;
END ;

Function FaitStCPTCEGID(Qui : Char ; Q : TQuery) : String ;
Var Dom,Etab,Guichet,NumCpt,Cle,LERIB,Cpt,RIBPrincipal : String ;
    Ct : TContact ;
    St1 : String ;
BEGIN
Result:='' ;
Case Qui Of
  'G' : BEGIN
        St1:='***CGE'+Format_String(Q.FindField('G_GENERAL').AsString,17)
                     +Format_String(Q.FindField('G_LIBELLE').AsString,35)
                     +Format_String(Q.FindField('G_NATUREGENE').AsString,3)
                     +Format_String(Q.FindField('G_LETTRABLE').AsString,1)
                     +Format_String(Q.FindField('G_POINTABLE').AsString,1)
                     +Format_String(Q.FindField('G_VENTILABLE1').AsString,1)
                     +Format_String(Q.FindField('G_VENTILABLE2').AsString,1)
                     +Format_String(Q.FindField('G_VENTILABLE3').AsString,1)
                     +Format_String(Q.FindField('G_VENTILABLE4').AsString,1)
                     +Format_String(Q.FindField('G_VENTILABLE5').AsString,1)
                     +Format_String(Q.FindField('G_TABLE0').AsString,3)
                     +Format_String(Q.FindField('G_TABLE1').AsString,3)
                     +Format_String(Q.FindField('G_TABLE2').AsString,3)
                     +Format_String(Q.FindField('G_TABLE3').AsString,3)
                     +Format_String(Q.FindField('G_TABLE4').AsString,3)
                     +Format_String(Q.FindField('G_TABLE5').AsString,3)
                     +Format_String(Q.FindField('G_TABLE6').AsString,3)
                     +Format_String(Q.FindField('G_TABLE7').AsString,3)
                     +Format_String(Q.FindField('G_TABLE8').AsString,3)
                     +Format_String(Q.FindField('G_TABLE9').AsString,3)
                     +Format_String(Q.FindField('G_ABREGE').AsString,17)
                     +Format_String(Q.FindField('G_SENS').AsString,3) ;
        END ;
  'X' : BEGIN
        RibPrincipal:='-' ;
        Cpt:=Q.FindField('T_AUXILIAIRE').AsString ;
        If ChercheRIBAux(Cpt,Dom,Etab,Guichet,NumCpt,Cle) Then RibPrincipal:='X' ;
        LERIB:=Format_String(Dom,24)
              +Format_String(Etab,5)+Format_String(Guichet,5)
              +Format_String(NumCpt,11)+Format_String(Cle,2) ;
        ChercheContactAux(Cpt,Ct) ;
        St1:='***CAU'+Format_String(Q.FindField('T_AUXILIAIRE').AsString,17)
                     +Format_String(Q.FindField('T_LIBELLE').AsString,35)
                     +Format_String(Q.FindField('T_NATUREAUXI').AsString,3)
                     +Format_String(Q.FindField('T_LETTRABLE').AsString,1)
                     +Format_String(Q.FindField('T_COLLECTIF').AsString,17)
                     +Format_String(Q.FindField('T_EAN').AsString,17)
                     +Format_String(Q.FindField('T_TABLE0').AsString,3)
                     +Format_String(Q.FindField('T_TABLE1').AsString,3)
                     +Format_String(Q.FindField('T_TABLE2').AsString,3)
                     +Format_String(Q.FindField('T_TABLE3').AsString,3)
                     +Format_String(Q.FindField('T_TABLE4').AsString,3)
                     +Format_String(Q.FindField('T_TABLE5').AsString,3)
                     +Format_String(Q.FindField('T_TABLE6').AsString,3)
                     +Format_String(Q.FindField('T_TABLE7').AsString,3)
                     +Format_String(Q.FindField('T_TABLE8').AsString,3)
                     +Format_String(Q.FindField('T_TABLE9').AsString,3)
                     +Format_String(Q.FindField('T_ADRESSE1').AsString,35)
                     +Format_String(Q.FindField('T_ADRESSE2').AsString,35)
                     +Format_String(Q.FindField('T_ADRESSE3').AsString,35)
                     +Format_String(Q.FindField('T_CODEPOSTAL').AsString,9)
                     +Format_String(Q.FindField('T_VILLE').AsString,35)
                     +LERIB
                     +Format_String(Q.FindField('T_PAYS').AsString,3)
                     +Format_String(Q.FindField('T_ABREGE').AsString,17)
                     +Format_String(Q.FindField('T_LANGUE').AsString,3)
                     +Format_String(Q.FindField('T_MULTIDEVISE').AsString,1)
                     +Format_String(Q.FindField('T_DEVISE').AsString,3)
                     +Format_String(Q.FindField('T_TELEPHONE').AsString,25)
                     +Format_String(Q.FindField('T_FAX').AsString,25)
                     +Format_String(Q.FindField('T_REGIMETVA').AsString,3)
                     +Format_String(Q.FindField('T_MODEREGLE').AsString,3)
                     +Format_String(Q.FindField('T_COMMENTAIRE').AsString,35)
                     +Format_String(Q.FindField('T_NIF').AsString,17)
                     +Format_String(Q.FindField('T_SIRET').AsString,17)
                     +Format_String(Q.FindField('T_APE').AsString,5)
                     +Format_String(Ct.Nom,35)
                     +Format_String(Ct.Service,35)
                     +Format_String(Ct.Fonction,35)
                     +Format_String(Ct.Tel,25)
                     +Format_String(Ct.Fax,25)
                     +Format_String(Ct.Telex,25)
                     +Format_String(Ct.RVA,50)
                     +Format_String(Ct.Civilite,3)
                     +Format_String(Ct.Principal,1)
                     +Format_String(Q.FindField('T_JURIDIQUE').AsString,3)
                     +Format_String(RibPrincipal,1)
                     +Format_String(Q.FindField('T_PAYEUR').AsString,17)
                     +Format_String(Q.FindField('T_ISPAYEUR').AsString,1)
                     +Format_String(Q.FindField('T_AVOIRRBT').AsString,1) ;
        END ;
  'S' : BEGIN
        St1:='***SAN'+Format_String(Q.FindField('S_SECTION').AsString,17)
                     +Format_String(Q.FindField('S_LIBELLE').AsString,35)
                     +Format_String(Q.FindField('S_AXE').AsString,3)
                     +Format_String(Q.FindField('S_TABLE0').AsString,3)
                     +Format_String(Q.FindField('S_TABLE1').AsString,3)
                     +Format_String(Q.FindField('S_TABLE2').AsString,3)
                     +Format_String(Q.FindField('S_TABLE3').AsString,3)
                     +Format_String(Q.FindField('S_TABLE4').AsString,3)
                     +Format_String(Q.FindField('S_TABLE5').AsString,3)
                     +Format_String(Q.FindField('S_TABLE6').AsString,3)
                     +Format_String(Q.FindField('S_TABLE7').AsString,3)
                     +Format_String(Q.FindField('S_TABLE8').AsString,3)
                     +Format_String(Q.FindField('S_TABLE9').AsString,3)
                     +Format_String(Q.FindField('S_ABREGE').AsString,17)
                     +Format_String(Q.FindField('S_SENS').AsString,3) ;
        END ;
  END ;
Result:=St1 ;
END ;


Function LibSousSection(Sect,Axe : String) : String ;
Var i,j : Integer ;
    fb : tFichierBase ;
    DebutSP,LgSP : Integer ;
    St,Code,Lib,CodeSP,LibSP : String ;
BEGIN
Result:='' ; fb:=AxeTofb(Axe) ; Lib:='' ;
For i:=1 To MaxSousPlan Do
  BEGIN
  DebutSP:=VH^.SousPlanAxe[fb,i].Debut ; LgSP:=VH^.SousPlanAxe[fb,i].Longueur ;
  CodeSP:=VH^.SousPlanAxe[fb,i].Code ;   LibSP:=VH^.SousPlanAxe[fb,i].Lib ;
  If (DebutSP=0) Or (LgSP=0) Then Continue ;
  Code:=Copy(Sect,DebutSP,LgSP) ;
  If (CodeSP<>'') And  (VH^.SousPlanAxe[fb,i].ListeSP<>NIL) Then
    for j:=0 to VH^.SousPlanAxe[fb,i].ListeSP.Count-1 Do
      BEGIN
      St:=VH^.SousPlanAxe[fb,i].ListeSP[j] ; CodeSP:='' ; LibSP:='' ;
      If St<>'' Then CodeSP:=readtokenSt(St) ; If St<>'' Then LibSP:=readtokenSt(St) ;
      If (CodeSP<>'') And (CodeSP=Code) Then BEGIN Lib:=Lib+LibSP+';' ; Break ; ENd ;
      END ;
    End ;
if Length(Lib)>35 then
   BEGIN
    While Length(Lib)>35 do
      BEGIN
        for i:=1 to Length(Lib) do
        if Lib[i]=';' then Delete(Lib,i-1,1) ;
      END ;
   END ;
While Pos(';',Lib) > 0 do Lib[Pos(';',Lib)]:=' ' ;
Result:=Lib ;
END ;

Function LitDebFormat(SFormatFic : String ; Var DebFormat : tDebFormat) : Boolean ;
Var FormatFic : Integer ;
BEGIN
Result:=TRUE ; Fillchar(DebFormat,SizeOf(DebFormat),#0) ; FormatFic:=-1 ;
If (SFormatFic='SAA') Or (SFormatFic='SN2') Then FormatFic:=0 Else
 If (SFormatFic='HLI') Then FormatFic:=1 Else
  If (SFormatFic='HAL') Then FormatFic:=2 Else
   If (SFormatFic='CGN') Then FormatFic:=3 Else
    If (SFormatFic='CGE') Then FormatFic:=4 ;
DebFormat.FormatFic:=FormatFic ;
With DebFormat Do
  Case FormatFic Of
      0 : BEGIN { SAARI COEUR DE GAMME ET NEGOCE V2 }
          DebNum:=106 ; DebMontant:=85 ; DebSens:=84 ; DebTC:=25 ; DebNTP:=10 ; LgNum:=7 ; DebGen:=12 ; LgGen:=13 ;
          END ;
      1 : BEGIN { HALLEY LIGHT }
          DebNum:=118 ; DebMontant:=97 ; DebSens:=96 ; DebTC:=31 ; DebNTP:=12 ; LgNum:=7 ; DebGen:=14 ; LgGen:=17 ;
          END ;
      2 : BEGIN { HALLEY ETENDU }
          DebNum:=118 ; DebMontant:=97 ; DebSens:=96 ; DebTC:=31 ; DebNTP:=12 ; LgNum:=7 ; DebGen:=14 ; LgGen:=17 ;
          END ;
      3,4 : BEGIN { CEGID NORMAL }
            DebNum:=152 ; DebMontant:=131 ; DebSens:=130 ; DebTC:=31 ; DebNTP:=12 ; LgNum:=8 ; DebGen:=14 ; LgGen:=17 ;
            END ;
      Else Result:=FALSE ;
    END ;
END ;


Function AffecteNumP(Var St : String ; Var EnRupture : Boolean ; Var PieceFictive,NoPiece : Integer ;
                     Var TotPieceD,TotPieceC : Double ; Var DF : tDebFormat ; Force : Boolean ;
                     Var Jal : String ; Var Periode : Word ;
                     Var InfoImp : TInfoImport ;
                     QFiche : TQFiche ; Var PasPiece : Boolean) : Boolean ;
Var StTest,Sens,LeJal : String ;
    OkNumP,OkMontant : Boolean ;
    Montant : Double ;
    CptLuJ : TCptLu ;
BEGIN
Result:=TRUE ;
PasPiece:=FALSE ;
{$IFDEF RECUPPCL}
If GetRecupPCL Then
  BEGIN
  Jal:=Trim(Copy(St,1,3)); Periode:=StrToInt(Copy(St,8,4))+StrToInt(Copy(St,6,2)) ;
  END ;
{$ELSE}
LeJal:=Trim(Copy(St,1,3));
CptLuJ.Cpt:=LeJal ;
If QFiche[3]<>NIL Then
  BEGIN
  If AlimLTabCptLu(3,QFiche[3],InfoImp.LJalLu,InfoImp.ListeCptFaux,CptLuJ) Then
    BEGIN
    If ModeJal(LeJal,InfoImp)<>mPiece Then
      BEGIN
      Jal:=Trim(Copy(St,1,3)); Periode:=StrToInt(Copy(St,8,4))+StrToInt(Copy(St,6,2)) ; PasPiece:=TRUE ;
      END ;
    END ;
  END ;
{$ENDIF}
StTest:=Trim(Copy(St,DF.DebNum,DF.LgNum)) ; OkNumP:=Force Or VerifEntier(StTest) ;
If Not Force Then
   BEGIN
   If (StTest<>'') And OkNumP Then NoPiece:=StrToInt(StTest) Else NoPiece:=0 ;
   END Else NoPiece:=0 ;
StTest:=Trim(Copy(St,DF.DebMontant,20)) ; OkMontant:=VerifMontant(StTest) ;
If (StTest<>'') And OkMontant Then Montant:=StrToFloat(StPoint(StTest)) Else Montant:=0 ;
If (Not OkMontant) Or (Not OkNumP) Then Result:=FALSE ;
Sens:=Copy(St,DF.DebSens,1) ; EnRupture:=FALSE ;
if (St[DF.DebTC]<>'A') And (St[DF.DebTC]<>'E') And (St[DF.DebTC]<>'D') then { pour shunter les écritures analytiques }
   BEGIN
   if (Arrondi(TotPieceD,V_PGI.OkDecV)=Arrondi(TotPieceC,V_PGI.OkDecV)) then
      BEGIN
      If PieceFictive>0 Then EnRupture:=TRUE ; Inc(PieceFictive) ; TotPieceD:=0 ; TotPieceC:=0 ;
      END ;
   If Sens='D' Then TotPieceD:=Arrondi(TotPieceD+Montant,V_PGI.OkDecV)
               Else TotPieceC:=Arrondi(TotPieceC+Montant,V_PGI.OkDecV) ;
   END ;
if Force Or ((NoPiece=0) And OkNumP) then
   BEGIN
   If DF.FormatFic in [3,4] Then St:=Insere(St,FormatFloat('00000000',PieceFictive),DF.DebNum,DF.LgNum)
                         Else St:=Insere(St,FormatFloat('0000000',PieceFictive),DF.DebNum,DF.LgNum) ;
   END ;
END ;

Function AffecteNatP(Var St : String ; Var DF : tDebFormat ; QFiche : TQFiche ; Var InfoImp : TInfoImport) : Boolean ;
Var NatP,NatG,Jal,Gen : String ;
    CptLuJ,CptLuG : TCptLu ;
BEGIN
Result:=FALSE ; If QFiche[3]=NIL Then Exit ;
//If Not QFiche[3].Prepared Then Exit ;
If InfoImp.LJalLu=NIL Then Exit ;
NatP:=Trim(Copy(St,DF.DebNTP,2)) ; If NatP<>'' Then Exit ;
Jal:=Trim(Copy(St,1,3)) ; NatP:='OD' ;
Fillchar(CptLuJ,SizeOf(CptLuJ),#0) ; CptLuJ.Cpt:=Trim(Jal) ;
If AlimLTabCptLu(3,QFiche[3],InfoImp.LJalLu,NIL,CptLuJ) Then
  BEGIN
  If CptLuJ.Nature='ACH' Then NatP:='FF' ;
  If CptLuJ.Nature='VTE' Then NatP:='FC' ;
  If (CptLuJ.Nature='BQE') Or (CptLuJ.Nature='CAI') Then
    BEGIN
    Gen:=Trim(Copy(St,DF.DebGen,DF.LgGen)) ; Gen:=BourreOuTronque(Gen,fbGene) ;
    If GetRecupSISCOPGI And (InfoImp.LGenLu<>NIL) And (QFiche[3]<>NIL) Then
      BEGIN
      Fillchar(CptLuG,SizeOf(CptLuG),#0) ; CptLuG.Cpt:=Gen ;
      If AlimLTabCptLu(0,QFiche[0],InfoImp.LGenLu,NIL,CptLuG) Then
        BEGIN
        NatG:=CptLuG.Nature ;
        END ;
      END Else NatG:=NatureCptImport(Gen,InfoImp.SC.FourchetteImport) ;
    If NatG='COC' Then NatP:='RC' Else
      If NatG='COF' Then NatP:='RF' ;
    END ;
  St:=Insere(St,NatP,DF.DebNTP,2) ;
  Result:=TRUE ;
  END ;
END ;

Procedure RetouchePieceAvantWrite(ListePiece : HTStringList ; Var DF : tDebFormat ; Var InfoImp : TInfoImport) ;
Var i : Integer ;
    St1 : String ;
//    OkOk : Boolean ;
    OkMajTP,ODDetecte : Boolean ;
    NatP,NatPGlobal : String ;
    CptLuJ : tCptLu ;
BEGIN
OkMajTP:=FALSE ; NatPGlobal:='' ; ODDetecte:=FALSE ;
For i:=0 To ListePiece.Count-1 Do
  BEGIN
//  OkOk:=TRUE ;
  NatP:=Trim(Copy(ListePiece[i],DF.DebNTP,2)) ;
  If NatP='OD' Then ODDetecte:=TRUE ;
  If (NatP='RC') Or (NatP='RF') Or (NatPGlobal<>'') Then
    BEGIN
    If NatPGlobal='' Then NatPGlobal:=NatP ;
    If (NatP<>NatPGlobal) Then
      BEGIN
      OkMajTP:=TRUE ;
      If ((NatP='RC') And (NatPGlobal='RF')) Or ((NatP='RF') And (NatPGlobal='RC')) Then NatPGlobal:='OD' ;
      END Else If ODDetecte Then OkMajTP:=TRUE ;
    END ;
  END ;
If OkMajTP And (InfoImp.LJalLu<>Nil) Then
  BEGIN
  Fillchar(CptLuJ,SizeOf(CptLuJ),#0) ; CptLuJ.Cpt:=Trim(Copy(ListePiece[0],1,3)) ;
  If ChercheCptLu(InfoImp.LJalLu,CptLuJ) Then
    BEGIN
    If CptLuJ.ModeSaisie='LIB' Then OkMajTP:=FALSE ;
    END ;
  END ;
If OkMajTP And ((NatPGlobal='RC') Or (NatPGlobal='RF') Or (NatPGlobal='OD')) Then
  BEGIN
  For i:=0 To ListePiece.Count-1 Do
    BEGIN
    St1:=ListePiece[i] ;
    St1:=Insere(St1,NatPGlobal,DF.DebNTP,2) ;
    ListePiece[i]:=St1 ;
    END ;
  END ;
END ;

Procedure EcritPiece(Var NewFichier : TextFile ; ListePiece : HtStringList) ;
Var i : Integer ;
BEGIN
for i:=0 to ListePiece.Count-1 do WriteLn(NewFichier,ListePiece[i]) ;
VideStringList(ListePiece) ;
END ;

FUNCTION TravailleFichier ( Force : Boolean ; Var StErr : String ; Var InfoImp : TInfoImport ; QFiche : TQFiche ; ignorel1 : boolean=False) : Boolean ;
Var Fichier,NewFichier : TextFile ;
    NewFileName : String ;
    St : String ;
    PieceFictive : Integer ;
    TotPieceD,TotPieceC : Double ;
    Longueur : integer ;
    OkOk,NatModifie : Boolean ;
    NumLigne : Integer ;
    ListePiece : HTStringList ;
    ListeCpt : HTStringList ;
    EnRupture : Boolean ;
    DebFormat : tDebFormat ;
    PasDeModif : Boolean ;
//    OkPourWriteAutre : Boolean ;
    NoPiece,OldNoPiece{,iio} : Integer ;
    Jal,OldJal : String ;
    Periode,OldPeriode : Word ;
    PremFois : Boolean ;
    PasPiece : Boolean ;
BEGIN
Result:=TRUE ; StErr:='' ; Longueur:=132 ; OldPeriode:= 0;
If Not LitDebFormat(InfoImp.Format,DebFormat) Then Exit ;
AssignFile(Fichier,InfoImp.NomFic) ; Reset(Fichier) ;
NewFileName:=FileTemp('.PNM') ;
AssignFile(NewFichier,NewFileName) ; Rewrite(NewFichier) ;
if not ignorel1 then begin ReadLn(Fichier,St) ; St:=Format_String(St,Longueur) ; WriteLn(NewFichier,St) ; end ;
PieceFictive:=0 ; TotPieceD:=0 ; TotPieceC:=0 ; NumLigne:=1 ; NoPiece:=0 ; oldNoPiece:=0 ;
ListePiece:=HTStringList.Create ; ListePiece.Sorted:=FALSE ; NatModifie:=FALSE ;
ListeCpt:=HTStringList.Create ; ListeCpt.Sorted:=FALSE ; PremFois:=TRUE ;
While Not EOF(Fichier) do
  BEGIN
  Inc(NumLigne) ;
  ReadLn(Fichier,St) ; If Trim(St)='' Then Continue ;
  PasDeModif:=EstUneLigneCpt(St) Or (St[DebFormat.DebTC] in ['L']) ;
//  If NumLigne=2 Then Longueur:=Length(St) ;
  Longueur:=Length(St) ;
  St:=Format1_String(St,Longueur) ;
  If Trim(St)<>'' Then
     BEGIN
//     OkPourWriteAutre:=FALSE ;
//     If EstUneLigneCpt(St) Or (St[DebFormat.DebTC] in ['L']) Then WriteLn(NewFichier,St) Else
     If PasDeModif Then ListeCpt.Add(St) Else
       BEGIN
       OkOk:=AffecteNumP(St,EnRupture,PieceFictive,NoPiece,TotPieceD,TotPieceC,DebFormat,Force,Jal,Periode,InfoImp,QFiche,PasPiece) ;
       If Not OkOk Then BEGIN StErr:=StErr+IntToStr(NumLigne)+';' ; Result:=FALSE ; END ;
       If AffecteNatP(St,DebFormat,QFiche,InfoImp) Then NatModifie:=TRUE ;
{$IFDEF RECUPPCL}
        If GetRecupPCL Then
          BEGIN
          If PremFois Then BEGIN OldNoPiece:=NoPiece ; OldJal:=Jal ; OldPeriode:=Periode ; END ;
          EnRupture:=(OldNoPiece<>NoPiece) Or (Jal<>OldJal) Or (Periode<>OldPeriode) ;
          If EnRupture Then BEGIN OldNoPiece:=NoPiece ; OldJal:=Jal ; OldPeriode:=Periode ; END ;
          END ;
{$ELSE}
        If PasPiece Then
          BEGIN
          If PremFois Then BEGIN OldNoPiece:=NoPiece ; OldJal:=Jal ; OldPeriode:=Periode ; END ;
          EnRupture:=(OldNoPiece<>NoPiece) Or (Jal<>OldJal) Or (Periode<>OldPeriode) ;
          If EnRupture Then BEGIN OldNoPiece:=NoPiece ; OldJal:=Jal ; OldPeriode:=Periode ; END ;
          END ;
{$ENDIF}
       PremFois:=FALSE ;
       If EnRupture Then
         BEGIN
         If NatModifie Then RetouchePieceAvantWrite(ListePiece,DebFormat,InfoImp) ;
         EcritPiece(NewFichier,ListeCpt) ;
         EcritPiece(NewFichier,ListePiece) ;
         NatModifie:=FALSE ;
         END ;
       ListePiece.Add(St) ;
       END ;
//     If OkPourWriteAutre Then WriteLn(NewFichier,St)
     END ;
  END ;
If NatModifie Then RetouchePieceAvantWrite(ListePiece,DebFormat,InfoImp) ;
EcritPiece(NewFichier,ListeCpt) ;
EcritPiece(NewFichier,ListePiece) ;
VideStringList(ListePiece) ; ListePiece.Free ;
VideStringList(ListeCpt) ; ListeCpt.Free ;
CloseFile(Fichier) ; CloseFile(NewFichier) ;
FichierOnDisk(InfoImp.NomFic,FALSE) ;
//Delay(3000) ;
{$i-} AssignFile(Fichier,InfoImp.NomFic) ; Erase(Fichier) ; {$i+}
{iio:=}IoResult ; //Delay(3000) ;
RenameFile(NewFileName,InfoImp.NomFic) ;
FichierOnDisk(InfoImp.NomFic,FALSE) ;
END ;

Procedure RecupMvtALettrer(Q : TQuery ; ChampR : String ; Var Mvt : TFMvtImport) ;
BEGIN
InitMvtImport(Mvt) ;
If Not Q.Eof Then With Mvt Do
  BEGIN
  IE_GENERAL:=Q.FindField('E_GENERAL').AsString ;
  IE_AUXILIAIRE:=Q.FindField('E_AUXILIAIRE').AsString ;
  IE_DATECOMPTABLE:=Q.FindField('E_DATECOMPTABLE').ASDateTime ;
  IE_DEVISE:=Q.FindField('E_DEVISE').AsString ;
  IE_LETTRAGE:=Q.FindField(ChampR).AsString ;
  IE_JOURNAL:=Q.FindField('E_JOURNAL').AsString ;
  IE_DEBIT:=Q.FindField('E_DEBIT').AsFloat ;
  IE_CREDIT:=Q.FindField('E_CREDIT').ASFloat ;
  IE_REFINTERNE:=Q.FindField('E_REFINTERNE').AsString ;
  IE_LIBELLE:=Q.FindField('E_LIBELLE').AsString ;
  IE_DEBITDEV:=Q.FindField('E_DEBITDEV').AsFloat ;
  IE_CREDITDEV:=Q.FindField('E_CREDITDEV').ASFloat ;
  IE_NATUREPIECE:=Q.FindField('E_NATUREPIECE').AsString ;
  END
END ;


Procedure PrepareLettrage(OkInc : Boolean ; Var InfoImp : TInfoImport ; TLett : TList ; Var NbLett : Integer ; Var Mvt : TFMvtImport) ;
Var DecDev : Integer ;
    Quotite : Double ;
    L : TL_Rappro ;
BEGIN
If OkInc Then Inc(NbLett) ;
RecupDevise(Mvt.IE_DEVISE,DecDev,Quotite,InfoImp.ListeDev) ;
L:=TL_Rappro.Create ;
L.General:=Mvt.IE_GENERAL ; L.Auxiliaire:=Mvt.IE_AUXILIAIRE ;
L.DateC:=Mvt.IE_DATECOMPTABLE ; L.DateE:=Mvt.IE_DATEECHEANCE ; L.DateR:=Mvt.IE_DATEREFEXTERNE ;
L.RefI:=Mvt.IE_REFINTERNE ; L.RefL:=Mvt.IE_REFLIBRE ;
L.RefE:=Mvt.IE_REFEXTERNE; L.Lib:=Mvt.IE_LIBELLE ;
L.Jal:=Mvt.IE_JOURNAL ; L.Numero:=Mvt.IE_NUMPIECE ;
L.NumLigne:=Mvt.IE_NUMLIGNE ; L.NumEche:=Mvt.IE_NUMECHE ;
L.CodeL:=Mvt.IE_LETTRAGE ;
L.TauxDEV:=Mvt.IE_TAUXDEV ;
L.CodeD:=Mvt.IE_DEVISE ;
L.Decim:=DecDev ;
L.Debit:=Mvt.IE_DEBIT ; L.Credit:=Mvt.IE_CREDIT ;
L.DebDev:=Mvt.IE_DEBITDEV ; L.CredDev:=Mvt.IE_CREDITDEV ;
L.Nature:=Mvt.IE_NATUREPIECE ;
L.Facture:=((L.Nature='FC') or (L.Nature='FF') or (L.Nature='AC') or (L.Nature='AF')) ;
L.Client:=((L.Nature='FC') or (L.Nature='AC') or (L.Nature='RC') or (L.Nature='OC')) ;
L.Solution:=0 ; L.Exo:=QUELEXODT(Mvt.IE_DATECOMPTABLE) ;
L.EditeEtatTva:=FALSE ;
TLett.Add(L) ;
END ;

procedure ChargeDev(Var InfoImp : TInfoImport) ;
var Q       : TQuery ;
    TDevise : TFDevise ;
BEGIN
Q := OpenSQL('SELECT D_DEVISE,D_QUOTITE,D_DECIMALE,H_TAUXREEL FROM DEVISE LEFt join CHANCELL on D_DEVISE=H_DEVISE '+
 ' AND H_DATECOURS=(SELECT MAX(H_DATECOURS) FROM CHANCELL WHERE D_DEVISE=H_DEVISE) ',True) ;

(*:=OpenSQL('SELECT D_DEVISE,D_QUOTITE,D_DECIMALE FROM DEVISE ORDER BY D_DEVISE',True) ;
St:='SELECT H_TAUXREEL FROM CHANCELL WHERE H_DEVISE=:DEV'
          +' AND H_DATECOURS=(SELECT MAX(H_DATECOURS) FROM CHANCELL WHERE H_DEVISE=:DEV)' ;
Q1:=PrepareSQL(St,TRUE) ;
Q1:=TQuery.Create(Application) ; Q1.DatabaseName:=DBSOC.DataBaseName ;
Q1.SQL.Add('SELECT H_TAUXREEL FROM CHANCELL WHERE H_DEVISE=:DEV'
          +' AND H_DATECOURS=(SELECT MAX(H_DATECOURS) FROM CHANCELL WHERE H_DEVISE=:DEV)') ;
ChangeSQL(Q1) ;Q1.Prepare ;
*)
InfoImp.ListeDev:=TList.Create ;
While not Q.Eof do
  BEGIN
  TDevise         := TFDevise.Create ;
  TDevise.Code    := Q.FindField ('D_DEVISE').AsString ;
  TDevise.Quotite := Q.FindField ('D_QUOTITE').AsFloat ;
  TDevise.Decimale:= Q.FindField ('D_DECIMALE').AsInteger ;
  TDevise.TauxDev:=Q.FindField('H_TAUXREEL').AsFloat ;
  InfoImp.ListeDev.Add(TDevise) ;
  Q.Next ;
  END ;
Ferme(Q) ;
END ;

procedure LettrageSurRegroupement(Var InfoImp : TInfoImport ; ChampR : String) ;
var SQL,OldLettrage : String ;
    QImp : TQuery ;
    TLett:TList ;
//    OkALettrer : Boolean ;
    {NbLettImp,}NbLett : integer ;
    Auxi,Gene,CodeLettre,Devise : String17 ;
    MvtImport : FMvtImport ;
BEGIN
SQL:='SELECT E_GENERAL,E_AUXILIAIRE,E_DATECOMPTABLE,E_DEVISE,'+ChampR+',E_LETTRAGE,E_JOURNAL,E_DEBIT,E_CREDIT,E_REFINTERNE,'
     +'E_LIBELLE,E_DEBITDEV,E_CREDITDEV,E_NATUREPIECE FROM ECRITURE WHERE E_REFEXTERNE<>"L" '
     +' ORDER BY E_AUXILIAIRE,E_GENERAL,E_DEVISE,'+ChampR+' , E_DATECOMPTABLE, E_NUMEROPIECE, IE_NUMLIGNE, IE_NUMECHE ' ;
QImp:=OpenSQL(SQL,True) ;
if (QImp.EOF) then BEGIN Ferme(QImp) ; Exit ; END ;
ChargeDev(InfoImp) ;
InitMove(RecordsCount(QImp),'') ;
TLett:=TList.Create ;
New(MvtImport) ;
While not QImp.Eof do
  BEGIN
  VideListe(TLett) ; OldLettrage:='' ; NbLett:=0 ; //NbLettImp:=0 ;
  RecupMvtALettrer(QImp,ChampR,MvtImport^) ;
  Auxi:=MvtImport^.IE_AUXILIAIRE ;
  Gene:=MvtImport^.IE_GENERAL ;
  CodeLettre:=MvtImport^.IE_LETTRAGE ;
  Devise:=MvtImport^.IE_DEVISE ;
//  OkALettrer:=True ;
  While not (QImp.Eof) and ((QImp.Fields[1].AsString=Auxi)
    and (QImp.Fields[0].AsString=Gene)
    and (QImp.Fields[4].AsString=CodeLettre)
    and (QImp.Fields[3].AsString=Devise)) do
    BEGIN
//    Inc(NbLettImp) ;
    RecupMvtALettrer(QImp,ChampR,MvtImport^) ;
    PrepareLettrage(TRUE,InfoImp,TLett,NbLett,MvtImport^) ;
    QImp.Next ; MoveCur(False) ;
    END ;
  LettrerUnPaquet(TLett,False,False) ;
  END ;
FiniMove ;
VideListe(TLett) ; TLett.Free ;
Ferme(QImp) ;
Dispose(MvtImport) ;
VideListe(InfoImp.ListeDev) ; InfoImp.ListeDev.Free ; InfoImp.ListeDev:=NIL ;
END ;

procedure ChargeDossierAuto ( Var FDossier : HTStrings ; VireRef : Boolean) ;
BEGIN
If FDossier=NIL Then
  FDossier:=HTStringList.Create ;
ChargeDossier(FDossier,VireRef) ;
END ;

procedure FreeDossierAuto ( Var FDossier : HTStrings ) ;
BEGIN
If FDossier<>NIL Then BEGIN FDossier.Clear ; FDossier.Free ; END ;
FDossier:=NIL ;
END ;

Function ChargeUnDossier(StNom : String ; NomChamp : String) : String ;
Var Q : TQuery ;
BEGIN
Result:='' ;
if DBSOC<>NIL then DeconnecteHalley ;
if ConnecteDB(StNom,DBSOC,'SOC',FALSE) then
  BEGIN
{$IFDEF SYSTEMU}
 {$IFDEF SPEC302}
  Q:=OpenSQL('SELECT SO_RVA,SO_CONTACT FROM SOCIETE',TRUE) ;
  If Not Q.Eof Then Result:=Q.Fields[0].AsString+Q.Fields[1].AsString ;
  Ferme(Q) ;
 {$ELSE}
  Q:=OpenSQL('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_RVA" ',TRUE) ;
  If Not Q.Eof Then Result:=Q.Fields[0].AsString ;
  Ferme(Q) ;
  Q:=OpenSQL('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_CONTACT" ',TRUE) ;
  If Not Q.Eof Then Result:=Result+Q.Fields[0].AsString ;
  Ferme(Q) ;
 {$ENDIF}
  DeconnecteHalley ;
{$ELSE}
 {$IFDEF SPEC302}
  Q:=OpenSQL('SELECT '+NomChamp+' FROM SOCIETE',TRUE) ;
  If Not Q.Eof Then Result:=Q.Fields[0].AsString ;
  Ferme(Q) ;
 {$ELSE}
  Q:=OpenSQL('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="'+NomChamp+'" ',TRUE) ;
  If Not Q.Eof Then Result:=Q.Fields[0].AsString ;
  Ferme(Q) ;
 {$ENDIF}
  DeconnecteHalley ;
{$ENDIF}

  END ;

END ;

Function TrouveNomDossier(CodeSoc,NomChamp : String) : String ;
Var FDossier : HTStrings ;
    CodeSocLu : String ;
    i : Integer ;
BEGIN
Result:='' ;  If CodeSoc='' Then Exit ; If NomChamp='' Then Exit ;
FDossier:=NIL ;
ChargeDossierAuto(FDossier,TRUE) ;
Initmove(FDossier.Count-1,'') ;
For i:=0 To FDossier.Count-1 Do
  BEGIN
  MoveCur(FALSE) ;
  CodeSocLu:=ChargeUnDossier(FDossier[i],NomChamp) ;
  If (CodeSocLu=Trim(CodeSoc)) Then BEGIN Result:=FDossier[i] ; Break ; END ;
  END ;
FiniMove ;
FreeDossierAuto(FDossier) ;
END ;

Procedure ChargeLCodeSoc(NomChamp : String ; LCodeSoc : HtStringList) ;
Var FDossier : HTStrings ;
    CodeSocLu : String ;
    i : Integer ;
BEGIN
If LCodeSoc=Nil Then Exit ; LCodeSoc.Clear ; If NomChamp='' Then Exit ;
FDossier:=NIL ;
ChargeDossierAuto(FDossier,TRUE) ;
Initmove(FDossier.Count-1,'') ;
For i:=0 To FDossier.Count-1 Do
  BEGIN
  MoveCur(FALSE) ;
  CodeSocLu:=ChargeUnDossier(FDossier[i],NomChamp) ;
  If CodeSocLu<>'' Then LCodeSoc.Add(FDossier[i]+';'+CodeSocLu+';') ;
  END ;
FreeDossierAuto(FDossier) ;
FiniMove ;
END ;

Function LitLCodeSoc (LCodeSoc : HtStringList ; CodeSoc : String) : String ;
Var St,StCode,StNom : String ;
    i : Integer ;
BEGIN
Result:='' ;
For i:=0 To LCodeSoc.Count-1 Do If LCodeSoc[i]<>'' Then
  BEGIN
  StCode:='' ; StNom:='' ;
  St:=LCodeSoc[i] ; StNom:=ReadTokenSt(St) ; If St<>'' Then StCode:=ReadTokenSt(St) ;
  If StCode=CodeSoc Then BEGIN Result:=StNom ; Break ; END ;
  END ;
END ;


Function ModeJal(Jal : String ; Var InfoImp : TInfoImport) : tModeJal ;
Var CptLuJ : tCptLu ;
    JalOk : Boolean ;
    Q : TQuery ;
    Mode : String ;
BEGIN
Result:=mPiece ; JalOk:=FALSE ; If GetRecupPCL Then Exit ; If Not ImportBor Then Exit ;
If InfoImp.LJalLu<>Nil Then
  BEGIN
  Fillchar(CptLuJ,SizeOf(CptLuJ),#0) ; CptLuJ.Cpt:=Jal ;
  If ChercheCptLu(InfoImp.LJalLu,CptLuJ) Then
    BEGIN
    If CptLuJ.ModeSaisie='BOR' Then Result:=mBor Else
     If CptLuJ.ModeSaisie='LIB' Then Result:=mLib ;
    JalOk:=TRUE ;
    END ;
  END ;
If Not JAlOk Then
  BEGIN
  Q:=OpenSQL('SELECT J_MODESAISIE FROM JOURNAL WHERE J_JOURNAL="'+Jal+'" ',TRUE) ;
  If Not Q.Eof Then
    BEGIN
    Mode:=Q.Fields[0].AsString ;
    If Mode='BOR' Then Result:=mBor Else
     If Mode='LIB' Then Result:=mLib ;
    END ;
  Ferme(Q) ;
  END ;
If Result<>MPiece Then InfoImp.AuMoinsUnBordereau:=TRUE Else InfoImp.AuMoinsUnePiece:=TRUE ;
END ;

Procedure AlimSoucheBor(LS : HtStringList) ;
Var Q : TQuery ;
    St{,StNum }: String ;
    Num : Integer ;
BEGIN
If LS=Nil Then Exit ; If Not ImportBor Then Exit ;
LS.Clear ;
St:='SELECT E_EXERCICE,E_JOURNAL,E_PERIODE,Max(E_NUMEROPIECE) AS MAXP FROM ECRITURE '+
    'WHERE (E_EXERCICE="'+VH^.EnCours.Code+'" OR E_EXERCICE="'+VH^.Suivant.Code+'") AND E_QUALIFPIECE="N" '+
    ' GROUP BY E_EXERCICE, E_JOURNAL, E_PERIODE ORDER BY E_EXERCICE, E_JOURNAL, E_PERIODE ' ;
Q:=OpenSQL(St,TRUE) ;
InitMove(RecordsCount(Q),'') ;
While Not Q.Eof Do
  BEGIN
  MoveCur(FALSE) ;
  Num:=Q.FindField('MAXP').AsInteger ;
  If Num<999000 Then Num:=999000 ;
  St:=Q.FindField('E_EXERCICE').AsString+';'+Q.FindField('E_JOURNAL').AsString+';'+
      Q.FindField('E_PERIODE').AsString+';'+IntToStr(Num)+';' ;
  LS.Add(St) ;
  Q.Next ;
  END ;
Ferme(Q) ;
FiniMove ;
END ;

Function TrouveSNum(LS : HtStringList ; Jal,Exo,Per : String ; OkInc : Boolean) : String ;
Var i : Integer ;
    St,StExo,StJal,StPer,StNum : String ;
    num : Integer ;
BEGIN
Result:='' ;
For i:=0 To LS.Count-1 Do
  BEGIN
  St:=LS[i] ;
  If St<>'' Then StExo:=ReadTokenSt(St) ;
  If St<>'' Then StJal:=ReadTokenSt(St) ;
  If St<>'' Then StPer:=ReadTokenSt(St) ;
  If St<>'' Then StNum:=ReadTokenSt(St) ;
  If (StExo=Exo) And (StJal=Jal) And (StPer=per) Then
    BEGIN
    If OkInc Then
      BEGIN
      Num:=StrToInt(StNum) ; Inc(Num) ; StNum:=IntToStr(Num) ;
      St:=StExo+';'+StJal+';'+StPer+';'+StNum+';' ;
      LS[i]:=St ;
      END ;
    Result:=StNum ; Exit ;
    END ;
  END ;
If Result='' Then
  BEGIN
  Num:=999001 ; Result:=IntToStr(Num) ;
  St:=Exo+';'+Jal+';'+Per+';'+IntToStr(Num)+';' ;
  LS.Add(St) ;
  END ;
END ;

Function SetIncNumSB(LS : HtStringList ; Jal : String ; DateP : TDateTime) : Integer;
Var Exo,Per : String ;
    Y,M,D : Word ;
    SNum : String ;
BEGIN
Result:=0 ;
If LS=Nil Then Exit ;
If (DateP>=VH^.EnCours.Deb) And (DateP<=VH^.EnCours.Fin) Then Exo:=VH^.EnCours.Code Else
 If (DateP>=VH^.Suivant.Deb) And (DateP<=VH^.Suivant.Fin) Then Exo:=VH^.Suivant.Code Else Exit ;
DecodeDate(DateP,Y,M,D) ;
per:=IntToStr(Y)+FormatFloat('00',M) ;
SNum:=TrouveSNum(LS,Jal,Exo,Per,TRUE) ;
If SNum<>'' Then Result:=StrToInt(SNum) ;
END ;


Procedure MajJalODA(Var StJal : String) ;
Var TobE,TobL : Tob ;
    i : Integer ;
    Q : TQuery ;
    LeJal : String ;
BEGIN
StJal:='' ;
  {JP 08/07/03 : Lors de l'importation de plusieurs fichiers, la requête ci-dessous ne
  renvoie rien ou ne renvoie que les éventuels nouveaux journaux sur cet exercice. On
  va donc récupérer les journaux 'ODA' des exercice pécédents}
  Q := OpenSQL('SELECT DISTINCT(J_JOURNAL) FROM JOURNAL WHERE J_NATUREJAL = "ODA"', True);
  while not Q.EOF do begin
    StJal := StJal + Q.Fields[0].AsString + ';';
    Q.Next;
  end;
  Ferme(Q);

TobE:=TOB.Create('',Nil,-1) ;
Q:=OpenSQL('SELECT * FROM JOURNAL WHERE J_COMPTEAUTOMAT="***" ',TRUE) ;
TOBE.LoadDetailDB('JOURNAL','','',Q,False,True) ;
Ferme(Q) ;
For i:=0 To TobE.Detail.Count-1 Do
  BEGIN
  TobL:=TobE.Detail[i] ;
  TobL.PutValue('J_COMPTEAUTOMAT','') ;
  TOBL.PutValue('J_AXE','A1') ;
  TOBL.PutValue('J_MODESAISIE','-') ;
  TOBL.PutValue('J_NATUREJAL','ODA') ;
  LeJal:=TobL.GetValue('J_JOURNAL') ;
  StJal:=StJal+LeJal+';' ;
  ExecuteSQL('UPDATE SOUCHE SET SH_ANALYTIQUE="X" WHERE SH_TYPE="CPT" AND SH_SOUCHE="'+LeJal+'" ') ;
  END ;
TobE.UpdateDB ;
TobE.Free ;
END ;

Procedure DeleteEcr(StJal : String) ;
Var St1,St2 : String ;
BEGIN
St1:=StJal ;
While St1<>'' Do
  BEGIN
  St2:=ReadTokenSt(St1) ;
  If St2<>'' Then ExecuteSQL('DELETE FROM ECRITURE WHERE E_JOURNAL="'+St2+'" ') ;
  END ;
END ;

Procedure Ana2ODA(StJal : String) ;
Var St1,St2 : String ;
    TobA,TobL : TOB ;
    i : Integer ;
    Gen,OldGen : String ;
    NumP,NumL : Integer ;
//    Rupt : Boolean ;
    Q : TQuery ;
BEGIN
St1:=StJal ; NumL:=0 ;
While St1<>'' Do
  BEGIN
  St2:=ReadTokenSt(St1) ;
  If St2<>'' Then
    BEGIN
    TobA:=TOB.Create('',Nil,-1) ;
    Q:=OpenSQL('SELECT * FROM ANALYTIQ WHERE Y_JOURNAL="'+St2+'" ',TRUE) ;
    TOBA.LoadDetailDB('ANALYTIQ','','',Q,False,True) ;
    Ferme(Q) ; NumP:=1 ; //Rupt:=FALSE ;
    For i:=0 To TobA.Detail.Count-1 Do
      BEGIN
      TOBL:=TobA.Detail[i] ;
      If i=0 Then OldGen:=TobL.GetValue('Y_GENERAL') ;
      Gen:=TobL.GetValue('Y_GENERAL') ;
      If OldGen<>Gen Then BEGIN Inc(NumP) ; NumL:=0 ; END ;
      Inc(NumL) ;
      TobL.PutValue('Y_TYPEANALYTIQUE','X') ;
      TobL.PutValue('Y_NUMEROPIECE',NumP) ;
      TobL.PutValue('Y_NUMLIGNE',0) ;
      TobL.PutValue('Y_POURCENTAGE',0) ;
      TobL.PutValue('Y_TOTALECRITURE',0) ;
      TobL.PutValue('Y_NUMVENTIL',NumL) ;
      TobL.PutValue('Y_TOTALDEVISE',0) ;
      TobL.PutValue('Y_TOTALQTE1',0) ;
      TobL.PutValue('Y_TOTALQTE2',0) ;
      TobL.PutValue('Y_TYPEMVT','AE') ;
      TobL.PutValue('Y_CONTREPARTIEGEN','') ;
      TobL.PutValue('Y_CONTREPARTIEAUX','') ;
      OldGen:=Gen ;
      END ;
    ExecuteSQL('DELETE FROM ANALYTIQ WHERE Y_JOURNAL="'+St2+'" ') ;
    TobA.SetAllModifie(TRUE) ;
    TobA.InsertDB(Nil,TRUE) ;
    TobA.Free ;
    ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART="'+IntToStr(NumP)+'" WHERE SH_TYPE="CPT" AND SH_SOUCHE="'+St2+'" ') ;
    END ;
  END ;
END ;

Procedure TraiteODA(Var InfoImp : TInfoImport)  ;
Var StJal : String ;
BEGIN
If Not InfoImp.AuMoinsUneODA Then Exit ;
MajJalODA(StJal) ;
DeleteEcr(StJal) ;
Ana2ODA(StJal) ;
END ;

Procedure TraiteEEXBQ(Var InfoImp : TInfoImport)  ;
Var //StJal : String ;
    St1,St2,Cpt,StPer : String ;
    Tob1,Tob2,TobL : Tob ;
    per : TDateTime ;
    OkInsert : Boolean ;
    Q : TQuery ;
BEGIN
If InfoImp.StEEXBQ='' Then Exit ;
OkInsert:=FALSE ;
Tob1:=TOB.Create('',Nil,-1) ;
Tob2:=TOB.Create('',Nil,-1) ;
Q:=OpenSQL('SELECT * FROM EEXBQ ',TRUE) ;
TOB1.LoadDetailDB('EEXBQ','','',Q,False,True) ;
Ferme(Q) ;
St1:=InfoImp.StEEXBQ ;
While St1<>'' Do
  BEGIN
  St2:=ReadTokenSt(St1) ; St2:=St2+',' ;
  Cpt:='' ; StPer:='' ;
  If St2<>'' Then Cpt:=ReadTokenPipe(St2,',') ;
  If St2<>'' Then StPer:=ReadTokenPipe(St2,',') ;
  If (Cpt<>'') And (StPer<>'') Then
    BEGIN
    Per:=StrToDate(Stper) ;
    TobL:=TOB1.FindFirst(['EE_GENERAL','EE_DATEPOINTAGE','EE_REFPOINTAGE','EE_NUMERO'],[Cpt,per,'POINTE',1],False) ;
    If TobL<>NIL Then Continue Else
      BEGIN
      TobL:=TOB2.FindFirst(['EE_GENERAL','EE_DATEPOINTAGE','EE_REFPOINTAGE','EE_NUMERO'],[Cpt,per,'POINTE',1],False) ;
      If TobL=NIL Then
        BEGIN
        TobL := TOB.Create('EEXBQ',TOB2,-1); TOBL.initValeurs ;
        TOBL.PutValue('EE_GENERAL',Cpt) ;
        TOBL.PutValue('EE_REFPOINTAGE','POINTE') ;
        TOBL.PutValue('EE_DATEPOINTAGE',Per) ;
        TOBL.PutValue('EE_NUMERO',1) ;
        OkInsert:=TRUE ;
        END ;
      END ;
    END ;
  END ;
If OkInsert Then TOB2.InsertDB(NIL,TRUE) ;
Tob2.Free ; Tob1.Free ;
END ;

Function CreateMPSISCO(CodeMP,LibMP,CatMP,AccMP : String) : Boolean ;
Var Q : tQuery ;
BEGIN
Result:=FALSE ;
Q:=OPENSQL('SELECT * FROM MODEPAIE WHERE MP_MODEPAIE="'+CodeMP+'" ',FALSE) ;
If Q.Eof Then
  BEGIN
  Result:=TRUE ;
  Q.Insert ; InitNEw(Q) ;
  Q.FindField('MP_MODEPAIE').AsString:=CodeMp ;
  Q.FindField('MP_LIBELLE').AsString:=Copy(LibMP,1,35) ;
  Q.FindField('MP_ABREGE').AsString:=Copy(LibMP,1,17) ;
  Q.FindField('MP_ENCAISSEMENT').AsString:='MIX' ;
  Q.FindField('MP_CATEGORIE').AsString:=CatMP ;
  Q.FindField('MP_CODEACCEPT').AsString:=AccMP ;
  Q.Post ;
  END ;
Ferme(Q) ;
END ;


{$ENDIF EAGLCLIENT}
{$ENDIF EAGLSERVER}


{$IFNDEF EAGLCLIENT}

Function EstDansFourchette(Cpt : String ; Var FCT : TFCha) : Boolean ;
Var i : Integer ;
BEGIN
Result:=FALSE ;
For i:=1 To 5 Do
  If (FCT[i].Deb<>'') And (FCT[i].Fin<>'') And (FCT[i].Deb<=Cpt) And (FCT[i].Fin>=Cpt) Then BEGIN Result:=TRUE ; Exit ; END ;
END ;

Function NatureCptImport(Cpt : String ; Var FI : TFourchetteImport) : String ;
BEGIN
Result:='---' ; If Trim(Cpt)='' Then Exit ;
If GetRecupSISCOPGI Then
  BEGIN
  END ;
If EstDansFourchette(Cpt,FI.Cha) Then Result:='CHA' Else
If EstDansFourchette(Cpt,FI.Pro) Then Result:='PRO' Else
If EstDansFourchette(Cpt,FI.DebD) Then Result:='TID' Else
If EstDansFourchette(Cpt,FI.DebC) Then Result:='TIC' Else
If EstDansFourchette(Cpt,FI.Imm) Then Result:='IMO' Else
If EstDansFourchette(Cpt,FI.Ban) Then Result:='BQE' Else
If EstDansFourchette(Cpt,FI.Cai) Then Result:='CAI' Else
If EstDansFourchette(Cpt,FI.Cli) Then Result:='COC' Else
If EstDansFourchette(Cpt,FI.Fou) Then Result:='COF' Else
If EstDansFourchette(Cpt,FI.Sal) Then Result:='COS' Else
If EstDansFourchette(Cpt,FI.Dvs) Then Result:='COD' ;
END ;

Function LitTRFParam(LeCode,LeNom : String) : String ;
Var Q : Tquery ;
BEGIN
Result:='' ;
Q:=OpenSQL('SELECT * FROM TRFPARAM WHERE TRP_CODE="'+LeCode+'" AND TRP_NOM="'+LeNom+'" ',TRUE) ;
If Not Q.Eof Then
  BEGIN
  Result:=Q.FindField('TRP_DATA').AsString ;
  END ;
Ferme(Q) ;
END ;

Procedure UpdateTRFParam(St,LeCode,LeNom : String) ;
Var Q : Tquery ;
BEGIN
Q:=OpenSQL('SELECT * FROM TRFPARAM WHERE TRP_CODE="'+LeCode+'" AND TRP_NOM="'+LeNom+'" ',FALSE) ;
If Not Q.Eof Then
  BEGIN
  Q.Edit ;
  Q.FindField('TRP_DATA').AsString:=St ;
  Q.Post ;
  END ;
Ferme(Q) ;
END ;

Procedure ChargeFourchetteCompte(Qui : String ; Var FI : TFourchetteImport) ;
BEGIN
DechiffreFour(FI.Cha,Qui,'CHA',5)  ; DechiffreFour(FI.Pro,Qui,'PRO',5) ;
DechiffreFour(FI.DebD,Qui,'LD%',5) ; DechiffreFour(FI.DebC,Qui,'LC%',5) ;
DechiffreFour(FI.Imm,Qui,'IM%',5)  ; DechiffreFour(FI.Ban,Qui,'BA%',3) ;
DechiffreFour(FI.Cai,Qui,'CA%',1)  ; DechiffreFour(FI.Cli,Qui,'CL%',5) ;
DechiffreFour(FI.Fou,Qui,'FO%',5)  ; DechiffreFour(FI.SAL,Qui,'SA%',5) ;
DechiffreFour(FI.Dvs,Qui,'DI%',5) ;
END ;

Procedure DechiffreFour(Var Four : TFCha ; Qui,St : String ; MaxE : Integer) ;
Var i : Integer ;
    Q : TQuery ;
BEGIN
Fillchar(Four,SizeOf(Four),#0) ;
{$IFDEF NOCONNECT}
Exit ;
{$ENDIF}
{$IFDEF RECUPPCL}
{$ELSE}
If Trim(St)='' Then Exit ; i:=0 ;
If (St='CHA') Or (St='PRO') Then
  BEGIN
{$IFDEF SPEC302}
  Q:=OpenSQL('SELECT * FROM SOCIETE',TRUE) ;
  If Not Q.Eof Then
    BEGIN
    For i:=1 To MaxE Do
      BEGIN
      If St='CHA' Then
        BEGIN
        Four[i].Deb:=Q.FindField('SO_CHADEB'+IntToStr(i)).AsString ; Four[i].Fin:=Q.FindField('SO_CHAFIN'+IntToStr(i)).AsString ;
        END Else
        BEGIN
        Four[i].Deb:=Q.FindField('SO_PRODEB'+IntToStr(i)).AsString ; Four[i].Fin:=Q.FindField('SO_PROFIN'+IntToStr(i)).AsString ;
        END ;
      END ;
    END ;
  Ferme(Q) ;
{$ELSE}
  For i:=1 To MaxE Do
      BEGIN
      If St='CHA' Then
         BEGIN
         Four[i].Deb:=GetParamSoc('SO_CHADEB'+IntToStr(i)) ; Four[i].Fin:=GetParamSoc('SO_CHAFIN'+IntToStr(i)) ;
         END Else
         BEGIN
         Four[i].Deb:=GetParamSoc('SO_PRODEB'+IntToStr(i)) ; Four[i].Fin:=GetParamSoc('SO_PROFIN'+IntToStr(i)) ;
         END ;
      END ;
{$ENDIF}
  Exit ;
  END ;
Q:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+Qui+'" AND CC_CODE LIKE "'+St+'" ORDER BY CC_CODE',TRUE) ;
While Not Q.Eof Do
  BEGIN
  Inc(i) ; If i>MaxE Then Break ;
  Four[i].Deb:=Q.FindField('CC_LIBELLE').AsString ;
  Four[i].Fin:=Q.FindField('CC_ABREGE').AsString ;
  Q.Next ;
  END ;
Ferme(Q) ;
{$ENDIF}
END ;

Procedure InitRequete(Var Q : TQuery ; What : Integer) ;
Var St : String ;
BEGIN
Q:=Nil ;
Case What Of
  0 : St:='SELECT G_NATUREGENE,G_REGIMETVA,G_DATEDERNMVT,G_NUMDERNMVT, '+
          'G_LIGNEDERNMVT,G_DEBITDERNMVT,G_CREDITDERNMVT,G_TOTALDEBIT,G_TOTALCREDIT, '+
          'G_TOTDEBE,G_TOTCREE,G_TOTDEBS,G_TOTCRES,G_TOTDEBANO,G_TOTCREANO,G_VENTILABLE, '+
          'G_VENTILABLE1,G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5, G_COLLECTIF, '+
          'G_LETTRABLE,G_POINTABLE,G_TVA,G_TPF,G_TVAENCAISSEMENT,G_TVASURENCAISS FROM GENERAUX '+
          'WHERE G_GENERAL=:CPTE' ;
  1 : St:='SELECT T_NATUREAUXI,T_REGIMETVA,T_DATEDERNMVT,T_NUMDERNMVT,T_COLLECTIF, '+
          'T_LIGNEDERNMVT,T_DEBITDERNMVT,T_CREDITDERNMVT,T_TOTALDEBIT,T_TOTALCREDIT, '+
          'T_TOTDEBE,T_TOTCREE,T_TOTDEBS,T_TOTCRES,T_TOTDEBANO,T_TOTCREANO,T_LETTRABLE, '+
          'T_TVAENCAISSEMENT,T_PAYEUR,T_ISPAYEUR, T_AUXILIAIRE FROM TIERS '+
          'WHERE T_AUXILIAIRE=:CPTE' ;
  2 : St:='SELECT S_SECTION,S_AXE,S_DATEDERNMVT, '+
          'S_NUMDERNMVT,S_LIGNEDERNMVT,S_DEBITDERNMVT,S_CREDITDERNMVT,S_TOTALDEBIT,S_TOTALCREDIT, '+
          'S_TOTDEBE,S_TOTCREE,S_TOTDEBS,S_TOTCRES,S_TOTDEBANO,S_TOTCREANO FROM SECTION '+
          'WHERE S_SECTION=:CPTE AND S_AXE=:AXE' ;
  3 : St:='SELECT J_JOURNAL,J_NATUREJAL,J_COMPTEURNORMAL,J_COMPTEURSIMUL,J_DATEDERNMVT,'+
          'J_NUMDERNMVT,J_DEBITDERNMVT,J_CREDITDERNMVT,J_TOTALDEBIT,J_TOTALCREDIT,'+
          'J_TOTDEBE,J_TOTCREE,J_TOTDEBS,J_TOTCRES,J_CONTREPARTIE,J_MODESAISIE FROM JOURNAL '+
          'WHERE J_JOURNAL=:CPTE' ;
  4 : St:='SELECT BJ_BUDJAL,BJ_NATJAL,BJ_COMPTEURNORMAL,BJ_COMPTEURSIMUL '+
          'FROM BUDJAL WHERE BJ_BUDJAL=:CPTE' ;
  5 : St:='SELECT CR_TYPE,CR_CORRESP,CR_LIBELLE,CR_LIBRETEXTE1 from CORRESP WHERE CR_TYPE=:CRTYPE AND CR_CORRESP=:CRCORRESP' ;
  END ;
Q:=PrepareSQL(St,TRUE) ;
END ;

Function HalteLa : Boolean ;
BEGIN
Result:=FALSE ;
If Not (GetRecupSISCOPGI or GetRecupPCL) Then Exit ;
{$IFNDEF NOVH}
If VH^.StopRSP Then Result:=TRUE ;
{$ENDIF}
END ;

Procedure UpdatePersoTL ;
Var St : String ;
BEGIN
St:='UPDATE CHOIXCOD SET CC_ABREGE="X" WHERE CC_TYPE="NAT" AND CC_CODE="E00" ' ;
ExecuteSQL(St) ;
St:='UPDATE CHOIXCOD SET CC_LIBRE="4" WHERE CC_TYPE="NAT" and CC_CODE="E00"' ;
ExecuteSQL(St) ;
{$IFNDEF NOVH}
VH^.LgTableLibre[6,1]:=4 ;
{$ENDIF}
END ;

Function RecupTL(St : String ; Var TL : TCptImport) : Boolean ;
BEGIN
Fillchar(TL,SizeOf(TL),#0) ;
With TL Do
  BEGIN
  Axe:=Copy(St,6,1) ;
  Cpt:=Trim(Copy(St,7,17)) ; Libelle:=Trim(Copy(St,24,35)) ;
  Nature:=Trim(Copy(St,59,3)) ;
  T0:=Trim(Copy(St,62,35))  ; T1:=Trim(Copy(St,97,35)) ;
  T2:=Trim(Copy(St,132,35))  ; T3:=Trim(Copy(St,167,35)) ;
  T4:=Trim(Copy(St,202,35))  ; T5:=Trim(Copy(St,237,35)) ;
  T6:=Trim(Copy(St,272,35))  ; T7:=Trim(Copy(St,307,35)) ;
  T8:=Trim(Copy(St,342,35))  ; T9:=Trim(Copy(St,377,35)) ;
  EAN:=TL.Nature[1]+FormatFloat('00',StrToInt(TL.Axe)) ;
  END ;
Result:=(TL.Cpt<>'') And (TL.Axe<>'') And (TL.Axe[1] in ['0'..'9']) And
        ((TL.Nature='GEN') Or (TL.Nature='TIE') Or (TL.Nature='SEC') Or (TL.Nature='ECR') Or (TL.Nature='ANA')) ;
END ;

Function TLVide(Var TL : TCptImport) : Boolean ;
Var St : String ;
    Q : TQuery ;
    Lg : Integer ;
    CodeNat : String ;
    i : Integer ;
    {$IFNDEF NOVH}
    j : integer ;
    {$ENDIF}
BEGIN
Result:=FALSE ;
CodeNat:=TL.EAN ;
St:='Select * FROM CHOIXCOD WHERE CC_TYPE="NAT" AND CC_CODE="'+CodeNat+'" AND CC_ABREGE="X"' ;
Q:=OpenSQL(St,TRUE) ;
If Q.Eof Then
  BEGIN
  St:='UPDATE CHOIXCOD SET CC_ABREGE="X" WHERE CC_TYPE="NAT" AND CC_CODE="'+CodeNat+'" ' ;
  ExecuteSQL(St) ;
  Lg:=Length(TL.Cpt) ; If Lg<3 Then Lg:=3 ;
  St:='UPDATE CHOIXCOD SET CC_LIBRE="'+IntToStr(Lg)+'" WHERE CC_TYPE="NAT" and CC_CODE="'+CodeNat+'"' ;
  ExecuteSQL(St) ;
  Case TL.Nature[1] Of
       'G' : i:=1 ; 'T' : i:=2 ; 'S' : i:=3 ; 'B' : i:=4 ;
       'D' : i:=5 ; 'E' : i:=6 ; 'A' : i:=7 ; 'U' : i:=8 ;
       'I' : i:=9 ;
       Else i:=0 ;
    END ;
  If i>0 Then
  BEGIN
{$IFNDEF NOVH}
       j:=StrToInt(TL.Axe)+1 ;
       VH^.LgTableLibre[i,j]:=Lg ;
{$ENDIF}
  END ;
  END ;
Ferme(Q) ;
TL.Cpt:=BourreLaDoncSurLaTable(CodeNat,TL.Cpt) ;
St:='SELECT NT_NATURE FROM NATCPTE WHERE NT_TYPECPTE="'+CodeNat+'" AND NT_NATURE="'+TL.Cpt+'" ' ;
Q:=OpenSQL(St,TRUE) ;
If Not Q.Eof Then Result:=TRUE ;
Ferme(Q) ;
END ;

Function InitSQLCpt1(What : Integer ; LeCpt,S1,S2 : String ; Existe : Boolean) : String ;
Var //Q : TQuery ;
    Table,Champ,StWhereSupp : String ;
BEGIN
Result:='' ; StWhereSupp:='' ;
If Not Existe Then LeCpt:=W_W ;
Case What Of
  0 : BEGIN
      Table:='GENERAUX' ; Champ:='G_GENERAL' ;
      END ;
  1 : BEGIN
      Table:='TIERS' ; Champ:='T_AUXILIAIRE' ;
      END ;
  2 : BEGIN
      Table:='SECTION' ; Champ:='S_SECTION' ; StWhereSupp:='AND S_AXE="'+S1+'" ' ;
      END ;
  3 : BEGIN
      Table:='JOURNAL' ; Champ:='J_JOURNAL' ;
      END ;
  10 : BEGIN
       If Not Existe Then S1:=W_W ;
       Table:='NATCPTE' ; Champ:='NT_TYPECPTE' ; StWhereSupp:='AND NT_NATURE="'+S1+'" ' ;
       END ;
  11 : BEGIN
      Table:='SSSTRUCR' ; Champ:='PS_CODE' ; StWhereSupp:='AND PS_AXE="'+S1+'" AND PS_SOUSSECTION="'+S2+'" ' ;
      END ;
  12 : BEGIN
      Table:='SOUCHE' ; Champ:='SH_SOUCHE' ; StWhereSupp:='AND SH_TYPE="CPT"' ;
      END ;
  END ;
Result:='SELECT * FROM '+Table+' WHERE '+Champ+'="'+LeCpt+'" '+StWhereSupp ;
END ;


Function ImportTL(St : String(* ; Var InfoImp : TInfoImport ; QFiche : TQFiche*)) : tResultImportCpte ;
Var Q : TQuery ;
    TL : TCptImport ;
    OkCreat,Existe : Boolean ;
BEGIN
Result:=ResRien ;
//Exit ;
OkCreat:=RecupTL(St,TL) ;
Existe:=TLVide(TL) ;
If Existe Or ((Not Existe) And OkCreat) Then
  BEGIN
  (*
  Q:=InitSQLCpt(10) ;
  Q.Params[0].AsString:=TL.EAN ;
  If Existe Then Q.Params[1].AsString:=TL.Cpt Else Q.Params[1].AsString:=w_w ;
  Q.Open ;
  *)
  Q:=OpenSQL(InitSQLCpt1(10,TL.EAN,TL.Cpt,'',Existe),FALSE) ;
  Result:=ResModifier ;
  If Existe Then Q.Edit Else BEGIN Q.Insert ; InitNew(Q) ; Result:=ResCreer ; END ;
  If Not Existe Then
    BEGIN
    Q.FindField('NT_NATURE').AsString:=TL.Cpt ;
    Q.FindField('NT_TYPECPTE').AsString:=TL.EAN ;
    END ;
  Q.FindField('NT_LIBELLE').AsString:=TL.Libelle ;
  If TL.T0<>'' Then Q.FindField('NT_TEXTE0').AsString:=TL.T0 ;
  If TL.T1<>'' Then Q.FindField('NT_TEXTE1').AsString:=TL.T1 ;
  If TL.T2<>'' Then Q.FindField('NT_TEXTE2').AsString:=TL.T2 ;
  If TL.T3<>'' Then Q.FindField('NT_TEXTE3').AsString:=TL.T3 ;
  If TL.T4<>'' Then Q.FindField('NT_TEXTE4').AsString:=TL.T4 ;
  If TL.T5<>'' Then Q.FindField('NT_TEXTE5').AsString:=TL.T5 ;
  If TL.T6<>'' Then Q.FindField('NT_TEXTE6').AsString:=TL.T6 ;
  If TL.T7<>'' Then Q.FindField('NT_TEXTE7').AsString:=TL.T7 ;
  If TL.T8<>'' Then Q.FindField('NT_TEXTE8').AsString:=TL.T8 ;
  If TL.T9<>'' Then Q.FindField('NT_TEXTE9').AsString:=TL.T9 ;
  Q.Post ;
  Q.Close ; Ferme(Q) ;
  END ;
END ;

Procedure CodeStat2TL(Var InfoImp :TInfoImport;  LS : HtStringList) ;
Var CS,St : String ;
    i : Integer ;
    LC : HtStringList ;
    Q : TQuery ;
    TS : TOB;
    stLibelle : string;
BEGIN
If LS=Nil Then Exit ; If LS.Count<=0 Then Exit ;
UpdatepersoTL ;
InitMove(Ls.count,'') ;
For i:=0 To LS.Count-1 Do
  BEGIN
  CS:=LS[i];
  TS := InfoImp.TOBStat.FindFirst(['CODE'],[CS],False);
  if TS <> nil then stLibelle := TS.GetValue('LIBELLE')
  else stLibelle := CS;
  MoveCur(FALSE) ;
  St:='***TL0'+Format_String(LS[i],17)+Format_String(stLibelle,35)+'ECR' ;
  ImportTL(St) ;
  END ;
FiniMove ;
// ajout me 11/07/2003 récupération direct dans comsx dans e_table0
if not GetRecupComSx  then
begin
      LC:=HTStringList.Create ; LC.Sorted:=TRUE ; LC.duplicates:=DupIgnore ;
      Q:=OpenSQL('SELECT DISTINCT E_GENERAL FROM ECRITURE WHERE E_LIBRETEXTE0<>"" ',TRUE) ;
      While Not Q.Eof Do BEGIN LC.Add(Q.Fields[0].AsString) ; Q.Next ; END ;
      Ferme(Q) ;
      InitMove(LC.count,'') ;
      For i:=0 To LC.Count-1 Do
        BEGIN
        MoveCur(FALSE) ;
        CS:=LC[i];
        St:='UPDATE ECRITURE SET E_TABLE0=E_LIBRETEXTE0 WHERE E_GENERAL="'+CS+'" AND E_LIBRETEXTE0<>""' ;
        ExecuteSQL(St) ;
        END ;
      FiniMove ;
      VideStringList(LC) ; LC.Free ;
end;
END ;

Procedure UpdateParamSect ;
Var _BourreAna : String ;
    _LgAna : String ;
    ll1 : Integer ;
    Q : TQuery ;
BEGIN
_BourreAna:='0' ;
If GetRecupSISCOPGI Then
  BEGIN
  If LitTRFParam('000','CP_LGOK')='X' Then
    BEGIN
    _LgAna:=LitTRFParam('000','CP_LGANA') ;
    If (_LgAna<>'') And (StrToInt(_LgAna)>4) Then
      BEGIN
      ll1:=StrToInt(_LgAna) ;
      _BourreAna:=LitTRFParam('000','CP_BOURREANA') ;
      If _BourreAna='' Then _BourreAna:='0' ;
      END Else ll1:=6 ;
    END Else
    BEGIN
    ll1:=6 ;
    END ;
  END Else
  BEGIN
  // Fiche 10514
  if GetInfoCpta(fbaxe1).Lg > 0 then
    ll1:= GetInfoCpta(fbaxe1).Lg
  else
    ll1 := 6;
  END ;
Q:=OpenSQL('SELECT * FROM AXE WHERE X_AXE="A1"',FALSE) ;
If Not Q.Eof Then
 BEGIN
 Q.Edit ;
 Q.FindField('X_LONGSECTION').AsInteger:=ll1 ;
 Q.FindField('X_BOURREANA').AsString:=_BourreAna ;
{$IFNDEF NOVH}
 VH^.Cpta[fbAxe1].Lg:=ll1 ; VH^.Cpta[fbAxe1].Cb:=_BourreAna[1] ;
{$ENDIF}
 Q.FindField('X_SECTIONATTENTE').AsString:=BourreEtLess(Q.FindField('X_SECTIONATTENTE').AsString,fbAxe1) ;
 Q.Post ;
 END ;
Ferme(Q) ;
END ;

Function RecupCptSect(St : String ; Var Sect : TCptImport ; Var InfoImp : TInfoImport ; TL17 : Boolean = FALSE) : Boolean ;
BEGIN
Fillchar(Sect,SizeOf(Sect),#0) ;
With Sect Do
  BEGIN
  Cpt:=Trim(Copy(St,7,17)) ; Libelle:=Trim(Copy(St,24,35)) ;

  {$IFDEF COMPTA}
  {JP 13/06/03 : Dans le cas d'une récupération de sisco vers la compta PGI, on force
                 lors de l'écriture des OD Analytique Y_AXE à "A1", ce qui n'est pas
                 compatible avec cet axe sur trois caractères dans la table SECTION}
  If GetRecupSISCOPGI Then
     Axe:= 'A1'
  else
     Axe:=Trim(Copy(St,59,3)) ;
  {$ELSE}
  Axe:=Trim(Copy(St,59,3)) ;
  {$ENDIF}

  If TL17 Then
    BEGIN
    T0:=Trim(Copy(St,62,17))  ; T1:=Trim(Copy(St,79,17)) ;
    T2:=Trim(Copy(St,96,17))  ; T3:=Trim(Copy(St,113,17)) ;
    T4:=Trim(Copy(St,130,17))  ; T5:=Trim(Copy(St,147,17)) ;
    T6:=Trim(Copy(St,164,17))  ; T7:=Trim(Copy(St,181,17)) ;
    T8:=Trim(Copy(St,198,17))  ; T9:=Trim(Copy(St,215,17)) ;
    Abrege:=Trim(Copy(St,232,17)) ; Sens:=Trim(Copy(St,249,3)) ;
    END Else
    BEGIN
    T0:=Trim(Copy(St,62,3))  ; T1:=Trim(Copy(St,65,3)) ;
    T2:=Trim(Copy(St,68,3))  ; T3:=Trim(Copy(St,71,3)) ;
    T4:=Trim(Copy(St,74,3))  ; T5:=Trim(Copy(St,77,3)) ;
    T6:=Trim(Copy(St,80,3))  ; T7:=Trim(Copy(St,83,3)) ;
    T8:=Trim(Copy(St,86,3))  ; T9:=Trim(Copy(St,89,3)) ;
    Abrege:=Trim(Copy(St,92,17)) ; Sens:=Trim(Copy(St,109,3)) ;
    END ;
  If (Sens='') Or ((Sens<>'M') And (Sens<>'D') And (Sens<>'C')) Then Sens:='M' ;
  If Axe='' Then Axe:='A1' ;
  If InfoImp.SC.Majuscule Then Cpt:=AnsiUpperCase(Cpt) ;
  END ;
Result:=(Sect.Cpt<>'') And (Sect.Axe<>'') ;
END ;


Function ImportSect(St : String ; Var InfoImp : TInfoImport ; QFiche : TQFiche ; TL17 : Boolean = FALSE) : tResultImportCpte ;
Var Q : TQuery ;
    CptLu : TCptLu ;
    Sect : TCptImport ;
    OkCreat,Existe : Boolean ;
    leFb : TFichierBase ;
    SectOrigine : String ;
BEGIN
Result:=ResRien ;
OkCreat:=RecupCptSect(St,Sect,InfoImp,TL17) ; SectOrigine:=Sect.Cpt ; TraiteCorrespCpt(StrToInt(Copy(Sect.Axe,2,1)),QFiche,Sect,InfoImp) ;
Lefb:=fbAxe1 ;
if (Sect.Axe='A2') then Lefb:=fbAxe2 else
if (Sect.Axe='A3') then Lefb:=fbAxe3 else
if (Sect.Axe='A4') then Lefb:=fbAxe4 else
if (Sect.Axe='A5') then Lefb:=fbAxe5 ;
Sect.Cpt:=BourreOuTronque(Sect.Cpt,Lefb) ;
CptLu.Cpt:=Sect.Cpt ; CptLu.Axe:=Sect.Axe ;
Existe:=AlimLTabCptLu(2,QFiche[2],InfoImp.LAnaLu,Nil,CptLu) ;
If Existe Or ((Not Existe) And OkCreat) Then
  BEGIN
  (*
  Q:=InitSQLCpt(2) ;
  If Existe Then Q.Params[0].AsString:=Sect.Cpt Else Q.Params[0].AsString:=w_w ;
  Q.Params[1].AsString:=Sect.Axe ;
  Q.Open ;
  *)
  Q:=OpenSQL(InitSQLCpt1(2,Sect.Cpt,Sect.Axe,'',Existe),FALSE) ;
  Result:=ResModifier ;
  If Existe Then Q.Edit Else BEGIN Q.Insert ; InitNew(Q) ; Result:=ResCreer ; END ;
  If Not Existe Then
    BEGIN
    Q.FindField('S_SECTION').AsString:=Sect.Cpt ;
    Q.FindField('S_AXE').AsString:=Sect.Axe ;
    If Sect.Sens<>'' Then Q.FindField('S_SENS').AsString:=Sect.Sens ;
    END ;
  Q.FindField('S_LIBELLE').AsString:=Sect.Libelle ;
  If Sect.Abrege<>'' Then Q.FindField('S_ABREGE').AsString:=Sect.Abrege
                     Else Q.FindField('S_ABREGE').AsString:=Copy(Sect.Libelle,1,17) ;
  If Sect.T0<>'' Then Q.FindField('S_TABLE0').AsString:=Sect.T0 ;
  If Sect.T1<>'' Then Q.FindField('S_TABLE1').AsString:=Sect.T1 ;
  If Sect.T2<>'' Then Q.FindField('S_TABLE2').AsString:=Sect.T2 ;
  If Sect.T3<>'' Then Q.FindField('S_TABLE3').AsString:=Sect.T3 ;
  If Sect.T4<>'' Then Q.FindField('S_TABLE4').AsString:=Sect.T4 ;
  If Sect.T5<>'' Then Q.FindField('S_TABLE5').AsString:=Sect.T5 ;
  If Sect.T6<>'' Then Q.FindField('S_TABLE6').AsString:=Sect.T6 ;
  If Sect.T7<>'' Then Q.FindField('S_TABLE7').AsString:=Sect.T7 ;
  If Sect.T8<>'' Then Q.FindField('S_TABLE8').AsString:=Sect.T8 ;
  If Sect.T9<>'' Then Q.FindField('S_TABLE9').AsString:=Sect.T9 ;
  If Q.FindField('S_CODEIMPORT')<>NIL Then Q.FindField('S_CODEIMPORT').AsString:=SectOrigine ;
//  If ((Not Existe) And OkCreat) Then Q.FindField('S_CREERPAR').AsString:='IMP' ;
  If ((Not Existe) And OkCreat) Then SetCreerPar(Q,'S_CREERPAR') ;
  Q.Post ;
  Q.Close ; Ferme(Q) ;
  END ;
END ;


Function WhereCptSupp(Pref : String) : String ;
BEGIN
//Result:=' AND '+Pref+'_GENERAL<>"'+VH^.EccEuroDebit+'" AND '+Pref+'_GENERAL<>"'+VH^.EccEuroCredit+'" '+
//        ' AND '+Pref+'_GENERAL<>"'+VH^.Cpta[fbGene].Attente+'" ' ;
  Result :=' AND '+Pref+'_GENERAL<>"'+GetInfoCpta(fbGene).Attente+'" ' ;
END ;

Function TrouveDevise(TobD : Tob ; Dev : String) : Integer ;
Var TOBL : TOB ;
BEGIN
Result:=V_PGI.OkDecV ;
TOBL:=TOBD.FindFirst(['D_DEVISE'],[Dev],FALSE) ;
if TOBL<>Nil then Result:=TOBL.GetValue('D_DECIMALE') ;
END ;

Procedure Cpt2Vent(LC :HtStringList) ;
Var i,j,Dec : Integer ;
    St,St1,Dev,OldDev : String ;
    Q : TQuery ;
    TOBE,TobL,TobD : Tob ;
    StSect : String ;

    procedure CreationVentilType;
    var Q : TQuery;
        TVentil, T, TDefVentil : TOB;
        NumVentil : integer;
        stSection : string;
    begin
      Q := OpenSQL ('SELECT E_GENERAL,E_LIBRETEXTE0 FROM ECRITURE WHERE E_LIBRETEXTE0<> "" GROUP BY E_GENERAL,E_LIBRETEXTE0',True);
      TVentil := TOB.Create('', nil, - 1);
      TVentil.LoadDetailDB('VENTIL','','',nil,False);
      TDefVentil := TOB.Create ('', nil, - 1);
      TDefVentil.LoadDetailDB('CHOIXCOD','"VTY"','',nil,True);
      try
        NumVentil := 1;
        while not Q.Eof do
        begin
          { Création de la définition de la ventilation type }
          T := TDefVentil.FindFirst(['CC_TYPE','CC_CODE'],['VTY',Format('%3.3d',[NumVentil])],False);
          { Création de la ventilation type }
          if T = nil then
          begin
            T := TOB.Create( 'CHOIXCOD',  TDefVentil, -1);
            T.PutValue('CC_TYPE','VTY');
            T.PutValue('CC_CODE',Format('%3.3d',[NumVentil]));
            T.PutValue('CC_LIBELLE', TraduireMemoire('Ventilation'+' '+Format('%3.3d',[NumVentil])));
          end;
          stSection := BourreOuTronque(Q.FindField('E_LIBRETEXTE0').AsString, fbAxe1);
          T := TVentil.FindFirst(['V_NATURE','V_SECTION'],['TY1',stSection],False);
          if T = nil then
          begin
            T := TOB.Create( 'VENTIL',  TVentil, -1);
            T.PutValue('V_NATURE','TY1');
            T.PutValue('V_COMPTE',Format('%3.3d',[NumVentil]));
            T.PutValue('V_SECTION',stSection);
            T.PutValue('V_TAUXMONTANT',100);
            T.PutValue('V_NUMEROVENTIL',1);
            Inc ( NumVentil );
          end;
          { Association d'une ventilation type au compte }
          T := TVentil.FindFirst(['V_NATURE','V_COMPTE'],['GE1',Q.FindField('E_GENERAL').AsString],False);
          if T = nil then
          begin
            T := TOB.Create( 'VENTIL',  TVentil, -1);
            T.PutValue('V_NATURE','GE1');
            T.PutValue('V_COMPTE',Q.FindField('E_GENERAL').AsString);
            T.PutValue('V_SECTION',stSection);
            T.PutValue('V_TAUXMONTANT',100);
            T.PutValue('V_NUMEROVENTIL',1);
          end;
        Q.Next;
      end;
      TVentil.InsertOrUpdateDB;
      TDefVentil.InsertOrUpdateDB;
      finally
        TDefVentil.Free;
        TVentil.Free;
        Ferme (Q);
      end;
    end;

BEGIN
// Suppression des écritures d'analytique éventuellement générées lors
// si des comptes étaient ventilables
{JP 31/10/03 : en PGE, on peut récupérer les fichiers en plusieurs sessions pas en PCL. Or RecupPCL est à True
               même lors d'une récupération PGE : d'où le rajout de FromPCL}
if GetRecupPCL and GetFromPCL then ExecuteSQL('DELETE FROM ANALYTIQ') ;
TobD:=TOB.Create('',Nil,-1) ;
Q:=OpenSQL('SELECT D_DEVISE,D_DECIMALE FROM DEVISE ',TRUE) ;
TOBD.LoadDetailDB('DEVISE','','',Q,False,True) ;
Ferme(Q) ;
InitMove(LC.count,'') ;
For i:=0 To LC.Count-1 Do
  BEGIN
  MoveCur(FALSE) ;
  St:=LC[i] ;
  St1:='UPDATE GENERAUX SET G_VENTILABLE="X", G_VENTILABLE1="X" WHERE G_GENERAL="'+St+'" '+WhereCptSupp('G') ;
  ExecuteSQL(St1) ;
  END ;
FiniMove ;

if GetRecupPCL and GetFromPCL then CreationVentilType;

TobE:=TOB.Create('',Nil,-1) ;
OldDev:='' ; Dec:=V_PGI.OkDecV ;
InitMove(LC.count,'') ;
For i:=0 To LC.Count-1 Do
  BEGIN
  MoveCur(FALSE) ;
  St:=LC[i] ;
//  Q:=OpenSQL('SELECT * FROM ECRITURE WHERE E_GENERAL="'+St+'" AND E_LIBRETEXTE0<>"" AND E_LIBRETEXTE1="" '+WhereCptSupp('E'),TRUE) ;
  Q:=OpenSQL('SELECT * FROM ECRITURE WHERE E_GENERAL="'+St+'" AND E_LIBRETEXTE1="" '+WhereCptSupp('E'),TRUE) ;
  TOBE.LoadDetailDB('ECRITURE','','',Q,False,True) ;
  Ferme(Q) ;
  For j:=0 To TobE.Detail.Count-1 Do
    BEGIN
    TobL:=TobE.Detail[j] ;
    Dev:=TobL.GetValue('E_DEVISE') ;
    If Dev<>OldDev Then BEGIN Dec:=TrouveDevise(TobD,Dev) ; OldDev:=Dev ; END ;
    if TOBL.Detail.Count<=0 then
      AlloueAxe( TOBL ) ;  // SBO 25/01/2006
    if GetRecupComSx  then  // ajout me 11/07/2003
    begin
         StSect := Trim(TobL.GetValue('E_TABLE0')) ;
         if StSect <> '' then
         begin
               // table correspondance  sur les axes
               Q := OpenSql('Select CR_CORRESP,CR_LIBELLE,CR_ABREGE from CORRESP where CR_TYPE="'+'IA1'+'" and CR_CORRESP="'
                          + StSect +'"', TRUE);
               if not Q.EOF then
                   StSect := BourreEtLess(Q.Findfield('CR_LIBELLE').asstring,fbAxe1)
               else
                   StSect:=BourreEtLess(TobL.GetValue('E_TABLE0'),fbAxe1) ;
               Ferme (Q);
         end;
    end
    else
    begin
         StSect:=Trim(TobL.GetValue('E_LIBRETEXTE0')) ;
         If StSect<>'' Then StSect:=BourreEtLess(TobL.GetValue('E_LIBRETEXTE0'),fbAxe1) ;
    end;
    VentilerTOBCreat(TOBL,Dec,'X----',StSect,FALSE) ;
    END ;
  TobE.InsertDBByNivel(TRUE,3,3) ;
  END ;
FiniMove ;
InitMove(LC.count,'') ;
For i:=0 To LC.Count-1 Do
  BEGIN
  MoveCur(FALSE) ;
  St:=LC[i] ;
  if GetRecupComSx  then  // ajout me 11/07/2003
     St1:='UPDATE ECRITURE SET E_TABLE0="",E_ANA="X",E_LIBRETEXTE1="X" WHERE E_GENERAL="'+St+'" AND E_LIBRETEXTE1="" '+WhereCptSupp('E')
  else
     St1:='UPDATE ECRITURE SET E_ANA="X",E_LIBRETEXTE1="X" WHERE E_GENERAL="'+St+'" AND E_LIBRETEXTE1="" '+WhereCptSupp('E') ;
  ExecuteSQL(St1) ;
  END ;
FiniMove ;
TOBE.Free ; TOBD.Free ;
END ;



// Création des sections et des ventilations analytiques depuis codes
// statistiques des fichiers TRT
// Si des sections ont été effectivement créées cette fonction renvoie vrai, sinon false
function CodeStat2Ana(Var InfoImp :TInfoImport;  bCreationSection : boolean) : boolean ;
Var QFiche : TQFiche ;
    LS,LC : HtStringList ;
    CS,St, StLibelle : String ;
    i : Integer ;
    Q : tQuery ;
    TS : TOB;
BEGIN
Result := False;
// Création des sections analytiques
if bCreationSection then
begin
  LS:=InfoImp.LCS ; If LS=Nil Then Exit ; If LS.Count<=0 Then Exit ;
  UpdateParamSect ;
  InitRequete(QFiche[2],2) ;
  For i:=0 To LS.Count-1 Do
  BEGIN
    CS:=LS[i];
    TS := InfoImp.TOBStat.FindFirst(['CODE'],[CS],False);
    if TS <> nil then StLibelle := TS.GetValue('LIBELLE')
    else StLibelle := CS;
//    St:='***SAN'+Format_String(LS[i],17)+Format_String(LS[i],35)+'A1' ;
    St:='***SAN'+Format_String(LS[i],17)+Format_String(StLibelle,35)+'A1' ;
    ImportSect(St,InfoImp,QFiche)
  END ;
  Ferme(QFiche[2]) ;
  Result := True;
end else
// Création des lignes de ventilation analytique
begin
  LC:=HTStringList.Create ; LC.Sorted:=TRUE ; LC.duplicates:=DupIgnore ;
  if GetRecupComSx  then
  Q:=OpenSQL('SELECT DISTINCT E_GENERAL FROM ECRITURE WHERE E_TABLE0<>"" ',TRUE)
  else
  Q:=OpenSQL('SELECT DISTINCT E_GENERAL FROM ECRITURE WHERE E_LIBRETEXTE0<>"" ',TRUE) ;
  While Not Q.Eof Do BEGIN LC.Add(Q.Fields[0].AsString) ; Q.Next ; END ;
  Ferme(Q) ;
  Cpt2Vent(LC) ;
  VideStringList(LC) ; LC.Free ;
end;
END ;



function TraiteCodeStatSISCO(Quoi : String ; Var InfoImp :TInfoImport; bCreationSection : boolean)  : boolean;
BEGIN
Result := False;
If Quoi='TL' Then CodeStat2TL(InfoImp, InfoImp.LCS)  Else If Quoi='ANA' Then Result := CodeStat2Ana(InfoImp, bCreationSection) ;
END ;

Function TraiteCorrespCpt(Quoi : Integer ; QFiche : TQfiche ; Var CptI : TCptImport ; Var InfoImp : TInfoImport) : Boolean ;
Var OkOk : Boolean ;
    CRType : String ;
    Cpt : String ;
    LeFb : TFichierBase ;
BEGIN
OkOk:=FALSE ; CRTYPE:='IGE' ; LeFb:=fbGene ; Result:=FALSE ; Cpt:=CptI.Cpt ; If InfoImp.SC.Majuscule Then Cpt:=AnsiUpperCase(CptI.Cpt) ;
Case Quoi Of
  0 : BEGIN
      OkOk:=InfoImp.Sc.CorrespGen AND (Cpt<>'') ; CRTYPE:='IGE' ; LeFb:=fbGene ;
      END ;
  1 : BEGIN
      OkOk:=InfoImp.Sc.CorrespAux AND (Cpt<>'') ; CRTYPE:='IAU' ; LeFb:=fbAux ;
      END ;
  2 : BEGIN
      OkOk:=InfoImp.Sc.CorrespSect1 AND (Cpt<>'') And (CptI.AXE='A1') ; CRTYPE:='IA1' ; LeFb:=fbAxe1 ;
      END ;
  3 : BEGIN
      OkOk:=InfoImp.Sc.CorrespSect2 AND (Cpt<>'') And (CptI.AXE='A2') ; CRTYPE:='IA2' ; LeFb:=fbAxe2 ;
      END ;
  4 : BEGIN
      OkOk:=InfoImp.Sc.CorrespSect3 AND (Cpt<>'') And (CptI.AXE='A3') ; CRTYPE:='IA3' ; LeFb:=fbAxe3 ;
      END ;
  5 : BEGIN
      OkOk:=InfoImp.Sc.CorrespSect4 AND (Cpt<>'') And (CptI.AXE='A4') ; CRTYPE:='IA4' ; LeFb:=fbAxe4 ;
      END ;
  6 : BEGIN
      OkOk:=InfoImp.Sc.CorrespSect5 AND (Cpt<>'') And (CptI.AXE='A5') ; CRTYPE:='IA5' ; LeFb:=fbAxe5 ;
      END ;
  (*
  7 : BEGIN
      OkOk:=InfoImp.Sc.CorrespMP AND (Cpt<>'') ; CRTYPE:='IPM' ; Cpt:=Mvt.IE_MODEPAIE ;
      END ;
  *)
  8 : BEGIN
      OkOk:=InfoImp.Sc.CorrespJal AND (Cpt<>'') ; CRTYPE:='IJA' ;
      END ;
  9 : BEGIN
      OkOk:=FALSE ;
      {$IFDEF SYSTEMU}
      OkOk:=TRUE ; CRTYPE:='IRM' ;
      {$ENDIF}
      END ;
  END ;
If Not OkOk Then Exit ;
QFiche[5].Params[0].AsString:=CRTYPE ;
QFiche[5].Params[1].AsString:=Cpt ;
QFiche[5].Open ;
If Not QFiche[5].Eof Then
  BEGIN
  Result:=TRUE ;
  Case Quoi Of
    0 : CptI.Cpt:=BourreOuTronque(QFiche[5].Fields[2].AsString,LeFb) ;
    1 : CptI.Cpt:=BourreOuTronque(QFiche[5].Fields[2].AsString,LeFb) ;
    2..6 : CptI.Cpt:=BourreOuTronque(QFiche[5].Fields[2].AsString,LeFb) ;
//    7 : Mvt.IE_MODEPAIE:=QFiche[5].Fields[2].AsString ;
    8 : CptI.Cpt:=QFiche[5].Fields[2].AsString ;
    9 : CptI.Cpt:=QFiche[5].Fields[3].AsString
    END ;
  END ;
QFiche[5].Close ;
END ;

Procedure SetCreerpar(Q : tQuery ; StWhat : String) ;
BEGIN
If GetRecupSISCOPGI Then Q.FindField(StWhat).AsString:='REC'
                     Else Q.FindField(StWhat).AsString:='IMP' ;

END ;

Function ChercheRibPrincipal(CptAux : String) : String ;
Var St : String ;
    QRib : TQuery ;
BEGIN
  Result:='' ; If Trim(CptAux)='' Then Exit ;
  St:='Select * from RIB Where R_AUXILIAIRE="'+CptAux+'" AND R_PRINCIPAL="X"' ;
  QRib:=OpenSQL(St,TRUE) ;
  {JP 21/10/03 : Si l'auxiliaire n'est pas français, on récupère l'IBAN}
  if not QRib.EOF then begin
    if UpperCase(Copy(QRib.Findfield('R_PAYS').AsString, 1, 2)) = 'FR' then
      Result := EncodeLeRib(QRib.Findfield('R_ETABBQ').ASString,QRib.Findfield('R_GUICHET').ASString,
                            QRib.Findfield('R_NUMEROCOMPTE').ASString,QRib.Findfield('R_CLERIB').ASString,
                            QRib.Findfield('R_DOMICILIATION').ASString)
    else
      Result := '*' + QRib.Findfield('R_CODEIBAN').AsString;
  end;
  Ferme(QRib) ;
END ;

function EncodeLeRIB ( Etab,Guichet,Numero,Cle,Dom : String ) : String ;
Var St : String ;
BEGIN
St:=Format_String(Etab,5)+'/'+Format_String(Guichet,5)+'/'+Format_String(Numero,11)
   +'/'+Format_String(Cle,2)+'/'+Format_String(Dom,24) ;
if ((Trim(Etab)='') or (Trim(Guichet)='') or (Trim(Numero)='') or (Trim(Cle)='')) then St:=Format_String('',Length(St)) ;
Result:=St ;
END ;

Procedure VerifUneSouche(CodeSouche : String ; Ana,Simu : Boolean) ;
Var Q : TQuery ;
BEGIN
If CodeSouche='' Then Exit ;
(*
Q:=InitSQLCpt(12) ;
Q.Params[0].AsString:=CodeSouche ;
Q.Open ;
*)
Q:=OpenSQL(InitSQLCpt1(12,CodeSouche,'','',TRUE),FALSE) ;
If Q.Eof Then
  BEGIN
  Q.Insert ; InitNew(Q) ;
  Q.FindField('SH_TYPE').AsString:='CPT' ;
  Q.FindField('SH_SOUCHE').AsString:=CodeSouche ;
  Q.FindField('SH_LIBELLE').AsString:='Souche '+CodeSouche ;
  Q.FindField('SH_ABREGE').AsString:=Q.FindField('SH_LIBELLE').AsString ;
  Q.FindField('SH_NUMDEPART').AsInteger:=1 ;
  Q.FindField('SH_NUMDEPARTS').AsInteger:=1 ;
  Q.FindField('SH_NUMDEPARTP').AsInteger:=1 ;
  Q.FindField('SH_SOUCHEEXO').AsString:='-' ;
  Q.FindField('SH_DATEDEBUT').AsDateTime:=iDate1900;
  Q.FindField('SH_DATEFIN').AsDateTime:=iDate2099 ;
  If Ana Then Q.FindField('SH_ANALYTIQUE').AsString:='X' Else Q.FindField('SH_ANALYTIQUE').AsString:='-' ;
  If Simu Then Q.FindField('SH_SIMULATION').AsString:='X' Else Q.FindField('SH_SIMULATION').AsString:='-' ;
  Q.Post ;
  END ;
Ferme(Q) ;
END ;

{$ENDIF}


end.

