unit ImpAuto;
(*
Lettrage Non repris
Alim param défaut lettrage auto
Alim flux écart de conversion avec compte écart de conversion : Compte non renseigné et montants à 0
 01349
Journal AC : Pb bordereau

*)

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, ComCtrls, HStatus, uDbxDataSet, ImpUtil, FmtChoix, Ent1, HEnt1,
  HCtrls,DB,MajTable,TImpFic,LicUtil,ed_tools,(*VerCpta,*) StdCtrls, ImpFicU,
  RappType,HMsgBox,PGIExec, UTOB, ParamSoc, HFLabel, HTB97,uLibStdCpta,
  galSystem, About, Variants, RecupUtil ;

Type SetOfByte = Set Of Byte ;

type
  TFImpauto = class(TForm)
    Status: THStatusBar;
    Panel1: TPanel;
    Timer1: TTimer;
    MessImport: THLabel;
    MessSoc: THLabel;
    Panel2: TPanel;
    TR: TLabel;
    XX: TLabel;
    TRSL: TLabel;
    YY: TLabel;
    L1: TLabel;
    L2: TLabel;
    L3: TLabel;
    L4: TLabel;
    EC1: TFlashingLabel;
    EC2: TFlashingLabel;
    EC3: TFlashingLabel;
    EC4: TFlashingLabel;
    Ok1: TImage;
    OK2: TImage;
    OK3: TImage;
    OK4: TImage;
    HM: THMsgBox;
    AF1: TFlashingLabel;
    AF2: TFlashingLabel;
    AF3: TFlashingLabel;
    AF4: TFlashingLabel;
    EBStop: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure EBStopClick(Sender: TObject);
  private
    { Déclarations privées }
    SocPCL : String ;
    RepPCl,MasquePCL,MonnaiePCL : String ;
    AnaPCL : Boolean ;
    CodeScenarioPCL : String ;
    ProfilPCL : String ;
    FicErrPCL : TextFile ;
    MultiSoc : Boolean ;
    FromPGE : Boolean ;
    _Stat : string;
    procedure LanceImportAuto ;
    Procedure ApresConnecte ;
    procedure InitDivers;
    Function RecupInfoDossier(Var InfoImp : TInfoImport) : Boolean ;
    Function  RecupFicheBase(InfoImp :PtTInfoImport) : Integer ;
    Function  RecupMouvement(InfoImp :PtTInfoImport) : Integer ;
    Function  FiniRecupInfoDossierAuto(Var InfoImp :TInfoImport) : Boolean ;
    procedure Affiche(i : Integer) ;
    Procedure InitMessage ;
    Procedure TraiteErreurGlobable(Ens : SetOfByte ; ErrGlobale : Integer ; Var InfoImp : TInfoImport) ;
    procedure RestaureBalSit;
    procedure OnInformationBalSit( Sender : TObject; Msg : string ; bErr : boolean);
    Procedure CreeCptGenParamsoc(LeCpt,NatCpt,ChampParamSoc1,ChampParamSoc2 : String ;
                                 Var ChampVH1 : String ; Ind : Integer) ;
    Procedure CreeCptCollParamsoc ;
    Procedure CreeCptAuxParamsoc(LeCpt,LeColl,NatCpt,ChampParamSoc1 : String ; Var ChampVH1 : String ; Ind : Integer) ;  public
  public
    { Déclarations publiques }
  end;

var
  FImpauto: TFImpauto;

implementation

uses SISCO,  USISCO, SoldeCpt , ImSaiCoef, ImpExpBDS, EntPGI;

{$R *.DFM}

Procedure TFImpauto.InitMessage ;
Var EC : tControl ;
    i : Integer ;
BEGIN
For i:=1 To 4 Do
  BEGIN
  EC:=TControl(FindComponent('EC'+IntToStr(i))) ;
  If EC<>NIL Then EC.Visible:=FALSE ;
  EC:=TControl(FindComponent('OK'+IntToStr(i))) ;
  If EC<>NIL Then EC.Visible:=FALSE ;
  EC:=TControl(FindComponent('AF'+IntToStr(i))) ;
  If EC<>NIL Then EC.Visible:=TRUE ;
  END ;
XX.Caption:='' ; YY.Caption:='' ;
Application.ProcessMessages ;
END ;

procedure TFImpauto.Affiche(i : Integer) ;
Var EC : tControl ;
BEGIN
EC:=TControl(FindComponent('AF'+IntToStr(i))) ;
If EC<>NIL Then EC.Visible:=FALSE ;
EC:=TControl(FindComponent('EC'+IntToStr(i))) ;
If EC<>NIL Then EC.Visible:=TRUE ;
If i>1 Then
  BEGIN
  EC:=TControl(FindComponent('EC'+IntToStr(i-1))) ;
  If EC<>NIL Then EC.Visible:=FALSE ;
  EC:=TControl(FindComponent('OK'+IntToStr(i-1))) ;
  If EC<>NIL Then EC.Visible:=TRUE ;
  END ;
Application.ProcessMessages ;
END ;

Function PCLACTIF : Boolean ;
BEGIN
Result:=((ctxPCL in V_PGI.PGIContexte)=TRUE) ;
END ;

procedure TFImpauto.FormShow(Sender: TObject);
Var i,j : Integer ;
    St,St1,Nom,Value : String ;
begin
  VStatus:=Status ;
  Status.Caption := Copyright ;
  VH^.ModeSilencieux:=TRUE ;
  SocPCL:='' ; RepPCl:='' ; MasquePCL:='' ; MonnaiePCL:='' ; ProfilPCL:='000' ;
  If VH^.FromPCL Then
  begin
    If (Not V_PGI.CegidApalatys) then
    begin
      HalSocIni := 'CEGIDPGI.INI';
    end;
    //# évite message bloquant
    If V_PGI.RunFromLanceur then V_PGI.MultiUserLogin := True;
//    V_PGI.UserName:='' ;
    V_PGI.UserLogin:='' ;
    V_PGI.PassWord:='' ;
    SocPCL:='' ;
  v_pgi.ModePCL := '0';
  v_pgi.NoDossier := '000000';
  v_pgi.InBaseCommune := false;
  for i := 1 to ParamCount do
  begin
    St := ParamStr(i);
    Nom := UpperCase(Trim(ReadTokenPipe(St, '=')));
    Value := UpperCase(Trim(St));
    //Paramètres de connexion
    if Nom = '/USER' then
    begin
      V_PGI.UserLogin :=Value ;
    end
    else if Nom = '/PASSWORD' then V_PGI.PassWord:= DecryptageSt(Value)
    else if Nom = '/DATE' then V_PGI.DateEntree := StrToDate(Value)
    else if Nom = '/DOSSIER' then
    begin
      SocPCL:=Value;
      V_PGI.CurrentAlias := Value;
      V_PGI.DefaultSectionName := '';
      j := pos('@', Value);
      if (j > 0) then
      begin
        V_PGI.DefaultSectionName := Copy(Value, j + 1, 255);
        V_PGI.CurrentAlias := Copy(Value, 1, j - 1);
        if (j > 3) and (Copy(Value, 1, 2) = 'DB') then
          V_PGI.NoDossier := Copy(Value, 3, j - 3); // sinon reste à '000000'
      end;
      V_PGI.RunFromLanceur := (V_PGI.DefaultSectionName <> '');
      if v_pgi.DefaultSectionName <> '' then
        if v_pgi.DBName = v_pgi.DefaultSectionDBName then
          v_pgi.InBaseCommune := true;
    end
    else if Nom = '/MODEPCL' then
    begin
      begin
        if assigned(v_pgi) then v_pgi.ModePCL := Value;
      end;
    end else if Nom='/TRF'  then
    BEGIN
        If Value[Length(Value)]<>';' Then Value:=Value+';' ;
        St1:=ReadTokenSt(Value) ; If St1<>'' Then RepPCL:=St1 ;
        St1:=ReadTokenSt(Value) ; If St1<>'' Then MasquePCL:=St1 ;
        St1:=ReadTokenSt(Value) ; If St1<>'' Then MonnaiePCL:=St1 ;
        St1:=ReadTokenSt(Value) ; If St1<>'' Then AnaPCL:=St1='O' ; // Env.Nodossier
        St1:=ReadTokenSt(Value) ; If St1<>'' Then ProfilPCL:=St1 ; // Env.Nodossier
    END;
  end;
  END Else
  BEGIN
  If FromPGE Then
    BEGIN
    END Else
    BEGIN
    If ParamCount>=2 Then SocPCL:=ParamStr(2) ;
//    If ParamCount>=3 Then V_PGI.UserName:=ParamStr(3) ;
    If ParamCount>=3 Then V_PGI.UserLogin:=ParamStr(3) ;
    If ParamCount>=4 Then RepPCl:=ParamStr(4) ;
    If ParamCount>=5 Then MasquePCL:=ParamStr(5) ;
    If ParamCount>=6 Then MonnaiePCL:=ParamStr(6) ;
    If ParamCount>=7 Then AnaPCL:=ParamStr(7)='O' ;
    If ParamCount>=8 Then ProfilPCL:=ParamStr(8) ;
    END ;
  END ;
  V_PGI.PassWord:=CryptageSt(DayPass(Date)) ;
end;

Procedure TFImpauto.ApresConnecte ;
BEGIN
If VH^.FromPCL Then
  BEGIN
    V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxPCL] ;
  END ;
InitMove(8,'') ;
If Not VH^.FromPCL Then
  BEGIN
  ExecuteSQL('DELETE FROM GENERAUX') ; MoveCur(FALSE) ;
  ExecuteSQL('DELETE FROM EXERCICE') ; MoveCur(FALSE) ;
  ExecuteSQL('DELETE FROM TIERS') ; MoveCur(FALSE) ;
  END ;
ExecuteSQL('DELETE FROM SECTION') ; MoveCur(FALSE) ;
ExecuteSQL('DELETE FROM JOURNAL') ; MoveCur(FALSE) ;
ExecuteSQL('DELETE FROM ECRITURE') ; MoveCur(FALSE) ;
ExecuteSQL('DELETE FROM ANALYTIQ') ; MoveCur(FALSE) ;
ExecuteSQL('DELETE FROM SOUCHE WHERE SH_TYPE="CPT"') ; MoveCur(FALSE) ;
ExecuteSQL('DELETE FROM CORRESP') ; MoveCur(FALSE) ;
If VH^.FromPCL Then  // ajout me 22-03-2006
   InitParamTablesLibres ('');
FiniMove ;
END ;

procedure init ;
BEGIN
V_PGI.Debug:=FALSE ;
V_PGI.Versiondev:=FALSE ;
V_PGI.Synap:=FALSE ;
VH^.GrpMontantMin:=0 ;
VH^.GrpMontantMax:=1000000 ;
V_PGI.DateEntree := Date ;
VH^.Mugler:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.NumVersion:='3' ;
V_PGI.SAV:=TRUE ;
V_PGI.Versiondev:=FALSE ;
V_PGI.Synap:=FALSE ;
VH^.GereSousPlan:=True ;
V_PGI.Halley:=TRUE ;
END ;

Function RepositionneNomFic(Var InfoImp : TInfoImport) : Boolean ;
Var LePath,LeNom,Lextension : String ;
    i : Integer ;
BEGIN
Result:=TRUE ;
(*
If InfoImp.NomFic='AUTO' Then Exit ;
pbSISCO:=FALSE ;
If InfoImp.FormatOrigine='SIS' Then
   BEGIN
   j:=TransfertSISCO(InfoImp.NomFic,FALSE) ;
   If j<>0 Then BEGIN PbSISCO:=TRUE ; Result:=FALSE ; END ;
   InfoImp.NomFicOrigine:=InfoImp.NomFic ;
   InfoImp.NomFic:=NewNomFic(InfoImp.NomFic,'CGE') ;
   END ;
*)
LePath:=ExtractFilePath(InfoImp.NomFic) ;
LExtension:=ExtractFileExt(InfoImp.NomFic) ;
LeNom:=ExtractFileName(InfoImp.NomFic) ;
i:=Pos('.',LeNom) ; If i>0 Then LeNom:=Copy(LeNom,1,i-1) ;
InfoImp.NomFicDoublon:=NewNomFicEtDir(InfoImp.NomFic,'Dbl','Doublons') ;
InfoImp.NomFicRejet:=NewNomFicEtDir(InfoImp.NomFic,'Err','Rejets') ;
InfoImp.NomFicRapport:=NewNomFicEtDir(InfoImp.NomFic,'Rap','Rapports') ;
CreateFicRapport(InfoImp) ;
END ;

Function CreateLeFicRapportGlobal(Var InfoImp : TInfoImport) : Boolean ;
Var LePath,LeNom,StDate : String ;
    YY,MM,DD : Word ;
BEGIN
Result:=TRUE ;
LePath:=InfoImp.EE.Rep ; If LePath[Length(LePath)]<>'\' Then LePath:=LePath+'\' ;
DecodeDate(Date,yy,mm,dd) ; StDate:=IntToStr(YY)+formatFloat('00',MM)+formatFloat('00',DD) ;
LeNom:=LePath+'RAPPORT_GLOBAL_'+HGetUserName+'.TXT' ;
InfoImp.NomFicRapportGlobal:=NewNomFicEtDir(LeNom,'','Rapports') ;
CreateFicRapportGlobal(InfoImp) ;
END ;

Function CreateTraceGlobal(Var InfoImp : TInfoImport ; i : Integer) : Boolean ;
Var LePath,LeNom,StDate : String ;
    YY,MM,DD : Word ;
    FicRap : TextFile ;
    St : String ;
BEGIN
Result:=TRUE ;
LePath:=InfoImp.EE.Rep ; If LePath[Length(LePath)]<>'\' Then LePath:=LePath+'\' ;
DecodeDate(Date,yy,mm,dd) ; StDate:=IntToStr(YY)+formatFloat('00',MM)+formatFloat('00',DD) ;
If VH^.FromPCL Then LeNom:=LePath+'CP'+V_PGI.NoDossier+'.'+IntToStr(i)
               Else LeNom:=InfoImp.NomFicOrigine ;
St:=NewNomFicEtDir(LeNom,'','Rapports') ;
AssignFile(FicRap,St) ;
{$i-} Rewrite(FicRap) ; {$i+}
Writeln(FicRap,IntToStr(i)) ;
CloseFile(FicRap) ;
END ;

Procedure TFImpauto.TraiteErreurGlobable(Ens : SetOfByte ; ErrGlobale : Integer ; Var InfoImp : TInfoImport) ;
Var b : Byte ;
    Msg  : string;
BEGIN
EcrireRapportGlobal(InfoImp,TRUE,'',TRUE) ;
For b:=1 To 20 Do If b In Ens Then
begin
  if VH^.RecupPCL and (b=12) then Msg := HM.Mess[b]+TraduireMemoire(' - NB : Si l''enregistrement 00 est présent, les dates d''exercice du fichier sont incohérentes avec celles du dossier.')
  else if VH^.RecupPCL and (b=16) then Msg := HM.Mess[b]+TraduireMemoire(' - NB : Si l''enregistrement 04 est présent, l''indicateur de clôture du fichier est incohérent avec le dossier.')
  else Msg := HM.Mess[b];
  EcrireRapportGlobal(InfoImp,FALSE,Msg,FALSE) ;
end;
CreateTraceGlobal(InfoImp,ErrGlobale) ;
MessImport.Caption:='PROBLEME SUR LE FICHIER '+InfoImp.NomFic ;
MessImport.Font.Style:=MessImport.Font.Style+[fsUnderline];
MessSoc.Font.Style:=MessSoc.Font.Style+[fsUnderline] ;
Delay(2000) ;
END ;

Function TrouveCodeSocFic(NomFic,St : String ; Var InfoImp : TInfoImport) : String ;
BEGIN
Result:=St ;
If InfoImp.Format='SU' Then
  BEGIN
  Result:=Copy(St,1,5) ;
  END Else
  BEGIN
  Result:=Copy(St,1,5) ;
  END ;
END ;

Function LitCodeSocFic(NomFic : String ; Var InfoImp : TInfoImport) : String ;
Var F : TextFile ;
    St :String ;
BEGIN
Result:='' ; Assign(F,NomFic) ;
{$i-}
Reset(F) ;
if IoResult=0 Then
  BEGIN
  Readln(F,St) ; If St<>'' Then Result:=TrouveCodeSocFic(NomFic,St,InfoImp) ;
  CloseFile(F) ;
  END ;
{$i+}
END ;

Procedure AlimNomFic(Var InfoImp : TInfoImport ; ListeFichier : TStringList ; LeChemin,LeNom : String ; Ind : tTypeMasque) ;
Var SearchRec : TSearchRec ;
    St,St2 : String ;
BEGIN
St2:='0' ;
Case Ind Of
  mMasqueEcr : St2:='1' ; mMasqueAux : St2:='2' ;
  END ;
If FindFirst(LeNom, 0, SearchRec)=0 Then
  BEGIN
  If (Pos('Dbl',SearchRec.Name)=0) And  (Pos('Err',SearchRec.Name)=0) Then
    BEGIN
    St:=LeChemin+SearchRec.Name ; //St1:=LitCodeSocFic(St,InfoImp) ;
//    ListeFichier.Add(St1+';'+St+';'+St2+';') ;
    ListeFichier.Add(St) ;
    END ;
  While (FindNext(SearchRec) = 0) Do
    If (Pos('Dbl',SearchRec.Name)=0) And  (Pos('Err',SearchRec.Name)=0) Then
      BEGIN
      St:=LeChemin+SearchRec.Name ; //St1:=LitCodeSocFic(St,InfoImp) ;
//      ListeFichier.Add(St1+';'+St+';'+St2+';') ;
      ListeFichier.Add(St) ;
      END ;
  END ;
FindClose(SearchRec);
END ;

Procedure RecupNomFic(Var InfoImp : TInfoImport ; ListeFichier : TStringList) ;
Var LeChemin,LePrefixe,LeSuffixe,LeNom : String ;
    Ind  : tTypeMasque ;
    i : Integer ;
BEGIN
For Ind:=mMasqueEcr to mMasqueAux Do
  For i:=0 To 9 Do If InfoImp.EE.MasqueFic[Ind,i]<>'' Then
  BEGIN
  LeChemin:=InfoImp.EE.Rep ;
  LePrefixe:='*' ; LeSuffixe:='*' ;
  If LeChemin<>'' Then If LeChemin[Length(LeChemin)]<>'\' Then LeChemin:=LeChemin+'\' ;
  (*
  LePrefixe:='*' ; LeSuffixe:='*' ;
  If InfoImp.Sc.Chemin<>'' Then LeChemin:=InfoImp.Sc.Chemin+'\' ;
  If InfoImp.Sc.Prefixe<>'' Then LePrefixe:=InfoImp.Sc.Prefixe+'*' ;
  If InfoImp.Sc.Suffixe<>'' Then LeSuffixe:=InfoImp.Sc.Suffixe ;
  *)
  LeNom:=LeChemin+InfoImp.EE.MasqueFic[Ind,i] ;
  DecomposeNomFic(LeNom,LeChemin,LePrefixe,LeSuffixe) ;
  If LeChemin<>'' Then InfoImp.Sc.Chemin:=LeChemin ;
  If LePrefixe<>'' Then InfoImp.Sc.Prefixe:=LePrefixe+'*' ;
  If LeSuffixe<>'' Then InfoImp.Sc.Suffixe:=LeSuffixe ;
  AlimNomFic(InfoImp,ListeFichier,LeChemin,LeNom,Ind) ;
  END ;
END ;

Function FichierSuivant(ListeFichier : TStringList ; Var InfoImp : TInfoImport ; Var NumFic : Integer) : Boolean ;
Var St,StCodeSoc,StNom,StTypeFic : String ;
BEGIN
Result:=FALSE ;
If NumFic<=ListeFichier.Count-1 Then
  BEGIN
  St:=ListeFichier[NumFic] ; StNom:=St ; StCodeSoc:='000' ;
  (*
  StCodeSoc:=ReadTokenSt(St) ; StNom:=ReadTokenSt(St) ; StTypeFic:=ReadTokenSt(St) ;
  *)
  InfoImp.NomFic:=StNom ; InfoImp.CodeSoc:=StCodeSoc ; InfoImp.TypeFic:=StTypeFic ;
//  If Not RepositionneNomFic(InfoImp) Then BEGIN Result:=FALSE ; Exit ;END ;
  Inc(NumFic) ;
  END Else Exit ;
Result:=TRUE ;
END ;

Function AlimInfoImp(Var InfoImp :TInfoImport) : Boolean ;
BEGIN
Result:=TRUE ;
//FillChar(InfoImp,SizeOf(InfoImp),#0) ;
//CreateListe(InfoImp) ;
InfoImp.Lequel:=InfoImp.EE.Import ;
InfoImp.Format:=InfoImp.EE.Format ;
If InfoImp.Format='ORLI' Then BEGIN InfoImp.Format:='SN2' ; Sn2Orli:=TRUE ; END ;
InfoImp.ImportAuto:=TRUE ;
//InfoImp.NomFic:=ParamStr(6) ;
InfoImp.CodeFormat:=InfoImp.EE.CodeFormat ;
PasDeBlanc:=FALSE ;
//If Not PCLActif Then V_PGI.UserName:=InfoImp.EE.User ;
If ParamCount>0 Then If ParamStr(ParamCount)='ORLIBLANCS' Then PasDeBlanc:=TRUE ;
InfoImp.ForceQualif:='' ;
InfoImp.ForcePositif:=FALSE ;
InfoImp.ForceBourrage:=TRUE ;
InfoImp.CtrlDB:=FALSE ;
InfoImp.ForceNumPiece:=FALSE ;
ChargeScenarioImport(InfoImp,True) ;
InfoImp.ForceNumPiece:=TRUE ;
If InfoImp.Sc.EstCharge Then
  BEGIN
  If InfoImp.SC.Doublon Then
    BEGIN
    InfoImp.CtrlDB:=TRUE ;
    If InfoImp.SC.ForcePiece Then InfoImp.ForceNumPiece:=TRUE ;
    END ;
  END ;
If InfoImp.Format='SIS' Then BEGIN InfoImp.FormatOrigine:='SIS' ; InfoImp.Format:='CGE' ;END ;
//If Not RepositionneNomFic(InfoImp) Then Result:=FALSE ;
END ;


Procedure initImporte(Var InfoImp :TInfoImport) ;
BEGIN
InfoImp.NbGenFaux:=0 ; InfoImp.NbAuxFaux:=0 ;
InfoImp.NbAnaFaux:=0 ; InfoImp.NbJalFaux:=0 ;
VideListeInfoImp(InfoImp,FALSE) ;
InfoImp.ListeEnteteDoublon.Sorted:=TRUE ;
InfoImp.ListeEnteteDoublon.Duplicates:=DupIgnore ;
END ;


Procedure initVerif(Var InfoImp :TInfoImport) ;
BEGIN
InfoImp.ListeEntetePieceFausse.Clear ;
InfoImp.ListeEntetePieceFausse.Sorted:=TRUE ;
InfoImp.ListeEntetePieceFausse.Duplicates:=DupIgnore ;
VideListe(InfoImp.ListePieceFausse) ;
InfoImp.ListePbAna.Clear ;
InfoImp.ListePbAna.Sorted:=TRUE ;
InfoImp.ListePbAna.Duplicates:=DupIgnore ;
END ;

Procedure InitIntegre(Var InfoImp :TInfoImport) ;
BEGIN
InfoImp.ListePieceIntegre.Clear ;
InfoImp.ListePieceIntegre.Sorted:=TRUE ;
InfoImp.ListePieceIntegre.Duplicates:=DupIgnore ;
END ;

Procedure DetruitLeFichier(St : String) ;
Var Fichier : TextFile ;
    FichierOk : Boolean ;
    i : Integer ;
BEGIN
i:=0 ;
Repeat
  Inc(i) ;
  AssignFile(Fichier,St) ;
  {$I-} Reset (Fichier) ; {$I+}
  FichierOk:=ioresult=0 ;
  If FichierOk Then Close(Fichier) ;
Until FichierOk Or (i>200) ;
If FichierOk Then
  BEGIN
  {$i-}
  AssignFile(Fichier,St) ; Erase(Fichier) ;
  {$i+}
  END ;
END ;

Procedure DetruitFichier(Var InfoImp :TInfoImport) ;
Var Fichier : TextFile ;
    FichierOk : Boolean ;
    i : Integer ;
BEGIN
If InfoImp.Sc.DetruitFic Then
  BEGIN
  DetruitLeFichier(InfoImp.NomFic) ;
  If (InfoImp.FormatOrigine='SIS') And (InfoImp.NomFicOrigine<>'') Then DetruitLeFichier(InfoImp.NomFicOrigine) ;
  END ;
END ;

Procedure AfficheRapportGlobal(Var InfoImp :TInfoImport) ;
Var St : String ;
BEGIN
St:=InfoImp.EE.Editeur ; If St='' Then Exit ;
St:=St+' "'+InfoImp.NomFicRapportGlobal+'"' ;
FileExec(St,FALSE,TRUE) ;
END ;

Procedure AlimNomFicToKill(Var InfoImp : TInfoImport ; ListeFichier : TStringList ; LeChemin,LeNom : String) ;
Var SearchRec : TSearchRec ;
    St,St1,St2 : String ;
BEGIN
If FindFirst(LeNom, faAnyFile, SearchRec)=0 Then
  BEGIN
  St:=LeChemin+SearchRec.Name ; ListeFichier.Add(St) ;
  While (FindNext(SearchRec) = 0) Do
    BEGIN
    St:=LeChemin+SearchRec.Name ; ListeFichier.Add(St) ;
    END ;
  END ;
FindClose(SearchRec);
END ;

Procedure RecupNomFicToKill(Var InfoImp : TInfoImport ; ListeFichier : TStringList) ;
Var SearchRec : TSearchRec ;
    LeChemin,LePrefixe,LeSuffixe,LeNom,St,St1 : String ;
    Ind  : tTypeMasque ;
    i : Integer ;
BEGIN
Exit ;
For Ind:=mMasqueEcr to mMasqueAux Do
  For i:=0 To 9 Do If InfoImp.EE.MasqueFic[Ind,i]<>'' Then
  BEGIN
  LeChemin:=InfoImp.EE.Rep ;
  LePrefixe:='*' ; LeSuffixe:='*' ;
  If LeChemin<>'' Then If LeChemin[Length(LeChemin)]<>'\' Then LeChemin:=LeChemin+'\' ;
  LeNom:=LeChemin+InfoImp.EE.MasqueFic[Ind,i] ;
  DecomposeNomFic(LeNom,LeChemin,LePrefixe,LeSuffixe) ;
  If LeChemin<>'' Then InfoImp.Sc.Chemin:=LeChemin ;
  If LePrefixe<>'' Then InfoImp.Sc.Prefixe:=LePrefixe+'*' ;
  If LeSuffixe<>'' Then InfoImp.Sc.Suffixe:=LeSuffixe ;
  AlimNomFicToKill(InfoImp,ListeFichier,LeChemin,LeNom) ;
  END ;
END ;

Procedure VideRepertoire(Var InfoImp :TInfoImport) ;
Var StNom : String ;
    LL : TStringList ;
    i : Integer ;
    F : TExtFile ;
BEGIN
If InfoImp.EE.Rep='' Then Exit ;
If Not InfoImp.EE.VideRep Then Exit ;
LL:=tStringList.Create ;
RecupNomFicToKill(InfoImp,LL) ;
For i:=0 To ll.Count-1 Do
  BEGIN
  {$i-}
  AssignFile(F,LL[i]) ; Erase(F) ;
  {$i+}
  END ;
LL.Clear ; LL.Free ;
END ;

Function InitEnv(Var InfoImp :TInfoImport ; RepPCL,MasquePCL : String) : Boolean ;
begin
  Result:=FALSE ;
  if Not VH^.FromPCL Then If Paramcount<6 Then Exit ;
  With InfoImp.EE Do
  BEGIN
    Rep:=RepPCL ;
    VideRep:=TRUE ;
    MasqueFic[mMasqueEcr,0]:=MasquePCL ;
    Import:='FEC' ; Format:='SIS' ;
    CodeFormat:='' ;
    EXE:='' ;
    Editeur:='' ;
  End ;
  If (InfoImp.EE.Rep= '' )  then exit;
  If InfoImp.EE.MasqueFic[mMasqueEcr,0]='' Then Exit ;
  Result:=TRUE ;
END ;

procedure InitEtablissement;
var  OB , OB_DETAIL : TOB;
begin
  // Création de l'établissement par défaut
  OB := TOB.Create('Les établissements',nil,-1);
  OB.LoadDetailDB('ETABLISS','','',nil,False);
  if OB.Detail.Count = 1 then  // Si un seul établissement, on prend celui-ci.
  begin
    SetParamSoc('SO_ETABLISDEFAUT',OB.Detail[0].GetValue('ET_ETABLISSEMENT'));
    VH^.EtablisDefaut:=OB.Detail[0].GetValue('ET_ETABLISSEMENT') ;
  end else
  begin
    OB_DETAIL := TOB.Create('ETABLISS',OB,-1);
    OB_DETAIL.PutValue('ET_ETABLISSEMENT',GetParamSocSecur('SO_SOCIETE',''));
    OB_DETAIL.PutValue('ET_SOCIETE',GetParamSocSecur('SO_SOCIETE',''));
    OB_DETAIL.PutValue('ET_LIBELLE',GetParamSocSecur('SO_LIBELLE',''));
  //  OB_DETAIL.PutValue('ET_ABREGE',GetParamSocSecur('SO_LIBELLE',''));
    OB_DETAIL.PutValue('ET_ABREGE',Copy (GetParamSocSecur('SO_LIBELLE',''),1,17));
    OB_DETAIL.PutValue('ET_ADRESSE1',GetParamSocSecur('SO_ADRESSE1',''));
    OB_DETAIL.PutValue('ET_ADRESSE2',GetParamSocSecur('SO_ADRESSE2',''));
    OB_DETAIL.PutValue('ET_ADRESSE3',GetParamSocSecur('SO_ADRESSE3',''));
    OB_DETAIL.PutValue('ET_CODEPOSTAL',GetParamSocSecur('SO_CODEPOSTAL',''));
    OB_DETAIL.PutValue('ET_VILLE',GetParamSocSecur('SO_VILLE',''));
    OB_DETAIL.PutValue('ET_PAYS',GetParamSocSecur('SO_PAYS',''));
    OB_DETAIL.PutValue('ET_TELEPHONE',GetParamSocSecur('SO_TELEPHONE',''));
    OB_DETAIL.PutValue('ET_FAX',GetParamSocSecur('SO_FAX',''));
    OB_DETAIL.PutValue('ET_TELEX',GetParamSocSecur('SO_TELEX',''));
    OB_DETAIL.PutValue('ET_SIRET',GetParamSocSecur('SO_SIRET',''));
    OB_DETAIL.PutValue('ET_APE',GetParamSocSecur('SO_APE',''));
    OB_DETAIL.PutValue('ET_JURIDIQUE',GetParamSocSecur('SO_NATUREJURIDIQUE',''));
    OB.InsertOrUpdateDB(True);
    // Etablissement par défaut = société = établissement unique.
    SetParamSoc('SO_ETABLISDEFAUT',GetParamSocSecur('SO_SOCIETE',''));
    VH^.EtablisDefaut:=GetParamSocSecur('SO_SOCIETE','') ;
  end;
  OB.Free;
end;

procedure TFImpauto.InitDivers;
Var Q : TQuery ;
    TTVA, T : TOB;
begin
  Q:=OPENSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="RTV" AND CC_CODE="FRA"',FALSE) ;
  If Q.Eof Then
  BEGIN
    Q.Insert ; InitNEw(Q) ;
    Q.FindField('CC_TYPE').AsString:='RTV' ;
    Q.FindField('CC_CODE').AsString:='FRA' ;
    Q.FindField('CC_LIBELLE').AsString:='France' ;
    Q.FindField('CC_ABREGE').AsString:='France' ;
    Q.FindField('CC_LIBRE').AsString:='' ;
    Q.Post ;
  END ;
  Ferme(Q) ;

  Q:=OPENSQL('SELECT * FROM MODEPAIE WHERE MP_MODEPAIE="DIV" ',FALSE) ;
  If Q.Eof Then
  BEGIN
    Q.Insert ; InitNEw(Q) ;
    Q.FindField('MP_MODEPAIE').AsString:='DIV' ;
    Q.FindField('MP_LIBELLE').AsString:='FRA' ;
    Q.FindField('MP_ABREGE').AsString:='France' ;
    Q.FindField('MP_ENCAISSEMENT').AsString:='MIX' ;
    Q.FindField('MP_CATEGORIE').AsString:='CHQ' ;
    Q.Post ;
  END ;
  Ferme(Q) ;

  Q:=OPENSQL('SELECT * FROM MODEREGL WHERE MR_MODEREGLE="002" ',FALSE) ;
  If Q.Eof Then
  BEGIN
    Q.Insert ; InitNEw(Q) ;
    Q.FindField('MR_MODEREGLE').AsString:='002' ;
    Q.FindField('MR_LIBELLE').AsString:='CHEQUE' ;
    Q.FindField('MR_ABREGE').AsString:='CHEQUE' ;
    Q.FindField('MR_APARTIRDE').AsString:='ECR' ;
    Q.FindField('MR_PLUSJOUR').AsInteger:=5 ;
    Q.FindField('MR_ARRONDIJOUR').AsString:='PAS' ;
    Q.FindField('MR_NOMBREECHEANCE').AsInteger:=1 ;
    Q.FindField('MR_SEPAREPAR').AsString:='QUI' ;
    Q.FindField('MR_MONTANTMIN').AsFloat:=999999 ;
    Q.FindField('MR_REMPLACEMIN').AsString:='002' ;
    Q.FindField('MR_MP1').AsString:='CHQ' ;
    Q.FindField('MR_TAUX1').AsFloat:=100 ;
    Q.FindField('MR_ESC1').AsString:='-' ;
    Q.Post ;
  END ;
  Ferme(Q) ;

  CreerDeviseTenue(MonnaiePCL) ;
  SetParamSoc('SO_I_CPTAPGI',False);
  SetParamSoc('SO_DATEDEBUTEURO',EncodeDate(1999,01,04)) ;
  SetParamSoc('SO_REGLEEQUILSAIS','CPT') ;
  SetParamSoc('SO_ZSAUVEFOLIOLOCAL',TRUE);
  SetParamSoc('SO_ZFOLIOTEMPSREEL',TRUE);
  SetParamSoc('SO_LETMODE','PCL');
  SetParamSoc('SO_BOUCLERSAISIECREAT',TRUE);
  SetParamSoc('SO_LETMVTOD',TRUE);
  SetParamSoc('SO_LETTOTAL',TRUE);
  SetParamSoc('SO_LETDC',TRUE);
  { Exercice de référence = 1er exercice ouvert }
  Q := OpenSQL ('SELECT EX_EXERCICE,EX_DATEDEBUT FROM EXERCICE WHERE EX_ETATCPTA="OUV" ORDER BY EX_DATEDEBUT',True);
  SetParamSoc ( 'SO_CPEXOREF',Q.FindField('EX_EXERCICE').AsString);
  Ferme (Q);
  { Régime par défaut = France }
  SetParamSoc ( 'SO_REGIMEDEFAUT','FRA');
  { Mode de réglement par défaut = chèque }
  SetParamSoc ( 'SO_GCMODEREGLEDEFAUT','002');
  { Barre de status : exercice en cours }
  SetParamSoc ( 'SO_CPSTATUSBARRE','EXO');
  // exigibilité de tva  fiche 12064
  SetParamSoc('SO_CODETVADEFAUT','TM');
  // code tva par défaut
  SetParamSoc('SO_CODETVAGENEDEFAULT', 'NOR');

  { TVA par régime fiscal }
  if not ExisteSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="TX1" AND CC_CODE="EXO"') then
    ExecuteSQL ('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) VALUES ("TX1","EXO","Exonéré","Exonéré",3)');
  if ExisteSQL('SELECT * FROM TXCPTTVA') then exit;
  Q := OpenSQL ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="RTV"', True);
  if not Q.Eof then
  begin
    TTVA := TOB.Create ('', nil , -1);
    while not Q.Eof do
    begin
      { Taux normal }
      T := TOB.Create ( 'TXCPTTVA',TTVA, - 1);
      T.PutValue('TV_TVAOUTPF','TX1');
      T.PutValue('TV_CODETAUX','NOR');
      T.PutValue('TV_REGIME',Q.FindField('CC_CODE').AsString);
      if Q.FindField('CC_CODE').AsString='FRA' then
      begin
        T.PutValue('TV_TAUXACH', 19.6);
        T.PutValue('TV_TAUXVTE', 19.6);
        T.PutValue('TV_CPTEACH', BourreEtLess('445661',fbGene));
        T.PutValue('TV_CPTEVTE', BourreEtLess('445710',fbGene));
      end;
      { Taux réduit }
      T := TOB.Create ( 'TXCPTTVA',TTVA, - 1);
      T.PutValue('TV_TVAOUTPF','TX1');
      T.PutValue('TV_CODETAUX','RED');
      T.PutValue('TV_REGIME',Q.FindField('CC_CODE').AsString);
      if Q.FindField('CC_CODE').AsString='FRA' then
      begin
        T.PutValue('TV_TAUXACH', 5.5);
        T.PutValue('TV_TAUXVTE', 5.5);
        T.PutValue('TV_CPTEACH', BourreEtLess('445661',fbGene));
        T.PutValue('TV_CPTEVTE', BourreEtLess('445710',fbGene));
      end;
      { Exonéré }
      T := TOB.Create ( 'TXCPTTVA',TTVA, - 1);
      T.PutValue('TV_TVAOUTPF','TX1');
      T.PutValue('TV_CODETAUX','EXO');
      T.PutValue('TV_REGIME',Q.FindField('CC_CODE').AsString);
      if Q.FindField('CC_CODE').AsString='FRA' then
      begin
        T.PutValue('TV_TAUXACH', 0);
        T.PutValue('TV_TAUXVTE', 0);
        T.PutValue('TV_CPTEACH', BourreEtLess('445661',fbGene));
        T.PutValue('TV_CPTEVTE', BourreEtLess('445710',fbGene));
      end;
      Q.Next;
    end;
    TTVA.InsertDB(nil);
    TTVA.Free;
  end;
  Ferme (Q);
end;

Procedure CreateJalCLO ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT * FROM JOURNAL WHERE J_JOURNAL="CLO" ',FALSE) ;
If Q.Eof Then
  BEGIN
  Q.Insert ; InitNew(Q) ;
  Q.FindField('J_JOURNAL').AsString:='CLO' ;
  Q.FindField('J_LIBELLE').AsString:='CLOTURE' ;
  Q.FindField('J_ABREGE').AsString:='CLOTURE' ;
  Q.FindField('J_NATUREJAL').AsString:='CLO' ;
  Q.FindField('J_COMPTEURNORMAL').AsString:='CLO' ;
  Q.FindField('J_COMPTEURSIMUL').AsString:='' ;
  Q.FindField('J_AXE').AsString:='' ;
  Q.FindField('J_MODESAISIE').AsString:='-' ;
  Q.FindField('J_MULTIDEVISE').AsString:='X' ;
  Q.FindField('J_COMPTEINTERDIT').AsString:='' ;
  Q.FindField('J_COMPTEAUTOMAT').AsString:='' ;
  Q.Post ;
  VerifUneSouche('CLO',FALSE,FALSE) ;
  END ;
Ferme(Q) ;
SetParamSoc('SO_JALFERME','CLO') ;
END ;

// Si aucun journal à nouveau , création du journal ANO
Procedure CreateJalANO ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT * FROM JOURNAL WHERE J_NATUREJAL="ANO" ',FALSE) ;
If Q.Eof Then
  BEGIN
  Q.Insert ; InitNew(Q) ;
  Q.FindField('J_JOURNAL').AsString:='ANO' ;
  Q.FindField('J_LIBELLE').AsString:='A NOUVEAU' ;
  Q.FindField('J_ABREGE').AsString:='A NOUVEAU' ;
  Q.FindField('J_NATUREJAL').AsString:='ANO' ;
  Q.FindField('J_COMPTEURNORMAL').AsString:='ANO' ;
  Q.FindField('J_COMPTEURSIMUL').AsString:='' ;
  Q.FindField('J_AXE').AsString:='' ;
  Q.FindField('J_MODESAISIE').AsString:='-' ;
  Q.FindField('J_MULTIDEVISE').AsString:='X' ;
  Q.FindField('J_COMPTEINTERDIT').AsString:='' ;
  Q.FindField('J_COMPTEAUTOMAT').AsString:='' ;
  Q.Post ;
  VerifUneSouche('ANO',FALSE,FALSE) ;
  SetParamSoc('SO_JALOUVRE','ANO') ;
  END ;
Ferme(Q) ;
END ;

Function CreateCptClo(Racine : String) : Boolean ;
Var Q : TQuery ;
    Cpt : String ;
BEGIN
Result:=FALSE ;
Q:=OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL LIKE "'+Racine+'" ',TRUE) ;
If Not Q.Eof Then
  BEGIN
  Cpt:=Q.Fields[0].AsString ;
  SetParamSoc('SO_OUVREPERTE',Cpt) ; SetParamSoc('SO_FERMEPERTE',Cpt) ;
  SetParamSoc('SO_OUVREBEN',Cpt) ; SetParamSoc('SO_FERMEBEN',Cpt) ;
  Result:=TRUE ;
  END ;
Ferme(Q) ;
END ;

Function TFImpauto.RecupInfoDossier(Var InfoImp : TInfoImport) : Boolean ;
Var StFichier,CptECC,Cpt : String ;
    StM : String ;
    Ens : SetOfByte ;
    b : Byte ;
BEGIN
Result:=TRUE;
MessImport.Visible:=TRUE ;
MessImport.Caption:='Importation en cours du fichier '+InfoImp.NomFic ;
InfoImp.NomFicOrigine:=InfoImp.NomFic ;
StFichier:=InfoImp.NomFicCpt ;
InitRecupDossier(StFichier,InfoImp) ;
InitEtablissement ;
// fiche 10194 ajout me 16-08-2003
if VH^.RecupPCL and (GetParamSocSecur('SO_DEVISEPRINC','') <> '') and (GetParamSocSecur('SO_DEVISEPRINC','')<> 'EUR')  then
begin
     Result := FALSE;
     EcrireRapportGlobal(InfoImp,FALSE,'Monnaie de tenue de la comptabilité est différente de l''EURO',FALSE) ;
     exit;
end;

InitDivers ;
ChargeMagHalley ;
CreateJalCLO ;
// CA - 10/01/2002 - Installation des coefficients dégressifs immo
InstalleLesCoefficientsDegressifs ;
SetParamSoc('SO_MONTANTNEGATIF',TRUE) ; VH^.MontantNegatif:=TRUE ;
SetParamSoc('SO_ZACTIVEPFU',TRUE) ; VH^.ZACTIVEPFU:=TRUE ;
SetParamSoc('SO_NBJECRAVANT',999) ; VH^.NbjEcrAV:=999 ;
SetParamSoc('SO_NBJECRAPRES',999) ; VH^.NbjEcrAP:=999 ;
SetParamSoc('SO_NBJECHAVANT',999) ; VH^.NbjEchAV:=999 ;
SetParamSoc('SO_NBJECHAPRES',999) ; VH^.NbjEchAP:=999;
Cpt:=BourreEtLess('60000000000000',fbGene) ; SetParamSoc('SO_CHADEB1',Cpt) ; VH^.FCha[1].Deb:=Cpt ;
Cpt:=BourreEtLess('69999999999999',fbGene) ; SetParamSoc('SO_CHAFIN1',Cpt) ; VH^.FCha[1].Fin:=Cpt ;
Cpt:=BourreEtLess('70000000000000',fbGene) ; SetParamSoc('SO_PRODEB1',Cpt) ; VH^.FPro[1].Deb:=Cpt ;
Cpt:=BourreEtLess('79999999999999',fbGene) ; SetParamSoc('SO_PROFIN1',Cpt) ; VH^.FPro[1].Fin:=Cpt ;
Cpt:=BourreEtLess('10000000000000',fbGene) ; SetParamSoc('SO_BILDEB1',Cpt) ; VH^.FBil[1].Deb:=Cpt ;
Cpt:=BourreEtLess('59999999999999',fbGene) ; SetParamSoc('SO_BILFIN1',Cpt) ; VH^.FBil[1].Fin:=Cpt ;
// Verif Dev tenu
Ens:=[] ;
If Not InfoImp.OkEnr[0] Then BEGIN Result:=FALSE ; Ens:=Ens+[12] ; END ;
If Not InfoImp.OkEnr[1] Then BEGIN Result:=FALSE ; Ens:=Ens+[13] ; END ;
If Not InfoImp.OkEnr[2] Then BEGIN Result:=FALSE ; Ens:=Ens+[14] ; END ;
If Not InfoImp.OkEnr[3] Then BEGIN Result:=FALSE ; Ens:=Ens+[15] ; END ;
If Not InfoImp.OkEnr[4] Then BEGIN Result:=FALSE ; Ens:=Ens+[16] ; END ;

If Not Result Then
  BEGIN
  (*
  EcrireRapportGlobal(InfoImp,TRUE,'',TRUE) ;
  For b:=12 To 16 Do If b In Ens Then EcrireRapportGlobal(InfoImp,FALSE,HM.Mess[b],FALSE) ;
  CreateTraceGlobal(InfoImp,3) ;
  *)
  TraiteErreurGlobable(Ens,3,InfoImp) ;
  END ;

END ;

function LireProfilInfoStat (ProfilPCL : string) : string;
var Q : TQuery;
begin
  Result := 'RIE';
  Q:=OpenSQL('SELECT * FROM TRFPARAM WHERE TRP_CODE="'+ProfilPCL+'" AND TRP_NOM="CP_STAT"',TRUE) ;
  If Not Q.Eof Then Result:=Q.FindField('TRP_DATA').AsString ;
  Ferme(Q) ;
end;

function OkSituation (ProfilPCL : string) : boolean;
var Q : TQuery;
begin
  Result := False ;
  Q:=OpenSQL('SELECT * FROM TRFPARAM WHERE TRP_CODE="'+ProfilPCL+'" AND TRP_NOM="CP_SITOK"',TRUE) ;
  If Not Q.Eof Then Result:=(Q.FindField('TRP_DATA').AsString ='X');
  Ferme(Q) ;
  Result := True;  // pour l'instant, on force la récupération des balances de situation
end;

Procedure RetoucheInfoImp(InfoImp :PtTInfoImport ; ProfilPCL : String) ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT * FROM TRFPARAM WHERE TRP_CODE="'+ProfilPCL+'" AND TRP_NOM="CP_CPTELETTRABLE"',TRUE) ;
If Not Q.Eof Then InfoImp^.RacineCptGenLet:=AnsiUpperCase(Q.FindField('TRP_DATA').AsString) ;
Ferme(Q) ;
Q:=OpenSQL('SELECT * FROM TRFPARAM WHERE TRP_CODE="'+ProfilPCL+'" AND TRP_NOM="CP_CPTEPOINTABLE"',TRUE) ;
If Not Q.Eof Then InfoImp^.RacineCptGenPoint:=AnsiUpperCase(Q.FindField('TRP_DATA').AsString) ;
Ferme(Q) ;
Q:=OpenSQL('SELECT * FROM TRFPARAM WHERE TRP_CODE="'+ProfilPCL+'" AND TRP_NOM="CP_JOURNALAN"',TRUE) ;
If Not Q.Eof Then InfoImp^.RacineJalAN:=Q.FindField('TRP_DATA').AsString ;
Ferme(Q) ;
END ;

Function TFImpauto.RecupFicheBase(InfoImp :PtTInfoImport) : Integer ;
Var SauveNom : String ;
    Ens : SetOfByte ;
    Q : TQuery ;
    k : Integer ;
BEGIN
Result:=0 ; Ens:=[] ;
RetoucheInfoImp(InfoImp,ProfilPCL) ;
InfoImp^.ProfilPCL:=ProfilPCL ;
TransfertSISCO(InfoImp^.NomFicCpt,FALSE,k,'',InfoImp) ;
If Not InfoImp^.OkFouFou Then BEGIN Result:=1 ; Ens:=Ens+[17] ; END ;
If Not InfoImp^.OkFouCli Then BEGIN Result:=2 ; Ens:=Ens+[18] ; END ;
If result<>0 Then
  BEGIN
  (*
  EcrireRapportGlobal(InfoImp^,TRUE,'',TRUE) ;
  If Result=1 Then EcrireRapportGlobal(InfoImp^,FALSE,HM.Mess[17],FALSE) ;
  If Result=2 Then EcrireRapportGlobal(InfoImp^,FALSE,HM.Mess[18],FALSE) ;
  CreateTraceGlobal(InfoImp^,3) ;
  *)
  TraiteErreurGlobable(Ens,3,InfoImp^) ;
  Exit ;
  END ;
SauveNom:=InfoImp^.NomFic ; InfoImp^.NomFic:=InfoImp^.NomFicCptCGE ;
ImporteLesEcritures(nil,InfoImp^) ;
InfoImp^.NomFic:=SauveNom ;
END ;

Function TFImpauto.FiniRecupInfoDossierAuto(Var InfoImp :TInfoImport) : Boolean ;
Var Cpt : String ;
    Q : TQuery ;
    St1 : String ;
BEGIN
  Result:=TRUE ;
  St1:='' ;
  CreeCptGenParamsoc('1200000000000','DIV','SO_OUVREBEN','SO_FERMEBEN',St1,0) ;
  CreeCptGenParamsoc('1290000000000','DIV','SO_OUVREPERTE','SO_FERMEPERTE',St1,0) ;
  CreeCptGenParamsoc('1280000000000','DIV','SO_RESULTAT','',St1,0) ;
  St1:=VH^.Cpta[fbGene].AxGenAttente ; CreeCptGenParamsoc('4718999999999','DIV','','',St1,37) ; VH^.Cpta[fbGene].AxGenAttente:=St1 ;
  St1:=VH^.OuvreBil ; CreeCptGenParamsoc('8900000000000','DIV','SO_OUVREBIL','SO_FERMEBIL',St1,1) ; VH^.OuvreBil:=St1 ;
  St1:=VH^.Cpta[fbGene].Attente ; CreeCptGenParamsoc('4710000000000','DIV','SO_GENATTEND','',St1,24) ; VH^.Cpta[fbGene].Attente:=St1 ;
  CreeCptCollparamSoc ;
  St1:=VH^.TiersDefCli ; CreeCptAuxParamsoc('CATTENTE00000',GetParamSocSecur('SO_DEFCOLCLI',''),'CLI','SO_CLIATTEND',St1,33) ; VH^.TiersDefCli:=St1 ;
  St1:=VH^.TiersDefFou ; CreeCptAuxParamsoc('FATTENTE00000',GetParamSocSecur('SO_DEFCOLFOU',''),'FOU','SO_FOUATTEND',St1,34) ; VH^.TiersDefFou:=St1 ;
  St1:=VH^.TiersDefSal ; CreeCptAuxParamsoc('SATTENTE00000',GetParamSocSecur('SO_DEFCOLSAL',''),'SAL','SO_SALATTEND',St1,35) ; VH^.TiersDefSal:=St1 ;
  St1:=VH^.TiersDefDiv ; CreeCptAuxParamsoc('DATTENTE00000',GetParamSocSecur('SO_DEFCOLDIV',''),'DIV','SO_DIVATTEND',St1,36) ; VH^.TiersDefDiv:=St1 ;

  If (Trim(VH^.Cpta[fbAxe1].Attente)='') Then
    Cpt:=BourreEtLess('9999999999999999',fbAxe1)
  Else Cpt:=Trim(VH^.Cpta[fbAxe1].Attente) ;
  Q:=OpenSQL('SELECT * FROM SECTION WHERE S_SECTION ="'+Cpt+'" ',FALSE) ;
  If Q.Eof Then
  BEGIN
    Q.Insert ; InitNew(Q) ;
    Q.FindField('S_SECTION').AsString:=Cpt ;
    Q.FindField('S_LIBELLE').AsString:=HM.Mess[24] ;
    Q.FindField('S_ABREGE').AsString:=Copy(HM.Mess[24],1,8) ;
    Q.FindField('S_AXE').AsString:='A1' ;
    Q.FindField('S_CREERPAR').AsString:='IMP' ;
    Q.FindField('S_SENS').AsString:='M' ;
    Q.Post ;
  END Else Cpt:=Q.FindField('S_SECTION').AsString ;
  Ferme(Q) ;
  VH^.Cpta[fbAxe1].Attente:=Trim(Cpt) ;
  MajParamSocDesImmos;
END ;

Function TFImpauto.RecupMouvement(InfoImp :PtTInfoImport) : Integer ;
Var SauveNom : String ;
    PbVerif : Boolean ;
    Ens : SetOfByte ;
    k : Integer ;
BEGIN
Result:=0 ; Ens:=[] ;
InfoImp^.ProfilPCL:=ProfilPCL ;
TransfertSISCO(InfoImp^.NomFicMvt,FALSE,k,'',InfoImp) ;
If InfoImp^.PbEnr[0] Then BEGIN Result:=1 ; Ens:=Ens+[19] ; END ; 
If result<>0 Then
  BEGIN
  (*
  EcrireRapportGlobal(InfoImp^,TRUE,'',TRUE) ;
  If Result=1 Then EcrireRapportGlobal(InfoImp^,FALSE,HM.Mess[19],FALSE) ;
  CreateTraceGlobal(InfoImp^,3) ;
  *)
  TraiteErreurGlobable(Ens,3,InfoImp^) ;
  Exit ;
  END ;
SauveNom:=InfoImp^.NomFic ; InfoImp^.NomFic:=InfoImp^.NomFicMvtCGE ;
RepositionneNomFic(InfoImp^) ;
//ImporteLesEcritures(nil,InfoImp^) ;
PbVerif:=FALSE ;
If ImporteLesEcritures(nil,InfoImp^) then
  BEGIN
  InitVerif(InfoImp^) ; PbVerif:=FALSE ;
//      if not VerPourImp2(InfoImp^.ListeEntetePieceFausse,InfoImp^.ListePieceFausse,TRUE,tvEcr) then
(*
  if not VerPourImp3(InfoImp.ListeEntetePieceFausse,InfoImp.ListePbAna,InfoImp.ListePieceFausse,TRUE,tvEcr,InfoImp.SC.ShuntPbAna,InfoImp.PbAna) then
    BEGIN
    MajImpErr(InfoImp.ListeEntetePieceFausse) ;
    PbVerif:=TRUE ;
    END ;
*)
//     if AucunMvtControlesOk then BEGIN HM.Execute(42,'','') ; Exit ; END ;
  If Not PbVerif Then
    BEGIN
    Affiche(4) ;
    InitIntegre(InfoImp^) ;
    IntegreEcr(Nil,InfoImp^) ;
    END ;
  END else VH^.ImportRL:=FALSE ;
InfoImp^.NomFic:=SauveNom ;
EcrireRapportGlobal(InfoImp^,PbVerif) ;
FaitFichierDoublon(InfoImp^) ; FaitFichierRejet(InfoImp^) ; DetruitFichier(InfoImp^) ;
MessImport.Visible:=FALSE;
END ;

procedure TFImpauto.LanceImportAuto ;
var Soc,OldSoc : String ;
    InfoImp : PtTInfoImport ;
    Qui : String ;
    Auto,PlusDeFichierAImporter,OkFic : Boolean ;
    ListeFichier : TStringList ;
    LCodeSoc : HtStringList ;
    Err,Err1: Integer ;
    OkOk,PbVerif,PremierAppel : Boolean ;
    Ens : SetOfByte ;
    OkSection : boolean;
    TImp : TOB;
    Nr   : string;
    Q    : TQuery;
  Label 0 ;
begin
  InitMessage ;
  ListeFichier:=NIL ;
  LCodeSoc:=Nil ;
  OkSection := False;
  PlusDeFichierAImporter:=TRUE ;
  Auto:=TRUE ;
  Init ;
  if DBSOC<>NIL Then DeconnecteHalley ;
  New(InfoImp) ;
  FillChar(InfoImp^,SizeOf(InfoImp^),#0) ;
  CreateListeImp(InfoImp^) ;
  If Not InitEnv(InfoImp^,REPPCL,MasquePCL) Then
  begin
    HM.Execute(3,Caption,'') ; ; Goto 0 ;
  end;
  VideRepertoire(InfoImp^) ;
  If InfoImp^.EE.Exe<>'' Then
  begin
    MessSoc.Visible:=TRUE ;
    MessSoc.Caption:=HM.Mess[2] ;
    FileExec(InfoImp^.EE.Exe,FALSE,TRUE) ;
    Delay(10000) ;
  end;
  ListeFichier:=TStringList.Create ; ListeFichier.Sorted:=TRUE ; ListeFichier.Duplicates:=dupIgnore ;
  InfoImp^.Sc.Chemin:='' ; InfoImp^.Sc.Prefixe:='' ; InfoImp^.Sc.Suffixe:='' ;
  MessSoc.Visible:=TRUE ;
  MessSoc.Caption:=HM.Mess[4] ;
  RecupNomFic(InfoImp^,ListeFichier) ;
  OldSoc:='' ;
  MessSoc.Visible:=TRUE ;
  MessSoc.Caption:=HM.Mess[5] ;
  If MultiSoc Then
  BEGIN
  LCodeSoc := HTStringList.Create ; LCodeSoc.Sorted:=TRUE ; ListeFichier.Duplicates:=dupIgnore ;
  ChargeLCodeSoc('SO_LIBELLE',LCodeSoc) ;
  END ;
InfoImp^.NumFic:=0 ;
CreateLeFicRapportGlobal(InfoImp^) ;
EBStop.Visible:=TRUE ; VH^.STOPRSP:=FALSE ;
Repeat
  InitMessage ;
  YY.Caption:=IntToStr(ListeFichier.Count) ;
  OkFic:=FichierSuivant(ListeFichier,InfoImp^,InfoImp^.NumFic) ;
  XX.Caption:=IntToStr(InfoImp^.NumFic) ;
//  OkFic:=TRUE ;
  If OkFic Then
    BEGIN
    If SOCPCL<>'' Then Soc:=SocPCL Else If MultiSoc Then Soc:=LitLCodeSoc(LCodeSoc,InfoImp^.CodeSoc) ;
    If (Soc<>OldSoc) And (Soc<>'') Then
      BEGIN
      MessSoc.Visible:=TRUE ;
      If VH^.FromPCL Then MessSoc.Caption:=HM.Mess[6]+' "'+V_PGI.NoDossier+'"'
                     Else MessSoc.Caption:=HM.Mess[6]+' "'+Soc+'"' ;
//      If Not PCLACTIF Then V_PGI.UserName:=InfoImp.EE.User ;
      if DBSOC<>NIL Then BEGIN Logout ; DeconnecteHalley ; END ;
      if ConnecteHalley(Soc,TRUE,@ChargeMagHalley,NIL,NIL,NIL) then
         BEGIN
         ApresConnecte ;
         END else
         BEGIN
         If (DBSOC<>NIL) AND (DBSOC.Connected) then DBSOC.Connected:=FALSE ;
         SourisNormale ; Exit ;
         END ;
      If Not AlimInfoImp(InfoImp^) Then Goto 0 ;
      END ;
    If VH^.FromPCL Then MessSoc.Caption:=HM.Mess[10]+': "'+V_PGI.NoDossier+'"'
                   Else MessSoc.Caption:=HM.Mess[10]+': "'+Soc+'"' ;
    OldSoc:=Soc ;
    PlusDeFichierAImporter:=FALSE ; PremierAppel:=TRUE ;
    Affiche(1) ;
    Un2DeuxSISCO(InfoImp^) ;
    If VH^.StopRSP Then Goto 0 ;
    If Not RecupInfoDossier(InfoImp^) Then Goto 0 ;
    If VH^.StopRSP Then Goto 0 ;
    Affiche(2) ;
    Err1:=RecupFicheBase(InfoImp) ;
    If VH^.StopRSP Then Goto 0 ;
    If Err1<>0 Then Goto 0 ;
    If PremierAppel Then FiniRecupInfoDossierAuto(InfoImp^) ;
    If VH^.StopRSP Then Goto 0 ;
    Affiche(3) ;
    Err1:=RecupMouvement(InfoImp) ;
    If VH^.StopRSP Then Goto 0 ;
    If Err1<>0 Then Goto 0 ;
    Affiche(5) ;
    PremierAppel:=FALSE ;
//    If Not RepositionneNomFic(InfoImp^) Then Goto 0 ;
    _Stat := LireProfilInfoStat ( ProfilPCL);
    OkSection := TraiteCodeStatSISCO(_Stat,InfoImp^, True) or  OkSection;
    If VH^.StopRSP Then Goto 0 ;
    VideListeInfoImp(InfoImp^,FALSE) ;
    END Else PlusDeFichierAImporter:=TRUE ;
Until PlusDeFichierAImporter ;
if OkSection then TraiteCodeStatSISCO(_Stat,InfoImp^, False) ;
TRSL.Visible:=FALSE ; XX.Visible:=FALSE ; YY.Visible:=FALSE ;TR.Caption:=HM.Mess[7] ;
Err1:=MajTotTousComptes(FALSE,'') ;
If Err1<>0 Then
  BEGIN
  Ens:=[] ;
  If Err1 >1000 Then BEGIN Ens:=Ens+[23] ; Err1:=Err1 Mod 1000 ; END ;
  If Err1 >100 Then BEGIN  Ens:=Ens+[22] ; Err1:=Err1 Mod 100 ; END ;
  If Err1 >10 Then BEGIN  Ens:=Ens+[21] ; Err1:=Err1 Mod 10 ; END ;
  If Err1 >0 Then Ens:=Ens+[20] ;
  TraiteErreurGlobable(Ens,2,InfoImp^) ;
  Nr := 'NRE';
  Goto 0 ;
  END
else Nr := 'ROK';
// update de trfs5 ajout me  Pour majlot
  if (VH^.RecupPCL) and (ctxPCL in V_PGI.PGIContexte)  then
  begin
        Q := OpenSQl ('SELECT * FROM TRFFICHIER Where TRF_FICHIER like "%' +ExtractFileName(InfoImp^.NomFic)+ '%" AND '+
        'TRF_REPERTOIRE="'+ RepPCl +'"', TRUE);
        if not Q.EOF then
        begin
             TImp :=TOB.Create('',Nil,-1) ;
             TImp.LoadDetailDB('TRFFICHIER', '', '', Q, TRUE, FALSE);
             TImp.detail[0].PutValue ('TRF_ETATREST', Nr);
             TImp.InsertOrUpdateDB(True);
             TImp.free;
        end;
        Ferme(Q);
  end;
CreateTraceGlobal(InfoImp^,0) ;
If Not MultiSoc Then AfficheRapportGlobal(InfoImp^) ;

//--- Suppression des monnaie 'In'  ajout me 17-08-2005
ExecuteSQL ('DELETE FROM DEVISE WHERE D_MONNAIEIN="X" AND D_FONGIBLE="-"');

SetParamSoc('SO_EXOV8',VH^.EnCours.Code) ;
if (VH^.RecupPCL) and (ctxPCL in V_PGI.PGIContexte) then CreateJalANO ;

// ajout me 19-08-2005 pour la création souche CPT par défaut
VerifUneSouche('CPT',FALSE,FALSE) ;

0:If ListeFichier<>NIL Then ListeFichier.Free ; If MultiSoc Then If LCodeSoc<>NIL Then LCodeSoc.Free ;
VideListeInfoImp(InfoImp^,TRUE) ;
ExecuteSQL('DELETE FROM IMPECR') ;
Dispose(InfoImp) ;
if OkSituation(ProfilPCL) then RestaureBalSit;
if DBSOC<>NIL Then
BEGIN
        if (V_PGI.NoDossier<>'') and  (V_PGI.Driver in [dbMssql, dbMssql2005]) then
        begin
          try
            ExecuteSql ('checkpoint');
            ExecuteSql ('backup log '+v_pgi.DBName+' with no_log');
            ExecuteSql ('dbcc shrinkdatabase('+v_pgi.DBName+')');
            ExecuteSql ('dbcc shrinkfile(2)');
          except;
          end;
        end;

     Logout ; DeconnecteHalley ;
END ;
SourisNormale ;
END ;

procedure TFImpauto.RestaureBalSit;
var ImpBDS : TImpExpBDS;
    i : integer;
    StFichier : string;
begin
  // Test existence du fichier de balances
  StFichier := RepPCL+'\'+Copy(MasquePCL,1,8)+'.DAD';
  if not FileExists(StFichier) then exit;

  // Inialisation de l'affichage
  L1.Visible := False;Ok1.Visible := False;  AF1.Visible := False;
  Ok2.Visible := False;  AF2.Visible := False;
  L3.Visible := False; Ok3.Visible := False;  AF3.Visible := False;
  L4.Visible := False; Ok4.Visible := False;  AF4.Visible := False;
  TR.Caption:=TraduireMemoire('Balances de situation');

  // Raz des balances de situation existantes
  ExecuteSQL('DELETE FROM CBALSIT WHERE BSI_CODEBAL LIKE "SISCO%" ') ;
  ExecuteSQL('DELETE FROM CBALSITECR WHERE BSE_CODEBAL LIKE "SISCO%" ') ;

  // Restauration des balances de situation
  ImpBDS := TImpExpBDS.Create ( miSiscoII ) ;
  try
    ImpBDS.OnInformation := OnInformationBalSit;
    ImpBDS.Importation ( StFichier , 'SISCO', ' Balance SISCO récupérée' ,'SISCO' ,TRUE);
{    if ImpBDS.GetLastError = 0 then  // Tout est OK
    else ;}
  finally
    ImpBDS.Free;
  end;
end;

procedure TFImpauto.OnInformationBalSit( Sender : TObject; Msg : string ; bErr : boolean);
begin
  L2.Caption := Msg;
  Application.ProcessMessages;
end;

procedure TFImpauto.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled:=FALSE ;
  LanceImportAuto ;
  Close ;
end;

procedure RenseigneSerie ;
var Buffer: array[0..1023] of Char;
    sWinPath,sIni : String ;
begin
HalSocIni:='CCS7.ini' ;
GetWindowsDirectory(Buffer,1023);
SetString(sWinPath, Buffer, StrLen(Buffer));
sIni:=sWinPath+'\'+HALSOCINI ;
if FileExists(sIni) then
   BEGIN
   HalSocIni:='CCS7.ini' ;
   NomHalley:= 'Comptabilité S7' ;
   TitreHalley := 'Imports / Exports S7' ;
   V_PGI.LaSerie:=S7 ;
   Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CCS7.HLP' ;
   END else
   BEGIN
   HalSocIni:='CCS5.ini' ;
   NomHalley:= 'Comptabilité S5' ;
   TitreHalley := 'Imports / Exports S5' ;
   V_PGI.LaSerie:=S5 ;
   Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CCS5.HLP' ;
   END ;
END ;

procedure TFImpauto.FormCreate(Sender: TObject);
begin
  VH^.FromPCL:=TRUE ;
  If (ParamCount>=1) And (ParamStr(1)='MDOFF') Then VH^.FromPCL:=FALSE ;
  VH^.RecupPCL:=TRUE ;
  RecupSISCO:=TRUE ;
  MultiSoc:=FALSE ;
  if V_PGI.CegidBureau then
  begin
    PGIAppAlone:=True ;
    CreatePGIApp ;
  end;
end;

procedure TFImpauto.EBStopClick(Sender: TObject);
begin
If Not VH^.STOPRSP Then
  BEGIN
  If PGIAsk('Confirmez-vous l''arrêt du traitement en cours ?',Self.Caption)=mrYes then
  BEGIN
    VH^.STOPRSP:=TRUE ;
    Application.ProcessMessages ;
  End ;
  END ;
end;

procedure TFImpauto.CreeCptAuxParamsoc(LeCpt, LeColl, NatCpt,
  ChampParamSoc1: String; var ChampVH1: String; Ind: Integer);
Var Cpt : String ;
    Q : tQuery ;
BEGIN
Cpt:=BourreEtLess(LeCpt,fbAux) ;
Q:=OpenSQL('SELECT * FROM TIERS WHERE T_AUXILIAIRE ="'+Cpt+'" ',FALSE) ;
If Q.Eof Then
  BEGIN
  Q.Insert ; InitNew(Q) ;
  Q.FindField('T_AUXILIAIRE').AsString:=Cpt ;
  Q.FindField('T_LIBELLE').AsString:=HM.Mess[Ind] ;
  Q.FindField('T_ABREGE').AsString:=Copy(HM.Mess[Ind],1,8) ;
  Q.FindField('T_NATUREAUXI').AsString:=NatCpt ;
  Q.FindField('T_CREERPAR').AsString:='IMP' ;
  Q.FindField('T_COLLECTIF').AsString:=LeColl ;
  Q.FindField('T_MODEREGLE').AsString:='002' ;
  Q.FindField('T_REGIMETVA').AsString:='FRA' ;
  Q.FindField('T_TVAENCAISSEMENT').AsString:='TD' ;
  Q.FindField('T_TIERS').AsString:=Cpt ;
  Q.FindField('T_LETTRABLE').AsString:='-' ;
  Q.Post ;
  END Else Cpt:=Q.FindField('T_AUXILIAIRE').AsString ;
Ferme(Q) ;
If ChampParamSoc1<>'' Then SetParamSoc(ChampParamSoc1,Trim(Cpt)) ;
// If ChampVH1<>'' Then ChampVH1:=Trim(Cpt) ; CA - 29/01/2003 Etonnant !
If ChampVH1='' Then ChampVH1:=Trim(Cpt) ;
END ;

procedure TFImpauto.CreeCptCollParamsoc;
Var CollCli,CollFou,CollSAl,CollDiv : String ;
    TC,TobL : TOB ;
    Cpt,Nat : String ;
    Q : tQuery ;
    St,St1 : String ;
    i : Integer ;
BEGIN
CollCli:='' ; CollFou:='' ; CollSAl:='' ; CollDiv:='' ;
TC:=TOB.Create('',Nil,-1) ;
St:='SELECT G_GENERAL,G_NATUREGENE FROM GENERAUX WHERE G_COLLECTIF="X" ' ;
Q:=OpenSQL(St,TRUE) ;
TC.LoadDetailDB('CPT','','',Q,False,True) ;
Ferme(Q) ;
For i:=0 To  TC.Detail.Count-1 Do
  BEGIN
  TOBL:=TC.Detail[i] ;
  If TOBL.GetValue('G_NATUREGENE')='COC' Then BEGIN If CollCli='' Then CollCli:=TOBL.GetValue('G_GENERAL') ; END Else
  If TOBL.GetValue('G_NATUREGENE')='COF' Then BEGIN If CollFou='' Then CollFou:=TOBL.GetValue('G_GENERAL') ; END Else
  If TOBL.GetValue('G_NATUREGENE')='COS' Then BEGIN If CollSal='' Then CollSal:=TOBL.GetValue('G_GENERAL') ; END Else
  If TOBL.GetValue('G_NATUREGENE')='COD' Then BEGIN If CollDiv='' Then CollDiv:=TOBL.GetValue('G_GENERAL') END ;
  END ;
TC.Free ;
If CollCli='' Then
  BEGIN
  St1:=VH^.DefautCli ;
  CreeCptGenParamsoc('4119999999999','COC','SO_DEFCOLCLI','',St1,29) ;
  VH^.DefautCli:=St1 ;
  END Else
  BEGIN
  SetParamSoc('SO_DEFCOLCLI',Trim(CollCli)) ; VH^.DefautCli:=Trim(CollCli) ;
  END ;
If CollFou='' Then
  BEGIN
  St1:=VH^.DefautFou ;
  CreeCptGenParamsoc('4019999999999','COF','SO_DEFCOLFOU','',St1,30) ;
  VH^.DefautFou:=St1 ;
  END Else
  BEGIN
  SetParamSoc('SO_DEFCOLFOU',Trim(CollFou)) ; VH^.DefautFou:=Trim(CollFou) ;
  END ;
If CollSal='' Then
  BEGIN
  St1:=VH^.DefautSal ;
  CreeCptGenParamsoc('4219999999999','COS','SO_DEFCOLSAL','',St1,31) ;
  VH^.DefautSal:=St1 ;
  END Else
  BEGIN
  SetParamSoc('SO_DEFCOLSAL',Trim(CollSal)) ; VH^.DefautSal:=Trim(CollSal) ;
  END ;
If CollDiv='' Then
  BEGIN
  St1:=VH^.DefautDivers ;
  CreeCptGenParamsoc('4719999999999','COD','SO_DEFCOLDIV','',St1,32) ;
  VH^.DefautDivers:=St1 ; 
  END Else
  BEGIN
  SetParamSoc('SO_DEFCOLDIV',Trim(CollDiv)) ; VH^.DefautDivers:=Trim(CollDiv) ;
  END ;
END ;

procedure TFImpauto.CreeCptGenParamsoc(LeCpt, NatCpt, ChampParamSoc1,
  ChampParamSoc2: String; var ChampVH1: String; Ind: Integer);
Var Cpt : String ;
    Q : tQuery ;
BEGIN
Cpt:=BourreEtLess(LeCpt,fbGene) ;
Q:=OpenSQL('SELECT * FROM GENERAUX WHERE G_GENERAL ="'+Cpt+'" ',FALSE) ;
If Q.Eof Then
  BEGIN
  Q.Insert ; InitNew(Q) ;
  Q.FindField('G_GENERAL').AsString:=Cpt ;
  Q.FindField('G_LIBELLE').AsString:=HM.Mess[Ind] ;
  Q.FindField('G_ABREGE').AsString:=Copy(HM.Mess[Ind],1,17) ;
  Q.FindField('G_NATUREGENE').AsString:=NatCpt ;
  If (NatCpt='COC') Or (NatCpt='COF') Or (NatCpt='COS') Or (NatCpt='COD') Then Q.FindField('G_COLLECTIF').AsString:='X' ;
    Q.FindField('G_CREERPAR').AsString:='IMP' ;
  Q.FindField('G_SENS').AsString:='M' ;
  Q.Post ;
  END Else Cpt:=Q.FindField('G_GENERAL').AsString ;
Ferme(Q) ;
If ChampParamSoc1<>'' Then SetParamSoc(ChampParamSoc1,Trim(Cpt)) ;
If ChampParamSoc2<>'' Then SetParamSoc(ChampParamSoc2,Trim(Cpt)) ;
// If ChampVH1<>'' Then ChampVH1:=Trim(Cpt) ; CA - 29/01/2003 . Etonnant !
If ChampVH1='' Then ChampVH1:=Trim(Cpt) ; 
END ;

Procedure InitLaVariablePGI;
Begin
  RenseigneLaSerie(ExeCCAUTO) ;

  Apalatys:='CEGID' ;
  Copyright:='© Copyright ' + Apalatys ;
  V_PGI.OutLook:=FALSE ;
  V_PGI.OfficeMsg:=FALSE ;
  V_PGI.VersionDemo:=FALSE ;
  V_PGI.SAV:=True ;
  V_PGI.BlockMAJStruct:=True ;
  V_PGI.EuroCertifiee:=False ;

  ChargeXuelib ;

  V_PGI.MenuCourant:=0 ;
  V_PGI.ImpMatrix := True ;
  V_PGI.OKOuvert:=FALSE ;
  V_PGI.Halley:=TRUE ;
  V_PGI.LaSerie:=S5 ;
  V_PGI.ParamSocLast:=True ;
  V_PGI.PGIContexte:=[ctxCompta, ctxPCL] ;
  V_PGI.CegidAPalatys:=FALSE ;
  V_PGI.CegidBureau:=TRUE ;
  V_PGI.StandardSurDP:=True ;
  V_PGI.MajPredefini:=False ;
end;

Initialization
  ProcChargeV_PGI :=  InitLaVariablePGI ;

end.

