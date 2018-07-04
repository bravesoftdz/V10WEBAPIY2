unit MainVerif;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, ComCtrls, HStatus, DBTables, Ent1, HEnt1,
  HCtrls, DB, MajTable, LicUtil, ed_tools, StdCtrls,
  HMsgBox, PGIEnv, PGIExec, UTOB, ParamSoc, HFLabel, FE_MAIN,
  ZEcriture,
  SaisUtil, Buttons;

type
  SetOfByte = set of Byte;

type
  TFMainVerif = class(TForm)
    Status: THStatusBar;
    Timer1: TTimer;
    LDossier: TLabel;
    Label2: TLabel;
    Animate: TAnimate;
    LTraitement: TLabel;
    BAnnuler: TBitBtn;
    procedure FormShow(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
  private
    { Déclarations privées }
    SocPCL: string;
    RepPCl, MasquePCL, MonnaiePCL: string;
    AnaPCL: Boolean;
    ProfilPCL: string;
    MultiSoc: Boolean;

    procedure LanceTraitement;
    procedure ApresConnecte;

    function GetFiltreFromDBCom( vNomFiltre, vLibelleFiltre : String ) : String ;
    procedure ExecuteTraitement ;
    procedure EnvoitLaSauce( vNomTraitement : String );

  public
    { Déclarations publiques }
  private
    FCleFiltre : String ;
    FNomFiltre : String ;
    ZEcriture : TZEcriture;
    FBoStop : Boolean;
  end;

var
  FMainVerif: TFMainVerif;

implementation

uses VerCpte,     // Vérification des comptes
     VerCpta,     // Vérification des mouvements
     uLibWindows, // GetWindowsTempPath
     RapSuppr,    // RapportErreurEnSerie
     uControlCP,
     SoldeCpt,    // MajTotTousComptes
     VerDPaq,     // DatePaquet
     ReparMvt,    // RepareMvt
     VerContr;    // VerifContreparties

{$R *.DFM}

procedure TFMainVerif.FormShow(Sender: TObject);
var
  i: Integer;
  St, St1, Nom, Value: string;
begin
  FBoStop := False;

  VStatus := Status;
  Status.Caption := Copyright;
  VH^.ModeSilencieux := TRUE;
  SocPCL := '';
  RepPCl := '';
  MasquePCL := '';
  MonnaiePCL := '';
  ProfilPCL := '000';
  if (not V_PGI.CegidApalatys) then
  begin
    // Création de l'instance unique
    V_PGI_Env := TPGIEnv.Create;
    //# lit param de PGIApp et màj V_PGI.RunFromLanceur
    InitPGIEnv();
    //# retour au mode APA
    if V_PGI_Env.ModeFonc = 'APA' then
    begin
      V_PGI_Env.Free;
      V_PGI_Env := nil;
    end
      //# routage halsocini
    else if Copy(V_PGI_Env.ModeFonc, 1, 1) = 'M' then
      HalSocIni := 'CEGIDPGI.INI';
  end;
  //# évite message bloquant
  if V_PGI.RunFromLanceur then
    V_PGI.MultiUserLogin := True;

  V_PGI.UserName := '';
  V_PGI.PassWord := '';
  SocPCL := '';
  for i := 1 to ParamCount do
  begin
    St := ParamStr(i);
    Nom := UpperCase(Trim(ReadTokenPipe(St, '=')));
    Value := UpperCase(Trim(St));
    
    //Paramètres de connexion
    if Nom = '/MAJSTRUCTURE' then
    begin
     (*kVerifStructure:=(Value<>'FALSE') ;*)
    end;

    if Nom = '/USER' then
    begin
      V_PGI.UserName := Value;
    end;

    if Nom = '/PASSWORD' then
    begin
      V_PGI.PassWord := DecryptageSt(Value);
    end;

    if Nom = '/DATE' then
    begin
      V_PGI.DateEntree := StrToDate(Value);
    end;

    if Nom = '/DOSSIER' then
    begin
      SocPCL := Value;
    end;

    if Nom = '/TRF' then
    begin
      if Value[Length(Value)] <> ';' then
        Value := Value + ';';
      St1 := ReadTokenSt(Value);
      if St1 <> '' then
        RepPCL := St1;
      St1 := ReadTokenSt(Value);
      if St1 <> '' then
        MasquePCL := St1;
      St1 := ReadTokenSt(Value);
      if St1 <> '' then
        MonnaiePCL := St1;
      St1 := ReadTokenSt(Value);
      if St1 <> '' then
        AnaPCL := St1 = 'O'; // Env.Nodossier
      St1 := ReadTokenSt(Value);
      if St1 <> '' then
        ProfilPCL := St1; // Env.Nodossier
    end;

    // GC - 07/02/2002
    // Paramètre de correction ou de vérification des dossiers Compta PGI
    if (Nom = '/CPVERIF') then
    begin
      //pgiinfo(Value,'');
      FCleFiltre := cFI_TableVerif;
      FNomFiltre := Value;
    end;

    if (Nom = '/CPCORRECT') then
    begin
      FCleFiltre := cFI_TableCorrect;
      FNomFiltre := Value;
    end;

    LDossier.Caption := TraduireMemoire('Dossier') + ' : ' + V_Pgi_Env.NoDossier + ' - ' + V_Pgi_Env.LibDossier;
    // Fin - GC

  end;
  V_PGI.PassWord := CryptageSt(DayPass(Date));
end;


procedure TFMainVerif.ApresConnecte;
begin
  if VH^.FromPCL then
  begin
    if V_PGI_Env <> nil then
    begin
      //# mono-entreprise : pointe l'environnement de la soc choisie
      if V_PGI_Env.ModeFonc = 'MONO' then
        V_PGI_Env.SocCommune := SOCPCL;
      //# mode multi-dossier : appli lancée sans passer par le lanceur...
      if (V_PGI_Env.ModeFonc = 'MULTI') and (not V_PGI.RunFromLanceur) then
        //manque le double choix : soc commune, dossier
        PGIInfo('Vous lancez une application en mode multi-dossier sans passer par le lanceur.', TitreHalley);
      //# mode multi-dossier : on est dans le lanceur => connecter la soc...
    end;
    if ((V_PGI_Env <> nil) and (V_PGI_Env.ModePCL = '1')) then
      V_PGI.PGIContexte := V_PGI.PGIContexte + [ctxPCL];
  end;
end;

procedure init;
begin
  V_PGI.Debug := FALSE;
  V_PGI.Versiondev := FALSE;
  V_PGI.Synap := FALSE;
  VH^.GrpMontantMin := 0;
  VH^.GrpMontantMax := 1000000;
  V_PGI.DateEntree := Date;
  VH^.Mugler := FALSE;
  V_PGI.Halley := TRUE;
  V_PGI.NumVersion := '3';
  V_PGI.SAV := TRUE;
  V_PGI.Versiondev := FALSE;
  V_PGI.Synap := FALSE;
  VH^.GereSousPlan := True;
  V_PGI.Halley := TRUE;
end;

procedure TFMainVerif.LanceTraitement;
label
  0;
begin
  if DBSOC <> nil then DeconnecteHalley;

  if DBSOC <> nil then
  begin
    Logout;
    DeconnecteHalley;
  end;

  if ConnecteHalley(SocPCL, TRUE, @ChargeMagHalley, nil, nil, nil) then
  begin
    ApresConnecte;
    ExecuteTraitement;
  end
  else
  begin
    if (DBSOC <> nil) and (DBSOC.Connected) then DBSOC.Connected := FALSE;
    SourisNormale;
    Exit;
  end;

  if DBSOC <> nil then
  begin
    Logout;
    DeconnecteHalley;
  end;
  SourisNormale;
end;

procedure TFMainVerif.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := FALSE;
  LanceTraitement;
  Close;
end;

procedure TFMainVerif.FormCreate(Sender: TObject);
begin
  VH^.FromPCL := TRUE;
  if (ParamCount >= 1) and (ParamStr(1) = 'MDOFF') then
    VH^.FromPCL := FALSE;
  VH^.RecupPCL := TRUE;
  MultiSoc := FALSE;
  if V_PGI.CegidBureau then
  begin
    PGIAppAlone := True;
    CreatePGIApp;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilmles COSTE
Créé le ...... : 07/02/2002
Modifié le ... : 07/02/2002
Description .. : Récupère les informations du filtre passé en paramètre dans
Suite ........ : la Base DB000000, alors que l' on est connecté à la base
Suite ........ : du Dossier
Mots clefs ... :
*****************************************************************}
function TFMainVerif.GetFiltreFromDBCom(vNomFiltre, vLibelleFiltre: String): String;
var lCom : TQuery;
begin
  Result := '';

  V_PGI_Env.Connected := True;

  lCom := TQuery.Create (nil);
  try
    lCom.DatabaseName := 'DBCOM';
    lCom.SQL.Text := 'SELECT FI_CRITERES FROM FILTRES WHERE FI_TABLE="' + vNomFiltre + '" AND FI_LIBELLE="' + vLibelleFiltre + '"';
    ChangeSQL(lCom);
    lCom.RequestLive := True;
    lCom.Open;
    if not lCom.IsEmpty then
    begin
      Result := LCom.Findfield('FI_CRITERES').AsString;
    end;
    lCom.Close;
  finally
    lCom.Free;
  end;

  V_PGI_Env.Connected := false;
end;

procedure TFMainVerif.ExecuteTraitement;
var lListeCritere : TStringList;
    i : Integer;
    lNomTraitement : String;
    lValeur : Boolean;
    lSqlPiece, lSqlEcr : TQuery;
    lTobPiece : Tob;
    lTobmEcr : Tobm;
begin
  lTobPiece := nil;
  lSqlEcr := nil;

  // Récupère le filtre spécifié en parammètre dans la DB000000 et Charge FListeCritere;
  lListeCritere := TStringList.Create;
  lListeCritere.Text := GetFiltreFromDBCom(FCleFiltre, FNomFiltre);

  {$IFDEF DEBUG}
    Showmessage( lListecritere.text );
  {$ENDIF}  

  try
    // Création du ZEcriture pour la vérification des écritures
    ZEcriture := TZEcriture.Create;
    ZEcriture.FBoParle := False;

    if lListeCritere.Count <> 0 then
    begin
      for i := 0 to lListeCritere.Count - 1 do
      begin
        if FBoStop then Exit;

        lNomTraitement := Copy( lListeCritere.Strings[i], 0, Pos(';',lListeCritere.Strings[i])-1);
        lValeur := Copy( lListeCritere.Strings[i], Pos(';',lListeCritere.Strings[i])+1, 1) = '1';
        if lValeur then EnvoitLaSauce(lNomTraitement);
      end;

      if Length( ZEcriture.FZEcritureCtx ) > 0 then
      begin
        lSqlPiece := OpenSql('select e_etablissement, '+
                             'e_exercice, '+
                             'e_journal, '+
                             'e_periode, ' +
                             'e_numeropiece from ecriture, exercice ' +
                             'where (e_qualifpiece = "N") and (e_exercice = ex_exercice) '+
                             'and (ex_etatcpta = "OUV") group by e_etablissement, e_exercice, e_journal, e_periode, e_numeropiece', True);

        lTobPiece := Tob.Create('', nil, -1);
        lTobPiece.LoadDetailDB('', '', '', lSqlPiece, False);
        Ferme( lSqlPiece );

        for i := 0 to lTobPiece.Detail.Count - 1 do
        begin
          try
             lSqlEcr := OpenSql('Select e_datecomptable, e_periode, e_semaine, e_numeropiece, ' +
                               'e_numligne, e_journal, e_refinterne, e_numeche, e_qualifpiece, ' +
                               'e_datemodif, e_paquetrevision, e_exercice, e_modesaisie, ' +
                               'e_ecranouveau, e_general, e_auxiliaire, e_etatlettrage, e_eche, e_ana from ecriture where ' +
                               'e_qualifpiece = "N" and ' +
                               'e_etablissement = "' + lTobPiece.Detail[i].GetValue('E_ETABLISSEMENT') + '" and ' +
                               'e_exercice = "' + lTobPiece.Detail[i].GetValue('E_EXERCICE') + '" and ' +
                               'e_journal = "' + lTobPiece.Detail[i].GetValue('E_JOURNAL') + '" and ' +
                               'e_periode = ' + IntToStr(lTobPiece.Detail[i].GetValue('E_PERIODE')) + ' and ' +
                               'e_numeropiece = ' + IntToStr(lTobPiece.Detail[i].GetValue('E_NUMEROPIECE')) + ' ' +
                               'order by e_etablissement, e_exercice,e_journal,e_periode, e_numeropiece', True);
          except
            on e: Exception do showmessage( E.message );
          end;

          if not lSqlEcr.IsEmpty then
          begin
            ZEcriture.ClearDetail;
            while not lSqlEcr.Eof do
            begin
              lTobmEcr := Tobm.Create( EcrGen, '', True, ZEcriture );
              lTobmEcr.SelectDB('', lSqlEcr, False);
              lSqlEcr.Next;
            end;
            Ferme( lSqlEcr );
            ZEcriture.Execute;
          end;
        end;
      end;
    end;

  finally
    FreeNil(lListeCritere);
    FreeNil(lTobPiece);
    FreeNil(ZEcriture);
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilmles COSTE
Créé le ...... : 07/02/2002
Modifié le ... : 07/02/2002
Description .. : Evoit le traitment correspondant en fonction du paramètre d'
Suite ........ : entrée
Suite ........ : NB : Cela peut être une vérification ou une correction
Mots clefs ... :
*****************************************************************}
procedure TFMainVerif.EnvoitLaSauce( vNomTraitement : String );
var NbNoPossible : Integer;
begin
  NbNoPossible := 0;

  // Fonction de vérification des dossiers Compta PGI
  if FCleFiltre = cFI_TableVerif then
  begin

    // VERIFICATION DES ERREURS SUR LES FICHES
    if vNomtraitement = 'FCBGENERAUX' then
    begin
      lTraitement.Caption := TraduireMemoire('Vérification des comptes généraux');
      VerCompteMAJ(1, NbNoPossible, True);
      lTraitement.Caption := TraduireMemoire('Vérification de la nature des comptes 6 et 7');
      ControleCpt67( False );
    end;

    if vNomTraitement = 'FCBAUXILIAIRES' then
    begin
      lTraitement.Caption := TraduireMemoire('Vérification des comptes auxiliaires');
      VerCompteMAJ(2, NbNoPossible, True);
    end;

    if VNomTraitement = 'FCBSECTIONS' then
    begin
      lTraitement.Caption := TraduireMemoire('Vérification des sections');
      VerCompteMaj(3, NbNoPossible, True);
    end;

    if vNomtraitement = 'FCBJOURNAUX' then
    begin
      lTraitement.Caption := TraduireMemoire('Vérification des journaux');
      VerCompteMaj(4, NbNoPossible, True);
    end;

    // VERIFICATION DES ERREURS SUR LES MOUVEMENTS COMPTABLES
    if vNomTraitement = 'FCBMOUVEMENTS' then
    begin
      lTraitement.Caption := TraduireMemoire('Vérification des mouvements');
      ControleMvt(1);
      lTraitement.Caption := TraduireMemoire('Vérification de la validité des champs des écritures');
      ZEcriture.AjouteCtx( cVPeriodeEcr );
      ZEcriture.AjouteCtx( cVJournalEcr );
      ZEcriture.AjouteCtx( cVCompteEcr );
      ZEcriture.AjouteCtx( cVEcheanceEcr );
    end;

    if vNomTraitement = 'FCBLETTRAGES' then
    begin
      lTraitement.Caption := TraduireMemoire('Vérification du lettrage');
      ControleMvt(2);
      //ZEcriture.AjouteCtx( cVLettrableEcr );
    end;

    (*
    if vNomTraitement = 'FCBANALYTIQECR' then
    begin
      lTraitement.Caption := TraduireMemoire('Vérification de l''existence de l''analytique');
      ZEcriture.AjouteCtx( cVAnalatytiqEcr );
    end;*)

  end;

  // Fonction de correction des dossiers Compta PGI
  if FCleFiltre = cFI_TableCorrect then
  begin
    // CORRECTION DES ERREURS SUR LES FICHES
    if vNomtraitement = 'FCBGENERAUX' then
    begin
      lTraitement.Caption := TraduireMemoire('Correction des comptes généraux');
      VerCompteMAJ(1, NbNoPossible, False) ;
      if NbNoPossible > 0 then AjouteErreurCor( 'Correction des comptes généraux', IntToStr(NbNoPossible) + ' compte(s) mouvementé(s)', False );
      lTraitement.Caption := TraduireMemoire('Correction de la nature des comptes 6 et 7');
      ControleCpt67( True );
    end;

    if vNomtraitement = 'FCBAUXILIAIRES' then
    begin
      lTraitement.Caption := TraduireMemoire('Correction des comptes auxiliaires');
      VerCompteMAJ(2, NbNoPossible, False) ;
      if NbNoPossible > 0 then AjouteErreurCor( 'Correction des comptes auxiliaires', IntToStr(NbNoPossible) + ' compte(s) mouvementé(s)', False );
    end;

    if vNomtraitement = 'FCBSECTIONS' then
    begin
      lTraitement.Caption := TraduireMemoire('Correction des sections');
      VerCompteMAJ(3, NbNoPossible, False) ;
      if NbNoPossible > 0 then AjouteErreurCor( 'Correction des sections', IntToStr(NbNoPossible) + ' compte(s) mouvementé(s)', False );
    end;

    if vNomtraitement = 'FCBJOURNAUX' then
    begin
      lTraitement.Caption := TraduireMemoire('Correction des journaux');
      VerCompteMAJ(4, NbNoPossible, False) ;
      if NbNoPossible > 0 then AjouteErreurCor( 'Correction des journaux', IntToStr(NbNoPossible) + ' compte(s) mouvementé(s)', False );
    end;

    if vNomTraitement = 'FCBSOLDECOMPTE' then
    begin
      lTraitement.Caption := TraduireMemoire('Recalcul des soldes des comptes');
      MajTotTousComptes(False,'');
    end;

    // CORRECTION DES ERREURS SUR LES MOUVEMENTS COMPTABLES
    if vNomTraitement = 'FCBEQUILIBREMVT' then
    begin
      lTraitement.Caption := TraduireMemoire('Correction de l''équilibrage des mouvements');
      RepareMvt;
      lTraitement.Caption := TraduireMemoire('Correction des comtes des mouvements');
      ZEcriture.AjouteCtx( cCCompteEcr );
      lTraitement.Caption := TraduireMemoire('Correction de l''échéance des mouvements');
      ZEcriture.AjouteCtx( cCEcheanceEcr );
    end;

    if vNomTraitement = 'FCBDATELETTRE' then
    begin
      lTraitement.Caption := TraduireMemoire('Correction des dates de paquet lettré');
      DatePaquet;
    end;

    if vNomTraitement = 'FCBCONTREPARTIE' then
    begin
      lTraitement.Caption := TraduireMemoire('Correction des contreparties');
      VerifContreparties;
    end;

    if vNomTraitement = 'FCBPERIODESEMAINE' then
    begin
      lTraitement.Caption := TraduireMemoire('Correction du couple Période/Semaine');
      ZEcriture.AjouteCtx( cCPeriodeEcr );
    end;

    if vNomTraitement = 'FCBJOURNAL' then
    begin
      lTraitement.Caption := TraduireMemoire('Correction du journal des écritures');
      ZEcriture.AjouteCtx( cCJournalEcr );
    end;

    //if vNomTraitement = 'FCBMONTANTEURO' then
    //if vNomTraitement = 'FCBMONTANTANO' then
    //if vNomTraitement = 'FCBCODEREGROUP' then
    //if vNomTraitement = 'FCBBASETVA' then
    //if vNomTraitement = 'FCBCODEACCEPT' then
    //if vNomTraitement = 'FCBCODEPOINTE' then
    //if vNomTraitement = 'FCBECART' then

  end;
  Application.ProcessMessages;
end;

procedure TFMainVerif.BAnnulerClick(Sender: TObject);
begin
  if PgiAsk('Confirmez vous l''arrêt du programme ?',FMainVerif.Caption) = MrYes then
  begin
    FBoStop := True;
  end;
end;

initialization
  RenseigneLaSerie(ExeCCAUTO);
  Apalatys := 'CEGID';
  Copyright := '© Copyright ' + Apalatys;
  V_PGI.OutLook := FALSE;
  V_PGI.OfficeMsg := FALSE;
  //V_PGI.ToolsBarRight:=TRUE ;
  V_PGI.VersionDemo := FALSE;
  V_PGI.SAV := True;
  V_PGI.BlockMAJStruct := True;
  V_PGI.EuroCertifiee := True;
  ChargeXuelib;
  V_PGI.MenuCourant := 0;
  V_PGI.NumVersion := '4.0.0';
  V_PGI.NumBuild := '9';
  V_PGI.NumVersionBase := 582;
  V_PGI.DateVersion := EncodeDate(2002, 05, 21);
  V_PGI.ImpMatrix := True;
  V_PGI.OKOuvert := FALSE;
  V_PGI.Halley := TRUE;
  //V_PGI.NiveauAcces:=1 ;
  V_PGI.LaSerie := S5;
  V_PGI.ParamSocLast := True;
  V_PGI.PGIContexte := [ctxCompta];
  V_PGI.CegidAPalatys := FALSE;
  V_PGI.CegidBureau := TRUE;
  V_PGI.StandardSurDP := True;
  V_PGI.MajPredefini := False;

  VH^.EnSerie := True;
  ProcGetVH := GetVH; // Corrige le proble sur Tablette TTEXOSAUFPRECEDENT
end.

