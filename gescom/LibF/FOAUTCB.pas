{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/12/2000
Modifié le ... : 23/07/2001
Description .. : Demande d'acceptation d'un paiement en carte bancaire à
Suite ........ : un TPE.
Mots clefs ... : FO
*****************************************************************}
unit FOAutCB;

interface

uses
  Forms, Graphics, ExtCtrls, StdCtrls, Hctrls, LCD_Lab, ComCtrls, FileCtrl,
  Controls, Windows, SysUtils, HmsgBox, HEnt1, HTB97, HPanel, Classes, UTob,
  ParamSoc;

function FOLanceAcceptationCB(TOBEche, TOBDev, TOBMdp: TOB): integer;

type
  TFAUTCB = class(TForm)
    DockBottom: TDock97;
    Outils97: TToolWindow97;
    pnlButtons: THPanel;
    BAbandon: TToolbarButton97;
    BDemande: TToolbarButton97;
    BManuel: TToolbarButton97;
    PMAIN: THPanel;
    ImageCB: TImage;
    FMONTANTREGLE: TLCDLabel;
    TTITRE: THLabel;
    TTITRE_: THLabel;
    AnimateCB: TAnimate;
    BStop: TToolbarButton97;
    TimerGen: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BDemandeClick(Sender: TObject);
    procedure BManuelClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure TimerGenTimer(Sender: TObject);
  private
    Resultat: TModalResult;
    Montant: double;
    EnCours: boolean;
    TOBEche: TOB;
    TOBDev: TOB;
    TOBMdp: TOB;
    procedure ActiveTransaction(Etat: boolean);
    function InitShow: integer;
    function LiaisonTPE(Caisse, TPVCARTEB, PORTCarteB, PARAMSTPE: string; Montant: double): boolean;
    function ICVGetDirectory: string;
    procedure ICVEcritFichier(Caisse: string);
    function ICVLitReponse(Caisse: string; var EtatTrans, ResultTPE: string): boolean;
    function LiaisonICVerify(Caisse, TPVCARTEB: string; Montant: double): integer;
  public
  end;

implementation

uses
  FOUtil, FODefi, MC_Lib, TPE_Base,
  UtilCB, MFODETAILCB_TOF;

const
  ICVFichier = 'ICVER'; // préfixe du fichier de transaction de ICVerify
  ICVTemp = '.TMP'; // extension nom du fichier temporaire de ICVerify
  ICVRequest = '.REQ'; // extension nom du fichier de demande de ICVerify
  ICVAnswer = '.ANS'; // extension nom du fichier de réponse de ICVerify

{$R *.DFM}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/12/2000
Modifié le ... : 20/11/2003
Description .. : Demande d'acceptation d'un paiement en carte bancaire à
Suite ........ : un TPE.
Suite ........ :
Suite ........ : Valeur retournée :
Suite ........ :   mrOk  : transaction acceptée par le TPE
Suite ........ :   mrYes : transaction forcée par l'utilisateur
Suite ........ :   autres: transaction refusée
Mots clefs ... : FO
*****************************************************************}

function FOLanceAcceptationCB(TOBEche, TOBDev, TOBMdp: TOB): integer;
var XX: TFAutCB;
begin
  Result := mrNone;
  if (not Assigned(TOBEche)) or (not Assigned(TOBDev)) then Exit;
  XX := TFAutCB.Create(Application);
  try
    //XX.TOBEche.Assign(TOBEche);
    //XX.TOBDev.Assign(TOBDev);
    //XX.TOBDev.Assign(TOBMdp);
    XX.TOBEche := TOBEche;
    XX.TOBDev := TOBDev;
    XX.TOBMdp := TOBMdp;
    Result := XX.InitShow;
  finally
    XX.Free;
  end;
  Application.ProcessMessages;
  SourisNormale;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/12/2000
Modifié le ... : 23/07/2001
Description .. : FormCreate
Mots clefs ... : FO
*****************************************************************}

procedure TFAUTCB.FormCreate(Sender: TObject);
begin
//  TOBEche := TOB.Create('PIEDECHE', nil, -1);
//  TOBDev := TOB.Create('DEVISE', nil, -1);
  Resultat := mrNone;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/12/2000
Modifié le ... : 23/07/2001
Description .. : FormShow
Mots clefs ... : FO
*****************************************************************}

procedure TFAUTCB.FormShow(Sender: TObject);
var Ok: boolean;
begin
  // Appel de la fonction d'empilage dans la liste des fiches
  AglEmpileFiche(Self);
  Ok := True;
  Ok := FOGetFromRegistry(REGSAISIETIC, REGAUTODIALTPE, Ok);
  if Ok then
  begin
    BManuel.Visible := False;
    BManuel.Enabled := False;
    BDemande.Visible := False;
    // lancement de la transmission du montant au TPE
    TimerGen.Interval := 100;
    TimerGen.Enabled := True;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/12/2000
Modifié le ... : 23/07/2001
Description .. : FormDestroy
Mots clefs ... : FO
*****************************************************************}

procedure TFAUTCB.FormDestroy(Sender: TObject);
begin
  if EnCours then Exit;
//  TOBEche.Free;
//  TOBDev.Free;
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/12/2000
Modifié le ... : 23/07/2001
Description .. : Affichage de la fiche
Mots clefs ... : FO
*****************************************************************}

function TFAUTCB.InitShow: integer;
begin
  Montant := VDouble(TOBEche.GetValue('GPE_MONTANTENCAIS'));
  FMONTANTREGLE.Caption := FOCadreDroite(StrfMontant(Montant, 15, VInteger(TOBDev.Getvalue('D_DECIMALE')), VString(TOBDev.GetValue('D_SYMBOLE')), True),
    FMONTANTREGLE.NoOfChars);
  ShowModal;
  Result := Resultat;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/12/2000
Modifié le ... : 23/07/2001
Description .. : Démarrage/Arrêt de la transaction avec le TPE
Mots clefs ... : FO
*****************************************************************}

procedure TFAUTCB.ActiveTransaction(Etat: boolean);
begin
  AnimateCB.Visible := Etat;
  AnimateCB.Active := Etat;
  EnCours := Etat;
  BDemande.Enabled := (not Etat);
  BManuel.Enabled := (not Etat);
  BAbandon.Enabled := (not Etat);
  BStop.Visible := False;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/12/2000
Modifié le ... : 23/07/2001
Description .. : BAbandonClick
Mots clefs ... : FO
*****************************************************************}

procedure TFAUTCB.BAbandonClick(Sender: TObject);
begin
  if EnCours then Exit;
  if not BAbandon.Enabled then Exit;
  //if Resultat = mrNone then Resultat := mrCancel ;
  ModalResult := mrCancel;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/12/2000
Modifié le ... : 23/07/2001
Description .. : BDemandeClick
Mots clefs ... : FO
*****************************************************************}

procedure TFAUTCB.BDemandeClick(Sender: TObject);
var
  Ok: boolean;
  NoErr: integer;
  TPVCARTEB, PORTCarteB, PARAMSTPE, Caisse, Stg: string;
begin
  if EnCours then Exit;
  if not BDemande.Enabled then Exit;
  Resultat := mrNone;
  Caisse := FOCaisseCourante;
  FODonneParamTPE(Caisse, TPVCARTEB, PORTCarteB, PARAMSTPE);

  if TPVCARTEB = '101' then
  begin
    NoErr := LiaisonICVerify(Caisse, TPVCARTEB, Montant);
    if NoErr < 0 then
    begin
      Close;
      Exit;
    end else
      Ok := (NoErr = 0);
  end else
  begin
    Ok := LiaisonTPE(Caisse, TPVCARTEB, PORTCarteB, PARAMSTPE, Montant);
  end;

  if (not Ok) and (FOJaiLeDroit(56, False, False)) then
  begin
    Stg := TraduireMemoire('Votre demande n''a pas été autorisée par le TPE.') + '#13#13'
      + TraduireMemoire('Acceptez-vous néanmoins ce règlement ?');
    //if PGIAsk(Stg, Caption) = mrYes then Resultat := mrOk;
    Resultat := PGIAsk(Stg, Caption);
  end;
  Close;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/12/2000
Modifié le ... : 23/07/2001
Description .. : BManuelClick
Mots clefs ... : FO
*****************************************************************}

procedure TFAUTCB.BManuelClick(Sender: TObject);
begin
  if not BManuel.Enabled then Exit;
  if EnCours then Exit;
  Resultat := mrNone;
  //Resultat := mrOk * Ord(PGIASK('Avez-vous reçu l''accord du centre d''autorisation ?', Caption) = mrYes);
  Resultat := PGIASK('Avez-vous reçu l''accord du centre d''autorisation ?', Caption);
  Close;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/12/2000
Modifié le ... : 23/07/2001
Description .. : FormKeyDown : gestion des touches du clavier
Mots clefs ... : FO
*****************************************************************}

procedure TFAUTCB.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if key = VK_Valide then
  begin
    if shift = [] then BDemande.Click else
      if ssShift in Shift then BManuel.Click;
  end;
  if key = VK_ESCAPE then BAbandon.Click;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/12/2000
Modifié le ... : 16/12/2003
Description .. : Lancement de la transmission du montant au TPE
Mots clefs ... : FO
*****************************************************************}

procedure TFAUTCB.TimerGenTimer(Sender: TObject);
begin
  TimerGen.Enabled := False;
  TimerGen.Interval := 0;
  BDemande.Click;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/12/2003
Modifié le ... : 12/12/2003
Description .. : Lance la communication avec un TPE
Mots clefs ... : FO
*****************************************************************}

function TFAUTCB.LiaisonTPE(Caisse, TPVCARTEB, PORTCarteB, PARAMSTPE: string; Montant: double): boolean;
var
  TypeTrans: TTPETransaction;
  EtatTrans, ResultTPE, NumCarte, Autres, Stg: string;
begin
  ActiveTransaction(True);
  TypeTrans := ttTPEAchat;
  if Montant < 0 then TypeTrans := ttTPERemboursement;

  Result := GereAutorisationCB(Caisse, 'CB', VString(TOBEche.GetValue('GPE_DEVISEESP')),
    TPVCARTEB, PORTCarteB, PARAMSTPE, TypeTrans, Montant, ResultTPE, NumCarte, Autres, EtatTrans);

  ActiveTransaction(False);
  if Result then
  begin
    Stg := '';
    if Trim(ResultTPE) <> '' then
    begin
      Stg := TraduireMemoire('Le résultat de la demande d''autorisation est');
      Stg := Format('%s :' + #13 + ' %s', [Stg, ResultTPE])
    end;
    if FOStrCmp(EtatTrans, '0;1;3') then
    begin
      // autorisation accordée
      Resultat := mrOk;
    end else
    begin
      // autorisation non accordée
      PGIBox(Stg, Caption);
      Result := False;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/12/2003
Modifié le ... : 12/12/2003
Description .. : Retourne le répertoire de stockage des fichiers ICVerify
Mots clefs ... : FO
*****************************************************************}

function TFAUTCB.ICVGetDirectory: string;
begin
  Result := Trim(FOGetParamCaisse('GPK_TPEREPERTOIRE'));
  if Result = '' then Result := GetCurrentDir;
  if not DirectoryExists(Result) then ForceDirectories(Result);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/12/2003
Modifié le ... : 12/12/2003
Description .. : Constitue à partir du règlement le fichier pour ICVerify.
Mots clefs ... : FO
*****************************************************************}

procedure TFAUTCB.ICVEcritFichier(Caisse: string);
var
  NomFicTmp, NomFicReq, NomFicAns, NomDir, Enreg, DateExp, sMont, sNumCarte: string;
  FichierICV: TextFile;
  Mnt: double;
  Ind: integer;
begin
  NomDir := ICVGetDirectory;
  NomFicReq := IncludeTrailingBackslash(NomDir) + ICVFichier + Caisse + ICVRequest;
  DeleteFile(NomFicReq);
  NomFicAns := IncludeTrailingBackslash(NomDir) + ICVFichier + Caisse + ICVAnswer;
  DeleteFile(NomFicAns);
  NomFicTmp := IncludeTrailingBackslash(NomDir) + ICVFichier + Caisse + ICVTemp;
  DeleteFile(NomFicTmp);

  AssignFile(FichierICV, NomFicTmp);
  if FileExists(NomFicTmp) then Reset(FichierICV) else Rewrite(FichierICV);
  try
    // enregistremet de la forme : "C1","Clerk John","Comment field","4005550000000019","0312","1.00"
    sNumCarte := DeCrypteNoCarteCB(vString(TOBEche.GetValue('GPE_CBINTERNET')), False, True, True);
    DateExp := vString(TOBEche.GetValue('GPE_DATEEXPIRE'));
    Mnt := VDouble(TOBEche.GetValue('GPE_MONTANTENCAIS'));
    sMont := Format('%.2f', [Abs(Mnt)]);
    Ind := Pos(',', sMont);
    if Ind > 0 then sMont[Ind] := '.'; // on remplace le séparateur des décimales par un point
    
    if Mnt < 0 then // leading field
      Enreg := '"C3",'
    else
      Enreg := '"C1",';
    Enreg := Enreg
      + '"' + vString(TOBEche.GetValue('GPE_CBLIBELLE')) + '",' // Clerk 32c. maxi
      + '"",' // Comment 32c. maxi
      + '"' + sNumCarte + '",' // Credit card number
      + '"' + Copy(DateExp, 5, 2) + Copy(DateExp, 1, 2) + '",' // Expiration date YYMM
      + '"' + sMont + '",'; // Amount with decimal point
    {
    if Mnt < 0 then // leading field  PC CHARGE
      Enreg := '"2"'
    else
      Enreg := '"1"';
    Enreg := Enreg
      + ' "' + vString(TOBEche.GetValue('GPE_CBINTERNET')) + '"' // Credit card number 20 digit
      + ' "' + Copy(DateExp, 1, 2) + Copy(DateExp, 5, 2) + '"' // Expiration date MMYY
      + ' "' + Format('%9.2f', [Abs(Mnt)]) + '",'; // Amount 6 digits for dollars and 2 digits for cents
      + ' ""' // Invoice number 9c.
      + ' ""' // Zip code 9c. maxi
      + ' ""' // Address 20c. maxi
    }
    Writeln(FichierICV, Enreg);
  finally
    CloseFile(FichierICV);
  end;
  RenameFile(NomFicTmp, NomFicReq);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 15/12/2003
Modifié le ... : 15/12/2003
Description .. : Interprête le fichier de réponse de ICVerify
Mots clefs ... : FO
*****************************************************************}

function TFAUTCB.ICVLitReponse(Caisse: string; var EtatTrans, ResultTPE: string): boolean;
var
  NomFicAns, NomDir, Enreg, Stg: string;
  FichierAns: TextFile;
  NbLec: integer;
begin
  Result := False;
  EtatTrans := '';
  ResultTPE := '';
  NbLec := 0;
  NomDir := ICVGetDirectory;
  NomFicAns := IncludeTrailingBackslash(NomDir) + ICVFichier + Caisse + ICVAnswer;

  BStop.Down := False;
  BStop.Visible := True;
  while not FileExists(NomFicAns) do
  begin
    Inc(NbLec);
    if (NbLec > 1000) or (BStop.Down) then Exit;
    Delay(1000);
  end;

  AssignFile(FichierAns, NomFicAns);
  FileMode := 0;
  Reset(FichierAns);
  try
    // enregistremet de la forme : "Y123456Y" / "NPICK UP CARD"
    Readln(FichierAns, Enreg);
    Stg := ReadTokenSTV(Enreg);
    if Copy(Stg, 1, 1) = '"' then
    begin
      EtatTrans := Copy(Stg, 2, 1); // Answer for transaction
      ResultTPE := Copy(Stg, 3, Length(Stg) -3); // Approval code 6 digits / Text of error
      Result := True;
    end;
  finally
    CloseFile(FichierAns);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/12/2003
Modifié le ... : 12/12/2003
Description .. : Lance la communication avec le logiciel ICVerify
Mots clefs ... : FO
*****************************************************************}

function TFAUTCB.LiaisonICVerify(Caisse, TPVCARTEB: string; Montant: double): integer;
var
  EtatTrans, ResultTPE, Stg: string;
  Ok: boolean;
begin
  Result := 1;
  if (vString(TOBEche.GetValue('GPE_CBINTERNET')) = '') or
    (vString(TOBEche.GetValue('GPE_DATEEXPIRE')) = '') then
  begin
    if FOGetparamCaisse('GPK_CLAVIERPISTE') = 'X' then
    begin
      // Lecture de la carte depuis le clavier
      Stg := 'SANSCLIENT;CAPTION=' + TOBMdp.GetValue('MP_LIBELLE') + ' : ' + Trim(FMONTANTREGLE.Caption);
      Stg := FOLanceFicheLectureCBFromKB('', '', Stg);
      if not FODecodeLectureCBFromKB(nil, TOBEche, TOBMdp, False, Stg) then
      begin
        Result := -1;
        Exit;
      end;
    end else
    begin
      // saisie des données de la carte
      if not FOSaisieDetailCB(nil, TOBEche, TOBMdp, '', True) then
      begin
        Result := -1;
        Exit;
      end;
    end;
  end;

  // Liaison avec ICVerify
  ActiveTransaction(True);
  ICVEcritFichier(Caisse);
  Ok := ICVLitReponse(Caisse, EtatTrans, ResultTPE);
  ActiveTransaction(False);

  if Ok then
  begin
    Stg := '';
    if Trim(ResultTPE) <> '' then
    begin
      Stg := TraduireMemoire('Le résultat de la demande d''autorisation est');
      Stg := Format('%s :' + #13 + ' %s', [Stg, ResultTPE])
    end;
    if EtatTrans = 'N' then
    begin
      // autorisation non accordée
      PGIBox(Stg, Caption);
      Result := 1;
    end else
    begin
      // autorisation accordée
      Result := 0;
      Resultat := mrOk;
      TOBEche.PutValue('GPE_CBNUMAUTOR', Copy(ResultTPE, 1, 6));
    end;
  end;
end;

end.
