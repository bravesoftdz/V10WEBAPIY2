unit RadS3S5;
{
 E_IO                   : indique . si l'écriture est révisée (ou en cours de révision) : Rx
                                  . sur quel site l'écriture peut être modifiée :
                                    > Chez le client :
                                      - RE : le client ne peut pas modifié la pièce ni pointer / lettrer : la pièce est en cours de révision chez l'expert
                                      - RC : le client ne peut pas modifié la pièce mais peut pointer / lettrer : la pièce est revenue d'une révision Expert
                                    > Chez l'expert :
                                      - RC : l'expert ne peut pas modifier la pièce : la pièce a déjà fait l'objet d'une révision retournée chez le client

 E_ETATREVISION         : Ce champ est flaggé à "M" dès qu'une pièce déjà révisée doit refaire un aller-retour.
                          Le "M" est sauvé automatiquement dès lors que le n° de paquet de révision < au dernier n° de paquet.
                          Pour assurer l'aller retour, il n'est pas enregistre dans le fichier exporté coté client :
                          - le client modifie => envoie chez l'expert, l'expert renvoie la pièce (dernier n° de paquet ) sans le "M"
                          - l'expert modifie  => envoie chez le client avec le "M", puis comme ci-dessus.
                          Contexte : lettrage (+ périmètre de lettrage) et pointage

 E_PAQUETREVISION       : Indique le N° de paquet d'écritures envoyer par le client chez l'expert.
                          Ce N° est incrémenté envoi par envoi par le client.
                          Chez l'expert, c'est toujours le dernier paquet reçu qui est envoyé


}
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, ExtCtrls, StdCtrls, Mask, Hctrls, ComCtrls, HSysMenu, hmsgbox,
  HTB97, HPanel, Hent1, DB, Hqry, HStatus, uTob, ParamSoc,
  dbtables, MailOL, HFLabel, UiUtil, Ent1;

procedure RevisionAdistance(Expert: boolean);

const
  nLigneLibre = 20;

  GenerauxS3 = 'G_GENERAL,          G_LIBELLE,          G_FERME,            G_NATUREGENE,'    +
               'G_UTILISATEUR,      G_PURGEABLE,        G_CENTRALISABLE,    G_DATECREATION,'  +
               'G_DATEMODIF,        G_TOTALDEBIT,       G_TOTALCREDIT,      G_VENTILABLE,'    +
               'G_LETTRABLE,        G_DERNLETTRAGE';

  TiersS3    = 'T_AUXILIAIRE,       T_NATUREAUXI,       T_LIBELLE,          T_ADRESSE1,'      +
               'T_ADRESSE2,         T_ADRESSE3,         T_CODEPOSTAL,       T_VILLE,'         +
               'T_PAYS,             T_LANGUE,           T_DEVISE,           T_SIRET,'         +
               'T_APE,              T_TELEPHONE,        T_FAX,              T_REMISE,'        +
               'T_MODEREGLE,        T_REGIMETVA,        T_SOUMISTPF,        T_DATECREATION,'  +
               'T_DATEMODIF,        T_CONFIDENTIEL,     T_UTILISATEUR,      T_RVA,'           +
               'T_DERNLETTRAGE,     T_CREDITACCORDE,    T_ESCOMPTE,'                          +
               'T_JURIDIQUE AS T_FORMEJURIDIQUE,        T_TABLE0 AS T_TABLELIBRE1,'           +
               'T_TABLE1 AS T_TABLELIBRE2,              T_TABLE2 AS T_TABLELIBRE3,'           +
               'T_TOTALCREDIT,      T_TOTALDEBIT,       T_SECTEUR,          T_FERME,'         +
               'T_COLLECTIF,        T_PRENOM';

  SectionS3  = 'S_SECTION,          S_LIBELLE,          S_DATECREATION,   S_DATEMODIF,'       +
               'S_UTILISATEUR,      S_TOTALCREDIT,      S_TOTALDEBIT' ;

  JournalS3  = 'J_JOURNAL,          J_LIBELLE,          J_NATUREJAL,      J_DATECREATION,'    +
               'J_DATEMODIF,        J_UTILISATEUR,      J_CONTREPARTIE,   J_COMPTEINTERDIT';

  DeviseS3   = 'D_DEVISE,           D_LIBELLE,          D_PARITEEURO AS D_COURS,'             +
               'D_DECIMALE AS D_NBDEC,                  D_MONNAIEIN AS D_INEURO';

  EcritureS3 = 'E_JOURNAL,          E_DATECOMPTABLE,    E_NUMEROPIECE,    E_NUMLIGNE,'        +
               'E_GENERAL,          E_AUXILIAIRE,       E_LIBELLE,        E_REFINTERNE AS E_REFERENCE,' +
               'E_DEBIT,            E_CREDIT,           E_LETTRAGE,       E_REFPOINTAGE,'     +
               'E_DEVISE,           E_DEBITDEV,         E_CREDITDEV,      E_TAUXDEV,'         +
               'E_DATEECHEANCE,     E_COUVERTURE,       E_ETATLETTRAGE,   E_ETAT,'            +
               'E_ANA,              E_REFEXTERNE,       E_NIVEAURELANCE,  E_DATECREATION,'    +
               'E_DATEMODIF,        E_UTILISATEUR,      E_QUALIFPIECE,    E_EXERCICE,'        +
               'E_PAQUETREVISION,   E_IO,               E_ETATREVISION,   E_MODEPAIE AS E_MODEP,' +
               'E_NUMECHE,          E_CONTREPARTIEGEN,  E_QTE1,           E_QTE2,'            +
               'E_QUALIFQTE1,       E_QUALIFQTE2,       E_TABLE0 AS E_TABLELIBRE1,'           +
               'E_TABLE1 AS E_TABLELIBRE2,              E_TABLE2 AS E_TABLELIBRE3,'           +
               'E_COTATION,'        +
               'E_MODESAISIE,       E_REFREVISION,      E_ECHE,           E_CREERPAR';
               {les champs 'E_DEBITEURO,E_CREDITEURO,E_SAISIEEURO,E_COTATION de S5(pour calculer debitdev, creditdev)
               n'existent pas dans S3, il faut les supprimer avant SaveToFile, E_MODESAISIE et E_REFREVISION aussi}

  AnalytiqS3 = 'Y_GENERAL,          Y_DATECOMPTABLE,    Y_NUMEROPIECE,    Y_NUMLIGNE,'        +
               'Y_SECTION,          Y_DEBIT,            Y_CREDIT,         Y_REFINTERNE AS Y_REFERENCE,' +
               'Y_LIBELLE,          Y_DEVISE,           Y_DEBITDEV,       Y_CREDITDEV,'       +
               'Y_TAUXDEV,          Y_JOURNAL,          Y_NUMVENTIL,      Y_QTE1,'            +
               'Y_QTE2,             Y_QUALIFQTE1,       Y_QUALIFQTE2,     Y_TABLE0 AS Y_TABLELIBRE1,' +
               'Y_TABLE1 AS Y_TABLELIBRE2,              Y_TABLE2 AS Y_TABLELIBRE3,'           +
               'Y_POURCENTAGE AS Y_PERCENT,'          +
               'Y_AXE,'             +
               'Y_EXERCICE,         Y_NATUREPIECE,      Y_QUALIFPIECE' ;
               {les derniers 4 champs de S5(pour calculer debitdev, creditdev)
               n'existent pas dans S3, il faut les supprimer avant SaveToFile}

type TRevision = (RevNone, RevS1, RevS1Pro, RevS3) ;
type TFill     = (boAna, boGen, boAux) ;

const FILL_RIGHT = '+' ;

type
  TFRevDist = class(TFAssist)
    Choix: TTabSheet;
    Periode: TTabSheet;
    Fichier: TTabSheet;
    Mail: TTabSheet;
    rg: TRadioGroup;
    FArrete: TCheckBox;
    Label3: TLabel;
    edPath: THCritMaskEdit;
    FMail: TCheckBox;
    Label2: TLabel;
    FEMail: THCritMaskEdit;
    FMode: TLabel;
    Bevel1: TBevel;
    Label5: TLabel;
    Label7: TLabel;
    Bevel2: TBevel;
    Label4: TLabel;
    Bevel3: TBevel;
    Bevel4: TBevel;
    Label8: TLabel;
    Label10: TLabel;
    Bevel5: TBevel;
    Label11: TLabel;
    FCorpsMail: TMemo;
    Label12: TLabel;
    Resume: TTabSheet;
    Bevel6: TBevel;
    Label13: TLabel;
    Label1: TLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    Label9: TLabel;
    Label15: TLabel;
    FLib1: TLabel;
    FVal1: TLabel;
    CM: TOpenDialog;
    FLib2: TLabel;
    FVal2: TLabel;
    FLib3: TLabel;
    FVal3: TLabel;
    FLib4: TLabel;
    FVal4: TLabel;
    Bevel7: TBevel;
    FAttClient: TLabel;
    bExec: TToolbarButton97;
    FHigh: TCheckBox;
    FlashLblInfo: TFlashingLabel;
    cbAddParametre: TCheckBox;
    MsgDiv: THMsgBox;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure rgClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bExecClick(Sender: TObject);
    procedure edPathElipsisClick(Sender: TObject);
    procedure PChange(Sender: TObject); override;
    procedure FMailClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);

  private
    { Déclarations privées }
    Expert: Boolean;
    NumPaquet: integer;
    VersionImport: Double;
    TiersDivers: TStringList;
    RadEtat:  integer;
    RevVersion : TRevision ;
    bAna : Boolean ;
    iStruct, iAna : integer ;
    Level, AnaLen, GenLen, AuxLen : integer ;
    SectionA : string ;
    function  UnFill(const sCode : string) : string ;
    function  Fill(const sCode : string; Fill: TFill) : string ;
    function  InitDateComptable: string;
    procedure InitRadioGroup;
    function  InitFEMail: string;
    function  InitChemin: string;
    procedure InitFCorpsMail;
    procedure SauvePathParamSoc;
    procedure SauveMailParamSoc;
    procedure Recevoir;
    procedure Envoyer;
    procedure AnnulerEnvoi;
    function  FichierCorrect(Fichier: string): boolean;
    procedure ImportFichier(Fichier: string);
    function  EnvoiPossible(Fichier: string; DateFin: TDateTime): boolean;
    function  ImportPossible(Fichier: string): boolean;
    function  RevisionEnCours(DateFin: TDateTime): boolean;
    function  InitRequete(DateDebut, DateFin: TDateTime): string;
    function  GetMaxNum(NomChamp: string): integer;
    function  GetDateDebut(DateFin: TDateTime): TDateTime;
    procedure InitFichierEntete(Fichier: string; DateFin: TDateTime);
    procedure ExportFichier(FichierSrc, FichierDes, sql: string);
    procedure ExportGeneraux(Fichier: string; DateDebut: TDateTime);
    procedure ExportSection(Fichier: string; DateDebut: TDateTime);
    procedure ExportJournal(Fichier: string; DateDebut: TDateTime);
    procedure ExportTiers(Fichier: string; DateDebut: TDateTime);
    procedure ExportDevise(Fichier: string; DateDebut: TDateTime);
    procedure ExportEcriture(q: TQuery; Fichier: String; NumPaquet: integer; DateDebut, DateFin: TDateTime; Perimetre: Boolean);
    procedure ExportPiece(Fichier, CodeJal: String; NumPiece, NumPaquet: integer; DateDebut, DateFin: TDateTime; Perimetre: Boolean);
    procedure AjoutChampSupS3(EcrLigne: Tob; Ajout: boolean);
    procedure ChangeNumeroPieceExport(EcrLigne: Tob);
    procedure ExportAnalytique(EcrLigne: Tob);
    procedure CalculDeviseImport(EcrLigne: Tob);
    Procedure SetMontants(Ligne: tob; CodeDev: string; TauxDev: double; ModeOppose: boolean; pf: string);
    procedure CalculDeviseExport(EcrLigne: Tob; cotation: double);
    procedure ExportPerimetreLettrage(pE_:TOB; Fichier:string; ajout:Boolean;NumPaquet:integer);
    procedure EnvoyerLeFichier(Fichier, MailTo, subject: string; MailText: TStrings; importance, EnvoiAuto: boolean);
    procedure MajEtatTransfert(CodeJal: String; NumPiece, NumPaquet: integer; DateDebit, DateFin: TDateTime);
    procedure TobToGeneraux;
    procedure TobToTiers;
    procedure TobToJournal;
    procedure TobToSection;
    procedure TobToDevise;
    procedure TobToEcriture;
    procedure AjusteCode(Fichier: string; aTob: Tob);
    procedure AjusteNature(aTob: Tob) ;
    function  GetLastNumeroPiece(DateComptable: TDatetime): integer;
    function  GetAxe(sSection : string) : string ;
    function  IsLettrable(sCode : string) : string ;
    procedure NatureGenJalS5ToS3(Fichier: string; aTob: Tob);
    procedure EcrTiersDiversS5ToS3(aTob: Tob);
    procedure GetTiersDivers;
    procedure TiersToGeneraux(Fichier: string);
    procedure GenerauxCreate(q: TQuery; Fichier: string);
    function  TableInclu(NomTable: string; DateRecep: TDateTime): boolean;
    procedure DatePaquetMinMax;
    procedure MajDatePaquetMinMax(NumPaquet: integer; CodeLettrage: string);
    function  StopImport(aTob: Tob): boolean;
    function  IsExist(Fichier, Champ, Val, sMsg: string): boolean;
    procedure InitBExec;
    procedure MajRadEtat(etat: integer);
  public
    { Déclarations publiques }
    LaTob: Tob;
    function  ImportTob(aTob: Tob): integer;
    function  PreviousPage : TTabSheet ; Override ;
    function  NextPage : TTabSheet ; Override ;
  end;


implementation

{$R *.DFM}

uses  SAISUTIL, utilPgi, CloPerio, SoldeCpt;

procedure RevisionAdistance(Expert: boolean);
Var
  fRad: TFRevDist;
  pInside: THPanel;
  sLien: string;
  lRev: TRevision ;
begin
  sLien := Uppercase(GetParamSocSecur('SO_CPLIENGAMME',''));
  lRev:=RevNone ;
  if (pos('S1', sLien)>0) then lRev:=RevS1 else
    if (pos('S3', sLien)>0) then lRev:=RevS1 ; // RevS3 ;
  if (lRev=RevNone) then
    begin
    PGIInfo('Pas de liaison avec d''autres comptabilités', 'Révision à distance') ;
    Exit ;
    end;
  SourisSablier;
  fRad := TFRevDist.Create(Application);
  fRad.Expert := TRUE ;
  fRad.VersionImport := 1;
  fRad.RadEtat := StrToInt(GetParamSocSecur('SO_CPRDETAT',''));
  fRad.RevVersion:=lRev ;
  PInside := FindInsidePanel;
  If pInside = nil then
  try
    fRad.ShowModal;
  finally
    fRad.Free;
  end
  else begin
    InitInside(fRad, pInside);
    fRad.Show;
  end;
  SourisNormale;
end;

procedure TFRevDist.FormShow(Sender: TObject);
var i : Integer ;
begin
  inherited;
  iStruct:=0 ; iAna:=0 ;
  Level:=0 ; AnaLen:=0  ; GenLen:=0 ; AuxLen:=17 ;
  if Expert then FMode.Caption:=MsgDiv.Mess[3];
  FlashLblInfo.Caption := '';
  InitRadioGroup;
  E_DATECOMPTABLE_.Text := InitDateComptable;
  FEMail.Text := InitFEMail;
  EdPath.text := InitChemin;
  InitFCorpsMail;
  FMailClick(nil);
  GetTiersDivers;
end;

procedure TFRevDist.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  rg.OnClick := nil;
  inherited;
  TiersDivers.Free;
  if IsInside(Self) then Action:=caFree ;
end;

procedure TFRevDist.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
try
  rg.OnClick := nil;
  inherited;
finally
  rg.OnClick := rgClick;
end;
end;

procedure TFRevDist.rgClick(Sender: TObject);
begin
try
  rg.OnClick := nil;
  if (RadEtat = -1) and (Rg.ItemIndex = 1) then
    MSG.Execute(23,Caption,'');                        //Etat envoi, demande réception
  if (RadEtat = 0) and (Rg.ItemIndex = 0) then begin
    MSG.Execute(24,Caption,'');                        //Etat jamais réception, demande envoi
    Rg.ItemIndex := 1;
  end;
  if (RadEtat = 1) and (Rg.ItemIndex = 0) then begin
    MSG.Execute(21,Caption,'');                        //Etat réception, demande envoi
    Rg.ItemIndex := 1;
  end;
  InitBExec;
finally
  rg.OnClick := rgClick;
end;
end;

procedure TFRevDist.bExecClick(Sender: TObject);
begin
  case rg.ItemIndex of
    0 : Transactions(Envoyer, 3);
    1 : Transactions(Recevoir, 3);
    2 : Transactions(AnnulerEnvoi, 3);
  end ;
end;

procedure TFRevDist.SauvePathParamSoc;
begin
  SetParamSoc('SO_CPRDREPERTOIRE', EdPath.text);
end;

procedure TFRevDist.SauveMailParamSoc;
begin
  SetParamSoc('SO_CPRDEMAILCLIENT', FEMail.Text)
end;

procedure TFRevDist.Recevoir;
var
  Fichier: string;
  q: TQuery;
begin
  Fichier := EdPath.text;
  if not ImportPossible(Fichier) then Exit;
  FlashLblInfo.caption := Traduirememoire('Réception en cours...');
  Application.ProcessMessages;
  SourisSablier;
  SauvePathParamSoc;
  SetParamSoc('SO_CPRDDATERECEPTION', NowH);
  ImportFichier(Fichier);
  if V_PGI.IoError=oeOk then
  begin
    DatePaquetMinMax;
    FlashLblInfo.caption:='';
    SourisNormale;
    MajRadEtat(-1);
    MajTotTousComptes(FALSE, VH^.Encours.Code) ;
    MSG.Execute(6,Caption,'');
  end
  else begin
    FlashLblInfo.caption:='';
    SourisNormale;
  end;
end;

procedure TFRevDist.Envoyer;
var
  DateDebut, DateFin, DateRecep : TDateTime;
  Fichier, sql: string;
  q: TQuery;
begin
  Fichier := EdPath.text;
  DateFin := VH^.EnCours.Fin ;// StrToDate(E_DATECOMPTABLE_.text);
  DateDebut := GetDateDebut(DateFin);
  if not EnvoiPossible(Fichier, DateFin) then
    exit;
  sql := InitRequete(DateDebut, DateFin);
  q := OpenSql(sql, true);
  try
    if q.Eof then begin
      if Expert then
        Msg.Execute(11, Caption,'')
      else
        Msg.Execute(10, Caption, E_DATECOMPTABLE_.Text);
      exit;
    end;
    SourisSablier;
    SauvePathParamSoc;
    SauveMailParamSoc;
    FlashLblInfo.caption := Traduirememoire('Envoi en cours...');
    InitFichierEntete(Fichier, DateFin);
    DateRecep := StrToDateTime(GetParamSocSecur('SO_CPRDDATERECEPTION',iDate1900));
    if TableInclu('GENERAUX', DateRecep) then
      ExportGeneraux(Fichier, DateDebut);
    if TableInclu('TIERS', DateRecep) then
      ExportTiers(Fichier, DateDebut);
    if TableInclu('JOURNAL', DateRecep) then
      ExportJournal(Fichier, DateDebut);
    if TableInclu('SECTION', DateRecep) then
      ExportSection(Fichier, DateDebut);
    ExportDevise(Fichier, DateDebut);
    ExportEcriture(q, Fichier, GetMaxNum('E_PAQUETREVISION')+1, DateDebut, DateFin, false);
  finally
    Ferme(q);
    SourisNormale;
  end;
  EnvoyerLeFichier(Fichier, FEMail.Text, MsgDiv.Mess[11], TStringList(FCorpsMail.Lines), FHigh.Checked, FMail.Checked);
  FlashLblInfo.caption := '';
  MajRadEtat(1);
end;

procedure TFRevDist.ExportPerimetreLettrage(pE_:TOB; Fichier:string; ajout:Boolean;NumPaquet:Integer);
var Tiers, Lett, stSql:string;
    i, NumLigne, NumeroPiece,CurrNum:integer;
    q:tquery;
begin
(*
if pE_.detail.count=0 then exit;
NumeroPiece:=pE_.Detail[0].GetValue('E_NUMEROPIECE') ;
for i:= 0 To pE_.detail.count-1 do
	begin
   Tiers:=vString(pE_.Detail[i].GetValue('E_AUXILIAIRE'));
   Lett:=vString(pE_.Detail[i].GetValue('E_LETTRAGE'));
   //Recherche du périmetre de lettrage
   if (Tiers<>'') and (Lett<>'') then
   	begin
      NumLigne:=pE_.Detail[i].GetValue('E_NUMLIGNE');
      stSql:='select * from ECRITURE '+
      		 'where E_AUXILIAIRE="'+Tiers+'" and E_LETTRAGE="'+Lett+'" ';
      //exclue les pièces repondant au critere d'envoi de révision en cours
      if V_PGI.ModeExpert then stSql:=stSql+ 'and E_ETATREVISION<>"M" and E_PAQUETREVISION<>0 and E_PAQUETREVISION<>'+intToStr(NumPaquet)+ ' '
      					  else stSql:=stSql+ 'and E_ETATREVISION<>"M" ';
      stSql:=stSql+ 'order by E_NUMEROPIECE,E_NUMLIGNE';
      q:=opensql(stsql, true);
      if not q.eof then
      	begin
         CurrNum:=0;
         while not q.eof do
         	begin
            if CurrNum<>q.FindField('E_NUMEROPIECE').asinteger then
            	begin
               CurrNum:=q.FindField('E_NUMEROPIECE').asinteger ;
               if (CurrNum=NumeroPiece)and(q.FindField('E_NUMLIGNE').AsInteger<>NumLigne)
                or (CurrNum<>NumeroPiece) then //Exclue la ligne de la pièce en cours
               	ExportPiece(Fichier, CurrNum, NumPaquet, false, True);
               end ;
            q.next;
            end;
         end;
      ferme(q);
      end;
   end;
*)
end;

procedure TFRevDist.edPathElipsisClick(Sender: TObject);
begin
if CM.Execute then  EdPath.Text := CM.Filename;
end;

{cette procedure ne fonctionne que pour post client}
procedure TFRevDist.AnnulerEnvoi;
var
  i: integer;
begin
  if MSG.Execute(14,Caption,'') <> mrYes then exit;
  if not Expert then begin
    i := GetMaxNum('E_PAQUETREVISION');
    if i > 0 then begin
      SourisSablier;
      try
        if ExecuteSQL('update ECRITURE set E_IO="",E_PAQUETREVISION=0,E_ETATREVISION="" ' +
                      'where E_PAQUETREVISION=' + IntToStr(i)) <= 0 then
          V_PGI.IoError := oeSaisie;
      finally
      SourisNormale;
      end;
    end;
  end;
  MSG.Execute(15,Caption,'');
end;

function TFRevDist.PreviousPage : TTabSheet ;
begin
  result:=nil;
  if GetPage = Periode then
    result:=Choix;
  if GetPage = Fichier then
  case rg.ItemIndex of
    0:    if Expert then result := Choix else result := Periode;    //si expert envoi à client, pas de page période
    1, 2: result:=Choix ;
  end;
  if GetPage = Mail then
  case rg.ItemIndex of
    0, 1: result := Fichier;
    2:    result := Choix;    //si Annuler envoi, pas de page période et fichier
  end;
  if GetPage = Resume then
  case rg.ItemIndex of
    0 : result:=Mail ;
    1 : result:=Fichier;      //si réception, pas de page mail
    2 : result:=Choix;        //si Annuler envoi, pas de page période, fichier et mail
  end;
end;

function TFRevDist.NextPage : TTabSheet ;
begin
  result:=nil ;
  if GetPage=Choix then
  case rg.ItemIndex of
    0: if Expert then result := Fichier else result := Periode;
    1: result := Fichier;
    2: result := Resume;
  end
  else if GetPage = Periode then
    result := Fichier
  else if GetPage = Fichier then begin
    if rg.ItemIndex = 0 then
      result := Mail
    else
      result := Resume;
  end
  else if GetPage = Mail then
    result:=Resume;
end;

procedure TFRevDist.PChange(Sender: TObject);
begin
  inherited;
  if GetPage = Resume then begin
    FLib2.Visible := (rg.ItemIndex = 0) and (not Expert);
    FVal2.Visible := FLib2.Visible;
    FLib3.Visible := (rg.ItemIndex < 2);
    FVal3.Visible := FLib3.Visible;
    FLib4.Visible := (rg.ItemIndex = 0) and (FMail.Checked);
    FVal4.Visible := FLib4.Visible;
    // Valeurs
    FVal1.Caption := rg.Items[rg.ItemIndex];
    FVal2.Caption := MsgDiv.Mess[12] + ' ' + E_DATECOMPTABLE_.Text;
    if FArrete.Checked then
      FVal2.Caption := FVal2.Caption + MsgDiv.Mess[13];
    FVal3.Caption := edPath.Text;
    FVal4.Caption := MsgDiv.Mess[14] + ' ' + FEMail.Text;
    bExec.Visible := True;
    bExec.Default := True;
  end
  else if GetPage = Periode then
    cbAddParametre.visible := (rg.ItemIndex = 0)
  else begin
    bExec.Visible := False;
    bExec.Default:=False;
  end
end;

procedure TFRevDist.FMailClick(Sender: TObject);
begin
  FEMail.Enabled := FMail.Checked;
  FCorpsMail.Enabled := FMail.Checked;
  FHigh.Enabled := FMail.Checked;
end;

procedure TFRevDist.InitRadioGroup;
begin
try
  rg.OnClick := nil;
  if Expert then begin
    if RadEtat = -1 then
      rg.ItemIndex := 0
    else
      rg.ItemIndex := 1;
    Rg.Items[0] := MsgDiv.Mess[5];
    Rg.Items[1] := MsgDiv.Mess[6];
  end else
    Rg.Items.Add(MsgDiv.Mess[4]);
  InitBExec;
finally
  rg.OnClick := rgClick;
end;
end;

procedure TFRevDist.InitBExec;
begin
  case Rg.ItemIndex of
    0 : bExec.Caption := MsgDiv.Mess[0];
    1 : bExec.Caption := MsgDiv.Mess[1];
    2 : bExec.Caption := MsgDiv.Mess[2];
  end;
  FAttClient.Visible:=(Rg.ItemIndex=0) and (not Expert);
end;

function TFRevDist.InitDateComptable;
var
  a, m, j : Word ;
begin
  DecodeDate(Date, a, m, j);
  if m > 1 then
    m := m - 1
  else begin
    m := 12;
    a := a - 1;
  end;
  result := DateToStr(FindeMois(EncodeDate(a, m, j)));
end;

function TFRevDist.InitFEMail: string;
begin
  result := GetParamSocSecur('SO_CPRDEMAILCLIENT','');
end;

function TFRevDist.InitChemin: string;
var
  Dir, NomFile: string;
//  d, m, y: word;
begin
//  DecodeDate(Now, y, m, d);
//  NomFile := V_PGI.CodeSociete + Format('%.4d%.2d%.2d',[y, m, d]) + '.PGI';
  NomFile := '0000000000' + IntToStr(GetMaxNum('E_PAQUETREVISION') + 1);
  NomFile := copy(NomFile, length(NomFile) - 9, 10);
  NomFile := V_PGI.CodeSociete + NomFile + '.PGI';
  Dir := GetParamSocSecur('SO_CPRDREPERTOIRE','');
  if Dir = '' then begin
    MSG.Execute(2,Caption,'');
    result := ExtractFilePath(Application.ExeName);
  end
  else begin
    if (Pos('.', Dir) = 0) and (Copy(Dir, Length(Dir), 1) <> '\') then
      Dir := Dir + '\';
    result := ExtractFilePath(Dir);
  end;
  result := result + NomFile;
end;

procedure TFRevDist.InitFCorpsMail;
var
  s: string;
begin
  s := MsgDiv.Mess[7] ;
  s := s + ' ' + V_PGI.NomSociete + ' ' + MsgDiv.Mess[8];
  if V_PGI.LaSerie = S5 then
    s := s + ' S5'
  else s := s + ' S7';
  with FCorpsMail.Lines do begin
    Clear;
    Add(s);
    Add('');
    Add(MsgDiv.Mess[9]);
    Add('');
    Add(MsgDiv.Mess[10]);
    Add('');
    Add(GetParamSocSecur('SO_CPRDNOMCLIENT',''));
  end;
end;

procedure TFRevDist.ImportFichier(Fichier: string);
begin
  LaTob := TOB.Create('', nil, -1);
  try
    TOBLoadFromFile(Fichier, ImportTOB);
  finally
    LaTob.Free;
  end;
end;

function TFRevDist.ImportTob(aTob: Tob): integer;
var
  NomTable: string;
begin
  LaTob.Dupliquer(aTob, true, true, true);
  if aTob.NomTable = 'GENERAUX' then
    TobToGeneraux
  else if aTob.NomTable = 'TIERS' then
    TobToTiers
  else if aTob.NomTable = 'JOURNAL' then
    TobToJournal
  else if aTob.NomTable = 'SECTION' then
    TobToSection
  else if aTob.NomTable = 'DEVISE' then
    TobToDevise
  else if aTob.NomTable = 'ECRITURE_S' then
    TobToEcriture;
  if V_PGI.IOError<>oeOK then begin
    result:=-1;
  end
  else result := 0;
  Application.ProcessMessages;  //pour flash info
end;

procedure TFRevDist.TobToGeneraux;
begin
// Ajout des champs nécessaires pour S5
LaTob.AddChampSup('G_COLLECTIF', TRUE) ;
LaTob.AddChampSup('G_ABREGE', TRUE) ;
LaTob.AddChampSup('G_SUIVITRESO', TRUE) ;
LaTob.AddChampSup('G_PURGEABLE', TRUE) ;
LaTob.AddChampSup('G_SOLDEPROGRESSIF', TRUE) ;
LaTob.AddChampSup('G_CONFIDENTIEL', TRUE) ;
LaTob.AddChampSup('G_SENS', TRUE) ;
LaTob.AddChampSup('G_MODELE', TRUE) ;
LaTob.AddChampSup('G_RISQUETIERS', TRUE) ;
LaTob.AddChampSup('G_VENTILABLE1', TRUE) ;
// Alimentation des champs
if (LaTob.GetValue('G_NATUREGENE')='COC') or (LaTob.GetValue('G_NATUREGENE')='COF')
  then LaTob.PutValue('G_COLLECTIF', 'X') ;
if (LaTob.GetValue('G_NATUREGENE')='TCO') or (LaTob.GetValue('G_NATUREGENE')='TDE')
  then LaTob.PutValue('G_NATUREGENE', 'DIV') ;
LaTob.PutValue('G_ABREGE', Copy(LaTob.GetValue('G_LIBELLE'), 1, 17)) ;
LaTob.PutValue('G_SUIVITRESO', 'RIE') ;
LaTob.PutValue('G_PURGEABLE', 'X') ;
LaTob.PutValue('G_SOLDEPROGRESSIF', 'X') ;
LaTob.PutValue('G_CONFIDENTIEL', '0') ;
LaTob.PutValue('G_SENS', 'M') ;
LaTob.PutValue('G_MODELE', '-') ;
LaTob.PutValue('G_RISQUETIERS', '-') ;
if (LaTob.GetValue('G_VENTILABLE')='X') then LaTob.PutValue('G_VENTILABLE1', 'X')
                                        else LaTob.PutValue('G_VENTILABLE1', '-') ;
// Vérifier la longueur du compte
LaTob.PutValue('G_GENERAL', Fill(LaTob.GetValue('G_GENERAL'), boGen)) ;
LaTob.InsertOrUpdateDB;
end;

procedure TFRevDist.TobToTiers;
var
  i: integer;
begin
with LaTob do
  begin
  if FieldExists('T_FORMEJURIDIQUE') then PutValue('T_JURIDIQUE', GetValue('T_FORMEJURIDIQUE')) ;
  if FieldExists('T_TABLELIBRE1')    then PutValue('T_TABLE0',    GetValue('T_TABLELIBRE1')) ;
  if FieldExists('T_TABLELIBRE2')    then PutValue('T_TABLE1',    GetValue('T_TABLELIBRE2')) ;
  if FieldExists('T_TABLELIBRE3')    then PutValue('T_TABLE2',    GetValue('T_TABLELIBRE3')) ;
  end;
// Ajout des champs nécessaires pour S5
LaTob.AddChampSup('T_SOLDEPROGRESSIF', TRUE) ;
LaTob.AddChampSup('T_CONFIDENTIEL', TRUE) ;
LaTob.AddChampSup('T_LETTRABLE', TRUE) ;
// Alimentation des champs
LaTob.PutValue('T_SOLDEPROGRESSIF', 'X') ;
LaTob.PutValue('T_CONFIDENTIEL', '0') ;
LaTob.PutValue('T_LETTRABLE', 'X') ;
// Vérifier la longueur du compte
LaTob.PutValue('T_AUXILIAIRE', Fill(LaTob.GetValue('T_AUXILIAIRE'), boAux)) ;
LaTob.PutValue('T_COLLECTIF',  Fill(LaTob.GetValue('T_COLLECTIF'),  boGen)) ;
LaTob.InsertOrUpdateDB;
end;

procedure TFRevDist.TobToJournal;
var Q : TQuery ;
begin
// Ajout des champs nécessaires pour S5
LaTob.AddChampSup('J_MODESAISIE', TRUE) ;
LaTob.AddChampSup('J_ABREGE', TRUE) ;
LaTob.AddChampSup('J_COMPTEURNORMAL', TRUE) ;
// Alimentation des champs
LaTob.PutValue('J_MODESAISIE', 'LIB') ;
LaTob.PutValue('J_ABREGE', Copy(LaTob.GetValue('J_LIBELLE'), 1, 17)) ;
// Première souche trouvée
Q:=OpenSQL('SELECT SH_SOUCHE FROM SOUCHE WHERE SH_SIMULATION="-" AND SH_TYPE="CPT" AND SH_ANALYTIQUE="-" ORDER BY SH_SOUCHE', TRUE) ;
LaTob.PutValue('J_COMPTEURNORMAL', Q.Fields[0].AsString) ;
Ferme(Q) ;
if LaTob.GetValue('J_NATUREJAL')='AN' then
  begin LaTob.PutValue('J_NATUREJAL', 'ANO') ; LaTob.PutValue('J_MODESAISIE', '-') ; end ;
if LaTob.GetValue('J_NATUREJAL')='TRE' then LaTob.PutValue('J_NATUREJAL', 'BQE') ;
if LaTob.GetValue('J_NATUREJAL')='VEN' then LaTob.PutValue('J_NATUREJAL', 'VTE') ;
LaTob.InsertOrUpdateDB;
end;

// Ne prendre que les sections de niveau 1 !
procedure TFRevDist.TobToSection;
const sRang : array[1..3] of string[3] = ('001', '002', '003') ;
const sAxes : array[1..3] of string[2] = ('A1',  'A2',  'A3') ;
//const sCols : array[1..3] of string[9] = ('Y_SECTION', 'Y_SECTION2', 'Y_SECTION3') ;
var Q : TQuery ; sSection : string ; sLen, i : Integer ;
begin
(*
sLen:=0 ; sSection:='' ;
if (RevVersion=RevS1Pro) and (bAna) then
  begin
  for i:=1 to iStruct do
    begin
    sLen:=sLen+Level[i] ;
    sSection:=sSection+LaTob.GetValue(sCols[i]) ;
    while length(sSection)<sLen do sSection:=sSection+'@' ;
    end ;
  end ;
LaTob.PutValue('Y_AXE', FirstAxe) ;
*)
if LaTob.GetValue('S_RANGSECTION')<>'001' then Exit ;
// Ajout des champs nécessaires pour S5
LaTob.AddChampSup('S_AXE', TRUE) ;
LaTob.AddChampSup('S_CONFIDENTIEL', TRUE) ;
LaTob.AddChampSup('S_SOLDEPROGRESSIF', TRUE) ;
LaTob.AddChampSup('S_SENS', TRUE) ;
LaTob.AddChampSup('S_ABREGE', TRUE) ;
// Alimentation des champs
LaTob.PutValue('S_AXE', 'A1') ;
LaTob.PutValue('S_CONFIDENTIEL', '0') ;
LaTob.PutValue('S_SOLDEPROGRESSIF', 'X') ;
LaTob.PutValue('S_SENS', 'M') ;
LaTob.PutValue('S_ABREGE', Copy(LaTob.GetValue('S_LIBELLE'), 1, 17)) ;
// Vérifier la longueur de la section
LaTob.PutValue('S_SECTION', Fill(LaTob.GetValue('S_SECTION'), boAna)) ;
LaTob.InsertOrUpdateDB ;
end;

procedure TFRevDist.TobToDevise;
var
  i: integer;
begin
(*
TaD_DECIMALE.AsInteger:=2 ;
TaD_QUOTITE.AsInteger:=1 ;
TaD_FONGIBLE.AsString:='-' ;
TaD_MONNAIEIN.AsString:='-' ;
TaD_PARITEEURO.AsFloat:=1 ;
*)
with LaTob do
  begin
  if FieldExists('D_COURS')  then PutValue('D_PARITEEURO', Valeur(VarToStr(GetValue('D_COURS'))));
  if FieldExists('D_NBDEC')  then PutValue('D_DECIMALE',   GetValue('D_NBDEC'));
  if FieldExists('D_INEURO') then PutValue('D_MONNAIEIN',  GetValue('D_INEURO'));
  end;
LaTob.InsertOrUpdateDB;
end;

function TFRevDist.GetAxe(sSection : string) : string ;
var Q : TQuery ;
begin
Result:='' ;
Q:=OpenSQL('SELECT S_AXE FROM SECTION WHERE S_SECTION="'+sSection+'"', TRUE) ;
if not Q.EOF then Result:=Q.Fields[0].AsString ;
Ferme(Q) ;
end ;

function TFRevDist.IsLettrable(sCode : string) : string ;
var Q : TQuery ;
begin
Result:='' ;
Q:=OpenSQL('SELECT G_LETTRABLE FROM GENERAUX WHERE G_GENERAL="'+sCode+'"', TRUE) ;
if not Q.EOF then Result:=Q.Fields[0].AsString ;
Ferme(Q) ;
end ;

function TFRevDist.UnFill(const sCode : string) : string ;
var i: integer ;
begin
i:=Length(sCode) ;
while (i>0) and (sCode[i]<=FILL_RIGHT) do dec(i) ;
Result:=Copy(sCode, 1, i) ;
end;

function TFRevDist.Fill(const sCode : string; Fill: TFill) : string ;
var iLenFill : integer ;
begin
Result:=sCode ; if sCode='' then Exit ;
iLenFill:=0 ;
case Fill of
  boAna: iLenFill:=AnaLen ;
  boGen: iLenFill:=GenLen ;
  boAux: iLenFill:=AuxLen ;
  end ;
if Length(sCode)<iLenFill then
  Result:=sCode+StringOfChar(FILL_RIGHT, iLenFill-Length(sCode)) ;
end ;

procedure TFRevDist.TobToEcriture;
var
  i, j, nRef, nPiece, iAxe: integer;
  modif: boolean; Ecr, Ana: TOB ;
begin
  nRef := 0;
  nPiece := 0;
  for i := 0 to LaTob.Detail.Count - 1 do
  with LaTob.Detail[i] do
  begin
    // Ajustement
    PutValue('E_GENERAL',    Fill(GetValue('E_GENERAL'),    boGen)) ;
    PutValue('E_AUXILIAIRE', Fill(GetValue('E_AUXILIAIRE'), boAux)) ;
    PutValue('E_ECHE', '-') ;
    PutValue('E_CREERPAR', 'SAI') ;
    if StopImport(LaTob.Detail[i]) then begin
      V_PGI.IOError := oeUnknown;
      Exit;
    end;
    if FieldExists('E_REFERENCE')   then PutValue('E_REFINTERNE', GetValue('E_REFERENCE'));
    if FieldExists('E_MODEP')       then PutValue('E_MODEPAIE',   GetValue('E_MODEP'));
    if FieldExists('E_TABLELIBRE1') then PutValue('E_TABLE0',     GetValue('E_TABLELIBRE1'));
    if FieldExists('E_TABLELIBRE2') then PutValue('E_TABLE1',     GetValue('E_TABLELIBRE2'));
    if FieldExists('E_TABLELIBRE3') then PutValue('E_TABLE2',     GetValue('E_TABLELIBRE3'));
    if FieldExists('E_TYPEPIECECOM') and (GetValue('E_TYPEPIECECOM') <> '') then
      PutValue('E_VALIDE', 'X');
//    if FieldExists('E_TYPESAISIE') and (GetValue('E_TYPESAISIE') = 1) then PutValue('E_MODESAISIE', 'LIB');
    PutValue('E_MODESAISIE', 'LIB');
    {E_QUALIFPIECE de s3 n'a pas la même utilisation que E_QUALIFPIECE de s5}
    if (GetValue('E_QUALIFPIECE') = 'AN') or (GetValue('E_QUALIFPIECE') = 'ANS')
      then begin PutValue('E_ECRANOUVEAU', 'H') ; PutValue('E_NATUREPIECE', 'OD') ; PutValue('E_MODESAISIE', '-') ; end
      else begin AjusteNature(LaTob.Detail[i]) ; PutValue('E_ECRANOUVEAU', 'N'); end ;
    PutValue('E_QUALIFPIECE',  'N');
    PutValue('E_PERIODE',      GetPeriode(GetValue('E_DATECOMPTABLE')));
    PutValue('E_SEMAINE',      NumSemaine(GetValue('E_DATECOMPTABLE')));
    PutValue('E_CONTROLETVA',  'RIE');
    PutValue('E_ETABLISSEMENT', VH^.EtablisDefaut);
    PutValue('E_EXERCICE',      VH^.EnCours.Code);
    if (GetValue('E_AUXILIAIRE')<>'') then PutValue('E_ETATLETTRAGE', 'AL') ;
    if GetValue('E_NUMECHE')>0 then begin PutValue('E_ECHE', 'X') ; PutValue('E_ETATLETTRAGE', 'AL') ; end ;
    if ((GetValue('E_AUXILIAIRE')='') and (IsLettrable(GetValue('E_GENERAL'))='X')) then PutValue('E_ETATLETTRAGE', 'AL') ;
  end;

  if LaTob.FieldExists('ACTIONFICHE') then
    modif := LaTob.GetValue('ACTIONFICHE') = 'MODIF'
  else
    modif := false;
  if not modif then begin
    nRef:=LaTob.Detail[0].GetValue('E_NUMEROPIECE') ;
    if LaTob.Detail[0].GetValue('E_MODESAISIE') = 'LIB' then
      nPiece := GetMaxNum('E_NUMEROPIECE') + 1
    else
      nPiece := GetLastNumeroPiece(LaTob.Detail[0].GetValue('E_DATECOMPTABLE'));
  end
  else begin
    nPiece := LaTob.Detail[0].GetValue('E_NUMEROPIECE');
    if LaTob.Detail[0].FieldExists('E_INFOS5') then
      nRef := StrToInt(LaTob.Detail[0].GetValue('E_INFOS5'));
    if nRef <> 0 then begin
      nPiece := nRef;
      nRef := LaTob.Detail[0].GetValue('E_NUMEROPIECE');
    end;
  end;
  if (LaTob.Detail[0].GetValue('E_TAUXDEV') = V_PGI.TauxEuro) and
     (LaTob.Detail[0].GetValue('E_DEVISE') <> GetParamSocSecur('SO_CPCODEEUROS3','')) then
    SetParamSoc('SO_CPCODEEUROS3', LaTob.Detail[0].GetValue('E_DEVISE'));

  for i := 0 to LaTob.Detail.Count - 1 do begin
    CalculDeviseImport(LaTob.Detail[i]);
    LaTob.Detail[i].PutValue('E_NUMEROPIECE', nPiece);
    if nRef <> 0 then
      LaTob.Detail[i].PutValue('E_REFREVISION', nRef);
    {import des lignes analytiques s'il y en a}
    if LaTob.Detail[i].Detail.Count > 0 then
    for j := 0 to LaTob.Detail[i].Detail.Count - 1 do
      begin
      with LaTob.Detail[i].Detail[j] do
        begin
        PutValue('Y_AXE', 'A1') ;
        PutValue('Y_SECTION', Fill(GetValue('Y_SECTION'), boAna)) ;
        if FieldExists('Y_REFERENCE')   then PutValue('Y_REFINTERNE',  GetValue('Y_REFERENCE'));
        if FieldExists('Y_TABLELIBRE1') then PutValue('Y_TABLE0',      GetValue('Y_TABLELIBRE1'));
        if FieldExists('Y_TABLELIBRE2') then PutValue('Y_TABLE1',      GetValue('Y_TABLELIBRE2'));
        if FieldExists('Y_TABLELIBRE3') then PutValue('Y_TABLE2',      GetValue('Y_TABLELIBRE3'));
        if FieldExists('Y_PERCENT')     then PutValue('Y_POURCENTAGE', GetValue('Y_PERCENT'));
        end;
      CalculDeviseImport(LaTob.Detail[i].Detail[j]);
      LaTob.Detail[i].Detail[j].PutValue('Y_NUMEROPIECE',   nPiece);
      LaTob.Detail[i].Detail[j].PutValue('Y_EXERCICE',      LaTob.Detail[i].GetValue('E_EXERCICE')) ;
      LaTob.Detail[i].Detail[j].PutValue('Y_NATUREPIECE',   LaTob.Detail[i].GetValue('E_NATUREPIECE')) ;
      LaTob.Detail[i].Detail[j].PutValue('Y_QUALIFPIECE',   LaTob.Detail[i].GetValue('E_QUALIFPIECE')) ;
      LaTob.Detail[i].Detail[j].PutValue('Y_PERIODE',       LaTob.Detail[i].GetValue('E_PERIODE'));
      LaTob.Detail[i].Detail[j].PutValue('Y_SEMAINE',       LaTob.Detail[i].GetValue('E_SEMAINE'));
      LaTob.Detail[i].Detail[j].PutValue('Y_ETABLISSEMENT', LaTob.Detail[i].GetValue('E_ETABLISSEMENT'));
      LaTob.Detail[i].Detail[j].PutValue('Y_TOTALECRITURE', LaTob.Detail[i].GetValue('E_DEBIT')+LaTob.Detail[i].GetValue('E_CREDIT'));
      if (LaTob.Detail[i].GetValue('E_ANA')='X') then
        begin
        Ecr:=TOB.Create('ECRITURE', nil, -1) ;
        Ecr.Dupliquer(LaTob.Detail[i], FALSE, TRUE) ;
        for iAxe:=1 to 5 do TOB.Create('A'+IntToStr(iAxe), Ecr, -1) ;
        Ana:=TOB.Create('ANALYTIQ', Ecr.Detail[0], -1) ;
        Ana.Assign(LaTob.Detail[i].Detail[j]) ;
        ProraterVentilsTOB(Ecr, 2) ; //2 à modifier
        end ;
    end;
  end;
  LaTob.InsertOrUpdateDB;
end;

{
  Au moment d'envoi, s'il y a des écritures dont la date comptable est inférieur à
  la date de DateFin, alors 2 cas:
  1. Si chez expert et E_IO = RC (la pièce a déjà fait l'objet d'une révision retournée chez le client)
     alors Revision en cours
  2. Si chez client et E_IO = RE (la pièce est en cours de révision chez l'expert)
     alors Revision en cours
  Dans les 2 cas, on considère la revision est dans cours, envoi ne peut pas être effectué
}
function TFRevDist.RevisionEnCours(DateFin: TDateTime): boolean;
var
  Etat, sql: string;
  q:tquery;
begin
  if Expert then
    sql := 'select E_IO from ECRITURE ' +
           'where E_IO="RC" and E_PAQUETREVISION=' + IntToStr(GetMaxNum('E_PAQUETREVISION'))
  else
    sql := 'select E_IO from ECRITURE ' +
           'where E_DATECOMPTABLE<="'+ usDateTime(DateFin) + '" and E_IO="RE"';
  q:=openSql(sql, true);
  result := (not q.eof);
  ferme(q);
end;

function TFRevDist.FichierCorrect(Fichier: string): boolean;
var
  F: TextFile;
begin
  result := true;
  if FileExists(Fichier) then
    exit;
  AssignFile(F, Fichier);
  {$I-}  ReWrite(F);   {$I+}
  if IoResult <> 0 then begin
    result := false;
    exit;
  end;
  CloseFile(F);
  DeleteFile(Fichier);
end;

function TFRevDist.EnvoiPossible(Fichier: string; DateFin: TDateTime): boolean;
begin
  Result:=False ;
  if RevisionEnCours(DateFin) then begin
    if Expert then
      MSG.Execute(21,Caption,'')
    else
      MSG.Execute(7,Caption,'');
    Exit;
  end;

  if not FichierCorrect(Fichier) then begin
    MSG.Execute(8,Caption,'');
    Exit;
  end;

  if FileExists(Fichier) then begin
    if MSG.Execute(9,Caption,'') <> mrYes then
      Exit;
    DeleteFile(Fichier) ;
  end;
  result := true;
end;

function TFRevDist.ImportPossible(Fichier: string): boolean;
var Q : TQuery ; f: TextFile; w, s, r: string; d1, d2: TDateTime; i : integer ;
begin
  Result:=False ;
  if (Fichier = '') or not FileExists(Fichier) then begin
    MSG.Execute(4, Caption, '');
    exit;
  end;
  AssignFile(f, Fichier);
  {$I-} Reset(f); {$I+}
  if IoResult <> 0 then begin
    CloseFile(f);
    MSG.Execute(16, Caption, '');
    Exit ;
  end;
  {Vérif SIRET}
  Readln(f,w);
  s := uppercase(GetParamSocSecur('SO_SIRET',''));
  if s <> '' then
  if (w <> '') and (uppercase(Trim(w)) <> s) then begin
    CloseFile(f);
    MSG.Execute(17, Caption, '');
    exit;
  end;
  {Vérif Devise de tenue}
  Readln(f,w) ;
  if (w <> '') and (trim(w) <> V_PGI.DevisePivot) then begin
    CloseFile(f);
    MSG.Execute(18, Caption, '');
    exit;
  end;
  {vérif Version d'import}
  Readln(f,w);
//  if (w <> '') and (trim(w) <> 'PGI '+FloatToStr(VersionImport)) then begin
  if (w <> '') and (Copy(trim(w), 1, 5) <> 'PGI 2') then
    begin
    CloseFile(f);
    MSG.Execute(19, Caption, '');
    exit;
    end;
  {Controle Exo}
  Readln(f,w) ;
  s := ReadTokenPipe(w,'|');
  d1 := StrToDate(copy(s,7,2) + '/' + copy(s,5,2) + '/' + copy(s,1,4));
  s := ReadTokenPipe(w,'|');
  d2 := StrToDate(copy(s,7,2) + '/' + copy(s,5,2) + '/' + copy(s,1,4));
  if (d1 < VH^.EnCours.Deb) or (d2 > VH^.EnCours.Fin) then begin
    CloseFile(f);
    MSG.Execute(20, Caption, '');
    exit;
  end;
  // Analytique
  Readln(f, w) ;
  s:=ReadTokenSt(w) ;
  if s[Length(s)]='X' then bAna:=TRUE else bAna:=FALSE ;
  if bAna then
    begin
    s:=ReadTokenSt(w) ;
    s:=ReadTokenSt(w) ;
    Level:=StrToInt(s) ;
    // Vérification adéquation
    Q:=OpenSQL('SELECT X_LONGSECTION FROM AXE WHERE X_AXE="A1"', TRUE) ;
    AnaLen:=Q.Fields[0].AsInteger ;
    Ferme(Q) ;
  end ;
  // Pointage
  Readln(f,w) ;
  // Contrôle compte de régul
  Readln(f,w) ;
  r:=Copy(w, 14, Length(w)-13) ;
  // Contrôle section d'attente
  Readln(f,w) ;
  r:=Copy(w, 17, Length(w)-16) ;
  SectionA:=ReadTokenSt(r) ;
  SectionA:=Fill(SectionA, boAna) ;
  ExecuteSQL('UPDATE AXE SET X_SECTIONATTENTE="'+SectionA+'" WHERE X_AXE="A1"') ;
  // Longueur des comptes
  Readln(f,w) ;
  r:=Copy(w, 17, Length(w)-16) ;
  s:=ReadTokenSt(r) ;
  GenLen:=StrToInt(s) ; //GenLen:=VH^.Cpta[fbGene].Lg ;
  Readln(f,w) ;
  //s:=ReadTokenSt(w) ;
  if w[Length(w)-1]='X' then AuxLen:=10 else AuxLen:=17 ;
  if AuxLen>VH^.Cpta[fbAux].Lg then AuxLen:=VH^.Cpta[fbAux].Lg ;
  // fin !
  CloseFile(f) ;
  result := true;
end;

function TFRevDist.InitRequete(DateDebut, DateFin: TDateTime): string;
var
  max: integer;
begin
  { E_PAQUETREVISION=0 => Ecritures ajoutées }
  if Expert then begin
    max := GetMaxNum('E_PAQUETREVISION');
    result := '(E_IO="" or E_ETATREVISION="M" or E_PAQUETREVISION=' + intToStr(max) +
              ' or E_PAQUETREVISION=0)';
  end
  else
    result := '(E_IO="" or E_PAQUETREVISION=0) and ' +
              'E_DATECOMPTABLE>="' + UsDatetime(DateDebut) + '" and E_DATECOMPTABLE<="' + UsDateTime(DateFin) + '"' +
              ' or E_ETATREVISION="M"';

  result := 'select E_JOURNAL, E_NUMEROPIECE from ECRITURE ' +
            'where ' + result + ' group by E_JOURNAL, E_NUMEROPIECE';
end;

procedure TFRevDist.InitFichierEntete(Fichier: string; DateFin: TDateTime);
var
  F: TextFile;
  DateD, DateF: TDateTime;
  i: integer;
begin
  SaisUtil.QuelDateExo(DateFin, DateD,DateF);
  AssignFile(F, Fichier);
  Rewrite(F);
  Writeln(F, GetParamSocSecur('SO_SIRET',''));
  Writeln(F, V_PGI.DevisePivot);
  Writeln(F, 'PGI 2') ; //+FloatToStr(VersionImport));
  Writeln(F, FormatDatetime('yyyymmdd', DateD) + '|' + FormatDatetime('yyyymmdd', DateFin)+ '|');
  for i := 1 to nLigneLibre do
    Writeln(F, '');
  CloseFile(F);
end;

function TFRevDist.GetDateDebut(DateFin: TDateTime): TDateTime;
var
  q : TQuery;
  bm : TBookmark;
begin
  bm := nil;
  q:=opensql('select * from EXERCICE order by EX_DATEDEBUT DESC', true);
  try
    if q.eof then begin
      result := iDate1900;
      exit;
    end;
    q.Last;
    result := q.findField('EX_DateDebut').AsDateTime;
    q.First;
    while not q.Eof do begin
      if (DateFin >= q.findField('EX_DateDebut').AsDateTime) and (DateFin <= q.findField('EX_DateFin').AsDateTime) then
      begin
        if q.findField('EX_ETATCPTA').AsString = 'CDE' then begin
          result := q.findField('EX_DateDebut').AsDateTime;
          Exit;
        end
        else begin
          bm := q.GetBookMark;
          q.Next;
          while not q.Eof do begin
            if q.findField('EX_ETATCPTA').AsString = 'CDE' then begin
              result := q.findField('EX_DateFin').AsDateTime + 1;
              exit;
            end
            else q.Next;
          end;
        end ;
      end ;
      if bm <> nil then begin
        q.GotoBookMark(bm);
        bm := nil;
      end;
      q.Next;
    end;
  finally
    q.FreeBookmark(bm);
    Ferme(q);
  end;
end;

procedure TFRevDist.ExportFichier(FichierSrc, FichierDes, sql: string);
var
  q : TQuery;
  aTob : Tob;
begin
  q := opensql(sql, true);
  try
    if not q.eof then begin
      aTob := Tob.create('_' + FichierSrc, nil,-1);
      InitMove(QCount(q), '');
      try
        while not q.eof do begin
          if aTob.SelectDB('', q, false) then
            begin
            NatureGenJalS5ToS3(FichierSrc, aTob);
            AjusteCode(FichierSrc, aTob) ;
            aTob.SaveToFile(FichierDes, true, true, true);
            end;
          q.next;
          MoveCur(false);
        end;
      finally
        finiMove;
        aTob.free;
      end;
    end;
  finally
    ferme(q);
  end;
end;

procedure TFRevDist.ExportGeneraux(Fichier: string; DateDebut: TDateTime);
var
  sql : string;
begin
//  if s5 then
//    sql := 'select * from GENERAUX where G_DATEMODIF>="'+usdateTime(DateDebut)+'" order by G_GENERAL';
//  else
  sql := 'select ' + GenerauxS3 + ' ' +
         'from GENERAUX where G_DATEMODIF>="'+usdateTime(DateDebut)+'" order by G_GENERAL';
  ExportFichier('GENERAUX', Fichier, sql);
  TiersToGeneraux(Fichier);
end;

procedure TFRevDist.ExportSection(Fichier: string; DateDebut: TDateTime);
var
  sql: string;
begin
//  if s5 then
//    sql := 'select * from SECTION where S_DATEMODIF>="'+usdateTime(DateDebut)+'" order by S_SECTION';
//  else
  sql := 'select ' + SectionS3 + ' ' +
         'from SECTION where S_DATEMODIF>="'+usdateTime(DateDebut)+'" order by S_SECTION';
  ExportFichier('SECTION', Fichier, sql);
end;

procedure TFRevDist.ExportJournal(Fichier: string; DateDebut: TDateTime);
var
  sql: string;
begin
//  if s5 then
//    sql := 'select * from JOURNAL where J_DATEMODIF>="'+usdateTime(DateDebut)+'" order by J_JOURNAL';
//  else
  sql := 'select ' + JournalS3 + ' ' +
         'from JOURNAL where J_DATEMODIF>="'+usdateTime(DateDebut)+'" order by J_JOURNAL';
  ExportFichier('JOURNAL', Fichier, sql);
end;

procedure TFRevDist.ExportTiers(Fichier: string; DateDebut: TDateTime);
var
  sql: string;
begin
//  if s5 then
//    sql := 'select * from TIERS where T_DATEMODIF>="'+usdateTime(DateDebut)+'" order by T_AUXILIAIRE';
//  else
  sql := 'select ' + TiersS3 + ' ' +
         'from TIERS where T_DATEMODIF>="'+usdateTime(DateDebut)+'" order by T_AUXILIAIRE';
  ExportFichier('TIERS', Fichier, sql);
end;

procedure TFRevDist.ExportDevise(Fichier: string; DateDebut: TDateTime);
var
  sql: string;
begin
//  if s5 then
//    sql := 'select * from TIERS where T_DATEMODIF>="'+usdateTime(DateDebut)+'" order by T_AUXILIAIRE';
//  else
  sql := 'select ' + DeviseS3 + ' ' +
         'from DEVISE order by D_DEVISE';
  ExportFichier('DEVISE', Fichier, sql);
end;

function TFRevDist.GetMaxNum(NomChamp: string): integer;
var
  q: TQuery;
begin
  result := 0;
  q := opensql('Select Max(' + NomChamp + ') from ECRITURE', true);
  try
    if not q.eof then
      result := q.Fields[0].Asinteger;
  finally
    ferme(q);
  end;
end;

procedure TFRevDist.ExportEcriture(q: TQuery; Fichier: String; NumPaquet: integer; DateDebut, DateFin: TDateTime; Perimetre: Boolean);
begin
  InitMove(QCount(q), '');
  q.First;
  try
    while not q.eof do begin
      ExportPiece(Fichier, q.FindField('E_JOURNAL').AsString, q.FindField('E_NUMEROPIECE').AsInteger, NumPaquet, DateDebut, DateFin, false);
      q.next;
      MoveCur(false);
    end;
  finally
    FiniMove;
  end;
end;

procedure TFRevDist.ExportPiece(Fichier, CodeJal: String; NumPiece, NumPaquet: integer; DateDebut, DateFin: TDateTime; Perimetre: Boolean);
var
  TOBEcr : Tob;
  q: TQuery;
  sql: string;
  i: integer;
  Ajout: boolean;
begin
  MajEtatTransfert(CodeJal, NumPiece, NumPaquet, DateDebut, DateFin);
  sql := 'select ' + EcritureS3 + ' from ECRITURE ' +
         'where E_JOURNAL="' + CodeJal + '"' + ' and E_NUMEROPIECE=' + IntToStr(NumPiece) + ' and E_DATECOMPTABLE>="' + UsDatetime(DateDebut) + '"';
  if not expert then sql := sql + ' and E_DATECOMPTABLE<="' + UsDateTime(DateFin) + '"';
  { Ecart euro n'envoie pas vers S3: E_DEBIT=0 and E_CREDIT=0 and (E_DEBITDEV<>0 or E_CREDITDEV<>0)   }
  sql := sql + ' and (E_DEBIT<>0 or E_CREDIT<>0)' ; // or (E_DEBITDEV=0 and E_CREDITDEV=0))';
  sql := sql + ' order by E_EXERCICE, E_NUMLIGNE, E_DATECOMPTABLE, E_JOURNAL';
  q := opensql(sql, true);
  TOBEcr := TOB.Create('ECRITURE_S', nil, -1);
  if not Expert or ((q.FindField('E_IO').asString <> '') and (q.FindField('E_PAQUETREVISION').asInteger <> 0)) then
  begin
    TOBEcr.AddChampSup('ACTIONFICHE',False);
    if (q.FindField('E_CREERPAR')<>nil) and (q.FieldByName('E_CREERPAR').asString = 'DET') then
      TOBEcr.PutValue('ACTIONFICHE','SUPPR')
    else
      TOBEcr.PutValue('ACTIONFICHE','MODIF');
  end;
  Ajout := TOBEcr.GetValue('ACTIONFICHE')='';
  try
    (*
       YCP 05/03/00 pas de details pour les suppr
    *)
    if TOBEcr.GetValue('ACTIONFICHE')='SUPPR' then
    begin
      TOBEcr.AddChampSup('E_NUMEROPIECE',True);
      TOBEcr.PutValue('E_NUMEROPIECE',q.FindField('E_REFREVISION').AsInteger) ;
    end
    else
    (* Fin Ajout *)
    begin
      if not TobEcr.LoadDetailDB('_ECRITURE','','',q, true, false) then Exit;
      for i := 0 to TobEcr.Detail.count - 1 do
      begin
        EcrTiersDiversS5ToS3(TOBEcr.Detail[i]);
        AjoutChampSupS3(TOBEcr.Detail[i], Ajout);
        TOBEcr.Detail[i].PutValue('E_GENERAL',    UnFill(TOBEcr.Detail[i].GetValue('E_GENERAL'))) ;
        TOBEcr.Detail[i].PutValue('E_AUXILIAIRE', UnFill(TOBEcr.Detail[i].GetValue('E_AUXILIAIRE'))) ;
        if not Ajout then
          ChangeNumeroPieceExport(TOBEcr.Detail[i]);
        if not expert then TOBEcr.Detail[i].PutValue('E_ETATREVISION','-');
        CalculDeviseExport(TOBEcr.Detail[i], TOBEcr.Detail[i].GetValue('E_COTATION'));
        ExportAnalytique(TOBEcr.Detail[i]);
      end;
      (*  à faire
       if Perimetre then
          ExportPerimetreLettrage(TOBEcr, Fichier, NewPiece, NumPaquet);
      *)
      with TobEcr.Detail[0] do begin
        DelChampSup('E_COTATION',True);
        DelChampSup('E_MODESAISIE',True);
        DelChampSup('E_REFREVISION',True);
        DelChampSup('E_ECHE',True);
        DelChampSup('E_CREERPAR',True);
      end;
    end ;
    TOBEcr.SaveToFile(Fichier, True, True, True);
  finally
    Ferme(q);
    TOBEcr.Free;
  end;
end;

{
  Si l'exportation d'une piéce existante dans S3, il ne faut pas envoyer E_TYPESAISIE
}
procedure TFRevDist.AjoutChampSupS3(EcrLigne: Tob; Ajout: boolean);
var
  iTypeSaisie: integer;
  sNumPiece: string;
begin
  iTypeSaisie := 0;
  sNumPiece := '';
  if (EcrLigne.GetValue('E_MODESAISIE') = 'BOR') or (EcrLigne.GetValue('E_MODESAISIE') = 'LIB') then
  begin
    iTypeSaisie := 1;
//    sNumPiece := intToStr(EcrLigne.GetValue('E_NUMEROPIECE'));
  end;
  if not Ajout then begin
    EcrLigne.AddChampSup('E_TYPESAISIE',True);
    EcrLigne.PutValue('E_TYPESAISIE', iTypeSaisie);
  end;
  EcrLigne.AddChampSup('E_INFOS5',True);
  EcrLigne.PutValue('E_INFOS5', '');
end;

procedure TFRevDist.ExportAnalytique(EcrLigne: Tob);
var
  sql, where, axe: string;
  q: TQuery;
  i, nRef, nP: integer;
begin
  if EcrLigne.GetValue('E_ANA') = 'X' then begin
    nRef := EcrLigne.GetValue('E_REFREVISION');
    if nRef <> 0 then
      nP := StrToInt(EcrLigne.GetValue('E_INFOS5'))
    else
      nP := EcrLigne.GetValue('E_NUMEROPIECE');
    where :=  'Y_EXERCICE="' + EcrLigne.GetValue('E_EXERCICE') + '"' +
         ' and Y_DATECOMPTABLE ="'+UsDateTime(EcrLigne.GetValue('E_DATECOMPTABLE')) + '"' +
         ' and Y_JOURNAL ="'+EcrLigne.GetValue('E_JOURNAL') + '"' +
         ' and Y_NUMEROPIECE=' + IntToStr(np) +
         ' and Y_NUMLIGNE=' + IntToStr(EcrLigne.GetValue('E_NUMLIGNE'));
    sql := 'select Y_AXE from ANALYTIQ where ' + where + ' order by Y_AXE';
    q := opensql(sql, true);
    if q.Eof then begin
      Ferme(q);
      exit;
    end;
    axe := q.FindField('Y_AXE').Asstring;   //importe que le premier Axe
    Ferme(q);
    sql := 'select ' + AnalytiqS3 + ' from ANALYTIQ where ' + where + ' and Y_AXE="' + axe + '"' +
    ' order by Y_NUMVENTIL';
    q := opensql(sql, true);
    try
      EcrLigne.LoadDetailDB('_ANALYTIQ','', '', q, true);
      for i := 0 to EcrLigne.Detail.count - 1 do
        begin
        CalculDeviseExport(EcrLigne.Detail[i], EcrLigne.GetValue('E_COTATION'));
        EcrLigne.PutValue('E_NUMEROPIECE', nRef);
        EcrLigne.Detail[i].PutValue('Y_SECTION', UnFill(EcrLigne.Detail[i].GetValue('Y_SECTION'))) ;
        end;
      with EcrLigne.Detail[0] do begin
        DelChampSup('Y_AXE', TRUE) ;
        DelChampSup('Y_EXERCICE', TRUE) ;
        DelChampSup('Y_NATUREPIECE', TRUE) ;
        DelChampSup('Y_QUALIFPIECE', TRUE) ;
      end;
    finally
      Ferme(q);
    end;
  end;
end;

procedure TFRevDist.SetMontants (Ligne: tob; CodeDev: string; TauxDev: double; ModeOppose: boolean; pf: string);
var
  XD, XC, SD, SC : double ;
BEGIN
XD := Ligne.GetValue(pf + 'DEBITDEV');
XC := Ligne.GetValue(pf + 'CREDITDEV');
if VH^.TenueEuro then begin
  {Compta tenue en Euro}
  if CodeDev = V_PGI.DevisePivot then begin
    if not ModeOppose then begin
      {Saisie en Euro}
      SD := Ligne.GetValue(pf + 'DEBIT');
      SC := Ligne.GetValue(pf + 'CREDIT');
      Ligne.PutValue(pf + 'DEBIT', XD);
      Ligne.PutValue(pf + 'CREDIT', XC);
      Ligne.PutValue(pf + 'DEBITDEV', XD);
      Ligne.PutValue(pf + 'CREDITDEV', XC);
    end
    else begin
      {Saisie en Franc}
      Ligne.PutValue(pf + 'DEBIT', PivotToEuro(XD));
      Ligne.PutValue(pf + 'CREDIT', PivotToEuro(XC));
      Ligne.PutValue(pf + 'DEBITDEV', PivotToEuro(XD));
      Ligne.PutValue(pf + 'CREDITDEV', PivotToEuro(XC));
    end;
  end
  else begin
    {Saisie en Devise}
    SD := Ligne.GetValue(pf + 'DEBITDEV');
    SC := Ligne.GetValue(pf + 'CREDITDEV');
    Ligne.PutValue(pf + 'DEBIT',  DeviseToEuro(XD, TauxDev, 1));
    Ligne.PutValue(pf + 'CREDIT', DeviseToEuro(XC, TauxDev, 1));
    Ligne.PutValue(pf + 'DEBITDEV', XD);
    Ligne.PutValue(pf + 'CREDITDEV',XC);
  end;
end  //if VH^.TenueEuro then
else begin
  {Compta tenue en Francs}
  if ModeOppose then begin
    {Saisie en Euros}
    Ligne.PutValue(pf + 'DEBIT',  EuroToPivot(XD));
    Ligne.PutValue(pf + 'CREDIT', EuroToPivot(XC));
    Ligne.PutValue(pf + 'DEBITDEV',  EuroToPivot(XD));
    Ligne.PutValue(pf + 'CREDITDEV', EuroToPivot(XC));
  end
  else begin
    {Saisie en Franc ou Devise}
    SD := Ligne.GetValue(pf + 'DEBITDEV');
    SC := Ligne.GetValue(pf + 'CREDITDEV');
    Ligne.PutValue(pf + 'DEBIT',  DeviseToPivot(XD, TauxDev, 1));
    Ligne.PutValue(pf + 'CREDIT', DeviseToPivot(XC, TauxDev, 1));
    Ligne.PutValue(pf + 'DEBITDEV',  XD);
    Ligne.PutValue(pf + 'CREDITDEV', XC);
  end;
end;
END ;

procedure TFRevDist.CalculDeviseImport(EcrLigne: Tob);
var
  TauxDev: double;
  CodeDev: string;
  SaisieEuro, ModeOppose: boolean;
  pf: string;
begin
  pf := EcrLigne.GetNomChamp(1001);
  pf := copy(pf, 1, pos('_', pf));
  CodeDev := EcrLigne.GetValue(pf + 'DEVISE');
  TauxDev := EcrLigne.GetValue(pf + 'TAUXDEV');
  SaisieEuro := CodeDev = GetParamSocSecur('SO_CPCODEEUROS3','');
  if SaisieEuro then begin
    EcrLigne.PutValue(pf + 'DEVISE', V_PGI.DevisePivot);
    EcrLigne.PutValue(pf + 'TAUXDEV', 1);
    CodeDev := V_PGI.DevisePivot;
    TauxDev := 1;
  end;
  ModeOppose := ((V_PGI.DevisePivot = 'FRF') and SaisieEuro) OR
                ((V_PGI.DevisePivot = 'EUR') and (CodeDev = 'FRF'));
  SetMontants(EcrLigne, CodeDev, TauxDev, ModeOppose, pf);
end;

procedure TFRevDist.CalculDeviseExport(EcrLigne: Tob; cotation: double);
var
  TauxDev: double;
  CodeDev: string;
  pf: string;
begin
  pf := EcrLigne.GetNomChamp(1001);
  pf := copy(pf, 1, pos('_', pf));
  CodeDev := EcrLigne.GetValue(pf + 'DEVISE');
  TauxDev := EcrLigne.GetValue(pf + 'TAUXDEV');
  if (CodeDev <> V_PGI.DevisePivot) and (not EstMonnaieIn(CodeDev)) then begin
    if VH^.TenueEuro then
      TauxDev := cotation;
  end ;
  EcrLigne.PutValue(pf + 'DEVISE', CodeDev);
  EcrLigne.PutValue(pf + 'TAUXDEV', TauxDev);
end;

procedure TFRevDist.EnvoyerLeFichier(Fichier, MailTo, subject: string; MailText: TStrings; importance, EnvoiAuto: boolean);
var
  high: integer;
begin
  if EnvoiAuto then begin
    if importance then
      high := 2
    else high := 1;
    SendMail(subject, MailTo, '', MailText, Fichier, true, high);
  end;
  if EnvoiAuto then
    Msg.Execute(12,Caption,'')
  else
    MSG.Execute(13,Caption,'');
end;

procedure TFRevDist.MajEtatTransfert(CodeJal: String; NumPiece, NumPaquet: integer; DateDebit, DateFin: TDateTime);
var
  sql, Etat, NumPaq, where: string;
begin
  if NumPiece = 0 then
    Exit;
  if Expert then begin
    Etat := 'RC';     //envoi pour révision chez client
    NumPaq := '';
    where := 'E_DATECOMPTABLE>="' + UsDateTime(DateDebit) + '"';
  end
  else begin
    Etat := 'RE';     //envoi pour révision chez expert
    NumPaq := intToStr(NumPaquet);
    where := 'E_DATECOMPTABLE>="' + UsDateTime(DateDebit) + '"' + ' and E_DATECOMPTABLE<="' + UsDateTime(DateFin) + '"';
  end;
  sql := 'Update ECRITURE Set E_IO="' + Etat + '", E_DATEMODIF="' + UsTime(NowH) + '"';
  if NumPaq <> '' then
    sql := sql + ' ,E_PAQUETREVISION="' + NumPaq + '"';
  sql := sql + ' where E_JOURNAL="' + CodeJal + '"' + ' and E_NUMEROPIECE=' + intToStr(NumPiece) + ' and ' + where;
  if ExecuteSQL(sql) <= 0 then
    V_PGI.IoError := oeSaisie;
end;

function TFRevDist.GetLastNumeroPiece(DateComptable: TDatetime): integer;
var
  mm: string17;
begin
  mm := '';
  result := GetNum(EcrGen, 'CPT', mm, DateComptable);
  SetNum(EcrGen, 'CPT', result + 1, result, DateComptable);
end;

procedure TFRevDist.AjusteNature(aTob: Tob) ;
var Q : TQuery ; sNatureGen, sNatureJal : string ;
begin
aTob.PutValue('E_NATUREPIECE', 'OD') ; sNatureGen:='' ; sNatureJal:='' ;
Q:=OpenSQL('SELECT G_NATUREGENE FROM GENERAUX WHERE G_GENERAL="'+aTob.GetValue('E_GENERAL')+'"', TRUE) ;
if not Q.EOF then sNatureGen:=Q.Fields[0].AsString ;
Ferme(Q) ;
Q:=OpenSQL('SELECT J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="'+aTob.GetValue('E_JOURNAL')+'"', TRUE) ;
if not Q.EOF then sNatureJal:=Q.Fields[0].AsString ;
Ferme(Q) ;
if (sNatureGen='') or (sNatureJal='') then Exit ;
if (sNatureJal='ACH') then
  begin
  if (sNatureGen='COF') and (aTob.GetValue('E_CREDIT')>0)
    then aTob.PutValue('E_NATUREPIECE', 'FF') ;
  if (sNatureGen='COF') and ((aTob.GetValue('E_CREDIT')<=0) or (aTob.GetValue('E_DEBIT')<>0))
    then aTob.PutValue('E_NATUREPIECE', 'AF') ;
  if (sNatureGen='CHA') and (aTob.GetValue('E_DEBIT')>0)
    then aTob.PutValue('E_NATUREPIECE', 'FF') ;
  if (sNatureGen='CHA') and ((aTob.GetValue('E_DEBIT')<=0) or (aTob.GetValue('E_CREDIT')<>0))
    then aTob.PutValue('E_NATUREPIECE', 'AF') ;
  end ;
if (sNatureJal='VTE') then
  begin
  if (sNatureGen='COC') and (aTob.GetValue('E_DEBIT')>0)
    then aTob.PutValue('E_NATUREPIECE', 'FC') ;
  if (sNatureGen='COC') and ((aTob.GetValue('E_DEBIT')<=0) or (aTob.GetValue('E_CREDIT')<>0))
    then aTob.PutValue('E_NATUREPIECE', 'AC') ;
  if (sNatureGen='PRO') and (aTob.GetValue('E_CREDIT')>0)
    then aTob.PutValue('E_NATUREPIECE', 'FC') ;
  if (sNatureGen='PRO') and ((aTob.GetValue('E_CREDIT')<=0) or (aTob.GetValue('E_DEBIT')<>0))
    then aTob.PutValue('E_NATUREPIECE', 'AC') ;
  end ;
if (sNatureJal='BQE') or (sNatureJal='CAI') then
  begin
  if ((sNatureGen='BQE') or (sNatureGen='CAI')) and (aTob.GetValue('E_DEBIT')>0)
    then aTob.PutValue('E_NATUREPIECE', 'RC') ;
  if ((sNatureGen='BQE') or (sNatureGen='CAI')) and ((aTob.GetValue('E_DEBIT')<=0) or (aTob.GetValue('E_CREDIT')<>0))
    then aTob.PutValue('E_NATUREPIECE', 'RF') ;
  if (sNatureGen='COC') and (aTob.GetValue('E_CREDIT')>0)
    then aTob.PutValue('E_NATUREPIECE', 'RC') ;
  if (sNatureGen='COF') and (aTob.GetValue('E_DEBIT')>0)
    then aTob.PutValue('E_NATUREPIECE', 'RF') ;
  end ;
end ;

procedure TFRevDist.AjusteCode(Fichier: string; aTob: Tob);
begin
if Fichier='GENERAUX' then aTob.PutValue('G_GENERAL', UnFill(aTob.GetValue('G_GENERAL'))) ;
if Fichier='TIERS' then
  begin
  aTob.PutValue('T_AUXILIAIRE', UnFill(LaTob.GetValue('T_AUXILIAIRE'))) ;
  aTob.PutValue('T_COLLECTIF',  UnFill(LaTob.GetValue('T_COLLECTIF'))) ;
  end ;
if Fichier='SECTION' then LaTob.PutValue('S_SECTION', UnFill(LaTob.GetValue('S_SECTION'))) ;
if Fichier='ECRITURE_S' then
  begin
  aTob.PutValue('T_AUXILIAIRE', UnFill(LaTob.GetValue('T_AUXILIAIRE'))) ;
  aTob.PutValue('T_COLLECTIF',  UnFill(LaTob.GetValue('T_COLLECTIF'))) ;
  end ;
end ;

procedure TFRevDist.NatureGenJalS5ToS3(Fichier: string; aTob: Tob);
var
  nat: string;
begin
  if Fichier = 'GENERAUX' then begin
    nat := aTob.GetValue('G_NATUREGENE');
    if (nat = 'COD') or (nat = 'COS') or (nat = 'EXT') or (nat = 'TIC') or (nat = 'TID') then
      aTob.PutValue('G_NATUREGENE', 'DIV');
  end
  else if Fichier = 'JOURNAL' then begin
    nat := aTob.GetValue('J_NATUREJAL');
    if nat = 'ANO' then
      nat := 'AN'
    else if nat = 'VTE' then
      nat := 'VEN'
    else if (nat = 'BQE') or (nat = 'CAI') then
      nat := 'TRE'
    else if (nat = 'ANA') or (nat = 'CLO') or (nat = 'ECC') or (nat = 'EXT') or (nat = 'ODA') or (nat = 'REG') then
      nat := 'OD';
    aTob.PutValue('J_NATUREJAL', nat);
  end;
end;

procedure TFRevDist.EcrTiersDiversS5ToS3(aTob: Tob);
var
  cpt: string;
begin
  cpt := aTob.GetValue('E_AUXILIAIRE');
  if TiersDivers.IndexOf(cpt) <> -1 then begin
    aTob.PutValue('E_REFEXTERNE', cpt);
    aTob.PutValue('E_AUXILIAIRE', '');
  end;
end;

procedure TFRevDist.GetTiersDivers;
var
  q: TQuery;
begin
  TiersDivers := TStringList.Create;
  q:=opensql('select T_AUXILIAIRE from Tiers where T_NATUREAUXI="DIV"', true);
  {??? comment faire pour d'autres natures qu'ils n'existe pas dans S3: AUC, AUD, NCP, SAL}
  try
    while not q.Eof do begin
      TiersDivers.Add(q.Fields[0].AsString);
      q.Next;
    end;
  finally
    Ferme(q);
  end;
end;

procedure TFRevDist.TiersToGeneraux(Fichier: string);
var
  i: integer;
  where: string;
  q: TQuery;
begin
  if TiersDivers.Count <= 0 then exit;
  where := ' T_AUXILIAIRE=';
  for i := 0 to TiersDivers.Count - 1 do
    where := where + '"' + TiersDivers[i] + '"' + 'or';
  where := copy(where, 1, length(where) - 2);
  q := opensql('select * from Tiers where ' + where, true);
  try
    while not q.eof do begin
      GenerauxCreate(q, Fichier);
      q.Next;
    end;
  finally
    Ferme(q);
  end;
end;

procedure TFRevDist.GenerauxCreate(q: TQuery; Fichier: string);
var
  i: integer;
  s, s1: string;
  aTob: Tob;
begin
  aTob := Tob.create('_GENERAUX', nil,-1);
  try
    s := GenerauxS3;
    while length(s) > 0 do
      aTob.AddChampSup(trim(ReadTokenPipe(s, ',')), False);
    for i := 0 to aTob.ChampsSup.Count - 1 do begin
      s := TCS(aTob.ChampsSup.Items[i]).Nom;
      s1 := 'T' + copy(s, 2, length(s) - 1);
      if q.FindField(s1) <> nil then
        aTob.PutValue(s, q.FieldValues[s1])
      else
        aTob.PutValue(s, Null);
    end;
    aTob.PutValue('G_GENERAL',      q.FieldValues['T_AUXILIAIRE']);
    aTob.PutValue('G_FERME',        '-');
    aTob.PutValue('G_NATUREGENE',   'DIV');
    aTob.PutValue('G_LETTRABLE',    'X');
    aTob.SaveToFile(Fichier, true, true, true);
  finally
    aTob.free;
  end;
end;

function TFRevDist.TableInclu(NomTable: string; DateRecep: TDateTime): boolean;
var
  q: TQuery;
  s: string;
begin
(*
  if NomTable = 'GENERAUX' then s := 'G_DATEMODIF';
  if NomTable = 'TIERS'    then s := 'T_DATEMODIF';
  if NomTable = 'JOURNAL'  then s := 'J_DATEMODIF';
  if NomTable = 'SECTION'  then s := 'S_DATEMODIF';
*)
  result := false;
  s := copy(NomTable, 1, 1) + '_DATEMODIF';
  s := 'select max(' + s + ') from ' + NomTable;
  q := OpenSql(s, true);
  try
    if not q.Eof then
      result := q.Fields[0].AsDateTime > DateRecep;
  finally
    Ferme(q);
  end;
end;

{ Pour les écritures possède le même code lettrage, les champs e_datepaquetmin et e_datepaquetmax
  doivent alimenter par la date comptable la plus petite et la plus grande de le paquet des écritures
}
procedure TFRevDist.DatePaquetMinMax;
var
  sql: string;
  q: TQuery;
begin
  NumPaquet := GetMaxNum('E_PAQUETREVISION');
  sql := 'select distinct E_LETTRAGE from ECRITURE where E_PAQUETREVISION=' + IntToStr(NumPaquet) + ' and E_LETTRAGE<>""';
  q := OpenSql(sql, true);
  try
    while not q.Eof do begin
      MajDatePaquetMinMax(NumPaquet, q.Fields[0].AsString);
      q.next;
    end;
  finally
    Ferme(q);
  end;
end;

procedure TFRevDist.MajDatePaquetMinMax(NumPaquet: integer; CodeLettrage: string);
var
  sql: string;
  q: TQuery;
  DateMin, DateMax: TDateTime;
begin
  sql := 'select E_DATECOMPTABLE from ECRITURE ' +
         'where E_PAQUETREVISION=' + IntToStr(NumPaquet) + ' and E_LETTRAGE="' + CodeLettrage + '" ' +
         'order by E_DATECOMPTABLE';
  q := OpenSql(sql, true);
  if not q.Eof then begin
    DateMin := q.Fields[0].AsDateTime;
    DateMax := q.Fields[0].AsDateTime;
    q.Last;
    DateMax := q.Fields[0].AsDateTime;
    Ferme(q);
    sql := 'Update ECRITURE set E_DATEPAQUETMIN="' + UsDateTime(DateMin) + '", E_DATEPAQUETMAX="' + UsDateTime(DateMax) + '" ' +
           'where E_PAQUETREVISION=' + IntToStr(NumPaquet) + ' and E_LETTRAGE="' + CodeLettrage + '"';
    if ExecuteSQL(sql) <= 0 then
      V_PGI.IoError := oeSaisie;

  end;
end;

function TFRevDist.IsExist(Fichier, Champ, Val, sMsg: string): boolean;
var
  s, sql: string;
  q: TQuery;
begin
  result := true;
  sql := 'select ' + Champ + ' from ' +  Fichier + ' where ' + Champ + '="' + Val + '"';
  q := OpenSql(sql, true);
  result := not q.eof;
  Ferme(q);
  if not result then
    MSG.Execute(22, Caption, sMsg + ' ' + Val);
end;


function TFRevDist.StopImport(aTob: Tob): boolean;
var
  s: string;
  i: integer;
begin
  result := not isExist('JOURNAL', 'J_JOURNAL', aTob.GetValue('E_JOURNAL'), 'Le code journal');
  if result then Exit;
  result := not isExist('GENERAUX', 'G_GENERAL',aTob.GetValue('E_GENERAL'), 'Le compte général');
  if result then Exit;
  s := aTob.GetValue('E_AUXILIAIRE');
  if s <> '' then begin
    result := not isExist('TIERS', 'T_AUXILIAIRE', aTob.GetValue('E_AUXILIAIRE'), 'Le compte auxiliaire');
    if result then Exit;
  end;
  result := not isExist('DEVISE', 'D_DEVISE', aTob.GetValue('E_DEVISE'), 'Le code devise');
  if result then Exit;
  if aTob.Detail.Count > 0 then
  for i := 0 to aTob.Detail.Count - 1 do begin
    result := not isExist('SECTION', 'S_SECTION', Fill(aTob.Detail[i].GetValue('Y_SECTION'), boAna), 'Le code section');
    if result then Exit;
  end;
end;

procedure TFRevDist.ChangeNumeroPieceExport(EcrLigne: Tob);
var
  nRef, nPiece: integer;
begin
  nRef := EcrLigne.GetValue('E_REFREVISION');
  if nRef <> 0 then begin
    nPiece := EcrLigne.GetValue('E_NUMEROPIECE');
    EcrLigne.PutValue('E_NUMEROPIECE', nRef);
    EcrLigne.PutValue('E_INFOS5', IntToStr(nPiece));
  end;
end;

procedure TFRevDist.MajRadEtat(etat: integer);
begin
  SetParamSoc('SO_CPRDETAT', IntToStr(etat));
end;

end.

