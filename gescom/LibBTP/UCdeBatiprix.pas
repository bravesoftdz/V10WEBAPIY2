unit UCdeBatiprix;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, ComCtrls, ExtCtrls, HPanel, HTB97,CBPPath,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,EdtREtat,uPDFBatch,
     fe_main,
{$else}
     eMul,UtileAGL,
     MainEagl,
{$ENDIF}
  ParamSoc,
  UTOB,
  AglInit,
  Hctrls, Mask, Spin, HDB,
  BatiprixDataClient_TLB, HFolders, GraphicEx
;

type
  TFCdeBatiprix = class(TFAssist)
    TS0: TTabSheet;
    TS1: TTabSheet;
    Sep0: TGroupBox;
    LTitreDeb0: THLabel;
    RAISONSOC: TEdit;
    TRAISONSOC: TLabel;
    NOM: TEdit;
    TNOM: TLabel;
    TPRENOM: TLabel;
    PRENOM: TEdit;
    TADRESSE: TLabel;
    ADRESS1: TEdit;
    ADRESS2: TEdit;
    TCODEPOST: TLabel;
    CODEPOST: TEdit;
    TVILLE: TLabel;
    VILLE: TEdit;
    TPAYS: TLabel;
    PAYS: TEdit;
    TTEL: TLabel;
    TEL: TEdit;
    TFAX: TLabel;
    FAX: TEdit;
    TEMAIL: TLabel;
    EMAIL: TEdit;
    BIDENTIFICATION: TToolbarButton97;
    TINFOBATIPRIX: THLabel;
    GroupBox2: TGroupBox;
    NBPOST: THDBSpinEdit;
    TEMPLACEMENT: TLabel;
    REPSRC: THCritMaskEdit;
    TNBPOST: TLabel;
    ODBATIPRIX: TBrowseFolder;
    Panel1: TPanel;
    Image1: TImage;
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BIDENTIFICATIONClick(Sender: TObject);
    procedure REPSRCElipsisClick(Sender: TObject);
  private
    { Déclarations privées }
    CdeMaster : TCommandeMaster;
    TheTOBetat : TOB;
    RepInst,NomCom,NumSer : string;
    procedure SetRepertBatiprix;
    procedure SetInfoCommande;
    procedure GetInfosIdBatiprix;
    procedure GetInfosCdeBatiprix;
    procedure GetRepertBatiprix;
    procedure EditeCommande;
    function GetMaster : boolean;
    function CdeTermine : boolean;
  public
    { Déclarations publiques }
  end;


procedure CommandeBiblioBatiprix;

implementation

{$R *.DFM}


procedure CommandeBiblioBatiprix;
var Cdebatiprix : TFCdeBatiprix;
begin
  Cdebatiprix := TFCdeBatiprix.create (application);
  TRY
    CdeBatiprix.ShowModal;
  finally
    Cdebatiprix.free;
  end;
end;

procedure TFCdeBatiprix.bSuivantClick(Sender: TObject);
begin
  inherited;
  if P.ActivePageIndex = P.PageCount-1 then
  begin
  	bfin.Enabled := true;
  	bSuivant.Enabled := false;
  end else if P.Pages[P.ActivePageIndex+1].tag = -1  then
  begin
  	bfin.Enabled := true;
  	bSuivant.Enabled := false;
  end else
  begin
  	bfin.Enabled := false;
  end;
end;

procedure TFCdeBatiprix.bPrecedentClick(Sender: TObject);
begin
  inherited;
  if P.ActivePageIndex < P.PageCount-1 then
  begin
  	bfin.Enabled := false;
  end;
end;


procedure TFCdeBatiprix.bFinClick(Sender: TObject);
begin
  inherited;
  SetInfoCommande;
  if not GetMaster then exit;
  if Not Cdetermine then exit;
  // On edite la commande
  EditeCommande;
end;

procedure TFCdeBatiprix.FormCreate(Sender: TObject);
begin
  inherited;
  CdeMaster := TCommandeMaster.Create(self);
  NomCom := getparamsoc('SO_NOMSOCBATIPRIX');
  NumSer := getparamsoc('SO_NUMSERBATIPRIX');
  GetInfosCdeBatiprix;
end;

procedure TFCdeBatiprix.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  CdeMaster.free;
  inherited;
end;

procedure TFCdeBatiprix.BIDENTIFICATIONClick(Sender: TObject);
begin
  inherited;
  AGLLanceFiche ('BTP','BTREFCDEBATIPRIX','','','ACTION=MODIFICATION');
  GetInfosIdBatiprix;
end;

procedure TFCdeBatiprix.REPSRCElipsisClick(Sender: TObject);
begin
  inherited;
  if ODBATIPRIX.Execute then
  begin
    REPSRC.Text := ODBATIPRIX.Directory;
    SetRepertBatiprix;
  end;
end;

procedure TFCdeBatiprix.SetRepertBatiprix;
begin
  SetParamSoc ('SO_BTSRCMASTER',RepSrc.text);
end;

procedure TFCdeBatiprix.GetInfosCdeBatiprix;
begin

  // emplacement de stockage des fichiers
//  RepInst := ExtractFilePath(Application.ExeName);
(*
  if Pos('\APP',UpperCase(RepInst))=0 then
     begin
     // Mode developpement
     RepInst := ExtractFileDrive (Application.exename)+'\PGI00\BATIPRIX\';
     end
  else
     begin
     RepInst := Copy (Repinst,1,Pos('\APP',UpperCase(RepInst))-1)+ '\BATIPRIX\';
     end;
*)
	RepInst := IncludeTrailingBackslash (TCBPPAth.GetCegidDataDistri)+IncludeTrailingBackslash('BATIPRIX');
  // Informations client
  GetInfosIdBatiprix;
  // rpertoire des ficheirs BATIPRIX;
  GetRepertBatiprix;
end;

procedure TFCdeBatiprix.GetInfosIdBatiprix;
var TOBSOciete : TOB;
    QQ : TQuery;
begin
  TOBSociete:= TOB.create('LA SOCIETE', nil, -1);
  QQ := OpenSql('SELECT * FROM SOCIETE',true);
  TOBSociete.SelectDb('SOCIETE', QQ);
  ferme(QQ);
//
  RaisonSoc.Text := TOBSociete.getvalue('SO_LIBELLE');
  Nom.text :=  TOBSociete.getvalue('SO_CONTACT');
  Prenom.Text :=  TOBSociete.getvalue('SO_PRENOMCONTACT');
  Adress1.text :=  TOBSociete.getvalue('SO_ADRESSE1');
  Adress2.text :=  TOBSociete.getvalue('SO_ADRESSE2');
  CodePost.text := TOBSociete.getvalue('SO_CODEPOSTAL');
  Ville.text :=  TOBSociete.getvalue('SO_VILLE');
  Pays.text := rechDom('TTPAYS',TOBSociete.getvalue('SO_PAYS'),false);
//  Pays.text := 'France';
  Tel.text := TOBSociete.getvalue('SO_TELEPHONE');
  Fax.text := TOBSociete.getvalue('SO_FAX');
  Email.text := TOBSociete.getvalue('SO_MAIL');
  //
  TOBSociete.free;
end;

procedure TFCdeBatiprix.GetRepertBatiprix;
begin
  RepSrc.text := getparamsoc('SO_BTSRCMASTER');
end;

function TFCdeBatiprix.GetMaster: boolean;
var Res : integer;
		ResMessage : Widestring;
begin
  result := true;
  CdeMaster.InstallEdition(RepInst, REPSRC.Text , Res, ResMessage);

  if Res <> 0 then
	begin
  	PgiBox(ResMessage);
    result := false;
	end;

end;

function TFCdeBatiprix.CdeTermine: boolean;
var Res : integer;
	  ResMessage : Widestring;
begin
  result := true;
  CdeMaster.GenereCommande(RepInst, Res, ResMessage);
  if Res <> 0 then
	begin
  	PgiBox(ResMessage);
    result := false;
	end;
end;

procedure TFCdeBatiprix.EditeCommande;
var LibelleDetail: string;
		prixDetail : double;
    Indice : integer;
    ListDetailCde : TListDetailCommande;
    Commande : Tcommande;
    PrixHT : double;
    Ret : integer;
    TOBEtatD,TOBETAT : TOB;
begin
   Commande := CdeMaster.InfoCommande;
	 TheTOBEtat := TOB.Create ('L ETAT',nil,-1);
	 TOBETAT := TOB.Create('PIECE',TheTOBEtat,-1);
   TOBETAT.AddChampSupValeur('EXP_NOMSOC',String(CdeMaster.RaisonSociale));
   TOBETAT.AddChampSupValeur('EXPNOMSOC_CTRL',String(CdeMaster.SommeRaisonSociale));
   TOBETAT.AddChampSupValeur('EXP_NOM',String(CdeMaster.Nom));
	 TOBETAT.AddChampSupValeur('EXPNOM_CTRL',String(CdeMaster.SommeNom));
	 TOBETAT.AddChampSupValeur('EXP_COMMANDE',String(Commande.Libelle));
   TOBETAT.AddChampSupValeur('EXP_PRENOM',String(CdeMaster.Prenom));
	 TOBETAT.AddChampSupValeur('EXP_ADRESSE1',String(CdeMaster.Adresse1));
   TOBETAT.AddChampSupValeur('EXP_ADRESSE2',String(CdeMaster.Adresse2));
   TOBETAT.AddChampSupValeur('EXP_CODEPOST',String(CdeMaster.CodePostal));
   TOBETAT.AddChampSupValeur('EXP_VILLE',String(CdeMaster.Ville));
	 TOBETAT.AddChampSupValeur('EXP_PAYS',String(CdeMaster.Pays));
   TOBETAT.AddChampSupValeur('EXP_TELEPHONE',String(CdeMaster.Telephone));
	 TOBETAT.AddChampSupValeur('EXP_FAX',String(CdeMaster.Fax));
	 TOBETAT.AddChampSupValeur('EXP_EMAIL',String(CdeMaster.Email));
	 TOBETAT.AddChampSupValeur('CLEF_CTRL',String(CdeMaster.SommeClefCommande));
	 TOBETAT.AddChampSupValeur('CLEF_CDE',String(CdeMaster.ClefCommande));
	 TOBETAT.AddChampSupValeur('NBPOSTES',integer(CdeMaster.NbPostes));
   TOBETAT.AddChampSupValeur('DATE_CDE',DATE);
	 TOBETAT.AddChampSupValeur('CEGID_NOMSOFT',String(CdeMaster.NomCommercial));
	 TOBETAT.AddChampSupValeur('CEGID_NOMSOFT_CTRL',String(CdeMaster.SommeNomCommercial));
	 TOBETAT.AddChampSupValeur('CEGID_SN',String(CdeMaster.NoSerie));
	 TOBETAT.AddChampSupValeur('CEGID_SN_CTRL',String(Cdemaster.SommeNumeroSerie));

   Commande := CdeMaster.InfoCommande;
	 ListDetailCde := CdeMaster.DetailCommande;
      //for indice:=0 to 9 do
  for Indice:=0 to ListDetailCde.count -1  do
  begin
    LibelleDetail := ListDetailCde.Items[Indice].Libelle;
    prixDetail := ListDetailCde.Items[Indice].prix;
    TOBETATD := TOB.Create ('LIGNE',TOBETAT,-1);
    TOBETATD.AddChampSupValeur('LOT',LibelleDetail);
    TOBETATD.AddChampSupValeur('PRIX_LOT',prixDetail);
  end;

   PrixHT := Commande.Prix;
   TOBETAT.AddChampSupValeur('TOTAL_HT',PrixHT);
   Ret := LanceEtatTob ('E','GPJ','BAT',TheTOBEtat,true,false,false,nil,'','Bon de commande BATIPRIX',false);

  TheTOBEtat.free;
  Commande.free;
  ListDetailCde.free;
  close;
end;

procedure TFCdeBatiprix.SetInfoCommande;
begin
  CdeMaster.NomCommercial := NomCom;
  CdeMaster.NoSerie := NumSer ;
  CdeMaster.RaisonSociale := RaisonSoc.Text;
  CdeMaster.Nom := nom.Text;
  CdeMaster.Prenom := Prenom.Text;
  CdeMaster.Adresse1 := Adress1.Text;
  CdeMaster.Adresse2 := Adress2.Text;
  CdeMaster.CodePostal := CodePost.Text;
  CdeMaster.Ville := Ville.Text;
  CdeMaster.Pays := Pays.Text;
  CdeMaster.Telephone := Tel.text;
  CdeMaster.Fax := Fax.Text;
  CdeMaster.Email := Email.Text;
  CdeMaster.NbPostes := NbPost.Value;
end;

end.
