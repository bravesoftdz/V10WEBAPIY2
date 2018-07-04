unit UintegreBatiprix;

interface

uses
Windows, 	Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist,	 HSysMenu, hmsgbox, StdCtrls, ComCtrls, ExtCtrls, HPanel, HTB97,
  Hctrls, Mask,UtilgestionChaine,
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
	UtilBatiprixFile,
  Hent1,
  UTOB,
  AglInit,
  Utilarticle,
  BatiprixDataClient_TLB, HFolders, Spin, HDB, GraphicEx,CBPPath,UtilsTOB,
  TntStdCtrls, TntComCtrls, TntExtCtrls
;

type
  TFIntegreBatiprix = class(TFAssist)
    TS0: TTabSheet;
    TS2: TTabSheet;
    LTitreDeb0: THLabel;
    Sep0: TGroupBox;
    THTRAITEMENT: THLabel;
    ANIMATION: TAnimate;
    THWAIT: THLabel;
    HLabel1: THLabel;
    GroupBox2: TGroupBox;
    TEMPLACEMENT: TLabel;
    REPSRC: THCritMaskEdit;
    VersionDLL: TLabel;
    NBPOST: THDBSpinEdit;
    Label1: TLabel;
    Panel1: TPanel;
    Image1: TImage;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure bSuivantClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
    CdeMaster : TCommandeMaster;
    RepInst : string;
    TheContainer : TCOntainerFichier;
    TOBDESARTICLES, TOBDESOUVRAGES : TOB;
    TOBCE:TOB;
    TOBGENREMAT:TOB;
    TOBQLFQTE, TOBQLFQTED : TOB;
    TOBPREST : TOB;
    Terreur : TintegrationError;
    NomCom : string;
    NumSer :String;
    RaisonSoc: string;
    Nom: string;
    Prenom: string;
    Adress1: string;
    Adress2: string;
    CodePost: string;
    Ville: string;
    Pays: string;
    Tel : string;
    Fax : string;
    Email : string;

    procedure integrationCE;
    procedure integrationFamilleMat;
    procedure integrationMateriaux;
    procedure integrationFamilleOuvrage;
    procedure integrationOuvrages;
    procedure EcritPrest;
    procedure EcritMEA;
    procedure EcritDomaineAct;
    function Decryptefichier : boolean;
//    function CdeTermine: boolean;
    procedure GetInfosIdBatiprix;
    procedure SetInfoCommande;
    procedure SupprimeSEQ;
    function FindInRef(TOBREF: TOB; CodeArticle: string): TOB;
    function FindFamilleBATIPRIX (TOBFAMILLE:TOB; IDFAM : string ; TypeA : string) : TOB;
  public
    { Déclarations publiques }
  end;

procedure IntegrationDonneeBatibrix;

implementation

{$R *.DFM}
procedure IntegrationDonneeBatibrix;
var XX: TFIntegreBatiprix;
begin
	XX := TFIntegreBatiprix.create (application);
  TRY
  	XX.ShowModal;
  FINALLY
  	XX.free;
  End;
end;



procedure TFIntegreBatiprix.SetInfoCommande;
begin
  CdeMaster.NomCommercial := NomCom;
  CdeMaster.NoSerie := NumSer ;
  CdeMaster.RaisonSociale := RaisonSoc;
  CdeMaster.Nom := nom;
  CdeMaster.Prenom := Prenom;
  CdeMaster.Adresse1 := Adress1;
  CdeMaster.Adresse2 := Adress2;
  CdeMaster.CodePostal := CodePost;
  CdeMaster.Ville := Ville;
  CdeMaster.Pays := Pays;
  CdeMaster.Telephone := Tel;
  CdeMaster.Fax := Fax;
  CdeMaster.Email := Email;
end;

procedure TFIntegreBatiprix.FormCreate(Sender: TObject);
var requete,requete1 : string;
begin
  inherited;
  ANIMATION.Active := true;
  CdeMaster := TCommandeMaster.Create(self);
  TheContainer := TCOntainerFichier.Create;
  Terreur := TintegrationError.create ;

  versionDll.Caption  := CdeMaster.VersionDll;
  // emplacement de stockage des fichiers
//  RepInst := ExtractFilePath(Application.ExeName);
(*  if Pos('\APP',UpperCase(RepInst))=0 then
  begin
    // Mode developpement
    RepInst := ExtractFileDrive (Application.exename)+'\PGI00\BATIPRIX\';
  end else
  begin
    RepInst := Copy (Repinst,1,Pos('\APP',UpperCase(RepInst))-1)+ '\BATIPRIX\';
  end;
*)
	RepInst := IncludeTrailingBackslash (TCBPPAth.GetCegidDataDistri)+IncludeTrailingBackslash('BATIPRIX');

  NomCom := getparamsoc('SO_NOMSOCBATIPRIX');
  NumSer := getparamsoc('SO_NUMSERBATIPRIX');
  GetInfosIdBatiprix;

  RepSrc.text := getparamsoc('SO_BTSRCMASTER');
  Bfin.visible := false;

  SetInfoCommande;
  SupprimeSEQ;
  TheContainer.FileDescriptor := RepInst + 'BATIPRIX.INI';
  Terreur.Open('');
  //
  TOBCE:=TOB.Create('LES CORPS D''ETAT',NIL,-1);
	TheContainer := TCOntainerFichier.create;
//  TheContainer.FileDescriptor := 'C:\PGI00\BATIPRIX\BATIPRIX.INI';
  TheContainer.FileDescriptor := IncludeTrailingBackslash (TCBPPAth.GetCegidDataDistri)+IncludeTrailingBackslash('BATIPRIX')+'BATIPRIX.INI';
	TOBGENREMAT:= TOB.Create('LES GENRES',Nil,-1);
  TOBQLFQTE := TOB.Create('LES UNITES',Nil,-1);
  //
  //  Chargement de TOBGENREMAT
  requete := 'SELECT * FROM BGENREMAT';
  TOBGENREMAT.LoadDetailDBFromSQL ('BGENREMAT',requete,false);

  // 	Chargement de TOBQLFQTE
  requete1 := 'SELECT * FROM MEA';
  TOBQLFQTE.LoadDetailDBFromSQL ('MEA',requete1,false);

  TOBDESARTICLES := TOB.Create ('NOS ARTICLES',nil,-1);
  TOBDESOUVRAGES := TOB.Create ('NOS OUVRAGES',nil,-1);
	TOBPREST := TOB.Create('LES PRESTATIONS',Nil,-1);

end;


procedure TFIntegreBatiprix.GetInfosIdBatiprix;
var TOBSOciete : TOB;
    QQ : TQuery;
begin
  TOBSociete:= TOB.create('LA SOCIETE', nil, -1);
  QQ := OpenSql('SELECT * FROM SOCIETE',true);
  TOBSociete.SelectDb('SOCIETE', QQ);
  ferme(QQ);
//
  RaisonSoc := TOBSociete.getvalue('SO_LIBELLE');
  Nom :=  TOBSociete.getvalue('SO_CONTACT');
  Prenom :=  TOBSociete.getvalue('SO_PRENOMCONTACT');
  Adress1 :=  TOBSociete.getvalue('SO_ADRESSE1');
  Adress2 :=  TOBSociete.getvalue('SO_ADRESSE2');
  CodePost := TOBSociete.getvalue('SO_CODEPOSTAL');
  Ville :=  TOBSociete.getvalue('SO_VILLE');
  Pays := rechDom('TTPAYS',TOBSociete.getvalue('SO_PAYS'),false);
//  Pays.text := 'France';
  Tel := TOBSociete.getvalue('SO_TELEPHONE');
  Fax := TOBSociete.getvalue('SO_FAX');
  Email := TOBSociete.getvalue('SO_MAIL');
  //
  TOBSociete.free;
end;

procedure TFIntegreBatiprix.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  CdeMaster.free;
  TheContainer.free;
  Terreur.close;
  Terreur.free;
  TOBQLFQTE.free;
  TOBCE.free;
  TOBGENREMAT.free;
  TOBDESARTICLES.free;
  TOBDESOUVRAGES.free;
  TOBPREST.Free;


  inherited;
end;

(*
function TFIntegreBatiprix.CdeTermine: boolean;
var Res : integer;
	  ResMessage : Widestring;
    ListDetailCde : TListDetailCommande;
    indice : integer;
		LibelleDetail: string;
		prixDetail : double;
    Commande : Tcommande;
begin
  result := true;
  CdeMaster.GenereCommande(RepInst, Res, ResMessage);
  if Res <> 0 then
	begin
  	PgiBox(ResMessage);
    result := false;
	end;
  Commande := CdeMaster.InfoCommande;
  Comlib.lines.Clear;
  ComLib.Lines.Add('Commande ' + Commande.Libelle );

  ListDetailCde := CdeMaster.DetailCommande;
  for Indice:=0 to ListDetailCde.Count -1 do
  begin
    LibelleDetail := ListDetailCde.Items[Indice].Libelle;
    prixDetail := ListDetailCde.Items[Indice].prix;
    ComLib.Lines.Add(LibelleDetail+' '+FloatToStr(PrixDetail)+' €');
  end;

end;
*)
procedure TFIntegreBatiprix.bSuivantClick(Sender: TObject);
begin
  inherited;
  CdeMaster.NbPostes := NbPost.Value ;
  (*
  if P.ActivePageIndex <> P.PageCount-1 then
  begin
    if not CdeTermine then BEGIN Bsuivant.enabled := false; bPrecedent.Visible := True; END;}
  end;
  *)
  // lancer la DLL pour decryptage des fichiers
  if Not Decryptefichier then BEGIN bFinClick(self); Exit; END;
	// si decryptage correct;
  TRY

    THTRAITEMENT.visible := true;
    ANIMATION.visible := true;
    sleep(1000);
    integrationCE;
    //
    EcritDomaineAct;
    //
    integrationFamilleMat;
    //
    integrationFamilleOuvrage;
    //
    integrationMateriaux;
    //
    integrationOuvrages;
    //
    EcritPrest;
    //
    EcritMEA;
    //
    SupprimeSEQ;
    //
    PgiBox('Traitement terminé avec succès');
  
  FINALLY
  //
  	bFinClick (self);
  END;
end;

procedure TFIntegreBatiprix.EcritDomaineAct;
var
	 TOBACTIV,TOBCED : TOB;
   indice : integer;
begin
   TOBACTIV := TOB.Create('BTDOMAINEACT',nil,-1);
	 for Indice := 0 to TOBCE.detail.count -1 do
   begin
   	TOBCED := TOBCE.detail[Indice];
    TOBACTIV.InitValeurs;
   	TOBACTIV.PutValue('BTD_CODE',TOBCED.GetValue('CEIN_ID'));
    TOBACTIV.PutValue('BTD_LIBELLE',TOBCED.GetValue('CEIN_LABEL'));
    TOBACTIV.PutValue('BTD_COEFFG',TOBCED.GetValue('CFG_MAT'));
    TOBACTIV.PutValue('BTD_COEFMARG',TOBCED.GetValue('CVENTE'));
   	TOBACTIV.InsertOrUpdateDB;
   end;
   TOBACTIV.free;
end;

procedure TFIntegreBatiprix.EcritMEA;
begin
  ExecuteSQL('DELETE FROM MEA');
  TOBQLFQTE.SetAllModifie (true);
	TOBQLFQTE.InsertOrUpdateDB;
end;

procedure TFIntegreBatiprix.EcritPrest;
var
  TOBENLIGD : TOB;
  TOBPRESTD : TOB;
  TOBO : TOB;
  CodeOuvrage, CodePrest: String;
  NBRLIGNE, indice : integer;
  MOA, MOC : double;
begin
  THTRAITEMENT.Caption := 'Traitement des DETAILS D''OUVRAGES ';
  self.refresh;
  TOBENLIGD := TOB.Create('NOMENLIG',Nil,-1);
 	for indice:=0 to TOBDESOUVRAGES.Detail.Count-1 do
  begin
    TOBO := TOBDESOUVRAGES.Detail[indice];
    MOC := TOBO.GetValue('QTEMOC');
    MOA := TOBO.GetValue('QTEMOA');
    if (MOC <> 0) or (MOA <> 0) then
    begin
      TOBENLIGD.InitValeurs;
      CodeOuvrage := TOBO.GetValue('OUVR_BATREF');
      ExecuteSql ('DELETE FROM NOMENLIG WHERE GNL_NOMENCLATURE="'+CodeOuvrage+'" AND GNL_CODEARTICLE IN ("'+TOBO.GetValue('CORDETAT')+'C000000000","'+TOBO.GetValue('CORDETAT')+'A000000000")');
      NBRLIGNE := TOBO.GetValue('NBRLIGNE');
      //
      if MOC <> 0 then
      begin
        TOBPRESTD := TOBPREST.FindFirst(['GA_CODEARTICLE'],[TOBO.GetValue('CORDETAT')+'C000000000'],true);
        if TOBPRESTD <> nil then
        begin
          CodePrest := TOBPRESTD.GetValue('GA_CODEARTICLE');
          TOBENLIGD.PutValue('GNL_NOMENCLATURE',CodeOuvrage);
          TOBENLIGD.PutValue('GNL_ARTICLE',CodeArticleUnique(CodePrest,'','','','',''));
          TOBENLIGD.PutValue('GNL_CODEARTICLE',CodePrest);
          TOBENLIGD.PutValue('GNL_NUMLIGNE',NBRLIGNE + 1);
          TOBENLIGD.PutValue('GNL_QTE',MOC);
          TOBENLIGD.PutValue('GNL_LIBELLE',TOBPRESTD.GetValue('GA_LIBELLE'));
          TOBENLIGD.InsertOrUpdateDB ;
        end;
      end;
      //
      if MOA <> 0 then
      begin
        TOBPRESTD := TOBPREST.FindFirst(['GA_CODEARTICLE'],[TOBO.GetValue('CORDETAT')+'A000000000'],true);
        if TOBPRESTD <> nil then
        begin
          CodePrest := TOBPRESTD.GetValue('GA_CODEARTICLE');
          TOBENLIGD.PutValue('GNL_NOMENCLATURE',CodeOuvrage);
          TOBENLIGD.PutValue('GNL_ARTICLE',CodeArticleUnique(CodePrest,'','','','',''));
          TOBENLIGD.PutValue('GNL_CODEARTICLE',CodePrest);
          TOBENLIGD.PutValue('GNL_NUMLIGNE',NBRLIGNE + 1 );
          TOBENLIGD.PutValue('GNL_QTE',MOA);
          TOBENLIGD.PutValue('GNL_LIBELLE',TOBPRESTD.GetValue('GA_LIBELLE'));
          TOBENLIGD.InsertOrUpdateDB ;
        end;
      end;
    end;
  end;
  TOBENLIGD.free;

	Terreur.SetError(2,'OK','');
end;

procedure TFIntegreBatiprix.integrationCE;
var
  MonFich : string;
	TOBDetailCE : TOB;
  TOBPRESTD : TOB;
  Prest,Lettre, Lib : string;
	I : integer;
begin
  THTRAITEMENT.Caption  := 'Traitement des CORPS D''ETATS';
  self.refresh;
  MonFich := RepInst+'CE.seq';
  TheContainer.OpenFile(MonFich);
  Terreur.SetError(1,'Corps_d_Etats','');
  repeat
    TOBDetailCE := TheContainer.ReadFile;
    if TOBDetailCE <> nil then
    begin
      TOBDetailCE.ChangeParent(TOBCE,-1);
      For I :=0 TO 1 do
      Begin
      	If I = 0 then
      	begin
        	Lettre := 'C';
        	Lib := TOBDetailCE.GetValue('CEIN_LABEL') + '_PRESTATION pour Main d''Oeuvre Chantier';
        end Else
        Begin
        	Lettre := 'A';
        	Lib := TOBDetailCE.GetValue('CEIN_LABEL') + '_PRESTATION pour Main d''Oeuvre Atelier';
        end;
        Prest := TOBDetailCE.GetValue('CEIN_ID') + Lettre + '000000000';
        //
      	TOBPRESTD := TOB.Create('ARTICLE',TOBPREST,-1);
        TOBPRESTD.PutValue('GA_ARTICLE',CodeArticleUnique(Prest,'','','','',''));
        TOBPRESTD.PutValue('GA_TYPEARTICLE','PRE');
      	TOBPRESTD.PutValue('GA_CODEARTICLE',Prest);
        TOBPRESTD.PutValue('GA_LIBELLE',Lib);
        TOBPRESTD.PutValue('GA_BLOCNOTE',Lib);
      	TOBPRESTD.PutValue('GA_PAHT',TOBDetailCE.GetValue('CETA_TH' + Lettre));
      	TOBPRESTD.PutValue('GA_COEFFG',TOBDetailCE.GetValue('CFG_MO' + Lettre));
      	TOBPRESTD.PutValue('GA_COEFCALCHT',TOBDetailCE.GetValue('CVENTE'));
      	TOBPRESTD.PutValue('GA_STATUTART','UNI');
        TOBPRESTD.PutValue('GA_NATUREPRES','MO');
        TOBPRESTD.PutValue('GA_QUALIFUNITEVTE','H');
        TOBPRESTD.PutValue('GA_QUALIFUNITESTO','H');
        TOBPRESTD.PutValue('GA_QUALIFMARGE','CO');
        TOBPRESTD.PutValue('GA_ACTIVITEREPRISE','F');
        TOBPRESTD.PutValue('GA_PRIXPOURQTE',1);
        TOBPRESTD.PutValue('GA_DATESUPPRESSION',Idate2099);
        TOBPRESTD.PutValue('GA_FAMILLETAXE1','TN');
        TOBPRESTD.PutValue('GA_CALCPRIXHT','DPR');
        TOBPRESTD.PutValue('GA_CALCPRIXPR','PAA');
        TOBPRESTD.PutValue('GA_DPRAUTO','X');
        TOBPRESTD.PutValue('GA_CALCAUTOHT','X');
        TOBPRESTD.PutValue('GA_DPA',TOBPRESTD.GetValue('GA_PAHT'));
        TOBPRESTD.PutValue('GA_DPR',Arrondi(TOBPRESTD.GetValue('GA_PAHT')*TOBPRESTD.GetValue('GA_COEFFG'),V_PGI.okdecP));
        TOBPRESTD.PutValue('GA_PVHT',Arrondi(TOBPRESTD.GetValue('GA_DPR')*TOBPRESTD.GetValue('GA_COEFCALCHT'),V_PGI.okdecP));
        TOBPRESTD.PutValue('GA_REMISEPIED','X');
        TOBPRESTD.PutValue('GA_REMISELIGNE','X');
        TOBPRESTD.PutValue('GA_ESCOMPTABLE','X');
        TOBPRESTD.PutValue('GA_COMMISSIONNABLE','X');
        TOBPRESTD.PutValue('GA_PRIXPASMODIF','X');
			end;
    end else
    begin
    	Terreur.SetError(3,MonFich,'');
      break;
  	end;
  until TheContainer.EofFile;
  TOBPREST.InsertOrUpdateDB(false);
  TheContainer.FermeFile;
  Terreur.SetError(2,'OK','');
end;

procedure TFIntegreBatiprix.integrationFamilleMat;

  function IncCodeFamilleArt (Niveau : Integer; Var valeur : string) : Boolean;
  var okok : Boolean;
      Sql,FNiv,Oniv : string;
  begin
    FNiv := 'FN'+InttoStr(Niveau);
    //
    okok := false;
    while Not okok do
    begin
      result :=  IncStringCode (valeur,3);
      if not result then break;
      // verification sur famille artible
      Sql := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+FNiv+'" AND CC_CODE="'+valeur+'"';
      if not ExisteSQL(Sql) then okok := True;
    end;
  end;

var
	Indice : integer;
	MonFich : string;
	TOBFAM,TOBFAMD : TOB;
  TOBN1,TOBN2,TOBN3,TOBN4,TOBN5 : TOB;
  TOBN1D,TOBN2D,TOBN3D,TOBN4D,TOBN5D : TOB;
  TOBN1P,TOBN2P,TOBN3P,TOBN4P,TOBLast : TOB;
  TOBCC,TOBCCD : TOB;
	requete : string;
  valeur : string;
begin
	THTRAITEMENT.Caption := 'Traitement des FAMILLES MATERIAUX';
  self.refresh;
	// initialisation
	TOBFAM:= TOB.create('LES FAMILLES',NIL,-1);
  TOBN1 := TOB.Create ('LES NIVEAUX 1',nil,-1);
  TOBN2 := TOB.Create ('LES NIVEAUX 2',nil,-1);
  TOBN3 := TOB.Create ('LES NIVEAUX 3',nil,-1);
  TOBN4 := TOB.Create ('LES NIVEAUX 4',nil,-1);
  TOBN5 := TOB.Create ('LES NIVEAUX 5',nil,-1);

  //  Chargement des TOBN1, TOBN2, TOBN3 contenant les enregistrement selon les  niveaux famille à partir de la table BFAMILLECHANGE
  requete := 'SELECT * FROM BFAMILLECHANGE WHERE BFG_TYPEIMPORT="BAT" AND '+
  					 'BFG_TYPEARTICLE="MAR" AND BFG_IMPFAMILLE2="" ORDER BY BFG_IMPFAMILLE1';
  TOBN1.LoadDetailDBFromSQL ('BFAMILLECHANGE',requete,false);

  requete := 'SELECT * FROM BFAMILLECHANGE WHERE BFG_TYPEIMPORT="BAT" AND '+
  					 'BFG_TYPEARTICLE="MAR" AND BFG_IMPFAMILLE2<>"" AND BFG_IMPFAMILLE3="" ORDER BY BFG_IMPFAMILLE1,BFG_IMPFAMILLE2';
  TOBN2.LoadDetailDBFromSQL ('BFAMILLECHANGE',requete,false);

  requete := 'SELECT * FROM BFAMILLECHANGE WHERE BFG_TYPEIMPORT="BAT" AND '+
  					 'BFG_TYPEARTICLE="MAR" AND BFG_IMPFAMILLE3<>"" AND BFG_IMPFAMILLE4=""  ORDER BY BFG_IMPFAMILLE1,BFG_IMPFAMILLE2,BFG_IMPFAMILLE3';
  TOBN3.LoadDetailDBFromSQL ('BFAMILLECHANGE',requete,false);

  requete := 'SELECT * FROM BFAMILLECHANGE WHERE BFG_TYPEIMPORT="BAT" AND '+
  					 'BFG_TYPEARTICLE="MAR" AND BFG_IMPFAMILLE4<>"" AND BFG_IMPFAMILLE5=""  ORDER BY BFG_IMPFAMILLE1,BFG_IMPFAMILLE2,BFG_IMPFAMILLE3,BFG_IMPFAMILLE4';
  TOBN4.LoadDetailDBFromSQL ('BFAMILLECHANGE',requete,false);

  requete := 'SELECT * FROM BFAMILLECHANGE WHERE BFG_TYPEIMPORT="BAT" AND '+
  					 'BFG_TYPEARTICLE="MAR" AND BFG_IMPFAMILLE5<>"" ORDER BY BFG_IMPFAMILLE1,BFG_IMPFAMILLE2,BFG_IMPFAMILLE3,BFG_IMPFAMILLE4,BFG_IMPFAMILLE5';
  TOBN5.LoadDetailDBFromSQL ('BFAMILLECHANGE',requete,false);

  // Lecture fichier séquentiel
	MonFich := RepInst+'MATE_FAMI.seq';
  TheContainer.OpenFile(MonFich);
  Terreur.SetError(1,'Famille_Matériaux','');
	repeat
  	TOBFamD := TheContainer.ReadFile;
    If TOBFamD <> nil then
    begin
      If TOBFamD.GetValue('MAFA_LEVEL') < 6 then
    	begin
      	TOBFamD.ChangeParent(TOBFAM, -1);
    	end;
    end else
    begin
    	Terreur.SetError(3,MonFich,'');
      break;
    end;
  until TheContainer.EofFile;
  TheContainer.FermeFile;

// Tri sur les niveaux dans la TOB
  TOBFAM.Detail.Sort('MAFA_LEVEL');
  Indice := 0;
  if TOBFAM.detail.count > 0 then
  begin
    repeat
      TOBFamD := TOBFAM.Detail[Indice];

      // Taitement des familles NIV 1 dans TOBN1
      if TOBFamD.getValue('MAFA_LEVEL')=1 then
      begin
        TOBN1D:= TOBN1.FindFirst(['BFG_IMPFAMILLE1'],[TOBFamD.getValue('MAFA_ID')],TRUE);
        If TOBN1D = nil then
        begin
          if TOBN1.detail.Count > 0 then
          begin
            TOBLast := TOBN1.detail[TOBN1.detail.count-1];
            valeur:=TOBLast.getvalue('BFG_FAMILLE')
          end else
          begin
            valeur:='000';
          end;

          TOBN1D := TOB.Create('BFAMILLECHANGE',TOBN1,-1);
          TOBN1D.PutValue('BFG_TYPEIMPORT','BAT');
          TOBN1D.PutValue('BFG_PARFOU','');
          TOBN1D.PutValue('BFG_TYPEARTICLE','MAR');
          TOBN1D.PutValue('BFG_IMPFAMILLE1',TOBFamD.getvalue('MAFA_ID'));
//          if IncStringCode(valeur,3) then
          if IncCodeFamilleArt(1,valeur) then
          begin
            TOBN1D.PutValue('BFG_FAMILLE',Valeur);
          end;
          TOBN1D.AddChampSupValeur ('LIBELLE',TOBFamD.getvalue('MAFA_LABEL'));
        end else
        begin
          TOBN1D.AddChampSupValeur ('LIBELLE',TOBFamD.getvalue('MAFA_LABEL'));
        end;

      // Taitement des familles NIV 2 dans TOBN2
      end else if TOBFamD.getValue('MAFA_LEVEL')=2 then
      begin
        // recherche du pere pour le NIV 2 dans TOBN2
        TOBN2D := TOBN2.findFirst (['BFG_IMPFAMILLE2'],
                                     [TOBFamD.getValue('MAFA_ID')],
                                     true);
        if TOBN2D = nil then
        begin
          TOBN1P := TOBN1.findFirst(['BFG_IMPFAMILLE1'],[TOBFamD.getValue('MAFA_PID')],true);
          If TOBN1P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,String(TOBFamD.getvalue('MAFA_ID')),2);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 2
          end;
          if TOBN2.detail.Count > 0 then
          begin
            TOBLast := TOBN2.detail[TOBN2.detail.count-1];
            valeur:=TOBLast.getvalue('BFG_FAMILLE')
          end else
          begin
            valeur:='000';
          end;
          TOBN2D := TOB.Create('BFAMILLECHANGE',TOBN2,-1);
          TOBN2D.PutValue('BFG_TYPEIMPORT','BAT');
          TOBN2D.PutValue('BFG_PARFOU','');
          TOBN2D.PutValue('BFG_TYPEARTICLE','MAR');
          TOBN2D.PutValue('BFG_IMPFAMILLE1',TOBFamD.getvalue('MAFA_PID'));
          TOBN2D.PutValue('BFG_IMPFAMILLE2',TOBFamD.getvalue('MAFA_ID'));
//          if IncStringCode(valeur,3) then
          if IncCodeFamilleArt(2,valeur) then
          begin
            TOBN2D.PutValue('BFG_FAMILLE',Valeur);
          end;
          TOBN2D.PutValue('BFG_FAMILLEPERE1',TOBN1P.GetValue('BFG_FAMILLE'));
          TOBN2D.AddChampSupValeur ('LIBELLE',TOBFamD.getvalue('MAFA_LABEL'));
        end else
        begin
          TOBN2D.PutValue('BFG_IMPFAMILLE1',TOBFamD.getValue('MAFA_PID'));
          TOBN2D.AddChampSupValeur ('LIBELLE',TOBFamD.getvalue('MAFA_LABEL'));
        end;

      //  Taitement des familles NIV 3 dans TOBN3
      end else if TOBFamD.getValue('MAFA_LEVEL')=3 then
      begin
        TOBN3D := TOBN3.findFirst (['BFG_IMPFAMILLE3'],
                                   [TOBFamD.getValue('MAFA_ID')],
                                     true);
        if TOBN3D = nil then
        begin
          TOBN2P := TOBN2.findFirst(['BFG_IMPFAMILLE2'],[TOBFamD.getValue('MAFA_PID')],true);
          If TOBN2P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,String(TOBFamD.getvalue('MAFA_ID')),3);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBN1P := TOBN1.findFirst(['BFG_IMPFAMILLE1'],[TOBN2P.getValue('BFG_IMPFAMILLE1')],true);
          If TOBN1P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,String(TOBFamD.getvalue('MAFA_PID')),2);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;

          if TOBN3.detail.Count > 0 then
          begin
            TOBLast := TOBN3.detail[TOBN3.detail.count-1];
            valeur:=TOBLast.getvalue('BFG_FAMILLE')
          end else
          begin
            valeur:='000';
          end;
          TOBN3D := TOB.Create('BFAMILLECHANGE',TOBN3,-1);
          TOBN3D.PutValue('BFG_TYPEIMPORT','BAT');
          TOBN3D.PutValue('BFG_PARFOU','');
          TOBN3D.PutValue('BFG_TYPEARTICLE','MAR');
          TOBN3D.PutValue('BFG_IMPFAMILLE1',TOBN1P.GetValue('BFG_IMPFAMILLE1'));
          TOBN3D.PutValue('BFG_IMPFAMILLE2',TOBFamD.getvalue('MAFA_PID'));
          TOBN3D.PutValue('BFG_IMPFAMILLE3',TOBFamD.getvalue('MAFA_ID'));
          if IncCodeFamilleArt(3,valeur) then
//          if IncStringCode(valeur,3) then
          begin
            TOBN3D.PutValue('BFG_FAMILLE',Valeur);
          end;
          TOBN3D.PutValue('BFG_FAMILLEPERE1',TOBN1P.GetValue('BFG_FAMILLE'));
          TOBN3D.PutValue('BFG_FAMILLEPERE2',TOBN2P.GetValue('BFG_FAMILLE'));
          TOBN3D.AddChampSupValeur ('LIBELLE',TOBFamD.getvalue('MAFA_LABEL'));
        end else
        begin
          TOBN2P := TOBN2.findFirst(['BFG_IMPFAMILLE2'],[TOBFamD.getValue('MAFA_PID')],true);
          If TOBN2P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,String(TOBFamD.getvalue('MAFA_ID')),3);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBN1P := TOBN1.findFirst(['BFG_IMPFAMILLE1'],[TOBN2P.getValue('BFG_IMPFAMILLE1')],true);
          If TOBN1P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBFamD.getvalue('MAFA_PID')),2);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          //
          TOBN3D.PutValue('BFG_IMPFAMILLE1',TOBN1P.GetValue('BFG_IMPFAMILLE1'));
          TOBN3D.PutValue('BFG_IMPFAMILLE2',TOBFamD.getvalue('MAFA_PID'));
          TOBN3D.AddChampSupValeur ('LIBELLE',TOBFamD.getvalue('MAFA_LABEL'));
        end;

  //  Traitement des familles niv 4
      end else if TOBFamD.getValue('MAFA_LEVEL')=4 then
      begin
        TOBN4D := TOBN4.findFirst (['BFG_IMPFAMILLE4'],
                                   [TOBFamD.getValue('MAFA_ID')],
                                     true);
        if TOBN4D = nil then
        begin
          TOBN3P := TOBN3.findFirst(['BFG_IMPFAMILLE3'],[TOBFamD.getValue('MAFA_PID')],true);
          If TOBN3P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBFamD.getvalue('MAFA_ID')),4);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 4
          end;
          TOBN2P := TOBN2.findFirst(['BFG_IMPFAMILLE2'],[TOBN3P.getValue('BFG_IMPFAMILLE2')],true);
          If TOBN2P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBFamD.getvalue('MAFA_PID')),3);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBN1P := TOBN1.findFirst(['BFG_IMPFAMILLE1'],[TOBN2P.getValue('BFG_IMPFAMILLE1')],true);
          If TOBN1P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBN2P.getvalue('BFG_IMPFAMILLE2')),2);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 2
          end;

          TOBN4D := TOB.Create('BFAMILLECHANGE',TOBN4,-1);
          TOBN4D.PutValue('BFG_TYPEIMPORT','BAT');
          TOBN4D.PutValue('BFG_PARFOU','');
          TOBN4D.PutValue('BFG_TYPEARTICLE','MAR');
          TOBN4D.PutValue('BFG_IMPFAMILLE1',TOBN1P.GetValue('BFG_IMPFAMILLE1'));
          TOBN4D.PutValue('BFG_IMPFAMILLE2',TOBN2P.getvalue('BFG_IMPFAMILLE2'));
          TOBN4D.PutValue('BFG_IMPFAMILLE3',TOBFamD.getvalue('MAFA_PID'));
          TOBN4D.PutValue('BFG_IMPFAMILLE4',TOBFamD.getvalue('MAFA_ID'));
          TOBN4D.PutValue('BFG_FAMILLE','');
          TOBN4D.PutValue('BFG_FAMILLEPERE1',TOBN1P.GetValue('BFG_FAMILLE'));
          TOBN4D.PutValue('BFG_FAMILLEPERE2',TOBN2P.GetValue('BFG_FAMILLE'));
          TOBN4D.PutValue('BFG_FAMILLEPERE3',TOBN3P.GetValue('BFG_FAMILLE'));
          TOBN4D.AddChampSupValeur ('LIBELLE',TOBFamD.getvalue('MAFA_LABEL'));
        end else
        begin
          TOBN3P := TOBN3.findFirst(['BFG_IMPFAMILLE3'],[TOBFamD.getValue('MAFA_PID')],true);
          If TOBN3P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBFamD.getvalue('MAFA_ID')),4);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBN2P := TOBN2.findFirst(['BFG_IMPFAMILLE2'],[TOBN3P.getValue('BFG_IMPFAMILLE2')],true);
          If TOBN2P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBN3P.getvalue('BFG_IMPFAMILLE2')),3);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBN1P := TOBN1.findFirst(['BFG_IMPFAMILLE1'],[TOBN2P.getValue('BFG_IMPFAMILLE1')],true);
          If TOBN1P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBN2P.getvalue('BFG_IMPFAMILLE1')),2);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBN4D.PutValue('BFG_IMPFAMILLE1',TOBN1P.GetValue('BFG_IMPFAMILLE1'));
          TOBN4D.PutValue('BFG_IMPFAMILLE2',TOBN2P.getvalue('BFG_IMPFAMILLE2'));
          TOBN4D.PutValue('BFG_IMPFAMILLE3',TOBFamD.getvalue('MAFA_PID'));
          TOBN4D.AddChampSupValeur ('LIBELLE',TOBFamD.getvalue('MAFA_LABEL'));
      end;

      // Traitement des familles niv 5
      end else if TOBFamD.getValue('MAFA_LEVEL')=5 then
      begin
        TOBN5D := TOBN5.findFirst (['BFG_IMPFAMILLE5'],
                                   [TOBFamD.getValue('MAFA_ID')],
                                     true);
        if TOBN5D = nil then
        begin
          TOBN4P := TOBN4.FindFirst(['BFG_IMPFAMILLE4'],[TOBFamD.getValue('MAFA_PID')],true);
          if TOBN4P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBFamD.getvalue('MAFA_ID')),5);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 5
          end;
          TOBN3P := TOBN3.findFirst(['BFG_IMPFAMILLE3'],[TOBN4P.getvalue('BFG_IMPFAMILLE3')],true);
          If TOBN3P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBFamD.getvalue('MAFA_PID')),4 );
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBN2P := TOBN2.findFirst(['BFG_IMPFAMILLE2'],[TOBN3P.getValue('BFG_IMPFAMILLE2')],true);
          If TOBN2P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBN3P.getvalue('BFG_IMPFAMILLE3')),3);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBN1P := TOBN1.findFirst(['BFG_IMPFAMILLE1'],[TOBN2P.getValue('BFG_IMPFAMILLE1')],true);
          If TOBN1P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBN2P.getvalue('BFG_IMPFAMILLE2')),2);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 2
          end;

          TOBN5D := TOB.Create('BFAMILLECHANGE',TOBN5,-1);
          TOBN5D.PutValue('BFG_TYPEIMPORT','BAT');
          TOBN5D.PutValue('BFG_PARFOU','');
          TOBN5D.PutValue('BFG_TYPEARTICLE','MAR');
          TOBN5D.PutValue('BFG_IMPFAMILLE1',TOBN1P.GetValue('BFG_IMPFAMILLE1'));
          TOBN5D.PutValue('BFG_IMPFAMILLE2',TOBN2P.getvalue('BFG_IMPFAMILLE2'));
          TOBN5D.PutValue('BFG_IMPFAMILLE3',TOBN3P.getvalue('BFG_IMPFAMILLE3'));
          TOBN5D.PutValue('BFG_IMPFAMILLE4',TOBFamD.getvalue('MAFA_PID'));
          TOBN5D.PutValue('BFG_IMPFAMILLE5',TOBFamD.getvalue('MAFA_ID'));
          TOBN5D.PutValue('BFG_FAMILLE','');
          TOBN5D.PutValue('BFG_FAMILLEPERE1',TOBN1P.GetValue('BFG_FAMILLE'));
          TOBN5D.PutValue('BFG_FAMILLEPERE2',TOBN2P.GetValue('BFG_FAMILLE'));
          TOBN5D.PutValue('BFG_FAMILLEPERE3',TOBN3P.GetValue('BFG_FAMILLE'));
          TOBN5D.AddChampSupValeur ('LIBELLE',TOBFamD.getvalue('MAFA_LABEL'));
        end else
        begin
          TOBN4P := TOBN4.FindFirst(['BFG_IMPFAMILLE4'],[TOBFamD.getValue('MAFA_PID')],true);
          if TOBN4P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBFamD.getvalue('MAFA_ID')),5);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 5
          end;
          TOBN3P := TOBN3.findFirst(['BFG_IMPFAMILLE3'],[TOBN4P.getvalue('BFG_IMPFAMILLE3')],true);
          If TOBN3P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBFamD.getvalue('MAFA_PID')),4);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBN2P := TOBN2.findFirst(['BFG_IMPFAMILLE2'],[TOBN3P.getValue('BFG_IMPFAMILLE2')],true);
          If TOBN2P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBN3P.getvalue('BFG_IMPFAMILLE3')),3);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBN1P := TOBN1.findFirst(['BFG_IMPFAMILLE1'],[TOBN2P.getValue('BFG_IMPFAMILLE1')],true);
          If TOBN1P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(10,string(TOBN2P.getvalue('BFG_IMPFAMILLE2')),2);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 2
          end;
          TOBN5D.PutValue('BFG_IMPFAMILLE1',TOBN1P.GetValue('BFG_IMPFAMILLE1'));
          TOBN5D.PutValue('BFG_IMPFAMILLE2',TOBN2P.getvalue('BFG_IMPFAMILLE2'));
          TOBN5D.PutValue('BFG_IMPFAMILLE3',TOBN3P.getvalue('BFG_IMPFAMILLE3'));
          TOBN5D.PutValue('BFG_IMPFAMILLE4',TOBFamD.getvalue('MAFA_PID'));
          TOBN5D.AddChampSupValeur ('LIBELLE',TOBFamD.getvalue('MAFA_LABEL'));
        end;
    end;
    inc(indice);
    until Indice >= TOBFam.detail.count;
  end;
// Intégration dans la table BFAMILLECHANGE
	ExecuteSQL('DELETE FROM BFAMILLECHANGE ' +
  					 'WHERE BFG_TYPEIMPORT="BAT" AND BFG_TYPEARTICLE="MAR"');
  TOBN1.InsertDb(NIL,FALSE);
  TOBN2.InsertDb(NIL,FALSE);
  TOBN3.InsertDb(NIL,FALSE);
  TOBN4.insertdb(NIL,FALSE);
	TOBN5.insertdb(NIL,FALSE);
// Intégration des informations dans les tablettes par NIVEAU
		TOBCC:=TOB.Create('les codes',NIL,-1);
// NIV1
	For indice:=0 to TOBN1.detail.count-1 DO
  BEGIN
  	TOBN1D:=TOBN1.Detail[indice];
    TOBCCD:=TOB.Create('CHOIXCOD',TOBCC,-1);
    TOBCCD.PutValue('CC_TYPE','FN1');
		TOBCCD.PutValue('CC_CODE',TOBN1D.GetValue('BFG_FAMILLE'));
    TOBCCD.PutValue('CC_LIBELLE',TOBN1D.GetValue('LIBELLE'));
    TOBCCD.PutValue('CC_ABREGE',copy(TOBN1D.GetValue('LIBELLE'),1,17));
    TOBCCD.PutValue('CC_LIBRE','');
	END;
//  ExecuteSQL('DELETE FROM CHOIXCOD' +
//  					 ' WHERE CC_TYPE ="FN1"');
	For indice:=0 to TOBCC.detail.count-1 DO
  begin
    TOBCC.detail[indice].InsertOrUpdateDB(FALSE);
  end;
  TOBCC.ClearDetail;
//NIV2
  For indice:=0 to TOBN2.detail.count-1 DO
  BEGIN
  	TOBN2D:=TOBN2.Detail[indice];
    TOBCCD:=TOB.Create('CHOIXCOD',TOBCC,-1);
    TOBCCD.PutValue('CC_TYPE','FN2');
		TOBCCD.PutValue('CC_CODE',TOBN2D.GetValue('BFG_FAMILLE'));
    TOBCCD.PutValue('CC_LIBELLE',TOBN2D.GetValue('LIBELLE'));
    TOBCCD.PutValue('CC_ABREGE',copy(TOBN2D.GetValue('LIBELLE'),1,17));
    TOBCCD.PutValue('CC_LIBRE',TOBN2D.GetValue('BFG_FAMILLEPERE1'));
	END;
//  ExecuteSQL('DELETE FROM CHOIXCOD' +
//  					 ' WHERE CC_TYPE ="FN2"');
	For indice:=0 to TOBCC.detail.count-1 DO
  begin
    TOBCC.detail[indice].InsertOrUpdateDB(FALSE);
  end;
// NIV3
  For indice:=0 to TOBN3.detail.count-1 DO
  BEGIN
  	TOBN3D:=TOBN3.Detail[indice];
    TOBCCD:=TOB.Create('CHOIXCOD',TOBCC,-1);
    TOBCCD.PutValue('CC_TYPE','FN3');
		TOBCCD.PutValue('CC_CODE',TOBN3D.GetValue('BFG_FAMILLE'));
    TOBCCD.PutValue('CC_LIBELLE',TOBN3D.GetValue('LIBELLE'));
    TOBCCD.PutValue('CC_ABREGE',copy(TOBN3D.GetValue('LIBELLE'),1,17));
    TOBCCD.PutValue('CC_LIBRE',TOBN3D.GetValue('BFG_FAMILLEPERE1')+TOBN3D.GetValue('BFG_FAMILLEPERE2'));
	END;
//  ExecuteSQL('DELETE FROM CHOIXCOD' +
//  					 ' WHERE CC_TYPE ="FN3"');
	For indice:=0 to TOBCC.detail.count-1 DO
  begin
    TOBCC.detail[indice].InsertOrUpdateDB(FALSE);
  end;
  TOBN1.free;
  TOBN2.free;
  TOBN3.Free;
  TOBN4.free;
  TOBN5.free;
  TOBFam.Free;
	Terreur.SetError(2,'OK','');
end;

procedure TFIntegreBatiprix.integrationFamilleOuvrage;

  function IncCodeFamilleOuv (Niveau : Integer; Var valeur : string) : Boolean;
  var okok : Boolean;
      Sql,FNiv,Oniv : string;
  begin
    FNiv := 'BO'+InttoStr(Niveau);
    //
    okok := false;
    while Not okok do
    begin
      result :=  IncStringCode (VALEUR,3);
      if not result then break;
      // verification sur famille artible
      Sql := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+FNiv+'" AND CC_CODE="'+valeur+'"';
      if not ExisteSQL(Sql) then okok := True;
    end;
  end;
  
var
	Indice : integer;
	MonFich : string;
	TOBOUVFAM,TOBOUVfamD : TOB;
  TOBOUVN1,TOBOUVN2,TOBOUVN3,TOBOUVN4 : TOB;
  TOBOUVN1D,TOBOUVN2D,TOBOUVN3D,TOBOUVN4D : TOB;
  TOBOUVN1P,TOBOUVN2P,TOBOUVN3P,TOBOUVLast : TOB;
  TOBOUVCC,TOBOUVCCD,TOBCED : TOB;
	requete : string;
  valeur : string;
begin
	THTRAITEMENT.Caption := 'Traitement des FAMILLES OUVRAGES';
  self.refresh;
	// initialisation
	TOBOUVFAM:= TOB.create('LES FAMILLES OUVRAGES',NIL,-1);
  TOBOUVN1 := TOB.Create ('LES NIVEAUX OUVRAGES 1',nil,-1);
  TOBOUVN2 := TOB.Create ('LES NIVEAUX OUVRAGES 2',nil,-1);
  TOBOUVN3 := TOB.Create ('LES NIVEAUX OUVRAGES 3',nil,-1);
  TOBOUVN4 := TOB.Create ('LES NIVEAUX OUVRAGES 5',nil,-1);
  //  Chargement des TOBN1, TOBN2, TOBN3, TOBN4 contenant les enregistrement selon les  niveaux famille à partir de la table BFAMILLECHANGE
  requete := 'SELECT * FROM BFAMILLECHANGE WHERE BFG_TYPEIMPORT="BAT" AND '+
  					 'BFG_TYPEARTICLE="OUV" AND BFG_IMPFAMILLE2="" ORDER BY BFG_IMPFAMILLE1';
  TOBOUVN1.LoadDetailDBFromSQL ('BFAMILLECHANGE',requete,false);

  requete := 'SELECT * FROM BFAMILLECHANGE WHERE BFG_TYPEIMPORT="BAT" AND '+
  					 'BFG_TYPEARTICLE="OUV" AND BFG_IMPFAMILLE2<>"" AND BFG_IMPFAMILLE3="" ORDER BY BFG_IMPFAMILLE1,BFG_IMPFAMILLE2';
  TOBOUVN2.LoadDetailDBFromSQL ('BFAMILLECHANGE',requete,false);

  requete := 'SELECT * FROM BFAMILLECHANGE WHERE BFG_TYPEIMPORT="BAT" AND '+
  					 'BFG_TYPEARTICLE="OUV" AND BFG_IMPFAMILLE3<>"" AND BFG_IMPFAMILLE4="" ORDER BY BFG_IMPFAMILLE1,BFG_IMPFAMILLE2,BFG_IMPFAMILLE3';
  TOBOUVN3.LoadDetailDBFromSQL ('BFAMILLECHANGE',requete,false);

  requete := 'SELECT * FROM BFAMILLECHANGE WHERE BFG_TYPEIMPORT="BAT" AND '+
  					 'BFG_TYPEARTICLE="OUV" AND BFG_IMPFAMILLE4<>"" ORDER BY BFG_IMPFAMILLE1,BFG_IMPFAMILLE2,BFG_IMPFAMILLE3,BFG_IMPFAMILLE4';
  TOBOUVN4.LoadDetailDBFromSQL ('BFAMILLECHANGE',requete,false);

  // Lecture fichier séquentiel
	MonFich := RepInst+'OUVR_FAMI.seq';
  TheContainer.OpenFile(MonFich);
  Terreur.SetError(1,'Famille_Ouvrages','');
  repeat
  	TOBOUVFamD := TheContainer.ReadFile;
    If TOBOUVFamD <> nil then
    begin
			If TOBOUVFamD.GetValue('OUFA_LEVEL') < 5 then
    	begin
      	TOBOUVFamD.ChangeParent(TOBOUVFAM, -1);
    	end;
  	end else
    begin
     	Terreur.SetError(3,MonFich,'');
     	break;
    end;
  until TheContainer.EofFile;
  TheContainer.FermeFile;
  // Tri sur les niveaux dans la TOB
  TOBOUVFAM.Detail.Sort('OUFA_LEVEL');

  // Enregistrement du corps d'etat dans le niveau 1 des familles d'ouvrages
  for Indice := 0 to TOBCE.detail.count -1 do
  begin
  	TOBCED := TOBCE.detail[Indice];
    TOBOUVN1D:= TOBOUVN1.FindFirst(['BFG_IMPFAMILLE1'],[TOBCED.getValue('CEIN_ID')],TRUE);
    If TOBOUVN1D = nil then
    begin
      if TOBOUVN1.detail.Count > 0 then
      begin
        TOBOUVLast := TOBOUVN1.detail[TOBOUVN1.detail.count-1];
        valeur:=TOBOUVLast.getvalue('BFG_FAMILLE')
      end else
      begin
        valeur:='000';
      end;
      TOBOUVN1D := TOB.Create('BFAMILLECHANGE',TOBOUVN1,-1);
      TOBOUVN1D.PutValue('BFG_TYPEIMPORT','BAT');
      TOBOUVN1D.PutValue('BFG_PARFOU','');
      TOBOUVN1D.PutValue('BFG_TYPEARTICLE','OUV');
      TOBOUVN1D.PutValue('BFG_IMPFAMILLE1',TOBCED.getvalue('CEIN_ID'));
      //if IncStringCode(valeur,3) then
      if IncCodeFamilleOuv (1,valeur) then
      begin
        TOBOUVN1D.PutValue('BFG_FAMILLE',Valeur);
      end;
      TOBOUVN1D.AddChampSupValeur ('LIBELLE',TOBCED.getvalue('CEIN_LABEL'));
      TOBOUVN1D.PutValue('BFG_CDETAT',TOBCED.getvalue('CEIN_ID'));
    end else
    begin
      TOBOUVN1D.AddChampSupValeur ('LIBELLE',TOBCED.getvalue('CEIN_LABEL'));
    end;
  end;
  //
  Indice := 0;
  if TOBOUVFAM.detail.count > 0 then
  begin
    repeat
      TOBOUVFamD := TOBOUVFAM.Detail[Indice];

      // Taitement des familles NIV 1 dans TOBOUVN1
      if TOBOUVFamD.getValue('OUFA_LEVEL')=1 then
      begin
        TOBOUVN1D:= TOBOUVN1.FindFirst(['BFG_IMPFAMILLE1'],[TOBOUVFamD.getValue('OUFA_ID')],TRUE);
        If TOBOUVN1D = nil then
        begin
          if TOBOUVN1.detail.Count > 0 then
          begin
            TOBOUVLast := TOBOUVN1.detail[TOBOUVN1.detail.count-1];
            valeur:=TOBOUVLast.getvalue('BFG_FAMILLE')
          end else
          begin
            valeur:='000';
          end;
          TOBOUVN1D := TOB.Create('BFAMILLECHANGE',TOBOUVN1,-1);
          TOBOUVN1D.PutValue('BFG_TYPEIMPORT','BAT');
          TOBOUVN1D.PutValue('BFG_PARFOU','');
          TOBOUVN1D.PutValue('BFG_TYPEARTICLE','OUV');
          TOBOUVN1D.PutValue('BFG_IMPFAMILLE1',TOBOUVFamD.getvalue('OUFA_ID'));
//          if IncStringCode(valeur,3) then
          if IncCodeFamilleOuv (1,valeur) then
          begin
            TOBOUVN1D.PutValue('BFG_FAMILLE',Valeur);
          end;
          TOBOUVN1D.AddChampSupValeur ('LIBELLE',TOBOUVFamD.getvalue('OUFA_LABEL'));
          TOBOUVN1D.PutValue('BFG_CDETAT',TOBOUVFamD.getvalue('CEIN_ID'));
        end else
        begin
          TOBOUVN1D.AddChampSupValeur ('LIBELLE',TOBOUVFamD.getvalue('OUFA_LABEL'));
        end;

      // Taitement des familles NIV 2 dans TOBOUVN2
      end else if TOBOUVFamD.getValue('OUFA_LEVEL')=2 then
      begin
        // recherche du pere pour le NIV 2 dans TOBN2
        TOBOUVN2D := TOBOUVN2.findFirst (['BFG_IMPFAMILLE2'],
                                     [TOBOUVFamD.getValue('OUFA_ID')],
                                     true);
        if TOBOUVN2D = nil then
        begin
          TOBOUVN1P := TOBOUVN1.findFirst(['BFG_IMPFAMILLE1'],[TOBOUVFamD.getValue('CEIN_ID')],true);
          If TOBOUVN1P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(20,string(TOBOUVFamD.getvalue('OUFA_ID')),2);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 2
          end;
          // -----------------------------------------------------------------
          if TOBOUVN2.detail.Count > 0 then
          begin
            TOBOUVLast := TOBOUVN2.detail[TOBOUVN2.detail.count-1];
            valeur:=TOBOUVLast.getvalue('BFG_FAMILLE')
          end else
          begin
            valeur:='000';
          end;
          TOBOUVN2D := TOB.Create('BFAMILLECHANGE',TOBOUVN2,-1);
          TOBOUVN2D.PutValue('BFG_TYPEIMPORT','BAT');
          TOBOUVN2D.PutValue('BFG_PARFOU','');
          TOBOUVN2D.PutValue('BFG_TYPEARTICLE','OUV');
          TOBOUVN2D.PutValue('BFG_IMPFAMILLE1',TOBOUVFamD.getvalue('CEIN_ID'));
          TOBOUVN2D.PutValue('BFG_IMPFAMILLE2',TOBOUVFamD.getvalue('OUFA_ID'));
//          if IncStringCode(valeur,3) then
          if IncCodeFamilleOuv (2,valeur) then
          begin
            TOBOUVN2D.PutValue('BFG_FAMILLE',Valeur);
          end;
          TOBOUVN2D.PutValue('BFG_FAMILLEPERE1',TOBOUVN1P.GetValue('BFG_FAMILLE'));
          TOBOUVN2D.AddChampSupValeur ('LIBELLE',TOBOUVFamD.getvalue('OUFA_LABEL'));
          TOBOUVN2D.PutValue('BFG_CDETAT',TOBOUVFamD.getvalue('CEIN_ID'));
        end else
        begin
          TOBOUVN2D.PutValue('BFG_IMPFAMILLE1',TOBOUVFamD.getValue('CEIN_ID'));
          TOBOUVN2D.AddChampSupValeur ('LIBELLE',TOBOUVFamD.getvalue('OUFA_LABEL'));
        end;

      //  Taitement des familles NIV 3 dans TOBOUVN3
      end else if TOBOUVFamD.getValue('OUFA_LEVEL')=3 then
      begin
        TOBOUVN3D := TOBOUVN3.findFirst (['BFG_IMPFAMILLE3'],
                                   [TOBOUVFamD.getValue('OUFA_ID')],
                                     true);
        if TOBOUVN3D = nil then
        begin
          TOBOUVN2P := TOBOUVN2.findFirst(['BFG_IMPFAMILLE2'],[TOBOUVFamD.getValue('OUFA_PID')],true);
          If TOBOUVN2P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(20,string(TOBOUVFamD.getvalue('OUFA_ID')),3);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBOUVN1P := TOBOUVN1.findFirst(['BFG_IMPFAMILLE1'],[TOBOUVN2P.getValue('BFG_IMPFAMILLE1')],true);
          If TOBOUVN1P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(20,string(TOBOUVFamD.getvalue('OUFA_PID')),2);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 2
          end;

          if TOBOUVN3.detail.Count > 0 then
          begin
            TOBOUVLast := TOBOUVN3.detail[TOBOUVN3.detail.count-1];
            valeur:=TOBOUVLast.getvalue('BFG_FAMILLE')
          end else
          begin
            valeur:='000';
          end;
          TOBOUVN3D := TOB.Create('BFAMILLECHANGE',TOBOUVN3,-1);
          TOBOUVN3D.PutValue('BFG_TYPEIMPORT','BAT');
          TOBOUVN3D.PutValue('BFG_PARFOU','');
          TOBOUVN3D.PutValue('BFG_TYPEARTICLE','OUV');
          TOBOUVN3D.PutValue('BFG_IMPFAMILLE1',TOBOUVN1P.GetValue('BFG_IMPFAMILLE1'));
          TOBOUVN3D.PutValue('BFG_IMPFAMILLE2',TOBOUVFamD.getvalue('OUFA_PID'));
          TOBOUVN3D.PutValue('BFG_IMPFAMILLE3',TOBOUVFamD.getvalue('OUFA_ID'));
//          if IncStringCode(valeur,3) then
          if IncCodeFamilleOuv (3,valeur) then
          begin
            TOBOUVN3D.PutValue('BFG_FAMILLE',Valeur);
          end;
          TOBOUVN3D.PutValue('BFG_FAMILLEPERE1',TOBOUVN1P.GetValue('BFG_FAMILLE'));
          TOBOUVN3D.PutValue('BFG_FAMILLEPERE2',TOBOUVN2P.GetValue('BFG_FAMILLE'));
          TOBOUVN3D.AddChampSupValeur ('LIBELLE',TOBOUVFamD.getvalue('OUFA_LABEL'));
          TOBOUVN3D.PutValue('BFG_CDETAT',TOBOUVFamD.getvalue('CEIN_ID'));
        end else
        begin
          TOBOUVN2P := TOBOUVN2.findFirst(['BFG_IMPFAMILLE2'],[TOBOUVFamD.getValue('OUFA_PID')],true);
          If TOBOUVN2P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(20,string(TOBOUVFamD.getvalue('OUFA_ID')),3);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBOUVN1P := TOBOUVN1.findFirst(['BFG_IMPFAMILLE1'],[TOBOUVN2P.getValue('BFG_IMPFAMILLE1')],true);
          If TOBOUVN1P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(20,string(TOBOUVFamD.getvalue('OUFA_PID')),2);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          //
          TOBOUVN3D.PutValue('BFG_IMPFAMILLE1',TOBOUVN1P.GetValue('BFG_IMPFAMILLE1'));
          TOBOUVN3D.PutValue('BFG_IMPFAMILLE2',TOBOUVFamD.getvalue('OUFA_PID'));
          TOBOUVN3D.AddChampSupValeur ('LIBELLE',TOBOUVFamD.getvalue('OUFA_LABEL'));
        end;

      //  Traitement des familles ouvrage niv 4
      end else if TOBOUVFamD.getValue('OUFA_LEVEL')=4 then
      begin
        TOBOUVN4D := TOBOUVN4.findFirst (['BFG_IMPFAMILLE4'],
                                   [TOBOUVFamD.getValue('OUFA_ID')],
                                     true);
        if TOBOUVN4D = nil then
        begin
          TOBOUVN3P := TOBOUVN3.findFirst(['BFG_IMPFAMILLE3'],[TOBOUVFamD.getValue('OUFA_PID')],true);
          If TOBOUVN3P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(20,string(TOBOUVFamD.getvalue('OUFA_ID')),4);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 4
          end;
          TOBOUVN2P := TOBOUVN2.findFirst(['BFG_IMPFAMILLE2'],[TOBOUVN3P.getValue('BFG_IMPFAMILLE2')],true);
          If TOBOUVN2P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(20,string(TOBOUVFamD.getvalue('OUFA_PID')),3);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBOUVN1P := TOBOUVN1.findFirst(['BFG_IMPFAMILLE1'],[TOBOUVN2P.getValue('BFG_IMPFAMILLE1')],true);
          If TOBOUVN1P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(20,string(TOBOUVN2P.getvalue('BFG_IMPFAMILLE2')),2);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 2
          end;

          TOBOUVN4D := TOB.Create('BFAMILLECHANGE',TOBOUVN4,-1);
          TOBOUVN4D.PutValue('BFG_TYPEIMPORT','BAT');
          TOBOUVN4D.PutValue('BFG_PARFOU','');
          TOBOUVN4D.PutValue('BFG_TYPEARTICLE','OUV');
          TOBOUVN4D.PutValue('BFG_IMPFAMILLE1',TOBOUVN1P.GetValue('BFG_IMPFAMILLE1'));
          TOBOUVN4D.PutValue('BFG_IMPFAMILLE2',TOBOUVN2P.getvalue('BFG_IMPFAMILLE2'));
          TOBOUVN4D.PutValue('BFG_IMPFAMILLE3',TOBOUVFamD.getvalue('OUFA_PID'));
          TOBOUVN4D.PutValue('BFG_IMPFAMILLE4',TOBOUVFamD.getvalue('OUFA_ID'));
          TOBOUVN4D.PutValue('BFG_FAMILLE','');
          TOBOUVN4D.PutValue('BFG_FAMILLEPERE1',TOBOUVN1P.GetValue('BFG_FAMILLE'));
          TOBOUVN4D.PutValue('BFG_FAMILLEPERE2',TOBOUVN2P.GetValue('BFG_FAMILLE'));
          TOBOUVN4D.PutValue('BFG_FAMILLEPERE3',TOBOUVN3P.GetValue('BFG_FAMILLE'));
          TOBOUVN4D.AddChampSupValeur ('LIBELLE',TOBOUVFamD.getvalue('OUFA_LABEL'));
        	TOBOUVN4D.PutValue('BFG_CDETAT',TOBOUVFamD.getvalue('CEIN_ID'));
        end else
        begin
          TOBOUVN3P := TOBOUVN3.findFirst(['BFG_IMPFAMILLE3'],[TOBOUVFamD.getValue('OUFA_PID')],true);
          If TOBOUVN3P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(20,string(TOBOUVFamD.getvalue('MAFA_ID')),2);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBOUVN2P := TOBOUVN2.findFirst(['BFG_IMPFAMILLE2'],[TOBOUVN3P.getValue('BFG_IMPFAMILLE2')],true);
          If TOBOUVN2P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(20,string(TOBOUVN3P.getvalue('BFG_IMPFAMILLE2')),3);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBOUVN1P := TOBOUVN1.findFirst(['BFG_IMPFAMILLE1'],[TOBOUVN2P.getValue('BFG_IMPFAMILLE1')],true);
          If TOBOUVN1P = nil then
          // gestion des erreurs
          begin
            Terreur.SetError(20,string(TOBOUVN2P.getvalue('BFG_IMPFAMILLE1')),2);
            inc(indice);
            continue; // pas de pere rattache a cette famille de niveau 3
          end;
          TOBOUVN4D.PutValue('BFG_IMPFAMILLE1',TOBOUVN1P.GetValue('BFG_IMPFAMILLE1'));
          TOBOUVN4D.PutValue('BFG_IMPFAMILLE2',TOBOUVN2P.getvalue('BFG_IMPFAMILLE2'));
          TOBOUVN4D.PutValue('BFG_IMPFAMILLE3',TOBOUVFamD.getvalue('OUFA_PID'));
          TOBOUVN4D.AddChampSupValeur ('LIBELLE',TOBOUVFamD.getvalue('OUFA_LABEL'));
        end;
      end;
      inc(indice);
    until Indice >= TOBOUVFam.detail.count;
  end;
// Intégration dans la table BFAMILLECHANGE
	ExecuteSQL('DELETE FROM BFAMILLECHANGE ' +
  					 'WHERE BFG_TYPEIMPORT="BAT" AND BFG_TYPEARTICLE="OUV"');
  TOBOUVN1.InsertDb(NIL,FALSE);
  TOBOUVN2.InsertDb(NIL,FALSE);
  TOBOUVN3.InsertDb(NIL,FALSE);
  TOBOUVN4.insertdb(NIL,FALSE);

// Intégration des informations dans les tablettes par NIVEAU
		TOBOUVCC:=TOB.Create('les codes ouvrage',NIL,-1);
// NIV1
	For indice:=0 to TOBOUVN1.detail.count-1 DO
  BEGIN
  	TOBOUVN1D:=TOBOUVN1.Detail[indice];
    TOBOUVCCD:=TOB.Create('CHOIXCOD',TOBOUVCC,-1);
    TOBOUVCCD.PutValue('CC_TYPE','BO1');
		TOBOUVCCD.PutValue('CC_CODE',TOBOUVN1D.GetValue('BFG_FAMILLE'));
    TOBOUVCCD.PutValue('CC_LIBELLE',TOBOUVN1D.GetValue('LIBELLE'));
    TOBOUVCCD.PutValue('CC_ABREGE',copy(TOBOUVN1D.GetValue('LIBELLE'),1,17));
    TOBOUVCCD.PutValue('CC_LIBRE','');
	END;
//  ExecuteSQL('DELETE FROM CHOIXCOD ' +
//  					 'WHERE CC_TYPE ="BO1"');
	For indice:=0 to TOBOUVCC.detail.count-1 DO
  begin
    TOBOUVCC.detail[indice].InsertOrUpdateDB(FALSE);
  end;
  TOBOUVCC.ClearDetail;
// NIV2
  For indice:=0 to TOBOUVN2.detail.count-1 DO
  BEGIN
  	TOBOUVN2D:=TOBOUVN2.Detail[indice];
    TOBOUVCCD:=TOB.Create('CHOIXCOD',TOBOUVCC,-1);
    TOBOUVCCD.PutValue('CC_TYPE','BO2');
		TOBOUVCCD.PutValue('CC_CODE',TOBOUVN2D.GetValue('BFG_FAMILLE'));
    TOBOUVCCD.PutValue('CC_LIBELLE',TOBOUVN2D.GetValue('LIBELLE'));
    TOBOUVCCD.PutValue('CC_ABREGE',copy(TOBOUVN2D.GetValue('LIBELLE'),1,17));
    TOBOUVCCD.PutValue('CC_LIBRE',TOBOUVN2D.GetValue('BFG_FAMILLEPERE1'));
	END;
//  ExecuteSQL('DELETE FROM CHOIXCOD ' +
//  					 'WHERE CC_TYPE ="BO2"');
	For indice:=0 to TOBOUVCC.detail.count-1 DO
  begin
    TOBOUVCC.detail[indice].InsertOrUpdateDB(FALSE);
  end;
  TOBOUVCC.ClearDetail;
// NIV3
  For indice:=0 to TOBOUVN3.detail.count-1 DO
  BEGIN
  	TOBOUVN3D:=TOBOUVN3.Detail[indice];
    TOBOUVCCD:=TOB.Create('CHOIXCOD',TOBOUVCC,-1);
    TOBOUVCCD.PutValue('CC_TYPE','BO3');
		TOBOUVCCD.PutValue('CC_CODE',TOBOUVN3D.GetValue('BFG_FAMILLE'));
    TOBOUVCCD.PutValue('CC_LIBELLE',TOBOUVN3D.GetValue('LIBELLE'));
    TOBOUVCCD.PutValue('CC_ABREGE',copy(TOBOUVN3D.GetValue('LIBELLE'),1,17));
    TOBOUVCCD.PutValue('CC_LIBRE',TOBOUVN3D.GetValue('BFG_FAMILLEPERE1')+TOBOUVN3D.GetValue('BFG_FAMILLEPERE2'));
	END;
//  ExecuteSQL('DELETE FROM CHOIXCOD ' +
//  					 'WHERE CC_TYPE ="BO3"');
	For indice:=0 to TOBOUVCC.detail.count-1 DO
  begin
    TOBOUVCC.detail[indice].InsertOrUpdateDB(FALSE);
  end;
  TOBOUVCC.ClearDetail;
  TOBOUVN1.free;
  TOBOUVN2.free;
  TOBOUVN3.Free;
  TOBOUVN4.Free;
	TOBOUVFam.Free;
	Terreur.SetError(2,'OK','');
end;


function TFIntegreBatiprix.FindInRef(TOBREF: TOB;CodeArticle : string) : TOB;
var QQ : TQuery;
		RefArticle : string;
begin
  RefArticle := CodeArticleUnique(CodeArticle,'','','','','');
	Result := TOBGetInto(TOBREF,['GA_ARTICLE'],[RefArticle]);
  if Result = nil then
  begin
		TOBREF.ClearDetail;
    QQ := OpenSql ('SELECT ##TOP 1000## GA_ARTICLE,GA_CODEARTICLE,GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,'+
                   'GA_LIBELLE,GA_DATECREATION,GA_PVHT '+
                   'FROM ARTICLE WHERE GA_ARTICLE>="'+RefArticle+'" ORDER BY GA_ARTICLE',true);
    TOBREF.loadDetailDb ('ARTICLE','','',QQ,false,false);
    Ferme(QQ);
		Result := TOBGetInto(TOBREF,['GA_ARTICLE'],[RefArticle]);
  end;
end;

procedure TFIntegreBatiprix.integrationMateriaux;
var
	genre,requete : string;
	MonFich : string;
	TOBMATD : TOB;
  TOBARTSI,TOBARTSM,TOBARTSD,TOBA : TOB;
	TOBFAMCHG,TOBFAMCHGD : TOB;
  TOBGENREMATD : TOB;
  TOBARTREF,TOBARTSF : TOB;
  QQ : TQuery;
  Indice : Integer;
  Famille1,Famille2,Famille3 : string;
  TitreLib : String;
  Compteur : integer;
begin
  Compteur := 0;
  THTRAITEMENT.caption := 'Traitement des MATERIAUX';
  self.refresh;
  // création des TOB
	TOBARTSI := TOB.Create ('LES ARTICLES A INSERER',nil,-1);
	TOBARTSM := TOB.Create ('LES ARTICLES A MODIFIER',nil,-1);
	TOBFAMCHG:= TOB.create ('LES FAMILLES',nil,-1);
  TOBARTREF := TOB.Create ('LES ARTICLES EN TABLE',nil,-1);

	// 	Chargement de TOBFAMCHG
  requete := 'SELECT * FROM BFAMILLECHANGE WHERE BFG_IMPFAMILLE5 <> ""';
  TOBFAMCHG.loadDetailDBFromSQL ('BFAMILLE',requete,false);
  (*
  QQ := OpenSql ('SELECT GA_ARTICLE,GA_CODEARTICLE,GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,'+
               'GA_LIBELLE,GA_DATECREATION,GA_PVHT '+
               'FROM ARTICLE WHERE GA_TYPEARTICLE<>"OUV"',true);
  TOBARTREF.loadDetailDb ('ARTICLE','','',QQ,false,false);
  ferme (QQ);
  *)
  // Lecture fichier séquentiel
	MonFich := RepInst+'MATE.seq';
  TheContainer.OpenFile(MonFich);
  Terreur.SetError(1,'Matériaux','');
  repeat
		TOBMATD := TheContainer.ReadFile;
    if TOBMATD <> nil then
    begin
      if TOBMATD.getvalue('MATE_BATREF')='' then BEGIN TOBMATD.free; continue; End;
		// recherche qualification quantité dans TOBQLFQTE
    	TOBQLFQTED := TOBQLFQTE.FindFirst(['GME_MESURE'],[TOBMATD.GetValue('MATE_UNFT_COD')],true);
      if TOBQLFQTED <> nil then
      begin
      	TOBQLFQTED.PutValue('GME_QUALIFMESURE','PIE');
      end else
      begin
      	TOBQLFQTED := TOB.Create('MEA',TOBQLFQTE,-1);
        TOBQLFQTED.PutValue('GME_QUALIFMESURE','PIE');
        TOBQLFQTED.PutValue('GME_MESURE',TOBMATD.GetValue('MATE_UNFT_COD'));
        TOBQLFQTED.PutValue('GME_LIBELLE',TOBMATD.GetValue('MATE_UNFT_COD'));
        TOBQLFQTED.PutValue('GME_QUOTITE','1');
      end;
    	genre := TOBMATD.getvalue('TFRT_COD');
      // recherche d'un GENRE dans la TOBGENREMAT
      TOBGENREMATD := TOBGENREMAT.findFirst(['BGM_GENRE'],[genre],true);
      if TOBGENREMATD <> nil then
      begin
        // position dans TOBFAMCHGD pour déterminer les niveaux pères
        TOBFAMCHGD := TOBFAMCHG.FindFirst(['BFG_IMPFAMILLE5'],[TOBMATD.GetValue('MAFA_ID')],true);
        if TOBFAMCHGD = nil then
        begin
        	TOBFAMCHGD := TOBFAMCHG.FindFirst(['BFG_IMPFAMILLE4','BFG_IMPFAMILLE5'],[TOBMATD.GetValue('MAFA_ID'),''],true);
        end;
        if TOBFAMCHGD = nil then
        begin
        	TOBFAMCHGD := TOBFAMCHG.FindFirst(['BFG_IMPFAMILLE3','BFG_IMPFAMILLE4','BFG_IMPFAMILLE5'],[TOBMATD.GetValue('MAFA_ID'),'',''],true);
        end;

        if TOBFAMCHGD = nil then
        begin
          TOBFAMCHGD := FindFamilleBATIPRIX (TOBFAMCHG,TOBMATD.GetValue('MAFA_ID'),'MAR');
        end;

        if TOBFAMCHGD = nil then
        begin
          Terreur.SetError(30,string(TOBMATD.getvalue('MATE_ID')),'');
          Famille1 := '';
          Famille2 := '';
          Famille3 := '';
        end else
        begin
          Famille1 := TOBFAMCHGD.getvalue('BFG_FAMILLEPERE1');
          Famille2 := TOBFAMCHGD.getvalue('BFG_FAMILLEPERE2');
          Famille3 := TOBFAMCHGD.getvalue('BFG_FAMILLEPERE3');
        end;
        
        	TitreLib := copy(TOBMATD.getvalue('MATE_LIB_LONG'),1,70);
//        TOBARTSF := TOBARTREF.FindFirst(['GA_CODEARTICLE'],[TOBMATD.getvalue('MATE_BATREF')],true);
        TOBARTSF := FindInRef(TOBARTREF,TOBMATD.getvalue('MATE_BATREF'));
        if TOBARTSF = nil then
        begin
          TOBARTSD := TOB.Create ('ARTICLE',TOBARTSI,-1);
          TOBARTSD.PutValue('GA_ARTICLE',CodeArticleUnique(TOBMATD.getvalue('MATE_BATREF'),'','','','',''));
          TOBARTSD.PutValue('GA_CODEARTICLE',TOBMATD.getvalue('MATE_BATREF'));
          TOBARTSD.PutValue('GA_TYPEARTICLE',TOBGENREMATD.getvalue('BGM_TYPEARTICLE'));
          TOBARTSD.PutValue('GA_FAMILLENIV1',Famille1);
          TOBARTSD.PutValue('GA_FAMILLENIV2',Famille2);
          TOBARTSD.PutValue('GA_FAMILLENIV3',Famille3);
          TOBARTSD.PutValue('GA_LIBELLE',TitreLib);
          TOBARTSD.PutValue('GA_BLOCNOTE',TOBMATD.getvalue('MATE_LIB_LONG'));
          TOBARTSD.PutValue('GA_NATUREPRES',TOBGENREMATD.getvalue('BGM_NATUREPRES'));
          TOBARTSD.PutValue('GA_QUALIFUNITEVTE',TOBMATD.getvalue('MATE_UNFT_COD'));
          TOBARTSD.PutValue('GA_QUALIFUNITESTO',TOBMATD.getvalue('MATE_UNFT_COD'));
          TOBARTSD.PutValue('GA_STATUTART','UNI');
          TOBARTSD.PutValue('GA_QUALIFMARGE','CO');
          TOBARTSD.PutValue('GA_ACTIVITEREPRISE','F');
          TOBARTSD.PutValue('GA_PAHT',TOBMATD.getvalue('MATE_PRIX'));
          TOBARTSD.PutValue('GA_PRIXPOURQTE',1);
          TOBARTSD.PutValue('GA_DATESUPPRESSION',Idate2099);
          TOBARTSD.PutValue('GA_FAMILLETAXE1','TN');
          TOBARTSD.PutValue('GA_COEFFG',1);
          TOBARTSD.PutValue('GA_CALCPRIXHT','DPR');
          TOBARTSD.PutValue('GA_CALCPRIXPR','PAA');
          TOBARTSD.PutValue('GA_DPRAUTO','X');
          TOBARTSD.PutValue('GA_CALCAUTOHT','X');
          TOBARTSD.PutValue('GA_COEFCALCHT',1);
          TOBARTSD.PutValue('GA_REMISEPIED','X');
          TOBARTSD.PutValue('GA_REMISELIGNE','X');
          TOBARTSD.PutValue('GA_ESCOMPTABLE','X');
          TOBARTSD.PutValue('GA_COMMISSIONNABLE','X');
          //
          TOBARTSD.PutValue('GA_DPA',TOBARTSD.GetValue('GA_PAHT'));
          TOBARTSD.PutValue('GA_DPR',TOBARTSD.GetValue('GA_PAHT'));
          TOBARTSD.PutValue('GA_PVHT',TOBARTSD.GetValue('GA_PAHT'));
          //
          TOBARTSD.SetAllModifie (true);
        end else
        begin
          TOBARTSD := TOB.Create ('ARTICLE',TOBARTSM,-1);
          TOBARTSD.Dupliquer (TOBARTSF,false,true,true);
          if TOBARTSD.GetValue('GA_FAMILLENIV1')<> Famille1 then
          	TOBARTSD.PutValue('GA_FAMILLENIV1',Famille1);
          if TOBARTSD.GetValue('GA_FAMILLENIV2')<> Famille2 then
          	TOBARTSD.PutValue('GA_FAMILLENIV2',Famille2);
          if TOBARTSD.GetValue('GA_FAMILLENIV3')<> Famille3 then
          	TOBARTSD.PutValue('GA_FAMILLENIV3',Famille3);
          if TOBARTSD.GetValue('GA_LIBELLE') <> copy(TOBMATD.getvalue('MATE_LIB_LONG'),1,70) then
          	TOBARTSD.PutValue('GA_LIBELLE',TitreLib);
          if TOBARTSD.GetValue('GA_DATECREATION') <> TOBMATD.getvalue('MATE_DATE_MAJ')then
          	TOBARTSD.PutValue('GA_DATECREATION',TOBMATD.getvalue('MATE_DATE_MAJ'));
          if TOBARTSD.GetValue('GA_PAHT')<>TOBMATD.getvalue('MATE_PRIX')then
          begin
          	TOBARTSD.PutValue('GA_PAHT',TOBMATD.getvalue('MATE_PRIX'));
            TOBARTSD.PutValue('GA_DPR',Arrondi(TOBARTSD.GetValue('GA_PAHT')*TOBARTSD.GetValue('GA_COEFFG'),V_PGI.okdecP));
            TOBARTSD.PutValue('GA_PVHT',TOBARTSD.GetValue('GA_PAHT'));
          end;

          if not TOBARTSD.IsOneModifie (false) then TOBARTSD.free else TOBARTSD.PutValue('GA_REMISELIGNE','X');
        end;
        //
        TOBA := TOB.Create ('UNE LIAISON ART',TOBDESARTICLES,-1);
        TOBA.AddChampSupValeur ('MATE_ID',TOBMATD.getValue('MATE_ID'));
        TOBA.AddChampSupValeur ('MATE_BATREF',TOBMATD.getvalue('MATE_BATREF'));
        TOBA.AddChampSupValeur ('LIBELLE',TitreLib);
        TOBA.AddChampSupValeur ('GA_ARTICLE',CodeArticleUnique(TOBMATD.getvalue('MATE_BATREF'),'','','','',''));
        //
        If TOBARTSI.Detail.Count >= 5000 then
        begin
        	Compteur := Compteur + 5000;
          THTRAITEMENT.Caption := 'Traitement des MATERIAUX: '+ IntToStr(Compteur) + ' enregistrements traités';
          self.refresh;
          //TOBARTSI.InsertDB(nil,true);
          TOBARTSI.InsertOrUpdateDB (true);
          TOBARTSI.ClearDetail;
        end;
        If TOBARTSM.Detail.Count >= 5000 then
        begin
        	Compteur := compteur + 5000;
          THTRAITEMENT.Caption := 'Traitement des MATERIAUX: ' + IntToStr(Compteur) + ' enregistrements traités' ;
          self.refresh;
          for Indice := 0 to TOBARTSM.detail.count -1 do
          begin
          	TOBARTSD := TOBARTSM.detail[Indice];
            if TOBARTSD.IsOneModifie (false) then
            begin
            	TOBARTSD.UpdateDB;
            end;
          end;
          //
        	TOBARTSM.ClearDetail;
        end;
      end else
      begin
        // gestion erreur
        Terreur.SetError(31,genre,'');
      end;
      TOBMATD.free;
    end else
    begin
    	Terreur.SetError(3,MonFich,'');
      break;
    end;
	until TheContainer.EofFile;
  if TOBARTSI.detail.count > 0 then
  begin
    TOBARTSI.InsertDB(nil,true);
    TOBARTSI.ClearDetail;
  end;
  if TOBARTSM.detail.count > 0 then
  begin
    for Indice := 0 to TOBARTSM.detail.count -1 do
    begin
      TOBARTSD := TOBARTSM.detail[Indice];
      if TOBARTSD.IsOneModifie (false) then
      begin
      	TOBARTSD.UpdateDB;
      end;
    end;
    TOBARTSM.ClearDetail;
  end;
  TheContainer.FermeFile;
  TOBARTREF.free;
  TOBFAMCHG.free;
  TOBARTSI.free;
  TOBARTSM.free;
	Terreur.SetError(2,'OK','');
end;

procedure TFIntegreBatiprix.integrationOuvrages;
var
	MonFich : string;
	TOBOuvrageD : TOB;
  TOBOUVARTS,TOBOUVARTSD : TOB;
  TOBENENT,TOBENENTD : TOB;   // table GNE_NOMENENT
	TOBENLIGD : TOB;  // table GNE_NOMENLIG
  TOBFAMCHG,TOBFAMCHGD : TOB;
  TOBOUVMATD,TOBOUVCOMPD : TOB;
  TOBO,TOBA,TOBOC : TOB;
  requete : string;
  CodeOuvComp,CodeOuvSimp : string;
  CodeArticle : string;
  Famille1,Famille2,Famille3,Cdetat : string;
  TitreMemo : wideString;
  TitreLib : string;
  OuvBATREF : string;
  Compteur : integer;
  NumLigne : integer;
begin
  Compteur := 0;
  THTRAITEMENT.caption := 'Traitement des OUVRAGES';
  self.refresh;
  // initialisation
  TOBOUVARTS := TOB.Create ('LES ARTICLES ',nil,-1);
	TOBFAMCHG:= TOB.create ('LES FAMILLES',nil,-1);
	TOBENENT:= TOB.create ('LES ENTETES OUVRAGES',nil,-1);
	// Chargement de TOBFAMCHG
  requete := 'SELECT * FROM BFAMILLECHANGE WHERE BFG_TYPEARTICLE="OUV" AND BFG_IMPFAMILLE3 <> ""';
  TOBFAMCHG.LoadDetailDBFromSQL ('BFAMILLECHANGE',requete,false);

  // lecture du fichier sequentiel
  MonFich := RepInst+'OUVR_OUVR.seq';
  TheContainer.OpenFile(MonFich);
  Terreur.SetError(1,'Ouvrages','');
  repeat
  	TOBOuvrageD := TheContainer.ReadFile;
    if TOBOuvrageD <> nil then
    begin
      if TOBOuvrageD.getValue('OUVR_BATREF')= '' then BEGIN TOBOuvrageD.free; continue; END;
		//recherche qualification quantité dans TOBQLFQTE
    	TOBQLFQTED := TOBQLFQTE.FindFirst(['GME_MESURE'],[TOBOuvrageD.GetValue('OUVR_UNFT_COD')],True);
      if TOBQLFQTED <> nil then
      begin
        TOBQLFQTED.PutValue('GME_QUALIFMESURE','PIE');
			end else
      begin
      	TOBQLFQTED := TOB.Create('MEA',TOBQLFQTE,-1);
        TOBQLFQTED.PutValue('GME_QUALIFMESURE','PIE');
        TOBQLFQTED.PutValue('GME_MESURE',TOBOuvrageD.GetValue('OUVR_UNFT_COD'));
        TOBQLFQTED.PutValue('GME_LIBELLE',TOBOuvrageD.GetValue('OUVR_UNFT_COD'));
        TOBQLFQTED.PutValue('GME_QUOTITE','1');
      end;
      TOBFAMCHGD := TOBFAMCHG.findfirst(['BFG_IMPFAMILLE4','BFG_IMPFAMILLE5'],[TOBOuvrageD.getvalue('OUFA_ID'),''],true);
    	if TOBFAMCHGD = nil then
    	Begin
				TOBFAMCHGD := TOBFAMCHG.findfirst(['BFG_IMPFAMILLE3','BFG_IMPFAMILLE4','BFG_IMPFAMILLE5'],[TOBOuvrageD.getvalue('OUFA_ID'),'',''],true);
      end;
      TitreLib :='';
      TitreMemo := '';
      Cdetat := Copy(TOBOuvrageD.getvalue('OUVR_BATREF'),1,2);

      if TOBFAMCHGD = nil then TOBFAMCHGD := FindFamilleBATIPRIX (TOBFAMCHG,TOBOuvrageD.GetValue('OUFA_ID'),'OUV');
      if TOBFAMCHGD = nil then
      begin
    		Terreur.SetError(40,string(TOBOuvrageD.getvalue('OUVR_ID')),'');
        Famille1 := '';
        Famille2 := '';
        Famille3 := '';
      end else
      begin
        Famille1 := TOBFAMCHGD.getvalue('BFG_FAMILLEPERE1');
        Famille2 := TOBFAMCHGD.getvalue('BFG_FAMILLEPERE2');
        Famille3 := TOBFAMCHGD.getvalue('BFG_FAMILLEPERE3');
      end;
      OuvBATREF := TOBOuvrageD.getValue('OUVR_BATREF');
      if TOBOuvrageD.GetValue('OUVR_DESC_DEB') = '' then
      begin
        TitreMemo := TOBOuvrageD.GetValue('OUVR_LIB_DISK');
        TitreLib := copy(TOBOuvrageD.GetValue('OUVR_LIB_DISK'),1,70);
      end else
      begin
        TitreMemo := TOBOuvrageD.GetValue('OUVR_DESC_DEB');
        TitreLib := copy(TOBOuvrageD.GetValue('OUVR_LIB_DISK'),1,70);
      end;
      TOBOUVARTSD := TOB.Create ('ARTICLE',TOBOUVARTS,-1);
      TOBOUVARTSD.PutValue('GA_ARTICLE',CodeArticleUnique(OuvBATREF,'','','','',''));
      TOBOUVARTSD.PutValue('GA_CODEARTICLE',OuvBATREF);
      TOBOUVARTSD.PutValue('GA_TYPEARTICLE','OUV');
      TOBOUVARTSD.PutValue('GA_FAMILLENIV1',Famille1);
      TOBOUVARTSD.PutValue('GA_FAMILLENIV2',Famille2);
      TOBOUVARTSD.PutValue('GA_FAMILLENIV3',Famille3);
      TOBOUVARTSD.PutValue('GA_LIBELLE',TitreLib);
      TOBOUVARTSD.PutValue('GA_BLOCNOTE',TitreMemo);
      TOBOUVARTSD.PutValue('GA_DATECREATION',TOBOuvrageD.getvalue('OUVR_DATE_MAJ'));
      TOBOUVARTSD.PutValue('GA_QUALIFUNITEVTE',TOBOuvrageD.getvalue('OUVR_UNFT_COD'));
      TOBOUVARTSD.putValue('GA_TYPENOMENC','OU1');
      TOBOUVARTSD.PutValue('GA_STATUTART','UNI');
      TOBOUVARTSD.PutValue('GA_QUALIFMARGE','CO');
      TOBOUVARTSD.PutValue('GA_ACTIVITEREPRISE','F');
      TOBOUVARTSD.PutValue('GA_PRIXPOURQTE',1);
      TOBOUVARTSD.PutValue('GA_DATESUPPRESSION',Idate2099);
      TOBOUVARTSD.PutValue('GA_FAMILLETAXE1','TN');
      TOBOUVARTSD.PutValue('GA_REMISEPIED','X');
      TOBOUVARTSD.PutValue('GA_DOMAINE',Cdetat);
      TOBOUVARTSD.PutValue('GA_ESCOMPTABLE','X');
      TOBOUVARTSD.PutValue('GA_COMMISSIONNABLE','X');
      TOBOUVARTSD.PutValue('GA_PRIXPASMODIF','X');
      //
      TOBENENTD := TOB.create('NOMENENT',TOBENENT,-1);
      TOBENENTD.PutValue('GNE_NOMENCLATURE',OuvBATREF);
      TOBENENTD.PutValue('GNE_LIBELLE',TitreLib);
      TOBENENTD.PutValue('GNE_ARTICLE',CodeArticleUnique(OuvBATREF,'','','','',''));
      TOBENENTD.PutValue('GNE_DATECREATION',TOBOuvrageD.getvalue('OUVR_DATE_MAJ'));
      TOBENENTD.PutValue('GNE_QTEDUDETAIL','1');
      TOBENENTD.PutValue('GNE_DOMAINE',Cdetat);

      //
      // Création d'une TOB fille VIRTUELLE de TOBDESOUVRAGES
      TOBO := TOB.Create ('UN LIEN OUVRAGE',TOBDESOUVRAGES,-1);
      TOBO.AddChampSupValeur ('OUVR_ID', TOBOuvrageD.getValue('OUVR_ID'));
      TOBO.AddChampSupValeur ('OUVR_BATREF', OuvBATREF);
      TOBO.AddChampSupValeur ('LIBELLE',TitreLib);
      TOBO.AddChampSupValeur ('GA_ARTICLE',CodeArticleUnique(TOBOuvrageD.getvalue('OUVR_BATREF'),'','','','',''));
      TOBO.AddChampSupValeur ('QTEMOA',TOBOuvrageD.getValue('TU_MOA'));
      TOBO.AddChampSupValeur ('QTEMOC',TOBOuvrageD.getValue('TU_MOC'));
      TOBO.AddChampSupValeur ('CORDETAT',Cdetat);
      TOBO.AddChampSupValeur ('NBRLIGNE',0);
      //
      if TOBOUVARTS.detail.count >=5000 then
      begin
        Compteur := Compteur + 5000;
        THTRAITEMENT.Caption := 'Traitement des OUVRAGES: '+ IntToStr(Compteur) + ' enregistrements traités';
        self.refresh;
        TOBOUVARTS.InsertOrUpdateDB;
        TOBENENT.InsertOrUpdateDB;
        TOBOUVARTS.cleardetail;
        TOBENENT.cleardetail;
      end;
    	TOBOuvrageD.free;
		end else
    begin
    	Terreur.SetError(3,MonFich,'');
      break;
    end;
	until TheContainer.EofFile;
  TheContainer.FermeFile;
  //
  // traitement des derniers enregistrements
  if TOBOUVARTS.detail.count > 0 then
  begin
    TOBOUVARTS.InsertOrUpdateDB;
    TOBENENT.InsertOrUpdateDB;
    TOBOUVARTS.cleardetail;
    TOBENENT.cleardetail;
  end;
  TOBOUVARTS.free;
  TOBFAMCHG.free;
	TOBENENT.free;
  Terreur.SetError(2,'OK','');

// Traitement des Détails d'ouvrages
  THTRAITEMENT.Caption := 'Traitement des DETAILS D''OUVRAGES ';
  self.refresh;
  //
  MonFich := RepInst+'OUVR_COMPO.seq';
  TheContainer.OpenFile(MonFich);
  TOBENLIGD := TOB.Create('NOMENLIG',nil,-1);
  Terreur.SetError(1,'Details_Ouvrages','');
  repeat
    CodeOuvComp := '';
    CodeOuvSimp := '';
  	TOBOUVCOMPD := TheContainer.ReadFile;
    if TOBOUVCOMPD <> nil then
    begin
			TOBOC:=TOBDESOUVRAGES.FindFirst(['OUVR_ID'],[TOBOUVCOMPD.getvalue('OUCO_COMP_ID')],true);
    	if TOBOC <> nil then
    	begin
    		CodeOuvComp := TOBOC.GetValue('OUVR_BATREF');
    	end else
    	begin
    		Terreur.SetError(50,string(TOBOUVCOMPD.getvalue('OUCO_COMP_ID')),'');
			end;
    	TOBO:=TOBDESOUVRAGES.FindFirst(['OUVR_ID'],[TOBOUVCOMPD.getvalue('OUCO_SIMP_ID')],true);
    	if TOBO <> nil then
    	begin
    		CodeOuvSimp := TOBO.GetValue('OUVR_BATREF');
    	end else
    	begin
    		Terreur.SetError(51,string(TOBOUVCOMPD.getvalue('OUCO_SIMP_ID')),'');
			end;
    	if (CodeOuvComp <> '') and (CodeOuvSimp<>'') then
    	begin
      	NumLigne := TOBOUVCOMPD.GetValue('OUCO_NUMORD')+1;
    		TOBENLIGD.InitValeurs;
      	TOBENLIGD.PutValue('GNL_NOMENCLATURE',CodeOuvComp);
      	TOBENLIGD.PutValue('GNL_CODEARTICLE',CodeOuvSimp);
      	TOBENLIGD.PutValue('GNL_SOUSNOMEN',CodeOuvSimp);
        TOBENLIGD.PutValue('GNL_ARTICLE',TOBO.GetValue('GA_ARTICLE'));
      	TOBENLIGD.PutValue('GNL_NUMLIGNE', NumLigne);
      	TOBENLIGD.PutValue('GNL_QTE',TOBOUVCOMPD.GetValue('OUCO_QTE'));
      	TOBENLIGD.PutValue('GNL_LIBELLE',TOBO.GetVAlue('LIBELLE'));
        If NumLigne > TOBOC.GetValue('NBRLIGNE') then
        begin
        	TOBOC.putValue('NBRLIGNE',NumLigne);
        end;
      	TOBENLIGD.InsertOrUpdateDB ;
    	end else
      begin
      	Terreur.SetError(53,string(TOBOUVCOMPD.getvalue('OUCO_COMP_ID')),string(TOBOUVCOMPD.getvalue('OUCO_SIMP_ID')));
      end;
      TOBOUVCOMPD.free;
    end else
    begin
    	Terreur.SetError(3,MonFich,'');
      break;
    end;
	until TheContainer.EofFile;

  TheContainer.FermeFile;
  //
  MonFich := RepInst+'OUVR_LIST_MAT.seq';
  TheContainer.OpenFile(MonFich);
	repeat
  CodeArticle := '';
  CodeOuvComp := '';
  TOBOUVMATD := TheContainer.ReadFile;
  if TOBOUVMATD <> nil then
  begin
	  TOBOC:=TOBDESOUVRAGES.FindFirst(['OUVR_ID'],[TOBOUVMATD.getvalue('OUVR_ID')],true);
  	if TOBOC <> nil then
  	begin
    	CodeOuvComp := TOBOC.GetValue('OUVR_BATREF');
    end else
    begin
    	Terreur.SetError(51,string(TOBOUVMATD.getvalue('OUVR_ID')),'');
    end;
  	TOBA:=TOBDESARTICLES.FindFirst(['MATE_ID'],[TOBOUVMATD.getvalue('MATE_ID')],true);
    if TOBA <> nil then
    begin
    	CodeArticle := TOBA.GetValue('MATE_BATREF');
    end else
    begin
			Terreur.SetError(52,string(TOBOUVMATD.getvalue('MATE_ID')),'');
    end;
    if (CodeOuvComp <> '') and (CodeArticle<>'') then
    begin
      NumLigne := TOBOUVMATD.GetValue('OULM_NUMORD') + 1;
    	TOBENLIGD.InitValeurs;
      TOBENLIGD.PutValue('GNL_NOMENCLATURE',CodeOuvComp);
      TOBENLIGD.PutValue('GNL_ARTICLE',TOBA.GetValue('GA_ARTICLE'));
      TOBENLIGD.PutValue('GNL_CODEARTICLE',CodeArticle);
      TOBENLIGD.PutValue('GNL_SOUSNOMEN','');
      TOBENLIGD.PutValue('GNL_NUMLIGNE', NumLigne);
      TOBENLIGD.PutValue('GNL_QTE',TOBOUVMATD.GetValue('OULM_QTE'));
      TOBENLIGD.PutValue('GNL_LIBELLE',TOBA.GetVAlue('LIBELLE'));
      If NumLigne > TOBOC.GetValue('NBRLIGNE') then
      begin
        TOBOC.putValue('NBRLIGNE',NumLigne);
      end;
      TOBENLIGD.InsertOrUpdateDB ;
    end else
    begin
    	Terreur.SetError(54,string(TOBOUVMATD.getvalue('OUVR_ID')),string(TOBOUVMATD.getvalue('MATE_ID')));
    end;
    TOBOUVMATD.free;
	end else
  begin
  	Terreur.SetError(3,MonFich,'');
    break;
  end;
  until TheContainer.EofFile;
	TheContainer.FermeFile;
	TOBENLIGD.free;
	Terreur.SetError(2,'OK','');
end;

function TFIntegreBatiprix.Decryptefichier: boolean;
var Res : integer;
		ResMessage : Widestring;
    RepertSRC : string;
begin
	RepertSRC := RepSRC.Text;
  result := true;
	CdeMaster.InstallCommande(RepertSRC, RepInst, Res, ResMessage);

  if Res <> 0 then
	begin
  	PgiBox(ResMessage);
    result := false;
	end;

end;

procedure TFIntegreBatiprix.bFinClick(Sender: TObject);
begin
  inherited;
	Close;
end;

procedure TFIntegreBatiprix.FormShow(Sender: TObject);
begin
  inherited;
	Bfin.Visible := false;
end;

procedure TFIntegreBatiprix.SupprimeSEQ;
var nomfic : string;
Rec : TSearchRec;
begin
  Nomfic:=RepInst+'\*.seq';
  if FindFirst (Nomfic,faAnyFile,Rec) = 0 then
   begin
     if (rec.name <> '.') and (rec.name <> '..') then deleteFile (RepInst+'\'+rec.Name);
     while FindNext (Rec) = 0 do
     begin
      if (rec.name <> '.') and (rec.name <> '..') then deleteFile (RepInst+'\'+rec.Name);
     end;
   end;
  FindClose (Rec);

end;

function TFIntegreBatiprix.FindFamilleBATIPRIX (TOBFAMILLE:TOB; IDFAM : string ; TypeA : string) : TOB;
var requete : string;
    QQ : TQuery;
    TT : TOB;
    OKOK : boolean;
begin
  result := nil;
  requete := 'SELECT * FROM BFAMILLECHANGE WHERE BFG_IMPFAMILLE4 = "'+IDFAM+'" AND BFG_IMPFAMILLE5="" AND BFG_TYPEARTICLE="'+TYPEA+'"';
  OKOK := false;
  if ExisteSQL(Requete) then
  begin
    OKOK := true;
  end else
  begin
    requete := 'SELECT * FROM BFAMILLECHANGE WHERE BFG_IMPFAMILLE3 <> "'+IDFAM+'" AND BFG_IMPFAMILLE4="" AND BFG_IMPFAMILLE5="" AND BFG_TYPEARTICLE="'+TYPEA+'"';
    if ExisteSQL(Requete) then
    begin
      OKOK := true;
    end else
    begin
      requete := 'SELECT * FROM BFAMILLECHANGE WHERE BFG_IMPFAMILLE3 <> "'+IDFAM+'" AND BFG_IMPFAMILLE4="" AND BFG_IMPFAMILLE5="" AND BFG_TYPEARTICLE="'+TYPEA+'"';
      if ExisteSQL(Requete) then
      begin
        OKOK := true;
      end;
    end;
  end;
  if OKOK then
  begin
    TT := TOB.Create ('BFAMILLECHANGE',TOBFAMILLE,-1);
    QQ := OpenSql (Requete,true,1,'',true);
    TT.SelectDB('',QQ);
    ferme (QQ);
    result := TT;
  end;
end;

end.
